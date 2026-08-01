# -*- coding: utf-8 -*-
"""
標案專案人力與組織圖 Web 產生器
==================================
讀取 Template_Staffing.xlsx（人員資料）與 Photos.zip（大頭照），
一鍵產出：
  Output 1：人員組織架構圖 (.pptx，原生可編輯圖形)
  Output 2：主要工作人力配置表 (.docx 表格 / .pptx 表格)
  Output 3：專案主要人員介紹（組長以上）(.docx)

作者：Claude (Anthropic) — 依需求規格開發
"""

import io
import math
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
from PIL import Image

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from docx import Document
from docx.shared import Cm, Pt as DocxPt, RGBColor as DocxRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn as docx_qn
from docx.oxml import OxmlElement

try:
    from google import genai as gemini_genai
except ImportError:
    gemini_genai = None


# =====================================================================
# 常數與樣式設定
# =====================================================================

REQUIRED_COLUMNS = [
    "Layer", "Role", "GroupName", "Name", "Title", "Badges", "PhotoName",
    "YearsOfExp", "Degree", "JobDescription", "Expertise", "BioNarrative",
]

# Excel 欄位名稱有時會被匯出成含中文說明的樣式，這裡做寬鬆對應
COLUMN_ALIASES = {
    "BioNarrative(經歷與描述）": "BioNarrative",
    "BioNarrative(經歷與描述)": "BioNarrative",
}

# 徽章代碼 -> 全名，可依專案需求自行調整
BADGE_MAP = {
    "技": "技師", "碩": "碩士", "博": "博士", "品": "品質管理人員",
    "安": "勞工安全衛生人員", "採": "採購專業人員", "乙": "乙級技術士",
    "甲": "甲級技術士", "景": "景觀技師", "土": "土木技師", "水": "水利技師",
}

BADGE_COLORS = {
    "技": "1F4E79", "碩": "2E7D32", "博": "6A1B9A", "品": "B8860B",
    "安": "C62828", "採": "00838F", "乙": "455A64", "甲": "37474F",
    "景": "558B2F", "土": "5D4037", "水": "0277BD",
}
DEFAULT_BADGE_COLOR = "607D8B"

# 主色系（淺色質感風格）
COLOR_BG = RGBColor(0xF7, 0xF9, 0xFC)
COLOR_NAVY = RGBColor(0x1F, 0x3A, 0x5F)
COLOR_NAVY_LIGHT = RGBColor(0x3E, 0x5C, 0x86)
COLOR_HEADER_GRAY = RGBColor(0x4A, 0x4A, 0x4A)
COLOR_CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_CARD_BORDER = RGBColor(0xD8, 0xDF, 0xE8)
COLOR_TEXT_DARK = RGBColor(0x22, 0x22, 0x22)
COLOR_TEXT_GRAY = RGBColor(0x66, 0x66, 0x66)
COLOR_FOOTER_BG = RGBColor(0xEE, 0xF1, 0xF6)
COLOR_PLACEHOLDER = RGBColor(0xC9, 0xCF, 0xD8)

LEADER_ROLE_ORDER = {"計畫顧問": 0, "計畫主持人": 1, "協同主持人": 2, "協同計畫主持人": 2}


# =====================================================================
# 資料載入與容錯處理
# =====================================================================

def load_staffing_excel(file) -> pd.DataFrame:
    """讀取 Excel 並做欄位對應、缺值容錯"""
    df = pd.read_excel(file, dtype=str)
    df = df.rename(columns={k: v for k, v in COLUMN_ALIASES.items() if k in df.columns})
    # 找不到的欄位一律補空字串，避免程式崩潰
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[REQUIRED_COLUMNS].copy()
    for col in REQUIRED_COLUMNS:
        df[col] = df[col].fillna("").astype(str).str.strip()
    # YearsOfExp 轉數值，轉換失敗補 0
    df["YearsOfExp_num"] = pd.to_numeric(df["YearsOfExp"], errors="coerce").fillna(0)
    return df


def load_photos_zip(file) -> dict:
    """讀取 Photos.zip，回傳 {檔名(不分大小寫): bytes}"""
    photos = {}
    try:
        with zipfile.ZipFile(file) as zf:
            for name in zf.namelist():
                if name.endswith("/"):
                    continue
                base = name.split("/")[-1]
                if not base or base.startswith("."):
                    continue
                try:
                    photos[base.lower()] = zf.read(name)
                except Exception:
                    continue
    except Exception as e:
        st.warning(f"照片 Zip 讀取失敗，將全部以預留方塊代替：{e}")
    return photos


def get_photo_bytes(photos: dict, photo_name: str):
    """依 PhotoName 找對應照片；找不到回傳 None（呼叫端需自行繪製灰色預留方塊）"""
    if not photo_name or photo_name.strip() in ("", "—", "-"):
        return None
    key = photo_name.strip().lower()
    if key in photos:
        return photos[key]
    # 容錯：嘗試補副檔名
    stem = key.rsplit(".", 1)[0]
    for ext in (".jpg", ".jpeg", ".png"):
        if (stem + ext) in photos:
            return photos[stem + ext]
    return None


def safe_open_image(raw_bytes):
    """驗證圖片是否可被 Pillow 開啟，避免匯出時崩潰"""
    if not raw_bytes:
        return None
    try:
        img = Image.open(io.BytesIO(raw_bytes))
        img.verify()
        return raw_bytes
    except Exception:
        return None


# =====================================================================
# 統計數據計算（頁尾時間軸）
# =====================================================================

def compute_footer_stats(df: pd.DataFrame) -> dict:
    avg_years = 0
    masters = 0
    technicians = 0
    groups = 0
    try:
        valid_years = df["YearsOfExp_num"]
        avg_years = int(round(valid_years.mean())) if len(valid_years) else 0
    except Exception:
        avg_years = 0
    try:
        masters = int(df["Degree"].str.contains("碩士|博士", na=False).sum())
    except Exception:
        masters = 0
    try:
        technicians = int(df["Badges"].apply(lambda b: "技" in [x.strip() for x in b.split(",")]).sum())
    except Exception:
        technicians = 0
    try:
        grp_df = df[df["Layer"].isin(["GroupLeader", "GroupMember"])]
        groups = grp_df[grp_df["GroupName"].str.strip().ne("") & grp_df["GroupName"].str.strip().ne("—")]["GroupName"].nunique()
    except Exception:
        groups = 0
    return {
        "avg_years": f"{avg_years} 平均年資",
        "masters": f"{masters} 碩士",
        "technicians": f"{technicians} 技師",
        "groups": f"{groups} 專業分組",
    }


# =====================================================================
# 共用小工具：徽章解析、姓名加註
# =====================================================================

def parse_badges(badge_str: str):
    if not badge_str or badge_str.strip() in ("", "—", "-"):
        return []
    return [b.strip() for b in badge_str.split(",") if b.strip()]


def name_with_cert_suffix(row) -> str:
    """依 Output2 樣式：若具技師徽章，姓名後加註 (技師)"""
    badges = parse_badges(row["Badges"])
    name = row["Name"] or "[資料待補]"
    if "技" in badges:
        return f"{name}(技師)"
    return name


def safe_text(value: str, fallback="[資料待補]") -> str:
    value = (value or "").strip()
    return value if value else fallback


# =====================================================================
# ============  Output 1：人員組織架構圖 (PPTX)  ======================
# =====================================================================

def _set_shape_fill(shape, color: RGBColor, line_color=None, line_width=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = line_width or Pt(0.75)
    shape.shadow.inherit = False


def _add_textbox(slide, x, y, w, h, text, size=11, bold=False, color=COLOR_TEXT_DARK,
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name="Microsoft JhengHei",
                  wrap=True, line_spacing=1.0):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    lines = text.split("\n") if isinstance(text, str) else [str(text)]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = font_name
        run.font.color.rgb = color
    return tb


def _add_photo_or_placeholder(slide, x, y, size, photo_bytes):
    """加入大頭照；若無照片則畫出可編輯的灰色預留方塊"""
    photo_bytes = safe_open_image(photo_bytes)
    if photo_bytes:
        try:
            slide.shapes.add_picture(io.BytesIO(photo_bytes), x, y, width=size, height=size)
            return
        except Exception:
            pass
    # 灰色原生預留方塊（可直接在 PowerPoint 中以「變更圖片」取代）
    ph = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, size, size)
    _set_shape_fill(ph, COLOR_PLACEHOLDER, line_color=RGBColor(0xAA, 0xAA, 0xAA), line_width=Pt(0.75))
    ph.text_frame.word_wrap = True
    ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = ph.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "照片\n待補"
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def _add_badge_row(slide, x, y, badges, chip_w=Inches(0.28), chip_h=Inches(0.24), gap=Inches(0.05)):
    cx = x
    for b in badges:
        color = RGBColor.from_string(BADGE_COLORS.get(b, DEFAULT_BADGE_COLOR))
        chip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, y, chip_w, chip_h)
        chip.adjustments[0] = 0.35
        _set_shape_fill(chip, color)
        chip.text_frame.word_wrap = False
        chip.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        chip.text_frame.margin_left = 0
        chip.text_frame.margin_right = 0
        p = chip.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = b
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cx += chip_w + gap
    return cx


def _add_leader_card(slide, cx, top, row, photos, card_w=Inches(2.5)):
    """中軸領導層卡片（計畫顧問 / 計畫主持人 / 協同主持人 / 計畫經理）"""
    photo_size = Inches(0.95)
    header_h = Inches(0.34)
    card_h = Inches(1.55)
    left = cx - card_w / 2

    # 白底卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, card_h)
    card.adjustments[0] = 0.06
    _set_shape_fill(card, COLOR_CARD_BG, line_color=COLOR_CARD_BORDER, line_width=Pt(1))

    # 深藍色職稱標頭
    header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, header_h)
    header.adjustments[0] = 0.5
    _set_shape_fill(header, COLOR_NAVY)
    header.text_frame.word_wrap = True
    header.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = header.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = safe_text(row["Role"], row["Role"] or "職稱待補")
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # 照片
    photo_x = left + (card_w - photo_size) / 2
    photo_y = top + header_h + Inches(0.08)
    _add_photo_or_placeholder(slide, photo_x, photo_y, photo_size, get_photo_bytes(photos, row["PhotoName"]))

    # 姓名
    name_y = photo_y + photo_size + Inches(0.02)
    _add_textbox(slide, left, name_y, card_w, Inches(0.26), safe_text(row["Name"]),
                 size=13, bold=True, color=COLOR_TEXT_DARK)

    # 職稱 / 學歷簡述
    title_y = name_y + Inches(0.26)
    _add_textbox(slide, left, title_y, card_w, Inches(0.2), safe_text(row["Title"]),
                 size=9.5, color=COLOR_TEXT_GRAY)

    # 徽章
    badges = parse_badges(row["Badges"])
    if badges:
        total_w = len(badges) * (Inches(0.28) + Inches(0.05)) - Inches(0.05)
        bx = left + (card_w - total_w) / 2
        by = title_y + Inches(0.22)
        _add_badge_row(slide, bx, by, badges)

    return left, top, card_w, card_h


def _add_group_block(slide, left, top, width, height, group_name, leader_row, member_rows, photos):
    header_h = Inches(0.34)
    # 外框
    outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    outer.adjustments[0] = 0.03
    _set_shape_fill(outer, COLOR_CARD_BG, line_color=COLOR_CARD_BORDER, line_width=Pt(1))

    # 組名深灰標頭
    header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, header_h)
    header.adjustments[0] = 0.5
    _set_shape_fill(header, COLOR_HEADER_GRAY)
    header.text_frame.word_wrap = True
    header.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = header.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = safe_text(group_name, "組別待補")
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    y = top + header_h + Inches(0.06)

    # 組長區塊
    if leader_row is not None:
        photo_size = Inches(0.62)
        _add_photo_or_placeholder(slide, left + Inches(0.12), y, photo_size, get_photo_bytes(photos, leader_row["PhotoName"]))
        info_x = left + Inches(0.12) + photo_size + Inches(0.1)
        info_w = width - (info_x - left) - Inches(0.1)
        _add_textbox(slide, info_x, y, info_w, Inches(0.16), "組長", size=8, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
        _add_textbox(slide, info_x, y + Inches(0.16), info_w, Inches(0.22),
                     f"{safe_text(leader_row['Name'])}　{safe_text(leader_row['Title'],'')}",
                     size=10.5, bold=True, color=COLOR_TEXT_DARK, align=PP_ALIGN.LEFT)
        badges = parse_badges(leader_row["Badges"])
        if badges:
            _add_badge_row(slide, info_x, y + Inches(0.38), badges, chip_w=Inches(0.22), chip_h=Inches(0.2), gap=Inches(0.04))
        y += photo_size + Inches(0.12)
    else:
        _add_textbox(slide, left + Inches(0.12), y, width - Inches(0.24), Inches(0.2), "組長：[資料待補]",
                     size=9.5, color=COLOR_TEXT_GRAY, align=PP_ALIGN.LEFT)
        y += Inches(0.28)

    # 分隔線（以細長矩形代替裝飾線，符合可編輯需求）
    divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left + Inches(0.12), y, width - Inches(0.24), Pt(1))
    _set_shape_fill(divider, COLOR_CARD_BORDER)
    y += Inches(0.08)

    _add_textbox(slide, left + Inches(0.12), y, width - Inches(0.24), Inches(0.16), "組員", size=8, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
    y += Inches(0.18)

    if member_rows:
        for m in member_rows:
            badges = parse_badges(m["Badges"])
            line_h = Inches(0.24)
            _add_textbox(slide, left + Inches(0.12), y, width - Inches(1.1), line_h,
                         f"{safe_text(m['Name'])} {safe_text(m['Title'],'')}",
                         size=9.5, color=COLOR_TEXT_DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)
            if badges:
                _add_badge_row(slide, left + width - Inches(0.98), y + Inches(0.01), badges,
                               chip_w=Inches(0.2), chip_h=Inches(0.18), gap=Inches(0.03))
            y += line_h
    else:
        _add_textbox(slide, left + Inches(0.12), y, width - Inches(0.24), Inches(0.2), "[資料待補]",
                     size=9, color=COLOR_TEXT_GRAY, align=PP_ALIGN.LEFT)


def build_org_chart_pptx(df: pd.DataFrame, stats: dict, photos: dict,
                          host_org: str = "", client_org: str = "") -> io.BytesIO:
    slide_w = Inches(13.333)

    # ---------- 第一階段：純數值試算版面高度，避免區塊互相重疊 ----------
    leader_top_calc = Inches(1.15) if (host_org or client_org) else Inches(0.55)
    middle_top_calc = leader_top_calc + Inches(1.85)
    groups_top_calc = middle_top_calc + Inches(1.85)

    group_names_calc = []
    for g in df[df["Layer"].isin(["GroupLeader", "GroupMember"])]["GroupName"]:
        g = g.strip()
        if g and g not in group_names_calc:
            group_names_calc.append(g)
    n_groups_calc = max(len(group_names_calc), 1)
    cols = min(4, n_groups_calc) if n_groups_calc > 0 else 1
    gap_y = Inches(0.25)

    row_max_h = {}
    for idx, gname in enumerate(group_names_calc):
        r, _c = divmod(idx, cols)
        member_rows = df[(df["Layer"] == "GroupMember") & (df["GroupName"] == gname)]
        block_h = Inches(0.95) + Inches(0.24) * max(len(member_rows), 1) + Inches(0.35)
        block_h = max(block_h, Inches(1.9))
        row_max_h[r] = max(row_max_h.get(r, 0), block_h)

    total_groups_h = sum(row_max_h.values()) + gap_y * max(len(row_max_h) - 1, 0)
    content_bottom = groups_top_calc + total_groups_h
    footer_h = Inches(0.75)
    needed_height = content_bottom + Inches(0.25) + footer_h

    slide_h = max(Inches(7.5), needed_height)

    prs = Presentation()
    prs.slide_width = slide_w
    prs.slide_height = slide_h
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版面

    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    _set_shape_fill(bg, COLOR_BG)
    bg.shadow.inherit = False

    # --- 頂部機關列 ---
    if host_org or client_org:
        top_texts = []
        if host_org:
            top_texts.append(("主辦機關", host_org))
        if client_org:
            top_texts.append(("委辦機關", client_org))
        box_w = Inches(3.4)
        gap = Inches(0.4)
        total_w = box_w * len(top_texts) + gap * (len(top_texts) - 1)
        start_x = (prs.slide_width - total_w) / 2
        for i, (label, val) in enumerate(top_texts):
            bx = start_x + i * (box_w + gap)
            box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, Inches(0.2), box_w, Inches(0.55))
            box.adjustments[0] = 0.25
            _set_shape_fill(box, COLOR_NAVY_LIGHT)
            box.text_frame.word_wrap = True
            box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = f"{label}：{val}"
            run.font.size = Pt(13)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        leader_top = Inches(1.15)
    else:
        leader_top = Inches(0.55)

    # --- 中軸領導層 ---
    advisors = df[df["Layer"] == "Advisor"].to_dict("records")
    tops = df[df["Layer"] == "Top"].to_dict("records")
    middles = df[df["Layer"] == "Middle"].to_dict("records")

    def top_sort_key(r):
        return LEADER_ROLE_ORDER.get(r["Role"], 1)

    leader_people = sorted(advisors + tops, key=top_sort_key)
    n = max(len(leader_people), 1)
    usable_w = slide_w - Inches(1.0)
    slot_w = usable_w / n
    start_x = Inches(0.5) + slot_w / 2

    leader_positions = []
    for i, person in enumerate(leader_people):
        cx = start_x + i * slot_w
        pos = _add_leader_card(slide, cx, leader_top, person, photos, card_w=min(Inches(2.7), slot_w - Inches(0.2)))
        leader_positions.append((cx, pos))

    middle_top = leader_top + Inches(1.85)
    middle_cx = slide_w / 2
    if middles:
        _add_leader_card(slide, middle_cx, middle_top, middles[0], photos, card_w=Inches(2.7))
    groups_top = middle_top + Inches(1.85)

    # 連接線（領導層 -> 計畫經理 -> 各組）
    try:
        for cx, _ in leader_positions:
            connector = slide.shapes.add_connector(2, int(cx), int(leader_top + Inches(1.55)),
                                                     int(middle_cx), int(middle_top))
            connector.line.color.rgb = COLOR_CARD_BORDER
            connector.line.width = Pt(1)
    except Exception:
        pass

    # --- 工作組別區 ---
    group_names = []
    for g in df[df["Layer"].isin(["GroupLeader", "GroupMember"])]["GroupName"]:
        g = g.strip()
        if g and g not in group_names:
            group_names.append(g)

    n_groups = max(len(group_names), 1)
    cols = min(4, n_groups) if n_groups > 0 else 1
    rows = math.ceil(n_groups / cols) if n_groups else 1

    margin_x = Inches(0.4)
    gap_x = Inches(0.25)
    gap_y = Inches(0.25)
    avail_w = slide_w - margin_x * 2 - gap_x * (cols - 1)
    block_w = avail_w / cols

    # 先算好每個群組區塊的資料與高度，再依「同一列取最大高度」計算 y 座標，避免不同高度區塊互相重疊
    group_infos = []
    for idx, gname in enumerate(group_names):
        r, c = divmod(idx, cols)
        leader_row = df[(df["Layer"] == "GroupLeader") & (df["GroupName"] == gname)]
        leader_row = leader_row.to_dict("records")[0] if len(leader_row) else None
        member_rows = df[(df["Layer"] == "GroupMember") & (df["GroupName"] == gname)].to_dict("records")

        block_h = Inches(0.95) + Inches(0.24) * max(len(member_rows), 1) + Inches(0.35)
        block_h = max(block_h, Inches(1.9))
        group_infos.append({"r": r, "c": c, "gname": gname, "leader_row": leader_row,
                             "member_rows": member_rows, "block_h": block_h})

    row_max_h = {}
    for info in group_infos:
        row_max_h[info["r"]] = max(row_max_h.get(info["r"], 0), info["block_h"])

    row_y_offset = {}
    cum = 0
    for r in sorted(row_max_h):
        row_y_offset[r] = cum
        cum += row_max_h[r] + gap_y

    for info in group_infos:
        bx = margin_x + info["c"] * (block_w + gap_x)
        by = groups_top + row_y_offset[info["r"]]
        _add_group_block(slide, bx, by, block_w, info["block_h"], info["gname"],
                          info["leader_row"], info["member_rows"], photos)

    footer_top = prs.slide_height - footer_h

    # --- 頁尾統計時間軸 ---
    footer = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, footer_top, slide_w, Inches(0.75))
    _set_shape_fill(footer, COLOR_FOOTER_BG)
    stat_items = [stats.get("avg_years", ""), stats.get("masters", ""),
                  stats.get("technicians", ""), stats.get("groups", "")]
    seg_w = slide_w / len(stat_items)
    for i, item in enumerate(stat_items):
        parts = item.split(" ", 1)
        number = parts[0] if parts else ""
        label = parts[1] if len(parts) > 1 else ""
        cx0 = i * seg_w
        _add_textbox(slide, cx0, footer_top + Inches(0.05), seg_w, Inches(0.4), number,
                     size=20, bold=True, color=COLOR_NAVY)
        _add_textbox(slide, cx0, footer_top + Inches(0.42), seg_w, Inches(0.28), label,
                     size=10, color=COLOR_TEXT_GRAY)
        if i > 0:
            divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx0, footer_top + Inches(0.12), Pt(1), Inches(0.5))
            _set_shape_fill(divider, RGBColor(0xC7, 0xCF, 0xDA))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


# =====================================================================
# =======  Output 2：主要工作人力配置表（DOCX / PPTX 表格）  ===========
# =====================================================================

LAYER_CATEGORY_MAP_DEFAULT = {
    "Advisor": "計畫顧問",
    "Middle": "計畫經理",
    "GroupLeader": "組長",
    "GroupMember": "組員",
}


def _category_for_row(row) -> str:
    if row["Layer"] == "Top":
        return "協同主持人" if "協同" in row["Role"] else "計畫主持人"
    return LAYER_CATEGORY_MAP_DEFAULT.get(row["Layer"], row["Role"] or "人員")


def build_staffing_table_rows(df: pd.DataFrame):
    """回傳依 Output2 樣式排序、格式化後的 list[dict]，供表格預覽/匯出共用"""
    order = {"Top": 0, "Advisor": 1, "Middle": 2, "GroupLeader": 3, "GroupMember": 4}
    df2 = df.copy()
    df2["_order"] = df2["Layer"].map(order).fillna(9)
    # Top 排序：主持人優先於協同主持人
    df2["_sub"] = df2.apply(lambda r: 0 if (r["Layer"] == "Top" and "協同" not in r["Role"]) else 1, axis=1)
    df2 = df2.sort_values(["_order", "_sub", "GroupName"], kind="stable")

    rows = []
    for _, r in df2.iterrows():
        rows.append({
            "類別": _category_for_row(r),
            "姓名": name_with_cert_suffix(r),
            "職稱": safe_text(r["Title"], ""),
            "最高學歷科系": safe_text(r["Degree"], ""),
            "擬任工作內容": safe_text(r["JobDescription"], ""),
            "相關經歷與專長": safe_text(r["Expertise"], ""),
        })
    return rows


def build_staffing_table_docx(df: pd.DataFrame) -> io.BytesIO:
    rows = build_staffing_table_rows(df)
    doc = Document()
    section = doc.sections[0]
    section.orientation = 0

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("主要工作人力配置表")
    run.font.size = DocxPt(16)
    run.font.bold = True

    headers = ["類別", "姓名", "職稱", "最高學歷科系", "擬任工作內容", "相關經歷與專長"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = DocxPt(11)
        _shade_cell(hdr_cells[i], "1F3A5F")
        run.font.color.rgb = DocxRGB(0xFF, 0xFF, 0xFF)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for r in rows:
        cells = table.add_row().cells
        for i, h in enumerate(headers):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i >= 3 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(r[h])
            run.font.size = DocxPt(10.5)
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


def _shade_cell(cell, hex_color):
    shd = OxmlElement("w:shd")
    shd.set(docx_qn("w:fill"), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def build_staffing_table_pptx(df: pd.DataFrame) -> io.BytesIO:
    rows = build_staffing_table_rows(df)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    headers = ["類別", "姓名", "職稱", "最高學歷科系", "擬任工作內容", "相關經歷與專長"]
    col_widths = [1.3, 1.5, 1.3, 2.6, 2.6, 4.0]

    # 每頁最多容納列數（含表頭）
    rows_per_slide = 10
    chunks = [rows[i:i + rows_per_slide] for i in range(0, len(rows), rows_per_slide)] or [[]]

    for chunk in chunks:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title_box = _add_textbox(slide, Inches(0.4), Inches(0.2), Inches(12.5), Inches(0.5),
                                  "主要工作人力配置表", size=22, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
        n_rows = len(chunk) + 1
        table_shape = slide.shapes.add_table(n_rows, len(headers), Inches(0.4), Inches(0.9),
                                              Inches(12.5), Inches(0.4) * n_rows)
        table = table_shape.table
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(w)
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_NAVY
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = Pt(11)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for ri, r in enumerate(chunk, start=1):
            for ci, h in enumerate(headers):
                cell = table.cell(ri, ci)
                cell.text = r[h]
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9.5)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


# =====================================================================
# ===========  Output 3：主要人員經歷敘述（組長以上）(DOCX)  ============
# =====================================================================

SENIOR_LAYERS = ["Top", "Advisor", "Middle", "GroupLeader"]

GEMINI_MODEL = "gemini-2.5-flash"


def get_secret_gemini_key() -> str:
    """安全讀取 Streamlit Secrets 中的 GEMINI_API_KEY；未設定 secrets.toml 時不拋出例外"""
    try:
        return (st.secrets.get("GEMINI_API_KEY", "") or "").strip()
    except Exception:
        return ""


def get_effective_gemini_key() -> str:
    """依規則取得實際呼叫用金鑰：優先 Streamlit Secrets，其次側邊欄手動輸入值"""
    return get_secret_gemini_key() or st.session_state.get("gemini_api_key", "").strip()


def refine_bio_with_gemini(api_key: str, name: str, job_description: str, bio_narrative: str) -> str:
    """呼叫 Gemini (google-genai) 依擬任工作內容，將原始履歷潤飾為 200 字以內備標敘述"""
    if gemini_genai is None:
        raise RuntimeError("尚未安裝 google-genai 套件，請確認 requirements.txt 是否包含 google-genai")

    job_description = job_description.strip() or "（未提供擬任工作內容）"
    bio_narrative = bio_narrative.strip() or "（未提供原始履歷）"

    prompt = (
        "你是一位政府水利工程標案專家，"
        f"請依據該人員（{name}）的【擬任工作】：{job_description}，"
        f"將其【原始履歷】：{bio_narrative}，"
        "潤飾精簡為一段200字以內、極具評審說服力且專業精準的水利工程備標履歷敘述。"
        "請只輸出潤飾後的履歷內文本身，不要加上任何前綴、標題、項目符號或引號。"
    )

    client = gemini_genai.Client(api_key=api_key)
    response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
    text = (getattr(response, "text", "") or "").strip()
    if not text:
        raise RuntimeError("Gemini 未回傳有效內容")
    return text


def build_bio_docx(df: pd.DataFrame, photos: dict) -> io.BytesIO:
    order = {"Top": 0, "Advisor": 1, "Middle": 2, "GroupLeader": 3}
    df2 = df[df["Layer"].isin(SENIOR_LAYERS)].copy()
    df2["_order"] = df2["Layer"].map(order).fillna(9)
    df2["_sub"] = df2.apply(lambda r: 0 if (r["Layer"] == "Top" and "協同" not in r["Role"]) else 1, axis=1)
    df2 = df2.sort_values(["_order", "_sub"], kind="stable")

    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("主要人員經歷")
    run.font.size = DocxPt(16)
    run.font.bold = True

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.columns[0].width = Cm(4.2)
    table.columns[1].width = Cm(10.5)

    hdr = table.rows[0].cells
    hdr[0].text = ""
    hdr[0].paragraphs[0].add_run("姓名職稱").font.bold = True
    hdr[1].text = ""
    hdr[1].paragraphs[0].add_run("專長說明").font.bold = True
    for c in hdr:
        _shade_cell(c, "1F3A5F")
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.color.rgb = DocxRGB(0xFF, 0xFF, 0xFF)

    for _, r in df2.iterrows():
        cells = table.add_row().cells
        photo_cell, text_cell = cells[0], cells[1]
        photo_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        photo_p = photo_cell.paragraphs[0]
        photo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        raw = safe_open_image(get_photo_bytes(photos, r["PhotoName"]))
        if raw:
            try:
                run = photo_p.add_run()
                run.add_picture(io.BytesIO(raw), width=Cm(3.2))
            except Exception:
                raw = None
        if not raw:
            # 原生可編輯灰色預留方塊（以有底色表格取代圖片，避免崩潰）
            ph_table = photo_cell.add_table(rows=1, cols=1)
            ph_cell = ph_table.rows[0].cells[0]
            _shade_cell(ph_cell, "C9CFD8")
            ph_cell.text = "照片待補"
            ph_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        badges = parse_badges(r["Badges"])
        cert_names = [BADGE_MAP.get(b, b) for b in badges]
        cert_line = "／".join(cert_names) if cert_names else ""

        role_line = f"{safe_text(r['Role'])}　{safe_text(r['Name'])}　{safe_text(r['Title'],'')}"
        if cert_line:
            role_line += f"／{cert_line}"

        p1 = text_cell.paragraphs[0]
        p1.text = ""
        run1 = p1.add_run(role_line)
        run1.font.bold = True
        run1.font.size = DocxPt(12)

        p2 = text_cell.add_paragraph()
        run2 = p2.add_run(safe_text(r["BioNarrative"]))
        run2.font.size = DocxPt(10.5)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# =====================================================================
# ============================  Streamlit UI  ==========================
# =====================================================================

# =====================================================================
# ============================  Streamlit UI  ==========================
# =====================================================================

st.set_page_config(page_title="標案專案人力與組織圖產生器", layout="wide", page_icon="📋")

# ---------------- 全域 CSS 美化 ----------------
st.markdown("""
<style>
/* 全站質感淺藍背景 */
.stApp { background-color: #FAFCFF; }

/* 頂部標題 Banner：淺藍漸層圓角卡片 */
.banner-box {
    background: linear-gradient(135deg, #EAF4FD 0%, #F7FBFF 100%);
    border: 1px solid #DCE9F7;
    border-radius: 18px;
    padding: 22px 28px;
    margin-bottom: 14px;
}
.banner-box h1 { margin: 0 0 4px 0; font-size: 28px; color: #123A63; }
.banner-box p { margin: 0; color: #4A5A6A; font-size: 14px; }

/* 提示訊息：淺黃質感提醒條 */
.hint-box {
    background: #FFF9E6;
    border: 1px solid #F3E3A0;
    border-radius: 12px;
    padding: 14px 18px;
    color: #6B5B18;
    font-size: 14px;
    margin: 6px 0 16px 0;
}

/* 主題藍：下載按鈕與一般按鈕 */
div[data-testid="stDownloadButton"] button, div[data-testid="stBaseButton-secondary"] {
    background-color: #007AC3 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    padding: 0.6em 1em !important;
}
div[data-testid="stDownloadButton"] button:hover { background-color: #005F99 !important; }
div[data-testid="stButton"] button {
    background-color: #EAF4FD !important;
    color: #123A63 !important;
    border: 1px solid #B9D8F0 !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
}

/* 側邊欄：縮小上傳區塊間距 */
section[data-testid="stSidebar"] .block-container { padding-top: 1.2rem; }
section[data-testid="stSidebar"] h3 { margin-bottom: 4px; }
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { padding: 0.6rem; }
section[data-testid="stSidebar"] .stFileUploader { margin-bottom: 0.3rem; }

/* Expander 標頭美化 */
details summary { font-weight: 700; color: #123A63; }
</style>
""", unsafe_allow_html=True)

# ---------------- Session State 初始化 ----------------
if "df" not in st.session_state:
    st.session_state.df = None
if "photos" not in st.session_state:
    st.session_state.photos = {}
if "stats" not in st.session_state:
    st.session_state.stats = {"avg_years": "0 平均年資", "masters": "0 碩士",
                               "technicians": "0 技師", "groups": "0 專業分組"}
if "host_org" not in st.session_state:
    st.session_state.host_org = ""
if "client_org" not in st.session_state:
    st.session_state.client_org = ""
if "gemini_api_key" not in st.session_state:
    st.session_state.gemini_api_key = get_secret_gemini_key()

# ---------------- ① 側邊欄：資料上傳 + Gemini AI 設定 ----------------
with st.sidebar:
    st.markdown("### ① 資料上傳")

    excel_file = st.file_uploader("Template_Staffing.xlsx", type=["xlsx"], label_visibility="visible")
    if excel_file is not None:
        try:
            st.session_state.df = load_staffing_excel(excel_file)
            st.success(f"已讀取 {len(st.session_state.df)} 筆人員資料", icon="✅")
        except Exception as e:
            st.error(f"Excel 讀取失敗：{e}")

    zip_file = st.file_uploader("Photos.zip（大頭照）", type=["zip"], label_visibility="visible")
    if zip_file is not None:
        st.session_state.photos = load_photos_zip(zip_file)
        st.success(f"已讀取 {len(st.session_state.photos)} 張照片", icon="✅")

    st.markdown("### ② Gemini AI 設定")
    _secret_key = get_secret_gemini_key()
    if _secret_key:
        st.caption("🔒 已從 Streamlit Secrets 自動讀取 GEMINI_API_KEY")
    st.session_state.gemini_api_key = st.text_input(
        "Gemini API Key", type="password",
        value=st.session_state.gemini_api_key or _secret_key,
        help="若已於 Streamlit Secrets 設定 GEMINI_API_KEY，將自動帶入並優先使用；"
             "此欄位僅作為備用手動輸入，金鑰僅存於本次瀏覽器工作階段，不會被儲存。",
    )

# ---------------- 頂部標題 Banner ----------------
st.markdown("""
<div class="banner-box">
    <h1>📋 標案專案人力與組織圖 Web 產生器</h1>
    <p>上傳人員資料 Excel 與照片 Zip，即可預覽並一鍵匯出組織圖、人力配置表、人員經歷說明書</p>
</div>
""", unsafe_allow_html=True)

df = st.session_state.df

if df is None:
    st.markdown("""
    <div class="hint-box">👉 請先於左側上傳人員資料 Excel（Template_Staffing.xlsx 格式）</div>
    """, unsafe_allow_html=True)
    st.stop()

# ---------------- ② 標案參數設定面板（可折疊） ----------------
with st.expander("⚙️ 標案參數設定面板", expanded=False):
    col_org, col_stats = st.columns(2)

    with col_org:
        st.markdown("**機關名稱**")
        st.session_state.host_org = st.text_input("主辦機關", value=st.session_state.host_org)
        st.session_state.client_org = st.text_input("委辦機關", value=st.session_state.client_org)

    with col_stats:
        st.markdown("**頁尾數據時間軸微調**")
        if st.button("🔄 自動從 Excel 計算", use_container_width=True):
            st.session_state.stats = compute_footer_stats(st.session_state.df)

        s1, s2 = st.columns(2)
        with s1:
            st.session_state.stats["avg_years"] = st.text_input("平均年資", st.session_state.stats["avg_years"])
            st.session_state.stats["technicians"] = st.text_input("技師執照數", st.session_state.stats["technicians"])
        with s2:
            st.session_state.stats["masters"] = st.text_input("碩博士人數", st.session_state.stats["masters"])
            st.session_state.stats["groups"] = st.text_input("專業分組數", st.session_state.stats["groups"])

host_org = st.session_state.host_org
client_org = st.session_state.client_org

# ---------------- ③ 匯出下載專區（主畫面頂部，三欄併排） ----------------
st.markdown("#### 📤 匯出下載專區")
exp_col1, exp_col2, exp_col3 = st.columns(3)

with exp_col1:
    try:
        pptx1 = build_org_chart_pptx(df, st.session_state.stats, st.session_state.photos,
                                      host_org, client_org)
        st.download_button("⬇️ Output 1　組織架構圖 (.pptx)", data=pptx1,
                            file_name="人員組織架構圖.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True)
    except Exception as e:
        st.error(f"Output 1 產生失敗：{e}")

with exp_col2:
    try:
        docx2 = build_staffing_table_docx(df)
        st.download_button("⬇️ Output 2　人力配置表 (.docx)", data=docx2,
                            file_name="主要工作人力配置表.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True)
    except Exception as e:
        st.error(f"Output 2 產生失敗：{e}")

with exp_col3:
    try:
        docx3 = build_bio_docx(df, st.session_state.photos)
        st.download_button("⬇️ Output 3　人員經歷說明書 (.docx)", data=docx3,
                            file_name="專案主要人員介紹.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True)
    except Exception as e:
        st.error(f"Output 3 產生失敗：{e}")

st.markdown("")

# ---------------- ④ 下方預覽區：Tab 1~3 ----------------
tab1, tab2, tab3 = st.tabs(["🗂️ 專案組織圖預覽 (Output 1)", "📊 人力配置表預覽 (Output 2)", "🧑‍💼 人員經歷敘述預覽 (Output 3)"])

# ---- Tab 1: 組織圖 HTML 概覽預覽 ----
with tab1:
    st.caption("以下為結構化概覽預覽，實際版面、字型與間距請以上方下載之 .pptx 檔案為準。")

    def badge_html(badges):
        chips = []
        for b in badges:
            color = "#" + BADGE_COLORS.get(b, DEFAULT_BADGE_COLOR)
            chips.append(f"<span style='background:{color};color:#fff;border-radius:6px;"
                         f"padding:1px 6px;font-size:11px;margin-right:3px;'>{b}</span>")
        return "".join(chips)

    def person_card(row, small=False):
        badges = parse_badges(row["Badges"])
        photo = get_photo_bytes(st.session_state.photos, row["PhotoName"])
        photo_html = ""
        if safe_open_image(photo):
            import base64
            b64 = base64.b64encode(photo).decode()
            photo_html = f"<img src='data:image/png;base64,{b64}' style='width:56px;height:56px;object-fit:cover;border-radius:6px;'/>"
        else:
            photo_html = "<div style='width:56px;height:56px;background:#C9CFD8;border-radius:6px;" \
                         "display:flex;align-items:center;justify-content:center;font-size:10px;color:#555;'>照片待補</div>"
        return f"""
        <div style='display:flex;gap:8px;align-items:center;background:#fff;border:1px solid #D8DFE8;
                    border-radius:8px;padding:6px 8px;margin-bottom:6px;'>
            {photo_html}
            <div>
                <div style='font-weight:700;font-size:13px;'>{safe_text(row['Name'])}</div>
                <div style='font-size:11px;color:#666;'>{safe_text(row['Title'],'')}</div>
                <div style='margin-top:2px;'>{badge_html(badges)}</div>
            </div>
        </div>
        """

    advisors = df[df["Layer"] == "Advisor"].to_dict("records")
    tops = sorted(df[df["Layer"] == "Top"].to_dict("records"),
                  key=lambda r: LEADER_ROLE_ORDER.get(r["Role"], 1))
    middles = df[df["Layer"] == "Middle"].to_dict("records")

    org_header = ""
    if host_org or client_org:
        org_header = "<div style='text-align:center;margin-bottom:10px;'>"
        if host_org:
            org_header += f"<span style='background:#3E5C86;color:#fff;padding:4px 14px;border-radius:14px;margin-right:8px;font-size:13px;'>主辦機關：{host_org}</span>"
        if client_org:
            org_header += f"<span style='background:#3E5C86;color:#fff;padding:4px 14px;border-radius:14px;font-size:13px;'>委辦機關：{client_org}</span>"
        org_header += "</div>"
    st.markdown(org_header, unsafe_allow_html=True)

    leader_cols = st.columns(max(len(advisors) + len(tops), 1))
    for i, person in enumerate(advisors + tops):
        with leader_cols[i]:
            st.markdown(f"<div style='text-align:center;background:#1F3A5F;color:#fff;border-radius:6px;"
                        f"padding:4px;font-size:12px;font-weight:700;margin-bottom:4px;'>{safe_text(person['Role'])}</div>",
                        unsafe_allow_html=True)
            st.markdown(person_card(person), unsafe_allow_html=True)

    if middles:
        mc1, mc2, mc3 = st.columns([1, 1, 1])
        with mc2:
            st.markdown(f"<div style='text-align:center;background:#1F3A5F;color:#fff;border-radius:6px;"
                        f"padding:4px;font-size:12px;font-weight:700;margin:10px 0 4px;'>{safe_text(middles[0]['Role'])}</div>",
                        unsafe_allow_html=True)
            st.markdown(person_card(middles[0]), unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("##### 工作組別")

    group_names = []
    for g in df[df["Layer"].isin(["GroupLeader", "GroupMember"])]["GroupName"]:
        g = g.strip()
        if g and g not in group_names:
            group_names.append(g)

    n_cols = min(4, max(len(group_names), 1))
    group_cols = st.columns(n_cols)
    for idx, gname in enumerate(group_names):
        col = group_cols[idx % n_cols]
        with col:
            st.markdown(f"<div style='background:#4A4A4A;color:#fff;text-align:center;border-radius:6px;"
                        f"padding:5px;font-size:12px;font-weight:700;margin-top:8px;'>{gname}</div>",
                        unsafe_allow_html=True)
            leader_row = df[(df["Layer"] == "GroupLeader") & (df["GroupName"] == gname)]
            if len(leader_row):
                st.markdown("<div style='font-size:10px;color:#1F3A5F;font-weight:700;margin-top:4px;'>組長</div>", unsafe_allow_html=True)
                st.markdown(person_card(leader_row.to_dict("records")[0]), unsafe_allow_html=True)
            members = df[(df["Layer"] == "GroupMember") & (df["GroupName"] == gname)]
            if len(members):
                st.markdown("<div style='font-size:10px;color:#1F3A5F;font-weight:700;'>組員</div>", unsafe_allow_html=True)
                for _, m in members.iterrows():
                    badges = parse_badges(m["Badges"])
                    st.markdown(f"<div style='display:flex;justify-content:space-between;background:#fff;"
                                f"border:1px solid #D8DFE8;border-radius:6px;padding:3px 6px;margin-bottom:4px;font-size:11px;'>"
                                f"<span>{safe_text(m['Name'])} {safe_text(m['Title'],'')}</span>"
                                f"<span>{badge_html(badges)}</span></div>", unsafe_allow_html=True)

    st.markdown("---")
    stats = st.session_state.stats
    fcols = st.columns(4)
    for i, key in enumerate(["avg_years", "masters", "technicians", "groups"]):
        val = stats.get(key, "")
        parts = val.split(" ", 1)
        num = parts[0] if parts else ""
        lbl = parts[1] if len(parts) > 1 else ""
        with fcols[i]:
            st.markdown(f"<div style='text-align:center;background:#EEF1F6;border-radius:8px;padding:10px;'>"
                        f"<div style='font-size:24px;font-weight:800;color:#1F3A5F;'>{num}</div>"
                        f"<div style='font-size:12px;color:#666;'>{lbl}</div></div>", unsafe_allow_html=True)

# ---- Tab 2: 人力配置表 ----
with tab2:
    rows = build_staffing_table_rows(df)
    preview_df = pd.DataFrame(rows)
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

    st.markdown("")
    try:
        pptx2 = build_staffing_table_pptx(df)
        st.download_button("⬇️ 另外匯出 Output 2：人力配置表 (.pptx 版本)", data=pptx2,
                            file_name="主要工作人力配置表.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        st.error(f"Output 2 (pptx) 產生失敗：{e}")

# ---- Tab 3: 人員經歷敘述（含 Gemini AI 智慧潤飾） ----
with tab3:
    top_l, top_r = st.columns([3, 2])
    with top_l:
        st.caption("僅呈現組長以上同仁（Layer: Top / Advisor / Middle / GroupLeader）之經歷敘述。")
    with top_r:
        refine_clicked = st.button(
            "✨ 一鍵使用 Gemini 依據擬任工作潤飾履歷",
            use_container_width=True,
            help="依據 JobDescription（擬任工作）潤飾 BioNarrative（履歷），需先於左側輸入 Gemini API Key",
        )

    if refine_clicked:
        effective_key = get_effective_gemini_key()
        if not effective_key:
            st.warning("請先輸入 Gemini API Key")
        elif gemini_genai is None:
            st.error("尚未安裝 google-genai 套件，請確認 requirements.txt 是否包含 google-genai 並重新部署")
        else:
            target_mask = st.session_state.df["Layer"].isin(SENIOR_LAYERS)
            target_idx = st.session_state.df[target_mask].index.tolist()
            total = len(target_idx)

            if total == 0:
                st.info("目前無組長以上人員資料可供潤飾")
            else:
                progress = st.progress(0, text=f"AI 潤飾處理中... (0/{total})")
                fail_count = 0
                for i, idx in enumerate(target_idx):
                    row = st.session_state.df.loc[idx]
                    try:
                        new_bio = refine_bio_with_gemini(
                            effective_key,
                            row["Name"], row["JobDescription"], row["BioNarrative"],
                        )
                        st.session_state.df.loc[idx, "BioNarrative"] = new_bio
                    except Exception as e:
                        fail_count += 1
                        st.warning(f"{safe_text(row['Name'])} 履歷潤飾失敗：{e}")
                    progress.progress((i + 1) / total, text=f"AI 潤飾處理中... ({i + 1}/{total})")
                progress.empty()

                if fail_count == 0:
                    st.toast("履歷 AI 潤飾完成！", icon="✨")
                else:
                    st.toast(f"履歷 AI 潤飾完成（{total - fail_count}/{total} 筆成功）", icon="⚠️")
                st.rerun()

    df = st.session_state.df  # 潤飾後重新取用最新資料

    senior_df = df[df["Layer"].isin(SENIOR_LAYERS)]
    order = {"Top": 0, "Advisor": 1, "Middle": 2, "GroupLeader": 3}
    senior_df = senior_df.copy()
    senior_df["_order"] = senior_df["Layer"].map(order).fillna(9)
    senior_df["_sub"] = senior_df.apply(lambda r: 0 if (r["Layer"] == "Top" and "協同" not in r["Role"]) else 1, axis=1)
    senior_df = senior_df.sort_values(["_order", "_sub"], kind="stable")

    for _, r in senior_df.iterrows():
        c1, c2 = st.columns([1, 4])
        with c1:
            raw = safe_open_image(get_photo_bytes(st.session_state.photos, r["PhotoName"]))
            if raw:
                st.image(raw, width=140)
            else:
                st.markdown("<div style='width:140px;height:140px;background:#C9CFD8;border-radius:8px;"
                            "display:flex;align-items:center;justify-content:center;color:#555;'>照片待補</div>",
                            unsafe_allow_html=True)
        with c2:
            badges = parse_badges(r["Badges"])
            cert_names = [BADGE_MAP.get(b, b) for b in badges]
            cert_line = "／".join(cert_names) if cert_names else ""
            st.markdown(f"**{safe_text(r['Role'])}　{safe_text(r['Name'])}　{safe_text(r['Title'],'')}**"
                        + (f"　/{cert_line}" if cert_line else ""))
            st.write(safe_text(r["BioNarrative"]))
        st.markdown("---")
