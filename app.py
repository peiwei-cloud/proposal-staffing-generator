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
    from google.genai import errors as gemini_errors
except ImportError:
    gemini_genai = None
    gemini_errors = None


# =====================================================================
# 常數與樣式設定
# =====================================================================

REQUIRED_COLUMNS = [
    "Layer", "Role", "GroupName", "Name", "Title", "Company", "Badges", "PhotoName",
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

# 全域中文字型（PPTX / DOCX 皆套用，含 EastAsia 屬性設定）
GLOBAL_FONT_NAME = "Microsoft JhengHei"


def classify_leadership_zone(row) -> str:
    """依 Layer / Role 規則將 Top/Advisor 人員分類至左/中/右三區
    左區：Layer=='Advisor' 或 Role 含「顧問」「品質督導」
    中區：Layer=='Top' 且 Role 含「計畫主持人」（未命中任何規則的 Top 預設歸中區）
    右區：Layer=='Top' 且 Role 含「協同」或「代表廠商」
    """
    role = row.get("Role", "") or ""
    layer = row.get("Layer", "")
    if layer == "Advisor" or ("顧問" in role) or ("品質督導" in role):
        return "left"
    if layer == "Top" and ("協同" in role or "代表廠商" in role):
        return "right"
    if layer == "Top" and ("計畫主持人" in role):
        return "center"
    if layer == "Top":
        return "center"  # 未命中關鍵字之 Top 人員，預設歸入中區，避免版面遺漏
    return "left"  # 理論上不會執行到此（僅 Top/Advisor 會呼叫本函式）


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
        valid_years = df["YearsOfExp_num"].replace(0, pd.NA).dropna()
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

def _set_pptx_run_ea_font(run, font_name=GLOBAL_FONT_NAME):
    """補上 PPTX 東亞字型設定（<a:ea typeface=.../>），確保中文於各檢視器正確套用指定字型"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", font_name)


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
                  align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=GLOBAL_FONT_NAME,
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
        _set_pptx_run_ea_font(run, font_name)
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
    run.font.size = Pt(9)
    run.font.name = GLOBAL_FONT_NAME
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    _set_pptx_run_ea_font(run)


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
        run.font.name = GLOBAL_FONT_NAME
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _set_pptx_run_ea_font(run)
        cx += chip_w + gap
    return cx


# 領導卡片版面常數（_add_leader_card 與 _leader_card_height 共用，確保繪製與試算高度一致）
_LC_HEADER_H = Inches(0.34)
_LC_PHOTO_GAP = Inches(0.08)
_LC_PHOTO_SIZE = Inches(0.95)
_LC_NAME_GAP = Inches(0.02)
_LC_NAME_H = Inches(0.24)
_LC_TITLE_H = Inches(0.19)
_LC_COMPANY_H = Inches(0.18)
_LC_BADGE_GAP = Inches(0.03)
_LC_BADGE_H = Inches(0.24)
_LC_BOTTOM_PAD = Inches(0.1)


def _leader_card_base_height(has_company: bool, has_badges: bool) -> Emu:
    h = (_LC_HEADER_H + _LC_PHOTO_GAP + _LC_PHOTO_SIZE + _LC_NAME_GAP + _LC_NAME_H
         + _LC_TITLE_H + _LC_BOTTOM_PAD)
    if has_company:
        h += _LC_COMPANY_H
    if has_badges:
        h += _LC_BADGE_GAP + _LC_BADGE_H
    return h


def _add_leader_card(slide, cx, top, row, photos, card_w=Inches(2.5)):
    """中軸領導層卡片（計畫顧問 / 計畫主持人 / 協同主持人 / 計畫經理 / SubTop 等）"""
    company = safe_text(row.get("Company", ""), "")
    badges = parse_badges(row["Badges"])
    card_h = _leader_card_base_height(bool(company), bool(badges))
    left = cx - card_w / 2

    # 白底卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, card_h)
    card.adjustments[0] = 0.06
    _set_shape_fill(card, COLOR_CARD_BG, line_color=COLOR_CARD_BORDER, line_width=Pt(1))

    # 深藍色職稱標頭
    header = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_w, _LC_HEADER_H)
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
    run.font.name = GLOBAL_FONT_NAME
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    _set_pptx_run_ea_font(run)

    # 照片
    photo_x = left + (card_w - _LC_PHOTO_SIZE) / 2
    photo_y = top + _LC_HEADER_H + _LC_PHOTO_GAP
    _add_photo_or_placeholder(slide, photo_x, photo_y, _LC_PHOTO_SIZE, get_photo_bytes(photos, row["PhotoName"]))

    # 姓名（PPTX 字級規範 9~12pt，取上限 12pt 突顯人名）
    name_y = photo_y + _LC_PHOTO_SIZE + _LC_NAME_GAP
    _add_textbox(slide, left, name_y, card_w, _LC_NAME_H, safe_text(row["Name"]),
                 size=12, bold=True, color=COLOR_TEXT_DARK)

    # 職稱 / 學歷簡述
    title_y = name_y + _LC_NAME_H
    _add_textbox(slide, left, title_y, card_w, _LC_TITLE_H, safe_text(row["Title"]),
                 size=9.5, color=COLOR_TEXT_GRAY)

    next_y = title_y + _LC_TITLE_H

    # 公司名稱（若 Excel 有填寫 Company 欄位才顯示）
    if company:
        _add_textbox(slide, left, next_y, card_w, _LC_COMPANY_H, company,
                     size=9, color=COLOR_TEXT_GRAY)
        next_y += _LC_COMPANY_H

    # 徽章
    if badges:
        total_w = len(badges) * (Inches(0.28) + Inches(0.05)) - Inches(0.05)
        bx = left + (card_w - total_w) / 2
        by = next_y + _LC_BADGE_GAP
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
    run.font.name = GLOBAL_FONT_NAME
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    _set_pptx_run_ea_font(run)

    y = top + header_h + Inches(0.06)

    # 組長區塊
    if leader_row is not None:
        photo_size = Inches(0.62)
        _add_photo_or_placeholder(slide, left + Inches(0.12), y, photo_size, get_photo_bytes(photos, leader_row["PhotoName"]))
        info_x = left + Inches(0.12) + photo_size + Inches(0.1)
        info_w = width - (info_x - left) - Inches(0.1)
        _add_textbox(slide, info_x, y, info_w, Inches(0.16), "組長", size=9, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
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

    _add_textbox(slide, left + Inches(0.12), y, width - Inches(0.24), Inches(0.16), "組員", size=9, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
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


def _leadership_zones(df: pd.DataFrame) -> dict:
    """依規則將 Top/Advisor 分類為 left/center/right；SubTop、Middle 另外歸類"""
    zones = {"left": [], "center": [], "right": []}
    for r in df[df["Layer"].isin(["Top", "Advisor"])].to_dict("records"):
        zone = classify_leadership_zone(r)
        zones.setdefault(zone, []).append(r)
    zones["subtop"] = df[df["Layer"] == "SubTop"].to_dict("records")
    zones["middle"] = df[df["Layer"] == "Middle"].to_dict("records")
    return zones


def _leader_card_height(row) -> Emu:
    """依是否填寫 Company 欄位與是否有徽章，動態計算單張領導卡片高度（與 _add_leader_card 繪製邏輯一致）"""
    company = safe_text(row.get("Company", ""), "")
    badges = parse_badges(row["Badges"])
    return _leader_card_base_height(bool(company), bool(badges))


CARD_STACK_GAP = Inches(0.2)


def _stack_height(rows) -> Emu:
    if not rows:
        return Inches(0)
    total = Inches(0)
    for i, r in enumerate(rows):
        total += _leader_card_height(r)
        if i < len(rows) - 1:
            total += CARD_STACK_GAP
    return total


def build_org_chart_pptx(df: pd.DataFrame, stats: dict, photos: dict,
                          host_org: str = "", client_org: str = "") -> io.BytesIO:
    slide_w = Inches(13.333)
    zones = _leadership_zones(df)
    center_stack = zones["center"] + zones["subtop"]  # SubTop 垂直排列於 Top 正下方，視為中欄同一堆疊

    # ---------- 第一階段：純數值試算版面高度，避免區塊互相重疊 ----------
    leader_top_calc = Inches(1.15) if (host_org or client_org) else Inches(0.55)
    leader_col_h_calc = max(_stack_height(zones["left"]), _stack_height(center_stack),
                             _stack_height(zones["right"]), Inches(0))
    gap_v = Inches(0.3)
    middle_top_calc = leader_top_calc + leader_col_h_calc + gap_v
    middle_h_calc = _leader_card_height(zones["middle"][0]) if zones["middle"] else Inches(0)
    groups_top_calc = middle_top_calc + middle_h_calc + (gap_v if zones["middle"] else Inches(0))
    if not zones["middle"]:
        groups_top_calc = leader_top_calc + leader_col_h_calc + gap_v

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
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.name = GLOBAL_FONT_NAME
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _set_pptx_run_ea_font(run)
        leader_top = Inches(1.15)
    else:
        leader_top = Inches(0.55)

    # --- 中軸領導層：標準三欄對稱結構（左／中／右） ---
    slot_w = (slide_w - Inches(1.0)) / 3
    zone_cx = {
        "left": Inches(0.5) + slot_w * 0.5,
        "center": Inches(0.5) + slot_w * 1.5,
        "right": Inches(0.5) + slot_w * 2.5,
    }
    card_w = min(Inches(2.7), slot_w - Inches(0.2))

    def _render_stack(zone_key, rows):
        """垂直堆疊繪製同一欄的人員卡片，回傳 (該欄第一張卡片中心x, 第一張卡片底部y, 該欄底部y)"""
        cx = zone_cx[zone_key]
        y = leader_top
        first_card_bottom = None
        for r in rows:
            h = _leader_card_height(r)
            _add_leader_card(slide, cx, y, r, photos, card_w=card_w)
            if first_card_bottom is None:
                first_card_bottom = y + h
            y += h + CARD_STACK_GAP
        col_bottom = y - CARD_STACK_GAP if rows else leader_top
        return cx, first_card_bottom, col_bottom

    left_cx, left_first_bottom, _ = _render_stack("left", zones["left"])
    center_cx, center_first_bottom, _ = _render_stack("center", center_stack)
    right_cx, right_first_bottom, _ = _render_stack("right", zones["right"])

    middle_top = leader_top + leader_col_h_calc + gap_v
    middle_cx = slide_w / 2
    if zones["middle"]:
        _add_leader_card(slide, middle_cx, middle_top, zones["middle"][0], photos, card_w=Inches(2.7))
    groups_top = groups_top_calc

    # 連接線：左／中／右三欄「第一張卡片」下緣 -> 計畫經理卡片上緣
    try:
        for cx, first_bottom in [(left_cx, left_first_bottom), (center_cx, center_first_bottom),
                                  (right_cx, right_first_bottom)]:
            if first_bottom is None:
                continue
            connector = slide.shapes.add_connector(2, int(cx), int(first_bottom),
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

    # --- 頁尾統計時間軸（大數字為刻意的視覺焦點設計，經確認不受 9-12pt 文字規範限制） ---
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
        zone = classify_leadership_zone(row)
        if zone == "right":
            return safe_text(row["Role"], "協同主持人")
        return "計畫主持人"
    if row["Layer"] == "SubTop":
        return safe_text(row["Role"], "次頂層人員")
    return LAYER_CATEGORY_MAP_DEFAULT.get(row["Layer"], row["Role"] or "人員")


def build_staffing_table_rows(df: pd.DataFrame):
    """回傳依 Output2 樣式排序、格式化後的 list[dict]，供表格預覽/匯出共用"""
    order = {"Top": 0, "SubTop": 1, "Advisor": 2, "Middle": 3, "GroupLeader": 4, "GroupMember": 5}
    df2 = df.copy()
    df2["_order"] = df2["Layer"].map(order).fillna(9)
    # Top 排序：主持人優先於協同主持人
    df2["_sub"] = df2.apply(lambda r: 0 if (r["Layer"] == "Top" and classify_leadership_zone(r) != "right") else 1, axis=1)
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


def _set_docx_run_font(run, size_pt=12, bold=None, font_name=GLOBAL_FONT_NAME):
    """設定 DOCX run 字型，含 w:eastAsia 屬性，確保中文正確套用指定字型"""
    run.font.name = font_name
    run.font.size = DocxPt(size_pt)
    if bold is not None:
        run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(docx_qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(docx_qn("w:eastAsia"), font_name)


def _set_docx_default_font(doc, font_name=GLOBAL_FONT_NAME):
    """設定文件預設樣式字型（保險機制，涵蓋未逐一設定的段落）"""
    try:
        style = doc.styles["Normal"]
        style.font.name = font_name
        rPr = style.element.get_or_add_rPr()
        rFonts = rPr.find(docx_qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(docx_qn("w:eastAsia"), font_name)
    except Exception:
        pass


def build_staffing_table_docx(df: pd.DataFrame) -> io.BytesIO:
    rows = build_staffing_table_rows(df)
    doc = Document()
    _set_docx_default_font(doc)
    section = doc.sections[0]
    section.orientation = 0

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("主要工作人力配置表")
    _set_docx_run_font(run, size_pt=16, bold=True)

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
        _set_docx_run_font(run, size_pt=12, bold=True)
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
            _set_docx_run_font(run, size_pt=12, bold=False)
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
                    run.font.name = GLOBAL_FONT_NAME
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    _set_pptx_run_ea_font(run)
        for ri, r in enumerate(chunk, start=1):
            for ci, h in enumerate(headers):
                cell = table.cell(ri, ci)
                cell.text = r[h]
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(9.5)
                        run.font.name = GLOBAL_FONT_NAME
                        _set_pptx_run_ea_font(run)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


# =====================================================================
# ===========  Output 3：主要人員經歷敘述（組長以上）(DOCX)  ============
# =====================================================================

SENIOR_LAYERS = ["Top", "SubTop", "Advisor", "Middle", "GroupLeader"]

# Gemini 模型選單（側邊欄下拉選單選項）；「自訂模型」為特殊值，選中後另跳出文字輸入框
GEMINI_MODEL_OPTIONS = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-flash",
                         "gemini-flash-latest", "自訂模型"]

# 備援退避鏈：使用者選擇的模型優先，其餘依序作為 404/NOT_FOUND 時的自動備援
GEMINI_FALLBACK_CHAIN = ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-flash", "gemini-flash-latest"]


def get_secret_gemini_key() -> str:
    """安全讀取 Streamlit Secrets 中的 GEMINI_API_KEY；未設定 secrets.toml 時不拋出例外"""
    try:
        return (st.secrets.get("GEMINI_API_KEY", "") or "").strip()
    except Exception:
        return ""


def get_effective_gemini_key() -> str:
    """依規則取得實際呼叫用金鑰：優先 Streamlit Secrets，其次側邊欄手動輸入值"""
    return get_secret_gemini_key() or st.session_state.get("gemini_api_key", "").strip()


def build_gemini_model_candidates(selected_model: str, custom_model: str) -> list:
    """組成備援模型清單：[使用者選取/自訂模型] + 預設備援鏈，並自動去除重複項（保留原順序）"""
    primary = custom_model.strip() if selected_model == "自訂模型" else selected_model.strip()
    raw_candidates = [primary] + GEMINI_FALLBACK_CHAIN
    seen = set()
    candidates = []
    for m in raw_candidates:
        m = (m or "").strip()
        if m and m not in seen:
            seen.add(m)
            candidates.append(m)
    return candidates


def _is_model_not_found_error(e: Exception) -> bool:
    """判斷例外是否為「模型不存在 / 404」類型錯誤，用以觸發自動退避切換下一個候選模型"""
    if gemini_errors is not None and isinstance(e, gemini_errors.ClientError):
        code = getattr(e, "code", None)
        if code == 404:
            return True
    msg = str(e).upper()
    return "404" in msg or "NOT_FOUND" in msg or "NOT FOUND" in msg


def refine_bio_with_gemini(api_key: str, name: str, job_description: str, bio_narrative: str,
                            model_candidates: list) -> tuple:
    """呼叫 Gemini (google-genai) 依擬任工作內容，將原始履歷潤飾為 200 字以內備標敘述。
    遇 404 / Model Not Found 會自動無感切換至下一個備援模型；其餘錯誤（金鑰、額度、網路等）直接中止並拋出。
    回傳 (潤飾後文字, 實際成功使用的模型 ID)。
    """
    if gemini_genai is None:
        raise RuntimeError("尚未安裝 google-genai 套件，請確認 requirements.txt 是否包含 google-genai")
    if not model_candidates:
        raise RuntimeError("沒有可用的 Gemini 模型可供呼叫，請確認模型選單設定")

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
    last_error = None
    for model_id in model_candidates:
        try:
            response = client.models.generate_content(model=model_id, contents=prompt)
            text = (getattr(response, "text", "") or "").strip()
            if not text:
                raise RuntimeError("Gemini 未回傳有效內容")
            return text, model_id
        except Exception as e:
            if _is_model_not_found_error(e):
                last_error = e
                continue  # 404 / 模型不存在：無感切換下一個備援模型
            raise  # 非 404 錯誤（金鑰無效、額度超限、網路逾時等）：立即中止，不逐一嘗試其餘模型

    raise RuntimeError(f"候選模型（{', '.join(model_candidates)}）皆回傳 404 Not Found，最後錯誤：{last_error}")


def build_bio_docx(df: pd.DataFrame, photos: dict) -> io.BytesIO:
    order = {"Top": 0, "SubTop": 1, "Advisor": 2, "Middle": 3, "GroupLeader": 4}
    df2 = df[df["Layer"].isin(SENIOR_LAYERS)].copy()
    df2["_order"] = df2["Layer"].map(order).fillna(9)
    df2["_sub"] = df2.apply(lambda r: 0 if (r["Layer"] == "Top" and classify_leadership_zone(r) != "right") else 1, axis=1)
    df2 = df2.sort_values(["_order", "_sub"], kind="stable")

    doc = Document()
    _set_docx_default_font(doc)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("主要人員經歷")
    _set_docx_run_font(run, size_pt=16, bold=True)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.columns[0].width = Cm(4.2)
    table.columns[1].width = Cm(10.5)

    hdr = table.rows[0].cells
    hdr[0].text = ""
    _set_docx_run_font(hdr[0].paragraphs[0].add_run("姓名職稱"), size_pt=12, bold=True)
    hdr[1].text = ""
    _set_docx_run_font(hdr[1].paragraphs[0].add_run("專長說明"), size_pt=12, bold=True)
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
            ph_cell.text = ""
            _set_docx_run_font(ph_cell.paragraphs[0].add_run("照片待補"), size_pt=12, bold=False)
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
        _set_docx_run_font(run1, size_pt=12, bold=True)

        p2 = text_cell.add_paragraph()
        run2 = p2.add_run(safe_text(r["BioNarrative"]))
        _set_docx_run_font(run2, size_pt=12, bold=False)

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
if "gemini_model_choice" not in st.session_state:
    st.session_state.gemini_model_choice = GEMINI_MODEL_OPTIONS[0]
if "gemini_custom_model" not in st.session_state:
    st.session_state.gemini_custom_model = ""

# ---------------- ① 側邊欄：資料上傳 + Gemini AI 設定 ----------------
with st.sidebar:
    st.markdown("### ① 資料上傳")

    excel_file = st.file_uploader("Template_Staffing.xlsx", type=["xlsx"], label_visibility="visible")
    if excel_file is not None:
        try:
            st.session_state.df = load_staffing_excel(excel_file)
            st.session_state.stats = compute_footer_stats(st.session_state.df)  # 上傳後自動計算，無需點按鈕
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
    st.session_state.gemini_model_choice = st.selectbox(
        "Gemini 模型選取", GEMINI_MODEL_OPTIONS,
        index=GEMINI_MODEL_OPTIONS.index(st.session_state.gemini_model_choice)
        if st.session_state.gemini_model_choice in GEMINI_MODEL_OPTIONS else 0,
        help="潤飾時若所選模型回傳 404 / Model Not Found，將自動無感切換至下一個備援模型。",
    )
    if st.session_state.gemini_model_choice == "自訂模型":
        st.session_state.gemini_custom_model = st.text_input(
            "自訂 Model ID", value=st.session_state.gemini_custom_model,
            placeholder="例如：gemini-2.5-pro",
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
        st.caption("Excel 上傳成功後已自動計算，如需重算或還原可點擊下方按鈕。")
        if st.button("🔄 重新計算", use_container_width=True):
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
        company = safe_text(row.get("Company", ""), "")
        company_html = f"<div style='font-size:10px;color:#888;'>{company}</div>" if company else ""
        return f"""
        <div style='display:flex;gap:8px;align-items:center;background:#fff;border:1px solid #D8DFE8;
                    border-radius:8px;padding:6px 8px;margin-bottom:6px;'>
            {photo_html}
            <div>
                <div style='font-weight:700;font-size:13px;'>{safe_text(row['Name'])}</div>
                <div style='font-size:11px;color:#666;'>{safe_text(row['Title'],'')}</div>
                {company_html}
                <div style='margin-top:2px;'>{badge_html(badges)}</div>
            </div>
        </div>
        """

    zones = _leadership_zones(df)
    center_stack = zones["center"] + zones["subtop"]
    middles = zones["middle"]

    org_header = ""
    if host_org or client_org:
        org_header = "<div style='text-align:center;margin-bottom:10px;'>"
        if host_org:
            org_header += f"<span style='background:#3E5C86;color:#fff;padding:4px 14px;border-radius:14px;margin-right:8px;font-size:13px;'>主辦機關：{host_org}</span>"
        if client_org:
            org_header += f"<span style='background:#3E5C86;color:#fff;padding:4px 14px;border-radius:14px;font-size:13px;'>委辦機關：{client_org}</span>"
        org_header += "</div>"
    st.markdown(org_header, unsafe_allow_html=True)

    st.caption("標準三欄對稱結構：左區（顧問／品質督導）｜中區（計畫主持人＋次頂層 SubTop）｜右區（協同／代表廠商）")
    zone_l, zone_c, zone_r = st.columns(3)
    zone_titles = {"left": "左區", "center": "中區", "right": "右區"}
    for zone_key, col, people in [("left", zone_l, zones["left"]),
                                   ("center", zone_c, center_stack),
                                   ("right", zone_r, zones["right"])]:
        with col:
            st.markdown(f"<div style='text-align:center;color:#8794A6;font-size:11px;font-weight:700;"
                        f"margin-bottom:4px;'>{zone_titles[zone_key]}</div>", unsafe_allow_html=True)
            if not people:
                st.markdown("<div style='text-align:center;color:#B7C0CC;font-size:11px;padding:8px;'>（無）</div>",
                            unsafe_allow_html=True)
            for person in people:
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
        st.caption("僅呈現組長以上同仁（Layer: Top / SubTop / Advisor / Middle / GroupLeader）之經歷敘述。")
    with top_r:
        refine_clicked = st.button(
            "✨ 一鍵使用 Gemini 依據擬任工作潤飾履歷",
            use_container_width=True,
            help="依據 JobDescription（擬任工作）潤飾 BioNarrative（履歷）；"
                 "若所選模型 404 / Model Not Found，將自動切換備援模型。",
        )

    if refine_clicked:
        effective_key = get_effective_gemini_key()
        if not effective_key:
            st.warning("請先輸入 Gemini API Key")
        elif gemini_genai is None:
            st.error("尚未安裝 google-genai 套件，請確認 requirements.txt 是否包含 google-genai 並重新部署")
        else:
            model_candidates = build_gemini_model_candidates(
                st.session_state.gemini_model_choice, st.session_state.gemini_custom_model,
            )
            if not model_candidates:
                st.warning("請選擇 Gemini 模型，或於「自訂模型」輸入 Model ID")
            else:
                target_mask = st.session_state.df["Layer"].isin(SENIOR_LAYERS)
                target_idx = st.session_state.df[target_mask].index.tolist()
                total = len(target_idx)

                if total == 0:
                    st.info("目前無組長以上人員資料可供潤飾")
                else:
                    progress = st.progress(0, text=f"AI 潤飾處理中... (0/{total})")
                    fail_count = 0
                    used_models = set()
                    for i, idx in enumerate(target_idx):
                        row = st.session_state.df.loc[idx]
                        try:
                            new_bio, used_model = refine_bio_with_gemini(
                                effective_key,
                                row["Name"], row["JobDescription"], row["BioNarrative"],
                                model_candidates,
                            )
                            st.session_state.df.loc[idx, "BioNarrative"] = new_bio
                            used_models.add(used_model)
                        except Exception as e:
                            fail_count += 1
                            st.warning(f"{safe_text(row['Name'])} 履歷潤飾失敗：{e}")
                        progress.progress((i + 1) / total, text=f"AI 潤飾處理中... ({i + 1}/{total})")
                    progress.empty()

                    model_note = f"（使用模型：{', '.join(sorted(used_models))}）" if used_models else ""
                    if fail_count == 0:
                        st.toast(f"履歷 AI 潤飾完成！{model_note}", icon="✨")
                    else:
                        st.toast(f"履歷 AI 潤飾完成（{total - fail_count}/{total} 筆成功）{model_note}", icon="⚠️")
                    st.rerun()

    df = st.session_state.df  # 潤飾後重新取用最新資料

    senior_df = df[df["Layer"].isin(SENIOR_LAYERS)]
    order = {"Top": 0, "SubTop": 1, "Advisor": 2, "Middle": 3, "GroupLeader": 4}
    senior_df = senior_df.copy()
    senior_df["_order"] = senior_df["Layer"].map(order).fillna(9)
    senior_df["_sub"] = senior_df.apply(lambda r: 0 if (r["Layer"] == "Top" and classify_leadership_zone(r) != "right") else 1, axis=1)
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
