# 標案專案人力與組織圖 Web 產生器

Streamlit 應用程式：上傳人員資料 Excel 與大頭照 Zip，即可預覽並一鍵匯出：

- **Output 1**：人員組織架構圖（.pptx，原生可編輯圖形）
- **Output 2**：主要工作人力配置表（.docx 表格 / .pptx 表格）
- **Output 3**：專案主要人員介紹（組長以上）（.docx）
- **✨ Gemini AI 智慧履歷潤飾**：一鍵依「擬任工作內容（JobDescription）」將組長以上人員的履歷（BioNarrative）潤飾為 200 字以內、專業精準的備標敘述。

## Gemini AI 智慧履歷潤飾（做法 B）

**API Key 讀取優先順序**：`st.secrets["GEMINI_API_KEY"]`（Streamlit Secrets）優先 → 若未設定，才使用側邊欄手動輸入的備用金鑰。

- **建議做法（部署到 Streamlit Community Cloud）**：於 App 的 Settings → Secrets 貼上：
  ```toml
  GEMINI_API_KEY = "你的 Gemini API Key"
  ```
  設定後側邊欄輸入框會自動帶入（並顯示「🔒 已從 Streamlit Secrets 自動讀取」提示），使用者不需再手動輸入，點擊潤飾按鈕也不會跳出警告。
- **備用做法（未設定 Secrets 時）**：於左側邊欄「② Gemini AI 設定」自行輸入 API Key（密碼輸入框，僅存於本次瀏覽器工作階段，不會被儲存或上傳）。

**使用流程**：

1. 於 Tab 3「人員經歷敘述預覽」點擊「✨ 一鍵使用 Gemini 依據擬任工作潤飾履歷」。
2. 系統會針對 Layer 為 `Top` / `Advisor` / `Middle` / `GroupLeader` 的人員，將其 `Name`、`JobDescription`、`BioNarrative` 傳給 `gemini-2.5-flash`，產生 200 字以內、專業精準的備標履歷敘述，並直接覆寫 `st.session_state.df` 中對應的 `BioNarrative` 欄位。
3. 完成後會以 `st.toast` 提示「履歷 AI 潤飾完成！」，頁面自動重新整理，Tab 3 與 Output 3 匯出檔案皆會反映潤飾後的最新內容。
4. 若 Secrets 與側邊欄皆未提供 API Key，點擊按鈕才會顯示 `st.warning("請先輸入 Gemini API Key")`；若單筆人員呼叫失敗（例如額度不足、網路逾時），該筆會顯示警告並保留原始履歷，不影響其餘人員的潤飾結果。

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 Streamlit Community Cloud

1. 將 `app.py`、`requirements.txt` 一併推送至 GitHub repository。
2. 於 [share.streamlit.io](https://share.streamlit.io) 新增 App，指向該 repo 與 `app.py`。
3. 部署完成後即可於瀏覽器上傳 Excel / Zip 使用。

## Excel 欄位格式（Template_Staffing.xlsx）

| 欄位 | 說明 |
|---|---|
| Layer | Top / Advisor / Middle / GroupLeader / GroupMember |
| Role | 職稱角色（如：計畫主持人、協同主持人、計畫顧問、計畫經理、組長、組員） |
| GroupName | 所屬工作組別名稱（Top/Advisor/Middle 可填「—」） |
| Name | 姓名 |
| Title | 職稱（如：資深協理） |
| Badges | 徽章代碼，逗號分隔（如：技,碩,採,品） |
| PhotoName | 對應 Photos.zip 內之照片檔名（如：林金龍.jpg），無照片可填「—」 |
| YearsOfExp | 年資（數字） |
| Degree | 最高學歷科系 |
| JobDescription | 擬任工作內容 |
| Expertise | 相關經歷與專長 |
| BioNarrative | 個人經歷完整敘述（僅組長以上於 Output 3 使用，留空自動顯示「[資料待補]」） |

## 徽章代碼對照（可於 app.py 的 `BADGE_MAP` / `BADGE_COLORS` 自行調整）

技=技師、碩=碩士、博=博士、品=品質管理人員、安=勞工安全衛生人員、
採=採購專業人員、乙=乙級技術士、甲=甲級技術士、景=景觀技師 …

## 容錯設計

- Excel 缺欄位、缺值：自動補空字串／「[資料待補]」，不會導致程式崩潰。
- 照片缺失：Output 1（PPTX）與 Output 3（DOCX）皆會繪製原生可編輯的灰色預留方塊，方便後續於 PowerPoint / Word 直接「變更圖片」取代。
- Output 1 版面高度會依人員數量動態調整，避免組別區塊與頁尾統計列重疊。

## 已知限制／可自訂項目

- `Degree`（最高學歷科系）欄位原樣輸出，未自動切分為「校名／科系＋學位」格式，可依需求於 Excel 來源先行整理。
- 主辦機關／委辦機關名稱為選填文字輸入框，非 Excel 欄位。
- 徽章全名對照表（BADGE_MAP）為範例對照，請依實際專案調整。
