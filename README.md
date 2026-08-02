# 標案專案人力與組織圖 Web 產生器

Streamlit 應用程式：上傳人員資料 Excel 與大頭照 Zip，即可預覽並一鍵匯出：

- **Output 1**：人員組織架構圖（.pptx，原生可編輯圖形）
- **Output 2**：主要工作人力配置表（.docx 表格 / .pptx 表格）
- **Output 3**：專案主要人員介紹（組長以上）（.docx）
- **✨ Gemini AI 智慧履歷潤飾**：一鍵依「擬任工作內容（JobDescription）」將組長以上人員的履歷（BioNarrative）潤飾為 200 字以內、專業精準的備標敘述。

## Gemini AI 智慧履歷潤飾（含模型自動退避機制）

**API Key 讀取優先順序**：`st.secrets["GEMINI_API_KEY"]`（Streamlit Secrets）優先 → 若未設定，才使用側邊欄手動輸入的備用金鑰。

- **建議做法（部署到 Streamlit Community Cloud）**：於 App 的 Settings → Secrets 貼上：
  ```toml
  GEMINI_API_KEY = "你的 Gemini API Key"
  ```
  設定後側邊欄輸入框會自動帶入（並顯示「🔒 已從 Streamlit Secrets 自動讀取」提示），使用者不需再手動輸入，點擊潤飾按鈕也不會跳出警告。
- **備用做法（未設定 Secrets 時）**：於左側邊欄「② Gemini AI 設定」自行輸入 API Key（密碼輸入框，僅存於本次瀏覽器工作階段，不會被儲存或上傳）。

**模型選取與自動退避（Rolling Fallback）**：

- 側邊欄「Gemini 模型選取」下拉選單：`gemini-3.6-flash` / `gemini-3.5-flash` / `gemini-2.5-flash` / `gemini-flash-latest` / `自訂模型`（選中後另跳出 Model ID 輸入框）。
- 候選鏈組成方式：`[使用者選取/自訂模型] + [gemini-3.6-flash, gemini-3.5-flash, gemini-2.5-flash, gemini-flash-latest]`，自動去除重複項。
- 呼叫時依序嘗試候選模型；**僅當遇到 404 / Model Not Found 時**才無感切換下一個備援模型。其餘錯誤（API Key 無效、額度超限、網路逾時等）會直接中止並顯示明確錯誤訊息，不會浪費時間逐一嘗試所有模型。
- ⚠️ 官方公告 `gemini-2.5-flash` 將於 2026/10/16 起停用，建議盡快改用 3.x 系列模型；清單中保留它僅作為過渡期備援。

**使用流程**：

1. 於 Tab 3「人員經歷敘述預覽」點擊「✨ 一鍵使用 Gemini 依據擬任工作潤飾履歷」。
2. 系統會針對 Layer 為 `Top` / `SubTop` / `Advisor` / `Middle` / `GroupLeader` 的人員，將其 `Name`、`JobDescription`、`BioNarrative` 傳給 Gemini，產生 200 字以內、專業精準的備標履歷敘述，並直接覆寫 `st.session_state.df` 中對應的 `BioNarrative` 欄位。
3. 完成後會以 `st.toast` 提示「履歷 AI 潤飾完成！」（並註明實際使用的模型），頁面自動重新整理，Tab 3 與 Output 3 匯出檔案皆會反映潤飾後的最新內容。
4. 若 Secrets 與側邊欄皆未提供 API Key，點擊按鈕才會顯示 `st.warning("請先輸入 Gemini API Key")`；若單筆人員呼叫失敗，該筆會顯示警告並保留原始履歷，不影響其餘人員的潤飾結果。

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
| Layer | `Top` / `SubTop` / `Advisor` / `Middle` / `GroupLeader` / `GroupMember` |
| Role | 職稱角色（如：計畫主持人、協同主持人、計畫顧問、代表廠商、設計負責人、計畫經理、組長、組員） |
| GroupName | 所屬工作組別名稱（Top/SubTop/Advisor/Middle 可填「—」） |
| Name | 姓名 |
| Title | 職稱（如：資深協理） |
| Company | （選填）公司名稱，僅顯示於 Output 1 組織圖卡片職稱下方 |
| Badges | 徽章代碼，逗號分隔（如：技,碩,採,品） |
| PhotoName | 對應 Photos.zip 內之照片檔名（如：林金龍.jpg），無照片可填「—」 |
| YearsOfExp | 年資（數字，留空或填 0 將不計入平均年資分母） |
| Degree | 最高學歷科系 |
| JobDescription | 擬任工作內容 |
| Expertise | 相關經歷與專長 |
| BioNarrative | 個人經歷完整敘述（僅組長以上於 Output 3 使用，留空自動顯示「[資料待補]」） |

## Output 1 組織圖：三欄對稱結構 + SubTop 次頂層

- **左區**：`Layer=='Advisor'` 或 `Role` 含「顧問」「品質督導」。
- **中區**：`Layer=='Top'` 且 `Role` 含「計畫主持人」（未命中任何規則的 `Top` 人員預設歸入中區，避免版面遺漏）。
- **右區**：`Layer=='Top'` 且 `Role` 含「協同」或「代表廠商」。
- **SubTop**（次頂層，如共同投標/設計分包之協同主持人、設計負責人）：垂直排列於中區「計畫主持人」卡片正下方。
- 版面高度（含每張卡片的動態高度、分組區塊、頁尾）皆為程式試算後動態決定，人員或徽章數量增加時投影片會自動加高，避免任何重疊。

## 頁尾統計數據

- **自動觸發**：Excel 上傳成功後立即自動計算，無需點擊按鈕；側邊欄另保留「🔄 重新計算」按鈕作為手動重算/還原的備用選項。
- **平均年資公式**：`df['YearsOfExp_num'].replace(0, pd.NA).dropna().mean()`，未填寫或填 0 的人員不計入分母，避免拉低團隊平均年資。

## 字型與字級規範

- 全域中文字型統一為 **Microsoft JhengHei（微軟正黑體）**，DOCX 各段落／表格文字皆已設定 `w:eastAsia` 屬性，PPTX 亦補上對應的東亞字型設定，確保中文於各種檢視器正確顯示。
- **Output 1（PPTX）**：一般文字字級落在 9–12pt 區間。例外：頁尾醒目大數字（如「12」「13」）維持較大字級（20pt），此為刻意保留的視覺焦點設計，避免失去原參考檔案的吸睛效果。
- **Output 2 / Output 3（DOCX）**：文件大標題維持 16pt 加粗；表格內容與 Output 3 內文統一為 12pt（表頭／姓名列加粗）。

## 徽章代碼對照（可於 app.py 的 `BADGE_MAP` / `BADGE_COLORS` 自行調整）

技=技師、碩=碩士、博=博士、品=品質管理人員、安=勞工安全衛生人員、
採=採購專業人員、乙=乙級技術士、甲=甲級技術士、景=景觀技師 …

## 容錯設計

- Excel 缺欄位、缺值：自動補空字串／「[資料待補]」，不會導致程式崩潰。
- 照片缺失：Output 1（PPTX）與 Output 3（DOCX）皆會繪製原生可編輯的灰色預留方塊，方便後續於 PowerPoint / Word 直接「變更圖片」取代。
- Output 1 版面高度會依人員數量、徽章數量、Company 欄位動態調整，避免任何區塊互相重疊。
- Gemini 呼叫失敗會逐筆記錄警告並保留原始履歷，不會中斷其餘人員的潤飾流程。

## 已知限制／可自訂項目

- `Degree`（最高學歷科系）欄位原樣輸出，未自動切分為「校名／科系＋學位」格式，可依需求於 Excel 來源先行整理。
- 主辦機關／委辦機關名稱為選填文字輸入框，非 Excel 欄位。
- `Company` 欄位目前僅顯示於 Output 1 組織圖卡片，未擴及 Output 2 表格與 Output 3 敘述。
- 徽章全名對照表（BADGE_MAP）為範例對照，請依實際專案調整。
