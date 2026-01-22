# CEO Project - Executive CV Automation Tool

自動化主管履歷資料蒐集與 PowerPoint 簡報生成工具。

---

## 快速開始

### 1. 安裝相依套件

```bash
pip install pandas openpyxl python-pptx requests python-dotenv ddgs
```

### 2. 設定 API 金鑰

在專案根目錄建立 `.env` 檔案：
```
PERPLEXITY_API_KEY=your_api_key_here
```

### 3. 執行自動化

直接按兩下對應的批次檔：

| 批次檔 | 用途 |
|--------|------|
| `run_automation.bat` | 完整流程（資料擴充 + 照片 + PPT）|
| `enrich_cell.bat` | 針對特定儲存格收集資料（如 H26）|
| `search_photos.bat` | 僅搜尋照片 + 審核報告 |
| `generate_ppt.bat` | 僅生成 PPT（套用已選擇的照片）|

---

## 檔案結構

```
CEO/
├── .env                          # API 金鑰設定
├── run_automation.bat            # 主程式入口（完整流程）
├── enrich_cell.bat               # 針對特定儲存格收集資料
├── search_photos.bat             # 照片搜尋 + 審核
├── generate_ppt.bat              # 僅生成 PPT
├── Standard Example.xlsx         # 原始資料檔
├── CV_標準範本.pptx              # PPT 範本
├── src/
│   ├── enrich_data.py           # 資料擴充程式
│   ├── enrich_cell.py           # 特定儲存格資料收集
│   ├── generate_ppt.py          # PPT 生成程式
│   ├── ppt/                     # PPT 模組
│   └── search/                  # 搜尋模組
└── output/
    ├── data/                    # 擴充後的資料
    │   ├── Standard_Example_Enriched.xlsx
    │   ├── photo_review.html    # 照片審核報告
    │   └── photo_selections.json # 照片選擇結果
    └── ppt/                     # 生成的 PPT（按專業分類）
        ├── 會計_財務類/
        ├── 法務類/
        ├── 商務_管理類/
        ├── 產業專業類/
        └── 未分類/
```

---

## 使用流程

### Step 1: 資料擴充

輸入列號後，系統會自動：
- 使用 Perplexity API 搜尋主管資訊
- 使用 DuckDuckGo 搜尋照片
- 將結果儲存到 Excel

### Step 2: 照片審核（可選）

如果需要審核照片：
1. 選擇「1」開啟照片審核報告
2. 在瀏覽器中點選正確的照片
3. 點「儲存選擇」下載 JSON
4. 將 JSON 移動到 `output/data/photo_selections.json`
5. 按任意鍵繼續

### Step 3: PPT 生成

系統會自動：
- 套用照片選擇（如果有）
- 根據專業分類建立資料夾
- 生成個人 CV 簡報

---

## 分享給同事

將整個資料夾壓縮成 ZIP 即可分享。收到的人只需要：

1. 解壓縮
2. 安裝 Python 3.10+
3. 執行 `pip install pandas openpyxl python-pptx requests python-dotenv ddgs`
4. 建立 `.env` 檔案並填入 API 金鑰
5. 雙擊 `run_automation.bat`

---

## 針對特定儲存格收集資料

使用 `enrich_cell.bat` 可以針對特定欄位和列號收集資料：

### 欄位對應表

| 編號 | 欄位代號 | 欄位名稱 |
|------|----------|----------|
| 1 | C | 年齡 |
| 2 | F | 專業分類 |
| 3 | G | 專業背景 |
| 4 | H | 學歷 |
| 5 | I | 主要經歷 |
| 6 | J | 現職/任 |
| 7 | K | 個人特質 |
| 8 | L | 現擔任獨董家數(年) |
| 9 | M | 擔任獨董年資(年) |
| 10 | N | 電子郵件 |
| 11 | O | 公司電話 |
| 12 | D | 照片 |

### 使用範例

```bash
# 使用儲存格參照
python src/enrich_cell.py --cell "H26"           # 第 26 列的學歷
python src/enrich_cell.py --cell "H26-H30"       # 第 26-30 列的學歷
python src/enrich_cell.py --cell "H26,I27,J28"   # 多個不同儲存格

# 使用欄位名稱 + 列號
python src/enrich_cell.py --field "學歷" --rows "26"
python src/enrich_cell.py --field "4" --rows "26-30"
python src/enrich_cell.py --field "H" --rows "26,27,28"

# 強制更新已有資料的欄位
python src/enrich_cell.py --cell "H26" --force
```

---

## 手動執行

```bash
# 完整資料擴充（含照片）
python src/enrich_data.py --rows "2, 5-10"

# 僅搜尋照片（不使用 Perplexity API）
python src/enrich_data.py --rows "2, 5-10" --photos-only

# 僅生成 PPT
python src/generate_ppt.py --rows "2, 5-10"
```

---

## 專業分類

系統會根據主管背景自動分類：

| 分類 | 說明 |
|------|------|
| 會計/財務類 | 會計師、財務長、CFO |
| 法務類 | 律師、法官、法務長 |
| 商務/管理類 | CEO、總經理、董事長 |
| 產業專業類 | 工程師、技術專家 |
| 其他專門職業 | 建築師、技師等 |
| 未分類 | 無法判定時 |
