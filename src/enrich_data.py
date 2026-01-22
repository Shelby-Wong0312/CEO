"""
enrich_data.py - 資料擴充腳本 (Executive Search & Private Investigator Quality)

根據 Standard Example.xlsx 中的姓名與公司資訊，
使用多重搜尋策略填補空缺欄位。

升級重點 (Phase 30):
1. Elite Executive Search Researcher & Private Investigator 角色
2. Age Heuristic - 從畢業年份/歷史新聞推算年齡
3. Education Deep Dive - 精確搜尋學歷
4. Contact Info Zero-Fail Zone - 100% 驗證才回傳
5. Zero Fabrication 原則 - 寧缺勿錯
6. Python 端照片搜尋 - DuckDuckGo 圖片搜尋繞過 LLM 限制

使用方式:
    python src/enrich_data.py --rows "2, 5-10, 15"
"""

import sys
import io

# 設定標準輸出為 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import argparse
import os
import re
import json
import time
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv

# 導入統一搜尋客戶端
try:
    from src.search import UnifiedSearchClient
    UNIFIED_SEARCH_AVAILABLE = True
except ImportError:
    UNIFIED_SEARCH_AVAILABLE = False

# 嘗試導入 DuckDuckGo 搜尋（作為 fallback）
try:
    from ddgs import DDGS
    DDGS_AVAILABLE = True
except ImportError:
    # 嘗試舊版套件名稱
    try:
        from duckduckgo_search import DDGS
        DDGS_AVAILABLE = True
    except ImportError:
        DDGS_AVAILABLE = False
        if not UNIFIED_SEARCH_AVAILABLE:
            print("警告: ddgs 未安裝，將僅使用 Perplexity API")

# 載入環境變數
load_dotenv()

# === 常數定義 ===
EXCEL_INPUT = "Standard Example.xlsx"
EXCEL_OUTPUT = "output/data/Standard_Example_Enriched.xlsx"
PHOTO_CANDIDATES_JSON = "output/data/photo_candidates.json"
PHOTO_REVIEW_HTML = "output/data/photo_review.html"

# 照片信心度門檻（分數 >= 此值才自動填入）
PHOTO_CONFIDENCE_THRESHOLD = 30

# 需要擴充的欄位（對應 Excel 欄位名稱）
ENRICHABLE_COLUMNS = [
    "年齡",
    "照片",
    "照片狀態",
    "專業分類",
    "專業背景",
    "學歷",
    "主要經歷",
    "現職/任",
    "個人特質",
    "現擔任獨董家數(年)",
    "擔任獨董年資(年)",
    "電子郵件",
    "公司電話"
]

# API 回傳欄位對應到 Excel 欄位
API_TO_EXCEL_MAPPING = {
    "age": "年齡",
    "photo_url": "照片",
    "professional_category": "專業分類",
    "professional_background": "專業背景",
    "education": "學歷",
    "key_experience": "主要經歷",
    "current_position": "現職/任",
    "personal_traits": "個人特質",
    "independent_director_count": "現擔任獨董家數(年)",
    "independent_director_tenure": "擔任獨董年資(年)",
    "email": "電子郵件",
    "phone": "公司電話"
}

# 需要結構化輸出的欄位（使用換行分隔）
STRUCTURED_FIELDS = ["學歷", "主要經歷", "現職/任", "個人特質"]


def parse_row_numbers(rows_str: str) -> list[int]:
    """解析 --rows 參數字串，轉換為 Excel 列號列表。"""
    rows_str = rows_str.replace('，', ',')
    result = set()
    parts = rows_str.replace(" ", "").split(",")

    for part in parts:
        if not part:
            continue
        if "-" in part:
            match = re.match(r"^(\d+)-(\d+)$", part)
            if match:
                start, end = int(match.group(1)), int(match.group(2))
                if start > end:
                    start, end = end, start
                result.update(range(start, end + 1))
            else:
                print(f"警告: 無法解析範圍 '{part}'，已跳過")
        else:
            try:
                result.add(int(part))
            except ValueError:
                print(f"警告: 無法解析數字 '{part}'，已跳過")

    result.discard(1)
    return sorted(result)


def excel_row_to_pandas_index(excel_row: int) -> int:
    """將 Excel 列號轉換為 pandas DataFrame 索引。"""
    return excel_row - 2


def search_with_ddg(query: str, max_results: int = 5) -> list[dict]:
    """
    使用 DuckDuckGo 進行網路搜尋。

    Returns:
        搜尋結果列表，每個結果包含 title, href, body
    """
    if not DDGS_AVAILABLE:
        return []

    try:
        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=max_results, region='tw-tzh'))
            return results
    except Exception as e:
        print(f"    DuckDuckGo 搜尋錯誤: {e}")
        return []


def extract_linkedin_url(results: list[dict]) -> str:
    """從搜尋結果中提取 LinkedIn URL。"""
    for result in results:
        href = result.get('href', '')
        if 'linkedin.com/in/' in href:
            return href
    return ""


def score_image_result(result: dict, name: str, company: str) -> int:
    """
    為圖片搜尋結果評分，分數越高越可靠。

    評分標準：
    - LinkedIn 來源: +50
    - 公司官網來源: +40
    - 新聞網站來源: +20
    - 圖片尺寸合適 (寬高 > 150px): +15
    - 圖片比例接近正方形或直式: +10
    - URL 包含人名: +10
    - 排除不良來源: -100

    Returns:
        評分 (整數)
    """
    score = 0
    image_url = result.get('image', '').lower()
    source_url = result.get('url', '').lower()
    title = result.get('title', '').lower()
    width = result.get('width', 0)
    height = result.get('height', 0)

    # === 來源評分 ===
    # LinkedIn 最可靠
    if 'linkedin.com' in source_url or 'linkedin' in image_url:
        score += 50

    # 公司官網
    company_domain_hints = ['company', 'corporate', 'about', 'team', 'leadership', 'management']
    if any(hint in source_url for hint in company_domain_hints):
        score += 40

    # 新聞網站
    news_sites = ['reuters', 'bloomberg', 'forbes', 'businessweek', 'cna.com', 'udn.com',
                  'ltn.com', 'chinatimes', 'ettoday', 'setn.com', 'bnext', 'technews']
    if any(site in source_url for site in news_sites):
        score += 20

    # === 圖片尺寸評分 ===
    if width > 0 and height > 0:
        # 尺寸合適 (至少 150x150)
        if width >= 150 and height >= 150:
            score += 15

        # 比例接近正方形或直式 (大頭照特徵)
        aspect_ratio = width / height if height > 0 else 0
        if 0.6 <= aspect_ratio <= 1.2:  # 直式或正方形
            score += 10

        # 太寬的圖片可能是橫幅
        if aspect_ratio > 2.0:
            score -= 20

    # === URL/標題包含人名 ===
    name_parts = name.lower().split()
    for part in name_parts:
        if len(part) > 1 and part in image_url:
            score += 10
            break
        if len(part) > 1 and part in title:
            score += 5
            break

    # === 排除不良來源 ===
    bad_keywords = ['logo', 'icon', 'banner', 'placeholder', 'avatar', 'default',
                    'stock', 'shutterstock', 'istockphoto', 'gettyimages', 'dreamstime',
                    'thumbnail', 'sprite', 'emoji', 'badge', 'button']
    if any(bad in image_url for bad in bad_keywords):
        score -= 100

    # 排除社交媒體預設頭像
    if 'default' in image_url and ('profile' in image_url or 'avatar' in image_url):
        score -= 100

    return score


def validate_image_url(url: str) -> bool:
    """
    驗證圖片 URL 是否有效。

    檢查：
    - 是否為有效圖片格式
    - 是否可訪問 (HEAD 請求)
    - Content-Type 是否為圖片
    """
    if not url:
        return False

    lower_url = url.lower()

    # 檢查副檔名
    valid_extensions = ['.jpg', '.jpeg', '.png', '.webp', '.gif']
    has_valid_ext = any(ext in lower_url for ext in valid_extensions)

    # 如果沒有明確副檔名，嘗試 HEAD 請求驗證
    if not has_valid_ext:
        try:
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
            response = requests.head(url, headers=headers, timeout=5, allow_redirects=True)
            content_type = response.headers.get('Content-Type', '')
            if not content_type.startswith('image/'):
                return False
        except:
            # 無法驗證時，假設有效
            pass

    return True


def find_executive_photo_python(name: str, company: str, job_title: str = "") -> dict:
    """
    使用多重策略搜尋高階主管照片 URL（增強版 + 審核模式）。

    改進方案：
    A. 多重搜尋策略 - 嘗試多種關鍵字組合
    B. 來源優先排序 - LinkedIn > 公司官網 > 新聞 > 其他
    C. 圖片尺寸驗證 - 過濾不合適的圖片
    D. 加入職稱搜尋 - 提高準確度
    E. 信心度門檻 - 高分自動填入，低分待審核

    Args:
        name: 主管姓名
        company: 公司名稱
        job_title: 職稱（可選，提高搜尋準確度）

    Returns:
        dict: {
            "best_url": 最佳照片 URL（低於門檻則為空）,
            "best_score": 最佳照片分數,
            "status": "待確認" | "待補充",
            "candidates": [{url, score, source}, ...]
        }
    """
    result = {
        "best_url": "",
        "best_score": 0,
        "status": "待補充",
        "candidates": []
    }

    if not DDGS_AVAILABLE:
        print("    [照片搜尋] DuckDuckGo 未安裝，跳過")
        return result

    # === 建立多重搜尋查詢 ===
    search_queries = []

    # 策略 1: LinkedIn 專屬搜尋（最可靠）
    search_queries.append(f'site:linkedin.com "{name}" {company}')

    # 策略 2: 加入職稱搜尋（如果有職稱）
    if job_title:
        # 取第一行職稱（可能有多行）
        first_title = job_title.split('\n')[0].strip()
        if first_title:
            search_queries.append(f'"{name}" "{first_title}" photo OR portrait')

    # 策略 3: 公司 + 姓名 + 大頭照關鍵字
    search_queries.append(f'"{name}" {company} 照片 OR headshot OR portrait')

    # 策略 4: 一般搜尋
    search_queries.append(f'{name} {company} profile photo')

    # === 收集所有搜尋結果 ===
    all_results = []
    seen_urls = set()

    print(f"    [照片搜尋] 執行 {len(search_queries)} 種搜尋策略...")

    try:
        with DDGS() as ddgs:
            for i, query in enumerate(search_queries, 1):
                try:
                    print(f"    [策略 {i}] {query[:50]}...")
                    results = list(ddgs.images(query, max_results=5))

                    for img in results:
                        image_url = img.get('image', '')
                        if image_url and image_url not in seen_urls:
                            seen_urls.add(image_url)
                            # 基本格式驗證
                            if validate_image_url(image_url):
                                all_results.append(img)

                    time.sleep(0.5)  # 避免請求過快

                except Exception as e:
                    print(f"    [策略 {i}] 搜尋失敗: {e}")
                    continue

    except Exception as e:
        print(f"    [照片搜尋] DuckDuckGo 錯誤: {e}")
        return result

    if not all_results:
        print("    [照片搜尋] 未找到任何結果")
        return result

    print(f"    [照片搜尋] 收集到 {len(all_results)} 張候選圖片，評分中...")

    # === 評分並排序 ===
    scored_results = []
    for img_result in all_results:
        score = score_image_result(img_result, name, company)
        if score > -50:  # 排除明顯不合適的
            scored_results.append((score, img_result))

    if not scored_results:
        print("    [照片搜尋] 所有候選圖片評分過低")
        return result

    # 按分數排序
    scored_results.sort(key=lambda x: x[0], reverse=True)

    # 輸出前 3 名的評分（除錯用）
    print(f"    [照片搜尋] 候選圖片評分:")
    for i, (score, img_result) in enumerate(scored_results[:3], 1):
        url = img_result.get('image', '')[:50]
        source = img_result.get('url', '')[:30]
        print(f"      #{i} 分數:{score:3d} | {url}... | 來源:{source}...")

    # 儲存所有候選照片（最多 5 張）
    for score, img_result in scored_results[:5]:
        result["candidates"].append({
            "url": img_result.get('image', ''),
            "score": score,
            "source": img_result.get('url', ''),
            "title": img_result.get('title', ''),
            "width": img_result.get('width', 0),
            "height": img_result.get('height', 0)
        })

    # 選擇最高分的圖片
    best_score, best_img = scored_results[0]
    best_url = best_img.get('image', '')
    result["best_score"] = best_score

    # 根據信心度門檻決定是否自動填入
    if best_score >= PHOTO_CONFIDENCE_THRESHOLD:
        result["best_url"] = best_url
        result["status"] = "待確認"
        print(f"    [照片搜尋] ✓ 分數 {best_score} >= {PHOTO_CONFIDENCE_THRESHOLD}，自動填入（待確認）")
    else:
        result["best_url"] = ""  # 不自動填入
        result["status"] = "待補充"
        print(f"    [照片搜尋] ✗ 分數 {best_score} < {PHOTO_CONFIDENCE_THRESHOLD}，需人工審核")

    return result


def extract_info_from_snippets(results: list[dict], name: str) -> dict:
    """
    從搜尋結果的摘要中提取資訊。

    注意：此函式已簡化，只提取 LinkedIn URL。
    學歷、年齡等結構化資訊改由 Perplexity API 處理，
    避免從不相關的搜尋片段中誤抓資料。

    Args:
        results: 搜尋結果列表
        name: 搜尋的人名（用於驗證相關性）

    Returns:
        提取的資訊字典（僅含 LinkedIn URL）
    """
    extracted = {}

    # 只提取 LinkedIn URL（這是可靠的）
    linkedin = extract_linkedin_url(results)
    if linkedin:
        extracted['LinkedIn'] = linkedin

    # 不再從搜尋片段中提取學歷、年齡等資訊
    # 這些資訊應由 Perplexity API 結構化回傳，以確保準確性

    return extracted


def build_executive_search_prompt(name: str, company: str) -> str:
    """
    建立 Executive Search Researcher & Private Investigator 品質的搜尋提示詞。

    Phase 29 升級：包含年齡推算邏輯 (Age Heuristic)
    """
    import datetime
    current_year = datetime.datetime.now().year

    prompt = f"""# Role
You are an elite Executive Search Researcher & Private Investigator (高階獵頭與徵信專家).
Your mission is to construct a **complete, verified profile** for the target executive.
You do not give up easily. You dig deep, infer intelligently, and verify strictly.

# Target Executive
Name: {name}
Company: {company}

# Search & Inference Protocol (MUST FOLLOW)

## 1. The "Age" Heuristic (Critical)
**Problem:** Direct "Age" is often missing.
**Solution:** You MUST attempt to **calculate** it if not found directly.
- **Step A:** Search for "Bachelor's degree year" or "University graduation year".
    - *Formula:* If they graduated Bachelor's in 1990 -> 1990 + 22 = Born approx 1968 -> Current Age = {current_year} - 1968.
- **Step B:** Search for "Date of birth" or "Born in".
- **Step C:** Search for old news (e.g., "Appointed in 2015 at age 45" -> Current Age = 45 + ({current_year} - 2015)).
- **Output:** Return the number (e.g., "55歲") only if you have a grounded estimation. Otherwise, null.

## 2. The "Education" Deep Dive
- Ignore generic bios. Search specifically for:
    - `"{name}" "{company}" education`
    - `"{name}" "{company}" alumni`
    - `"{name}" "{company}" LinkedIn`
    - `"{name}" "{company}" 畢業`
- **Requirement:** Must list Degree + School (e.g., "國立台灣大學 電機系 學士").

## 3. CONTACT INFO: The "Zero-Fail" Zone (Strict Rules)
- **Email & Phone:** These are **High-Risk Fields**.
- **Rule:** You are FORBIDDEN from guessing. You are FORBIDDEN from constructing emails like `name@company.com` unless you find it indexed on the web.
- **Search Targets:** Look for PDF presentations, conference attendee lists, or official press contacts.
- **Verification:** If you find `info@company.com`, IGNORE IT. Only personal work emails (e.g., `john.doe@company.com`) are accepted.
- **Output:** If 100% sure, return the string. If 99% sure, return `""`. **Accuracy > Availability.**

## 4. Professional Category (專業分類) - REQUIRED
Classify the person into ONE of the following categories based on their PRIMARY expertise:

**Categories (MUST choose exactly ONE):**
- "會計/財務類" - For: 會計師、財務長、CFO、財會學者、審計師
- "法務類" - For: 律師、法官、檢察官、法學教授、法務長
- "商務/管理類" - For: 企業經營者、管理學者、商學院教授、CEO、總經理、董事長
- "產業專業類" - For: 工程師、技術專家、科技業主管、金融專業人員、醫療專業等
- "其他專門職業" - For: 建築師、技師、國考及格之專業人員

**Classification Rules:**
1. Look at their education background (學歷) - 會計系/財金系 → 會計/財務類, 法律系 → 法務類
2. Look at their professional certifications - CPA/會計師 → 會計/財務類, 律師 → 法務類
3. Look at their career path - CFO roles → 會計/財務類, Legal roles → 法務類
4. If multiple categories apply, choose the ONE that best represents their PRIMARY expertise
5. Default to "商務/管理類" if they are primarily a business executive without specific professional background

**Output:** Return ONLY the category name (e.g., "會計/財務類")

## 5. The "Professional Background" Summary (專業背景)
This is a ONE-PARAGRAPH executive summary of the person's expertise.
It should describe:
- Total years of experience
- Key industry verticals they've worked in
- Their functional expertise areas
- The type/level of roles they've held

**Format (REQUIRED):**
"約 X 年在[產業1]、[產業2]、[產業3]等領域經歷，專長於[專業領域]，長期在[公司類型]擔任[職位層級]職務。"

**Example:**
"約 30 年在產品管理、市場行銷、銷售、業務發展和資通訊與科技產業經歷，長期在跨國企業擔任高階專業經理人有豐富的經驗。"

**Another Example:**
"約 25 年在金融科技、企業軟體、數據分析等領域深耕，擅長業務拓展與策略規劃，曾於多家跨國企業擔任台灣區最高負責人。"

- MUST be a single paragraph (no bullet points)
- MUST start with "約 X 年在..."
- MUST be in 繁體中文
- Calculate years from their career start to now

## 6. The "Personal Traits" Analysis (個人特質) - MUST BE DETAILED
This describes WHO THIS PERSON IS, not their achievements.
Focus on: personality, leadership style, work habits, interpersonal skills.

**Format (STRICTLY REQUIRED):**
Return a SINGLE STRING with numbered items and sub-items using this exact format:
"1.[特質名稱]\\n- [具體事蹟或描述，說明這個特質如何展現]\\n- [為何此特質對公司/董事會有價值]\\n2.[特質名稱]\\n- [具體事蹟或描述]\\n3.[特質名稱]\\n- [具體事蹟或描述]"

**Good Example:**
"1.高度行動力與執行力\\n- 過去在台灣微軟，被形容為開會一整天依然精神飽滿，說話速度快、思路清楚，帶領團隊推進策略及檢討執行細節。\\n- 對於需要「推動轉型」與「落地執行」的公司，是非常強的執行型董事人選。\\n2.強烈的成就導向與數字導向\\n- 溝通能力強、具號召力的「女強人」形象：媒體與同事形容她總是帶著笑容、充滿能量，善於在內部與夥伴間建立信任與合作。\\n3.前瞻科技與變革思維\\n- 對於需要導入 AI、數位轉型或國際化策略的公司，她可以在董事會層級帶來戰略視野。"

**BAD Example (DO NOT DO THIS):**
"企業家精神\\n管理長才\\n客戶導向\\n創新推動者"

**Search Strategy:**
- Search for interviews, speeches, media coverage about their personality
- Look for quotes from colleagues, partners, or media describing their style
- Search: `"{name}" 領導風格 OR 管理風格 OR 個性 OR 工作態度`

**Requirements:**
- MUST have numbered items (1. 2. 3.)
- MUST have sub-items with "-" prefix after each numbered item
- MUST include specific evidence or descriptions
- MUST be 3-5 traits
- MUST be in 繁體中文
- DO NOT just list keywords
- DO NOT include "[1]" or citation markers in the output
- Return as a SINGLE STRING with \\n for line breaks

## 7. Independent Director Stats
- Search for "{name} 獨立董事" or "{name} 獨董 年資" or check 公開資訊觀測站.
- If found, return count and tenure. Otherwise, null.

# Output Format
Return **ONLY** a raw JSON object (no markdown, no extra text):
{{
  "company_industry": "String (公司產業別)",
  "chamber_of_commerce": "String (所屬商會/協會)",
  "age": "String (e.g. '54歲') or null",
  "professional_category": "String (會計/財務類 | 法務類 | 商務/管理類 | 產業專業類 | 其他專門職業)",
  "professional_background": "String (約 X 年在[領域]經歷，專長於[專業]，長期在[公司類型]擔任[職位]。)",
  "education": ["String (學校 科系 學位)", ...],
  "key_experience": ["String (公司: 職位 (成就/地區))", ...],
  "current_position": ["String (現任職位)", ...],
  "personal_traits": "String (1.特質一\\n- 具體描述\\n2.特質二\\n- 具體描述)",
  "independent_director_count": Integer or null,
  "independent_director_tenure": "String (e.g. '5年') or null",
  "email": "String or null (STRICT: 100% verified only)",
  "phone": "String or null (STRICT: 100% verified only)",
  "photo_search_term": "String (最佳圖片搜尋關鍵字)"
}}

# Example Output
{{
  "company_industry": "半導體 / 人工智慧",
  "chamber_of_commerce": "台北市美國商會",
  "age": "55歲",
  "professional_category": "產業專業類",
  "professional_background": "約 30 年在產品管理、市場行銷、銷售、業務發展和資通訊與科技產業經歷，長期在跨國企業擔任高階專業經理人有豐富的經驗。",
  "education": ["史丹佛大學 電機工程學系 碩士", "國立台灣大學 資訊工程學系 學士"],
  "key_experience": [
    "Google 台灣: 董事總經理 (帶領台灣團隊成長300%)",
    "Microsoft 亞太區: 副總裁 (負責企業解決方案)",
    "IBM 台灣: 總經理 (主導數位轉型專案)"
  ],
  "current_position": ["NVIDIA 台灣區總經理", "台灣人工智慧學校 董事"],
  "personal_traits": "1.高度行動力與執行力\\n- 過去在台灣微軟，被形容為開會一整天依然精神飽滿，說話速度快、思路清楚。\\n- 對於需要「推動轉型」與「落地執行」的公司，是非常強的執行型董事人選。\\n2.強烈的成就導向與數字導向\\n- 溝通能力強、具號召力，善於在內部與夥伴間建立信任與合作。\\n3.前瞻科技與變革思維\\n- 對於需要導入 AI、數位轉型或國際化策略的公司，她可以在董事會層級帶來戰略視野。",
  "independent_director_count": 2,
  "independent_director_tenure": "5年",
  "email": "",
  "phone": "",
  "photo_search_term": "{name} {company} headshot portrait"
}}

CRITICAL REMINDERS:
1. Age: Use the heuristic formula if direct age is not found. Show your calculation logic internally.
2. Contact: Return "" if not 100% verified. Never guess.
3. All text in Traditional Chinese (繁體中文) for the final output.
4. Return ONLY the JSON object. No markdown, no explanations."""

    return prompt


def _clean_value(value) -> str:
    """
    清理欄位值，將 null、NaN、placeholder 等無效值轉為空字串。

    這確保 Excel 中不會出現 "null"、"NaN"、"已略過" 等 placeholder 文字，
    而是顯示為空白儲存格。

    Args:
        value: 原始值（任意類型）

    Returns:
        清理後的字串，或空字串
    """
    if value is None:
        return ""

    # 處理 pandas/numpy 的 NaN
    if isinstance(value, float):
        import math
        if math.isnan(value):
            return ""

    str_value = str(value).strip()

    # 空字串直接返回
    if not str_value:
        return ""

    # 要過濾的 placeholder 值（不分大小寫）
    placeholder_values = [
        "null", "none", "nan", "n/a", "na", "undefined",
        "已略過", "待補充", "(待補充)", "（待補充）",
        "無", "無資料", "找不到", "未知", "不明",
        "暫無", "尚無", "缺", "空", "nil"
    ]

    if str_value.lower() in [p.lower() for p in placeholder_values]:
        return ""

    # 檢查是否為以特定 placeholder 開頭的文字
    skip_prefixes = ["無法", "找不到", "查無", "尚未", "暫無法"]
    for prefix in skip_prefixes:
        if str_value.startswith(prefix):
            return ""

    return str_value


def _is_valid_age(age_str: str, professional_background: str = None) -> bool:
    """
    驗證年齡是否合理。

    規則：
    1. 年齡必須在 35-85 歲之間（高階主管的合理範圍）
    2. 如果有專業背景，年齡必須與工作年資一致
       - 假設 22 歲開始工作，若有 N 年經驗，年齡至少要 22 + N

    Args:
        age_str: 年齡字串，如 "55歲" 或 "55"
        professional_background: 專業背景字串（可選，用於交叉驗證）

    Returns:
        True 如果年齡合理
    """
    if not age_str:
        return False

    # 提取數字
    age_match = re.search(r'(\d+)', str(age_str))
    if not age_match:
        return False

    age = int(age_match.group(1))

    # 1. 基本範圍檢查：高階主管通常 35-85 歲
    if age < 35 or age > 85:
        return False

    # 2. 如果有專業背景，進行交叉驗證
    if professional_background:
        # 從專業背景中提取工作年資
        years_match = re.search(r'約\s*(\d+)\s*年', professional_background)
        if years_match:
            experience_years = int(years_match.group(1))
            # 假設 22 歲大學畢業開始工作
            min_age_required = 22 + experience_years
            if age < min_age_required:
                # 年齡與經驗不符，年齡資料可能有誤
                return False

    return True


def _extract_experience_years(professional_background: str) -> int:
    """
    從專業背景中提取工作年資。

    Args:
        professional_background: 專業背景字串

    Returns:
        工作年資（整數），找不到則返回 0
    """
    if not professional_background:
        return 0

    years_match = re.search(r'約\s*(\d+)\s*年', professional_background)
    if years_match:
        return int(years_match.group(1))
    return 0


def _is_valid_education_entry(text: str) -> bool:
    """
    驗證學歷條目是否為有效格式，過濾掉垃圾資料。

    有效學歷格式範例：
    - "國立台灣大學 電機系 學士"
    - "史丹佛大學 MBA"
    - "政治大學 企業管理研究所 碩士"

    無效格式（會被過濾）：
    - "1 day ago · 22歲就讀嶺南大學..."（新聞片段）
    - 超過 100 字的長文
    - 包含明顯非學歷關鍵字

    Returns:
        True 如果是有效學歷，False 如果是垃圾資料
    """
    if not text or not isinstance(text, str):
        return False

    text = text.strip()

    # 1. 長度檢查：學歷條目通常不超過 80 字元
    if len(text) > 100:
        return False

    # 2. 排除新聞片段特徵
    garbage_patterns = [
        r'\d+\s*(day|hour|minute|second)s?\s*ago',  # "1 day ago"
        r'\d+\s*(天|小時|分鐘)前',  # "1 天前"
        r'·',  # 新聞分隔符
        r'》',  # 中文引號（常見於新聞標題）
        r'《',  # 中文引號
        r'http[s]?://',  # URL
        r'申請來港',
        r'虛報',
        r'涉',
        r'被捕',
        r'起訴',
        r'判刑',
        r'詐騙',
        r'偽造',
        r'總經理',  # 學歷欄位不應包含職稱
        r'董事長',
        r'執行長',
        r'CEO',
        r'請培養',
        r'- 星島',
        r'- 蘋果',
        r'- Yahoo',
        r'- ETtoday',
        r'- 聯合',
        r'- 中時',
        r'- 自由',
    ]

    for pattern in garbage_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return False

    # 3. 必須包含至少一個學歷關鍵字
    edu_keywords = [
        '大學', '學院', '研究所', '學系', '系',
        '學士', '碩士', '博士', '畢業',
        'University', 'College', 'Institute', 'School',
        'Bachelor', 'Master', 'MBA', 'EMBA', 'PhD', 'Doctor',
        'B.S.', 'M.S.', 'B.A.', 'M.A.', 'B.B.A.', 'M.B.A.'
    ]

    has_edu_keyword = any(kw in text for kw in edu_keywords)
    if not has_edu_keyword:
        return False

    # 4. 不應該是純數字或太短
    if len(text) < 5:
        return False

    return True


def process_api_response(api_data: dict) -> dict:
    """
    將 API 回傳的資料轉換為 Excel 欄位格式。

    包含交叉驗證：
    - 年齡必須與工作年資一致（年齡 >= 22 + 工作年資）
    - 學歷必須是結構化格式，非新聞片段

    Args:
        api_data: API 回傳的 JSON 資料

    Returns:
        對應到 Excel 欄位的字典
    """
    result = {}

    # === 先提取專業背景（用於後續交叉驗證）===
    professional_background = None
    if api_data.get("professional_background"):
        bg = api_data["professional_background"]
        if isinstance(bg, str) and bg.strip():
            professional_background = bg.strip()
            result["專業背景"] = professional_background

    # === 年齡（含交叉驗證）===
    if api_data.get("age"):
        age_str = str(api_data["age"])
        if _is_valid_age(age_str, professional_background):
            result["年齡"] = age_str
        else:
            # 年齡不合理，記錄警告但不填入
            age_num = re.search(r'(\d+)', age_str)
            if age_num:
                age_val = int(age_num.group(1))
                exp_years = _extract_experience_years(professional_background) if professional_background else 0
                if age_val < 35:
                    print(f"    ⚠ 年齡 {age_val} 歲對高階主管不合理，已略過")
                elif exp_years > 0 and age_val < 22 + exp_years:
                    print(f"    ⚠ 年齡 {age_val} 歲與 {exp_years} 年經驗不符，已略過")

    # 專業分類 - 單一字串
    if api_data.get("professional_category"):
        cat = api_data["professional_category"]
        if isinstance(cat, str) and cat.strip():
            # 驗證是否為有效分類
            valid_categories = ["會計/財務類", "法務類", "商務/管理類", "產業專業類", "其他專門職業"]
            cat_clean = cat.strip()
            if cat_clean in valid_categories:
                result["專業分類"] = cat_clean
            else:
                # 嘗試模糊匹配
                for valid_cat in valid_categories:
                    if valid_cat in cat_clean or cat_clean in valid_cat:
                        result["專業分類"] = valid_cat
                        break

    # 專業背景已在上面處理，這裡跳過
    # （保留原本的條件檢查以防重複）
    if api_data.get("professional_background") and "專業背景" not in result:
        bg = api_data["professional_background"]
        if isinstance(bg, str) and bg.strip():
            result["專業背景"] = bg.strip()

    # 學歷 - 陣列轉換為換行分隔（含驗證）
    if api_data.get("education"):
        edu = api_data["education"]
        if isinstance(edu, list):
            # 過濾掉垃圾資料
            valid_edu = []
            for item in edu:
                if isinstance(item, str) and _is_valid_education_entry(item):
                    valid_edu.append(item.strip())
            if valid_edu:
                result["學歷"] = "\n".join(valid_edu)
        elif isinstance(edu, str) and _is_valid_education_entry(edu):
            result["學歷"] = edu.strip()

    # 主要經歷 - 陣列轉換為換行分隔
    if api_data.get("key_experience"):
        exp = api_data["key_experience"]
        if isinstance(exp, list):
            result["主要經歷"] = "\n".join(exp)
        else:
            result["主要經歷"] = str(exp)

    # 現職/任 - 陣列轉換為換行分隔
    if api_data.get("current_position"):
        pos = api_data["current_position"]
        if isinstance(pos, list):
            result["現職/任"] = "\n".join(pos)
        else:
            result["現職/任"] = str(pos)

    # 個人特質 - 陣列轉換為換行分隔
    if api_data.get("personal_traits"):
        traits = api_data["personal_traits"]
        if isinstance(traits, list):
            result["個人特質"] = "\n".join(traits)
        else:
            result["個人特質"] = str(traits)

    # 獨董家數
    if api_data.get("independent_director_count") is not None:
        result["現擔任獨董家數(年)"] = str(api_data["independent_director_count"])

    # 獨董年資
    if api_data.get("independent_director_tenure"):
        result["擔任獨董年資(年)"] = str(api_data["independent_director_tenure"])

    # 電子郵件 - 嚴格驗證
    email = api_data.get("email")
    if email and isinstance(email, str) and "@" in email and email.lower() not in ["", "null", "none"]:
        # 過濾通用信箱
        generic_patterns = ["info@", "contact@", "service@", "support@", "admin@", "hello@"]
        is_generic = any(pattern in email.lower() for pattern in generic_patterns)
        if not is_generic:
            result["電子郵件"] = email

    # 公司電話 - 嚴格驗證
    phone = api_data.get("phone")
    if phone and isinstance(phone, str) and phone.lower() not in ["", "null", "none"]:
        # 驗證是否包含數字
        if re.search(r'\d{6,}', phone.replace("-", "").replace(" ", "")):
            result["公司電話"] = phone

    # 照片搜尋關鍵字 - 用於後續搜尋
    if api_data.get("photo_search_term"):
        result["_photo_search_term"] = api_data["photo_search_term"]

    # === 最終清理：移除所有 placeholder 值 ===
    # 確保 Excel 中不會出現 "null"、"NaN"、"已略過" 等 placeholder 文字
    cleaned_result = {}
    for key, value in result.items():
        # 跳過內部欄位（以 _ 開頭）
        if key.startswith("_"):
            cleaned_result[key] = value
            continue

        # 對字串值進行清理
        if isinstance(value, str):
            cleaned_value = _clean_value(value)
            if cleaned_value:  # 只保留非空值
                cleaned_result[key] = cleaned_value
        elif value is not None:
            # 非字串值保留原樣
            cleaned_result[key] = value

    return cleaned_result


def search_with_perplexity(name: str, company: str) -> dict:
    """
    使用 Perplexity API 進行 Executive Search Researcher 品質的深度搜尋。

    Returns:
        找到的資訊字典（已轉換為 Excel 欄位格式）
    """
    api_key = os.getenv("PERPLEXITY_API_KEY")

    if not api_key:
        print("    警告: PERPLEXITY_API_KEY 未設定")
        return {}

    # 使用專業的 Executive Search Researcher 提示詞
    prompt = build_executive_search_prompt(name, company)

    # Executive Search Researcher & Private Investigator 系統提示詞
    system_prompt = """You are an elite Executive Search Researcher & Private Investigator (高階獵頭與徵信專家) with 20+ years of experience in C-suite recruitment, due diligence, and investigative research.

Your expertise includes:
- Deep investigative research on senior executives
- Age inference from graduation years and historical news
- Due diligence for board appointments
- Verification of career histories and credentials
- Analysis of leadership styles and personal traits

CRITICAL RULES:
1. **Age Heuristic:** If direct age not found, CALCULATE from graduation year (grad_year + 22 = birth_year) or historical news mentions.
2. **Zero Fabrication:** NEVER guess contact info. If not 100% verified, return empty string.
3. **Contact Info Zero-Fail:** FORBIDDEN from constructing emails. Only return if explicitly found indexed online.
4. **Executive Tone:** Use concise, professional Traditional Chinese (繁體中文).
5. **Source Priority:** LinkedIn > Official Company Website > Bloomberg/Reuters > News Articles > Conference PDFs.

You do not give up easily. You dig deep, infer intelligently, and verify strictly.
Respond ONLY with valid JSON. No markdown, no explanations."""

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = requests.post(
                "https://api.perplexity.ai/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": "sonar-pro",
                    "messages": [
                        {
                            "role": "system",
                            "content": system_prompt
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    "temperature": 0.1,
                    "max_tokens": 4000
                },
                timeout=120
            )

            if response.status_code == 200:
                result = response.json()
                content = result.get("choices", [{}])[0].get("message", {}).get("content", "")

                # 清理 markdown 格式
                content = content.strip()
                if content.startswith("```json"):
                    content = content[7:]
                if content.startswith("```"):
                    content = content[3:]
                if content.endswith("```"):
                    content = content[:-3]
                content = content.strip()

                # 提取 JSON
                json_match = re.search(r'\{[\s\S]*\}', content)
                if json_match:
                    try:
                        api_data = json.loads(json_match.group())

                        # 轉換為 Excel 欄位格式
                        excel_data = process_api_response(api_data)

                        # 統計找到的欄位
                        found_fields = [k for k, v in excel_data.items() if v and not k.startswith("_")]
                        if found_fields:
                            print(f"    → 找到 {len(found_fields)} 個欄位: {', '.join(found_fields)}")

                        return excel_data

                    except json.JSONDecodeError as e:
                        print(f"    JSON 解析錯誤 ({attempt + 1}/{max_retries}): {e}")
                        # 嘗試修復常見的 JSON 錯誤
                        try:
                            # 嘗試替換單引號為雙引號
                            fixed_content = content.replace("'", '"')
                            api_data = json.loads(fixed_content)
                            return process_api_response(api_data)
                        except:
                            pass

            else:
                print(f"    Perplexity API 錯誤 ({attempt + 1}/{max_retries}): {response.status_code}")
                if response.status_code == 429:
                    print("    → API 請求過於頻繁，等待 10 秒...")
                    time.sleep(10)

        except requests.exceptions.Timeout:
            print(f"    API 請求超時 ({attempt + 1}/{max_retries})")
        except Exception as e:
            print(f"    搜尋錯誤 ({attempt + 1}/{max_retries}): {e}")

        if attempt < max_retries - 1:
            time.sleep(3)

    return {}


def multi_search_executive(name: str, company: str, missing_fields: list[str], search_client=None) -> dict:
    """
    使用多重搜尋策略獲取主管資訊。

    策略:
    1. 統一搜尋客戶端（SerpAPI + DuckDuckGo fallback）搜尋 LinkedIn 檔案
    2. 統一搜尋客戶端搜尋中文簡歷/介紹
    3. Perplexity API Executive Search Researcher 深度搜尋

    Args:
        name: 主管姓名
        company: 所屬公司
        missing_fields: 需要搜尋的欄位
        search_client: UnifiedSearchClient 實例（可選）

    Returns:
        找到的資訊字典
    """
    result = {field: "" for field in missing_fields}

    # 使用統一搜尋客戶端或 fallback 到直接 DuckDuckGo
    use_unified = search_client is not None

    # === 搜尋策略 A: LinkedIn 檔案 ===
    print(f"    [策略 A] 搜尋 LinkedIn...")
    query_linkedin = f'"{name}" "{company}" LinkedIn'

    if use_unified:
        search_results_linkedin = search_client.search(query_linkedin, num_results=5)
    else:
        search_results_linkedin = search_with_ddg(query_linkedin, max_results=5)

    if search_results_linkedin:
        linkedin_info = extract_info_from_snippets(search_results_linkedin, name)
        for key, value in linkedin_info.items():
            if key in result and not result[key]:
                result[key] = value

        # 提取 LinkedIn URL
        linkedin_url = extract_linkedin_url(search_results_linkedin)
        if linkedin_url:
            print(f"    → 找到 LinkedIn: {linkedin_url[:60]}...")

    time.sleep(1)  # 避免請求過快

    # === 搜尋策略 B: 中文簡歷/介紹 ===
    print(f"    [策略 B] 搜尋中文資料...")
    query_bio = f'"{name}" "{company}" 簡歷 OR 介紹 OR 經歷 OR 學歷'

    if use_unified:
        search_results_bio = search_client.search(query_bio, num_results=5)
    else:
        search_results_bio = search_with_ddg(query_bio, max_results=5)

    if search_results_bio:
        bio_info = extract_info_from_snippets(search_results_bio, name)
        for key, value in bio_info.items():
            if key in result and not result[key]:
                result[key] = value

    time.sleep(1)

    # === 搜尋策略 C: Perplexity API Executive Search Researcher 深度搜尋 ===
    # 只對還沒找到的欄位進行深度搜尋
    still_missing = [f for f in missing_fields if not result.get(f)]

    if still_missing:
        print(f"    [策略 C] Perplexity Executive Search Researcher...")
        print(f"    → 搜尋欄位: {', '.join(still_missing)}")
        perplexity_result = search_with_perplexity(name, company)

        for key, value in perplexity_result.items():
            # 跳過內部欄位
            if key.startswith("_"):
                continue
            if key in result and not result[key] and value:
                result[key] = value

    time.sleep(1)

    # === 搜尋策略 D: Python 端照片搜尋 (增強版 + 審核模式) ===
    # 繞過 LLM 限制，使用多重策略搜尋照片
    photo_result = {"best_url": "", "best_score": 0, "status": "待補充", "candidates": []}

    if "照片" in missing_fields:
        print(f"    [策略 D] Python 端照片搜尋（增強版 + 審核模式）...")

        # 取得已找到的職稱（用於提高搜尋準確度）
        job_title = result.get("現職/任", "")

        if use_unified:
            # 使用統一搜尋客戶端（簡化版，不含完整評分）
            search_queries = [
                f'site:linkedin.com "{name}" {company}',
                f'"{name}" {company} portrait OR headshot'
            ]
            if job_title:
                first_title = job_title.split('\n')[0].strip()
                if first_title:
                    search_queries.insert(1, f'"{name}" "{first_title}" photo')

            all_urls = []
            for query in search_queries:
                print(f"    [照片搜尋] 搜尋: {query[:50]}...")
                urls = search_client.search_images(query, num_results=3)
                all_urls.extend(urls)
                if len(all_urls) >= 10:
                    break

            # 建立候選清單
            for i, url in enumerate(all_urls[:5]):
                score = 50 if 'linkedin' in url.lower() else 20
                photo_result["candidates"].append({
                    "url": url,
                    "score": score,
                    "source": "",
                    "title": "",
                    "width": 0,
                    "height": 0
                })

            # 選擇最佳照片
            if all_urls:
                # 優先 LinkedIn
                for url in all_urls:
                    if 'linkedin' in url.lower():
                        photo_result["best_url"] = url
                        photo_result["best_score"] = 50
                        photo_result["status"] = "待確認"
                        print(f"    [照片搜尋] ✓ 選擇 LinkedIn 來源圖片（待確認）")
                        break

                if not photo_result["best_url"]:
                    photo_result["best_url"] = all_urls[0]
                    photo_result["best_score"] = 20
                    photo_result["status"] = "待確認"
                    print(f"    [照片搜尋] 選擇第一張候選圖片（待確認）")
        else:
            # 使用增強版 DuckDuckGo 搜尋（含完整評分系統）
            photo_result = find_executive_photo_python(name, company, job_title)

        # 填入照片 URL（根據信心度門檻）
        if photo_result["best_url"]:
            result["照片"] = photo_result["best_url"]

        # 填入照片狀態
        result["照片狀態"] = photo_result["status"]

    # 將照片候選資訊附加到結果中（供後續儲存到 JSON）
    result["_photo_candidates"] = photo_result

    return result


def generate_photo_review_html(photo_data: dict):
    """
    生成照片審核 HTML 報告。

    Args:
        photo_data: 照片候選資料 {excel_row: {name, company, candidates, ...}}
    """
    html_content = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>照片審核報告 - CEO Project</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Microsoft JhengHei", sans-serif;
            background: #f5f5f5;
            margin: 0;
            padding: 20px;
        }
        .container { max-width: 1200px; margin: 0 auto; }
        h1 {
            color: #333;
            border-bottom: 3px solid #007bff;
            padding-bottom: 10px;
        }
        .instructions {
            background: #e7f3ff;
            border: 1px solid #b3d9ff;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .instructions h3 { margin-top: 0; color: #0056b3; }
        .instructions ol { margin-bottom: 0; }
        .person-card {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            overflow: hidden;
        }
        .person-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .person-header h2 { margin: 0; font-size: 1.3em; }
        .status-badge {
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: bold;
        }
        .status-pending { background: #ffc107; color: #333; }
        .status-confirm { background: #28a745; color: white; }
        .status-selected { background: #17a2b8; color: white; }
        .candidates-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
            gap: 15px;
            padding: 20px;
        }
        .candidate {
            border: 3px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
            cursor: pointer;
            transition: all 0.2s;
            position: relative;
        }
        .candidate:hover { border-color: #007bff; transform: translateY(-2px); }
        .candidate.selected { border-color: #28a745; background: #e8f5e9; }
        .candidate.selected::after {
            content: "✓";
            position: absolute;
            top: 10px;
            right: 10px;
            background: #28a745;
            color: white;
            width: 30px;
            height: 30px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
        }
        .candidate img {
            width: 100%;
            height: 180px;
            object-fit: cover;
            display: block;
        }
        .candidate-info {
            padding: 10px;
            font-size: 0.85em;
        }
        .candidate-score {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 10px;
            font-weight: bold;
            font-size: 0.8em;
        }
        .score-high { background: #d4edda; color: #155724; }
        .score-medium { background: #fff3cd; color: #856404; }
        .score-low { background: #f8d7da; color: #721c24; }
        .no-select {
            border: 3px dashed #ccc;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 220px;
            color: #666;
            cursor: pointer;
        }
        .no-select:hover { border-color: #999; background: #f9f9f9; }
        .no-select.selected { border-color: #dc3545; background: #fff5f5; }
        .actions {
            padding: 15px 20px;
            background: #f8f9fa;
            border-top: 1px solid #eee;
        }
        .url-input {
            width: 100%;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 0.9em;
        }
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
            margin-right: 10px;
        }
        .btn-primary { background: #007bff; color: white; }
        .btn-primary:hover { background: #0056b3; }
        .btn-success { background: #28a745; color: white; }
        .btn-success:hover { background: #1e7e34; }
        .save-section {
            position: sticky;
            bottom: 0;
            background: white;
            padding: 20px;
            box-shadow: 0 -2px 10px rgba(0,0,0,0.1);
            text-align: center;
        }
        .img-error {
            background: #f8f9fa;
            display: flex;
            align-items: center;
            justify-content: center;
            height: 180px;
            color: #999;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📸 照片審核報告</h1>

        <div class="instructions">
            <h3>使用說明</h3>
            <ol>
                <li>點擊正確的照片選擇它（綠框 + 勾勾表示已選擇）</li>
                <li>如果所有照片都不對，點擊「都不正確」</li>
                <li>可以在下方輸入框手動貼上正確的照片 URL</li>
                <li>完成後點擊「儲存選擇」按鈕，會下載一個 JSON 檔案</li>
                <li>將下載的 JSON 檔案放到 <code>output/data/</code> 資料夾中覆蓋原檔案</li>
            </ol>
        </div>

        <div id="persons-container">
"""

    # 生成每個人的卡片
    for row_str, data in sorted(photo_data.items(), key=lambda x: int(x[0])):
        name = data.get("name", "未知")
        company = data.get("company", "")
        status = data.get("status", "待補充")
        best_url = data.get("best_url", "")
        candidates = data.get("candidates", [])

        status_class = "status-confirm" if status == "待確認" else "status-pending"

        html_content += f"""
        <div class="person-card" data-row="{row_str}">
            <div class="person-header">
                <h2>[列 {row_str}] {name} - {company}</h2>
                <span class="status-badge {status_class}">{status}</span>
            </div>
            <div class="candidates-grid">
"""

        # 候選照片
        for i, candidate in enumerate(candidates):
            url = candidate.get("url", "")
            score = candidate.get("score", 0)
            source = candidate.get("source", "")

            # 分數顏色
            if score >= 40:
                score_class = "score-high"
            elif score >= 20:
                score_class = "score-medium"
            else:
                score_class = "score-low"

            # 預選最高分的（如果是 best_url）
            selected_class = "selected" if url == best_url and best_url else ""

            html_content += f"""
                <div class="candidate {selected_class}" data-url="{url}" onclick="selectCandidate(this, '{row_str}')">
                    <img src="{url}" alt="候選照片 {i+1}" onerror="this.parentElement.innerHTML='<div class=img-error>圖片載入失敗</div>'">
                    <div class="candidate-info">
                        <span class="candidate-score {score_class}">分數: {score}</span>
                        <div style="margin-top:5px;color:#666;font-size:0.8em;word-break:break-all;">
                            {source[:50] + '...' if len(source) > 50 else source}
                        </div>
                    </div>
                </div>
"""

        # "都不正確" 選項
        html_content += f"""
                <div class="no-select" data-url="" onclick="selectCandidate(this, '{row_str}')">
                    <span>❌ 都不正確</span>
                </div>
            </div>
            <div class="actions">
                <label>手動輸入照片 URL：</label>
                <input type="text" class="url-input" id="url-{row_str}" placeholder="貼上正確的照片 URL..." onchange="updateManualUrl('{row_str}', this.value)">
            </div>
        </div>
"""

    html_content += """
        </div>

        <div class="save-section">
            <button class="btn btn-success" onclick="saveSelections()">💾 儲存選擇（下載 JSON）</button>
            <button class="btn btn-primary" onclick="copyToClipboard()">📋 複製到剪貼簿</button>
        </div>
    </div>

    <script>
        // 儲存所有選擇
        let selections = {};

        // 初始化選擇（使用預設的 best_url）
        document.querySelectorAll('.person-card').forEach(card => {
            const row = card.dataset.row;
            const selected = card.querySelector('.candidate.selected');
            if (selected) {
                selections[row] = selected.dataset.url;
            }
        });

        function selectCandidate(element, row) {
            // 移除同組的其他選擇
            const card = element.closest('.person-card');
            card.querySelectorAll('.candidate, .no-select').forEach(c => c.classList.remove('selected'));

            // 選擇當前
            element.classList.add('selected');

            // 儲存選擇
            selections[row] = element.dataset.url;

            // 更新狀態標籤
            const badge = card.querySelector('.status-badge');
            if (element.dataset.url) {
                badge.textContent = '已選擇';
                badge.className = 'status-badge status-selected';
            } else {
                badge.textContent = '待補充';
                badge.className = 'status-badge status-pending';
            }

            // 清空手動輸入框（如果選擇了候選照片）
            if (element.dataset.url) {
                document.getElementById('url-' + row).value = '';
            }
        }

        function updateManualUrl(row, url) {
            if (url.trim()) {
                // 取消其他選擇
                const card = document.querySelector(`.person-card[data-row="${row}"]`);
                card.querySelectorAll('.candidate, .no-select').forEach(c => c.classList.remove('selected'));

                // 儲存手動輸入的 URL
                selections[row] = url.trim();

                // 更新狀態
                const badge = card.querySelector('.status-badge');
                badge.textContent = '已手動輸入';
                badge.className = 'status-badge status-selected';
            }
        }

        function saveSelections() {
            // 讀取當前的照片候選 JSON 結構
            const output = {};

            document.querySelectorAll('.person-card').forEach(card => {
                const row = card.dataset.row;
                const selectedUrl = selections[row] || '';
                const manualUrl = document.getElementById('url-' + row).value.trim();
                const finalUrl = manualUrl || selectedUrl;

                output[row] = {
                    selected_url: finalUrl,
                    status: finalUrl ? '已確認' : '待補充'
                };
            });

            // 下載 JSON
            const blob = new Blob([JSON.stringify(output, null, 2)], {type: 'application/json'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'photo_selections.json';
            a.click();
            URL.revokeObjectURL(url);

            alert('已下載 photo_selections.json\\n\\n請執行以下命令套用選擇：\\npython src/apply_photo_selections.py');
        }

        function copyToClipboard() {
            const output = {};
            document.querySelectorAll('.person-card').forEach(card => {
                const row = card.dataset.row;
                const selectedUrl = selections[row] || '';
                const manualUrl = document.getElementById('url-' + row).value.trim();
                output[row] = manualUrl || selectedUrl;
            });

            navigator.clipboard.writeText(JSON.stringify(output, null, 2))
                .then(() => alert('已複製到剪貼簿！'))
                .catch(err => alert('複製失敗: ' + err));
        }
    </script>
</body>
</html>
"""

    # 儲存 HTML 檔案
    html_path = Path(PHOTO_REVIEW_HTML)
    html_path.parent.mkdir(parents=True, exist_ok=True)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)


def search_photos_only(rows_str: str):
    """
    僅搜尋照片模式 - 跳過 Perplexity API，只做照片搜尋和審核。

    Args:
        rows_str: 要處理的 Excel 列號字串
    """
    print("=" * 60)
    print("照片搜尋程序啟動 (Photos Only Mode)")
    print("=" * 60)

    # 1. 解析列號
    target_rows = parse_row_numbers(rows_str)
    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    print(f"\n目標 Excel 列號: {target_rows}")
    print(f"共 {len(target_rows)} 列待處理")

    # 2. 讀取 Excel 檔案
    try:
        # 優先讀取擴充後的檔案
        if Path(EXCEL_OUTPUT).exists():
            df = pd.read_excel(EXCEL_OUTPUT)
            print(f"\n讀取 '{EXCEL_OUTPUT}'")
        else:
            df = pd.read_excel(EXCEL_INPUT)
            print(f"\n讀取 '{EXCEL_INPUT}'")

        # 確保照片相關欄位存在
        if "照片" not in df.columns:
            df["照片"] = None
        if "照片狀態" not in df.columns:
            df["照片狀態"] = None

        df["照片"] = df["照片"].astype(object)
        df["照片狀態"] = df["照片狀態"].astype(object)

    except FileNotFoundError:
        print(f"錯誤: 找不到 Excel 檔案")
        sys.exit(1)
    except Exception as e:
        print(f"錯誤: 讀取 Excel 失敗 - {e}")
        sys.exit(1)

    # 3. 驗證列號範圍
    max_excel_row = len(df) + 1
    invalid_rows = [r for r in target_rows if r > max_excel_row or r < 2]
    if invalid_rows:
        print(f"警告: 以下列號超出範圍: {invalid_rows}")
        target_rows = [r for r in target_rows if r <= max_excel_row and r >= 2]

    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    # 4. 讀取現有照片候選資料
    existing_photo_candidates = {}
    if Path(PHOTO_CANDIDATES_JSON).exists():
        try:
            with open(PHOTO_CANDIDATES_JSON, 'r', encoding='utf-8') as f:
                existing_photo_candidates = json.load(f)
        except:
            pass

    # 5. 處理每一列 - 只搜尋照片
    all_photo_candidates = {}
    updated_count = 0

    for excel_row in target_rows:
        pandas_idx = excel_row_to_pandas_index(excel_row)
        row_data = df.iloc[pandas_idx]

        name = row_data.get("姓名（中英）", "")
        company = row_data.get("所屬公司", "")
        job_title = row_data.get("現職/任", "")

        if pd.isna(name) or not name:
            print(f"\n[列 {excel_row}] 跳過 - 無姓名資料")
            continue

        print(f"\n[列 {excel_row}] 搜尋照片: {name} ({company})")
        print("-" * 50)

        # 使用增強版照片搜尋
        photo_result = find_executive_photo_python(name, company, str(job_title) if pd.notna(job_title) else "")

        # 儲存候選資料
        if photo_result.get("candidates"):
            all_photo_candidates[excel_row] = {
                "name": name,
                "company": company,
                "best_url": photo_result.get("best_url", ""),
                "best_score": photo_result.get("best_score", 0),
                "status": photo_result.get("status", "待補充"),
                "candidates": photo_result.get("candidates", [])
            }

            # 更新 DataFrame
            if photo_result["best_url"]:
                df.at[pandas_idx, "照片"] = photo_result["best_url"]
                updated_count += 1
            df.at[pandas_idx, "照片狀態"] = photo_result["status"]

        # 避免請求過快（增加間隔減少被限制）
        time.sleep(2)

    # 6. 儲存更新後的 Excel
    output_path = Path(EXCEL_OUTPUT)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 在儲存前清理整個 DataFrame 中的 placeholder 值
    for col in ENRICHABLE_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: _clean_value(x) if pd.notna(x) else "")

    try:
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n{'=' * 60}")
        print(f"照片搜尋完成！")
        print(f"  - 處理列數: {len(target_rows)}")
        print(f"  - 找到照片: {updated_count} 筆")
        print(f"  - 輸出檔案: {output_path}")
    except PermissionError:
        # 檔案被鎖定，嘗試儲存到備用檔案
        backup_path = output_path.with_name("Standard_Example_Enriched_backup.xlsx")
        try:
            df.to_excel(backup_path, index=False, engine='openpyxl')
            print(f"\n{'=' * 60}")
            print(f"照片搜尋完成！")
            print(f"  - 處理列數: {len(target_rows)}")
            print(f"  - 找到照片: {updated_count} 筆")
            print(f"\n  ⚠️  原檔案被鎖定，已儲存到備用檔案:")
            print(f"      {backup_path}")
            print(f"\n  請關閉 Excel 後，手動將備用檔案改名為:")
            print(f"      {output_path.name}")
        except Exception as e2:
            print(f"錯誤: 儲存 Excel 失敗 - {e2}")
            print("請關閉 Excel 後重新執行")
            sys.exit(1)
    except Exception as e:
        print(f"錯誤: 儲存 Excel 失敗 - {e}")
        sys.exit(1)

    # 7. 合併並儲存照片候選 JSON
    if all_photo_candidates:
        for row, data in all_photo_candidates.items():
            existing_photo_candidates[str(row)] = data

        json_path = Path(PHOTO_CANDIDATES_JSON)
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(existing_photo_candidates, f, ensure_ascii=False, indent=2)
            print(f"\n照片候選資料已儲存: {json_path}")
        except Exception as e:
            print(f"警告: 儲存照片候選 JSON 失敗 - {e}")

        # 8. 生成照片審核 HTML 報告
        try:
            generate_photo_review_html(existing_photo_candidates)
            print(f"照片審核報告已生成: {PHOTO_REVIEW_HTML}")
        except Exception as e:
            print(f"警告: 生成照片審核報告失敗 - {e}")

        # 統計照片狀態
        confirmed_count = sum(1 for d in all_photo_candidates.values() if d.get("status") == "待確認")
        pending_count = sum(1 for d in all_photo_candidates.values() if d.get("status") == "待補充")

        print(f"\n照片審核狀態:")
        print(f"  - 待確認（高信心度）: {confirmed_count} 筆")
        print(f"  - 待補充（需人工選擇）: {pending_count} 筆")

    print(f"\n{'=' * 60}")
    print(f"請開啟 {PHOTO_REVIEW_HTML} 審核照片")
    print(f"{'=' * 60}")


def enrich_data(rows_str: str, photos_only: bool = False):
    """主要資料擴充函式。

    Args:
        rows_str: 要處理的 Excel 列號字串
        photos_only: 如果為 True，則只搜尋照片，跳過 Perplexity API
    """
    # 如果是 photos_only 模式，使用專用函式
    if photos_only:
        search_photos_only(rows_str)
        return

    print("=" * 60)
    print("資料擴充程序啟動 (Executive Search Researcher Quality)")
    print("=" * 60)

    # 初始化統一搜尋客戶端
    search_client = None
    if UNIFIED_SEARCH_AVAILABLE:
        search_client = UnifiedSearchClient()
        status = search_client.get_status()
        print("\n搜尋引擎狀態:")
        print(f"  主要引擎: {status['primary_engine']}")
        if status['serpapi']['available']:
            print(f"  SerpAPI: {status['serpapi']['remaining']}/{status['serpapi']['quota']} 次額度剩餘 ({status['serpapi']['month']})")
        else:
            print("  SerpAPI: 未設定")
        print(f"  DuckDuckGo: {'可用' if status['duckduckgo']['available'] else '未安裝'}")
    else:
        print("\n搜尋引擎狀態:")
        print("  使用: DuckDuckGo (直接模式)")

    print("\n核心原則:")
    print("  1. Zero Fabrication - 寧缺勿錯")
    print("  2. Contact Info 100% Accuracy - 嚴格驗證")
    print("  3. Executive Tone - 專業簡報品質")

    # 1. 解析列號
    target_rows = parse_row_numbers(rows_str)
    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    print(f"\n目標 Excel 列號: {target_rows}")
    print(f"共 {len(target_rows)} 列待處理")

    # 2. 讀取 Excel 檔案（同步原始檔案結構 + 保留擴充資料）
    try:
        # 一定要先讀取原始檔案（取得最新欄位結構）
        df_original = pd.read_excel(EXCEL_INPUT)
        print(f"\n讀取原始檔案 '{EXCEL_INPUT}'")
        print(f"原始結構: {len(df_original)} 列 x {len(df_original.columns)} 欄")

        # 如果已有擴充檔案，則合併資料
        if Path(EXCEL_OUTPUT).exists():
            df_enriched = pd.read_excel(EXCEL_OUTPUT)
            print(f"讀取已擴充檔案 '{EXCEL_OUTPUT}'")

            # 以原始檔案為基礎，合併擴充資料
            df = df_original.copy()

            # 將擴充檔案中有的欄位資料複製過來
            for col in df_enriched.columns:
                if col in df.columns:
                    # 對於可擴充欄位，優先使用擴充檔案的資料
                    if col in ENRICHABLE_COLUMNS:
                        for idx in range(min(len(df), len(df_enriched))):
                            enriched_val = df_enriched.iloc[idx].get(col)
                            if pd.notna(enriched_val) and enriched_val != "" and enriched_val != 0:
                                df.at[idx, col] = enriched_val
                elif col not in df.columns:
                    # 新增擴充檔案特有的欄位（如 照片狀態、專業分類）
                    df[col] = df_enriched[col] if len(df) == len(df_enriched) else None

            # 確保新增的欄位也存在
            for col in ["照片狀態", "專業分類"]:
                if col not in df.columns:
                    df[col] = None
                if col in df_enriched.columns:
                    for idx in range(min(len(df), len(df_enriched))):
                        val = df_enriched.iloc[idx].get(col)
                        if pd.notna(val) and val != "":
                            df.at[idx, col] = val

            print(f"已合併資料（保留原始結構 + 擴充資料）")
        else:
            df = df_original.copy()
            # 確保新增欄位存在
            for col in ["照片狀態", "專業分類"]:
                if col not in df.columns:
                    df[col] = None

        print(f"最終結構: {len(df)} 列 x {len(df.columns)} 欄")

        # 將需要擴充的欄位轉換為 object 類型，避免寫入字串時出錯
        for col in ENRICHABLE_COLUMNS + ["照片狀態", "專業分類"]:
            if col in df.columns:
                df[col] = df[col].astype(object)

    except FileNotFoundError:
        print(f"錯誤: 找不到檔案 '{EXCEL_INPUT}'")
        sys.exit(1)
    except Exception as e:
        print(f"錯誤: 讀取 Excel 失敗 - {e}")
        sys.exit(1)

    # 3. 驗證列號範圍
    max_excel_row = len(df) + 1
    invalid_rows = [r for r in target_rows if r > max_excel_row or r < 2]
    if invalid_rows:
        print(f"警告: 以下列號超出範圍: {invalid_rows}")
        target_rows = [r for r in target_rows if r <= max_excel_row and r >= 2]

    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    # 4. 處理每一列
    updated_count = 0
    total_fields = 0
    all_photo_candidates = {}  # 收集所有照片候選資料 {excel_row: {...}}

    for excel_row in target_rows:
        pandas_idx = excel_row_to_pandas_index(excel_row)
        row_data = df.iloc[pandas_idx]

        name = row_data.get("姓名（中英）", "")
        company = row_data.get("所屬公司", "")

        if pd.isna(name) or not name:
            print(f"\n[列 {excel_row}] 跳過 - 無姓名資料")
            continue

        print(f"\n[列 {excel_row}] 處理中: {name} ({company})")
        print("-" * 50)

        # 找出空缺欄位
        missing_fields = []
        for col in ENRICHABLE_COLUMNS:
            if col in df.columns:
                val = row_data.get(col)
                if pd.isna(val) or val == "" or val == 0:
                    missing_fields.append(col)

        if not missing_fields:
            print(f"  → 所有欄位已有資料，跳過")
            continue

        print(f"  空缺欄位 ({len(missing_fields)}): {', '.join(missing_fields)}")
        total_fields += len(missing_fields)

        # 多重搜尋（傳入搜尋客戶端）
        found_data = multi_search_executive(name, company, missing_fields, search_client)

        # 提取照片候選資料（如果有）
        photo_candidates_info = found_data.pop("_photo_candidates", None)
        if photo_candidates_info and photo_candidates_info.get("candidates"):
            all_photo_candidates[excel_row] = {
                "name": name,
                "company": company,
                "best_url": photo_candidates_info.get("best_url", ""),
                "best_score": photo_candidates_info.get("best_score", 0),
                "status": photo_candidates_info.get("status", "待補充"),
                "candidates": photo_candidates_info.get("candidates", [])
            }

        # 更新 DataFrame
        fields_filled = 0
        for field, value in found_data.items():
            # 跳過內部欄位
            if field.startswith("_"):
                continue
            if field in df.columns and value:
                # 對於 missing_fields 中的欄位，或是新增的 "照片狀態" 欄位
                if field in missing_fields or field == "照片狀態":
                    df.at[pandas_idx, field] = value

                    # 顯示值（處理換行符號）
                    display_value = str(value).replace('\n', ' | ')
                    if len(display_value) > 60:
                        display_value = display_value[:60] + "..."
                    print(f"  ✓ [{field}]: {display_value}")

                    if field in missing_fields:
                        updated_count += 1
                        fields_filled += 1

        print(f"\n  → 本列填入 {fields_filled}/{len(missing_fields)} 個欄位")

        # 避免 API 限制
        time.sleep(2)

    # 5. 儲存結果
    output_path = Path(EXCEL_OUTPUT)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 在儲存前清理整個 DataFrame 中的 placeholder 值
    for col in ENRICHABLE_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: _clean_value(x) if pd.notna(x) else "")

    try:
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n{'=' * 60}")
        print(f"擴充完成！")
        print(f"  - 處理列數: {len(target_rows)}")
        print(f"  - 總空缺欄位: {total_fields}")
        print(f"  - 成功填入欄位: {updated_count}")
        print(f"  - 填入率: {updated_count/total_fields*100:.1f}%" if total_fields > 0 else "  - 填入率: N/A")
        print(f"  - 輸出檔案: {output_path}")
        print(f"{'=' * 60}")
    except Exception as e:
        print(f"錯誤: 儲存 Excel 失敗 - {e}")
        sys.exit(1)

    # 6. 儲存照片候選資料到 JSON
    if all_photo_candidates:
        json_path = Path(PHOTO_CANDIDATES_JSON)
        try:
            # 讀取既有的 JSON（如果存在）
            existing_data = {}
            if json_path.exists():
                with open(json_path, 'r', encoding='utf-8') as f:
                    existing_data = json.load(f)

            # 合併新資料（使用 excel_row 作為 key）
            for row, data in all_photo_candidates.items():
                existing_data[str(row)] = data

            # 儲存更新後的 JSON
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(existing_data, f, ensure_ascii=False, indent=2)

            print(f"\n照片候選資料已儲存: {json_path}")
        except Exception as e:
            print(f"警告: 儲存照片候選 JSON 失敗 - {e}")

        # 7. 生成照片審核 HTML 報告
        try:
            generate_photo_review_html(existing_data if existing_data else all_photo_candidates)
            print(f"照片審核報告已生成: {PHOTO_REVIEW_HTML}")
        except Exception as e:
            print(f"警告: 生成照片審核報告失敗 - {e}")

        # 統計照片狀態
        confirmed_count = sum(1 for d in all_photo_candidates.values() if d.get("status") == "待確認")
        pending_count = sum(1 for d in all_photo_candidates.values() if d.get("status") == "待補充")

        print(f"\n照片審核狀態:")
        print(f"  - 待確認（高信心度）: {confirmed_count} 筆")
        print(f"  - 待補充（需人工選擇）: {pending_count} 筆")
        if confirmed_count + pending_count > 0:
            print(f"\n請開啟 {PHOTO_REVIEW_HTML} 審核照片後再生成 PPT")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="資料擴充腳本 (Executive Search Researcher Quality) - 高品質主管資料搜尋",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
    python src/enrich_data.py --rows "2"
    python src/enrich_data.py --rows "2, 5, 10"
    python src/enrich_data.py --rows "2-10"

核心原則:
    1. Zero Fabrication - 找不到資料時回傳空值，絕不捏造
    2. Contact Info 100% Accuracy - 電子郵件與電話必須 100% 確認來源
    3. Executive Tone - 繁體中文專業簡報品質

搜尋策略:
    A. DuckDuckGo 搜尋 LinkedIn 檔案
    B. DuckDuckGo 搜尋中文簡歷/介紹
    C. Perplexity API Executive Search Researcher 深度搜尋

輸出格式:
    - 學歷、主要經歷、現職、個人特質 使用換行符號分隔
    - 每個項目包含具體年份、成就等詳細資訊
        """
    )
    parser.add_argument(
        "--rows",
        type=str,
        required=True,
        help="要處理的 Excel 列號"
    )
    parser.add_argument(
        "--photos-only",
        action="store_true",
        help="僅搜尋照片（跳過 Perplexity API 資料擴充）"
    )

    args = parser.parse_args()
    enrich_data(args.rows, photos_only=args.photos_only)
