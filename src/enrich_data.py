"""
enrich_data.py - 資料擴充腳本 (Executive Search & Private Investigator Quality)

根據 Standard Example.xlsx 中的姓名與公司資訊，
使用多重搜尋策略填補空缺欄位。

升級重點 (Phase 31 - 修復版):
1. 修復 Excel 讀取 dtype 問題（URL 被誤判為 float64）
2. 改進 DuckDuckGo 搜尋錯誤處理
3. 增加網路連線測試

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
    from duckduckgo_search import DDGS
    DDGS_AVAILABLE = True
except ImportError:
    try:
        from ddgs import DDGS
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


def test_network_connection() -> bool:
    """
    測試網路連線是否正常。
    
    Returns:
        True 如果網路正常
    """
    test_urls = [
        "https://www.google.com",
        "https://duckduckgo.com",
        "https://www.bing.com"
    ]
    
    for url in test_urls:
        try:
            response = requests.get(url, timeout=5)
            if response.status_code == 200:
                return True
        except:
            continue
    
    return False


def read_excel_safe(filepath: str) -> pd.DataFrame:
    """
    安全讀取 Excel 檔案，避免 dtype 問題。
    
    Args:
        filepath: Excel 檔案路徑
        
    Returns:
        DataFrame
    """
    # 定義所有可能包含字串的欄位為 str 類型
    dtype_spec = {
        "姓名（中英）": str,
        "所屬公司": str,
        "年齡": str,
        "照片": str,
        "照片狀態": str,
        "專業分類": str,
        "專業背景": str,
        "學歷": str,
        "主要經歷": str,
        "現職/任": str,
        "個人特質": str,
        "現擔任獨董家數(年)": str,
        "擔任獨董年資(年)": str,
        "電子郵件": str,
        "公司電話": str
    }
    
    try:
        # 先嘗試用指定 dtype 讀取
        df = pd.read_excel(filepath, dtype=dtype_spec)
        return df
    except Exception as e1:
        print(f"    注意: 使用指定 dtype 讀取失敗，嘗試自動偵測...")
        try:
            # 退回到自動偵測，但之後轉換欄位類型
            df = pd.read_excel(filepath)
            
            # 將所有 ENRICHABLE_COLUMNS 轉為 object 類型
            for col in df.columns:
                if col in ENRICHABLE_COLUMNS or col in ["姓名（中英）", "所屬公司"]:
                    df[col] = df[col].astype(object)
            
            return df
        except Exception as e2:
            raise Exception(f"讀取 Excel 失敗: {e2}")


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
    使用 DuckDuckGo 進行網路搜尋（增強錯誤處理版）。

    Returns:
        搜尋結果列表，每個結果包含 title, href, body
    """
    if not DDGS_AVAILABLE:
        return []

    max_retries = 3
    for attempt in range(max_retries):
        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=max_results, region='tw-tzh'))
                return results
        except Exception as e:
            error_msg = str(e).lower()
            
            # 判斷錯誤類型
            if 'ratelimit' in error_msg or 'rate' in error_msg:
                print(f"    DuckDuckGo 請求過於頻繁，等待 {(attempt + 1) * 5} 秒...")
                time.sleep((attempt + 1) * 5)
            elif 'timeout' in error_msg:
                print(f"    DuckDuckGo 連線逾時，重試中 ({attempt + 1}/{max_retries})...")
                time.sleep(2)
            elif 'no results' in error_msg:
                # 沒有結果不是錯誤，直接返回空列表
                return []
            else:
                print(f"    DuckDuckGo 搜尋錯誤 ({attempt + 1}/{max_retries}): {e}")
                time.sleep(2)
    
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
    """
    score = 0
    image_url = result.get('image', '').lower()
    source_url = result.get('url', '').lower()
    title = result.get('title', '').lower()
    width = result.get('width', 0)
    height = result.get('height', 0)

    # === 來源評分 ===
    if 'linkedin.com' in source_url or 'linkedin' in image_url:
        score += 50

    company_domain_hints = ['company', 'corporate', 'about', 'team', 'leadership', 'management']
    if any(hint in source_url for hint in company_domain_hints):
        score += 40

    news_sites = ['reuters', 'bloomberg', 'forbes', 'businessweek', 'cna.com', 'udn.com',
                  'ltn.com', 'chinatimes', 'ettoday', 'setn.com', 'bnext', 'technews']
    if any(site in source_url for site in news_sites):
        score += 20

    # === 圖片尺寸評分 ===
    if width > 0 and height > 0:
        if width >= 150 and height >= 150:
            score += 15
        aspect_ratio = width / height if height > 0 else 0
        if 0.6 <= aspect_ratio <= 1.2:
            score += 10
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

    if 'default' in image_url and ('profile' in image_url or 'avatar' in image_url):
        score -= 100

    return score


def validate_image_url(url: str) -> bool:
    """驗證圖片 URL 是否有效。"""
    if not url:
        return False

    lower_url = url.lower()
    valid_extensions = ['.jpg', '.jpeg', '.png', '.webp', '.gif']
    has_valid_ext = any(ext in lower_url for ext in valid_extensions)

    if not has_valid_ext:
        try:
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
            response = requests.head(url, headers=headers, timeout=5, allow_redirects=True)
            content_type = response.headers.get('Content-Type', '')
            if not content_type.startswith('image/'):
                return False
        except:
            pass

    return True


def find_executive_photo_python(name: str, company: str, job_title: str = "") -> dict:
    """
    使用多重策略搜尋高階主管照片 URL（增強錯誤處理版）。
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

    # 先測試網路連線
    if not test_network_connection():
        print("    [照片搜尋] 網路連線異常，跳過")
        return result

    # === 建立多重搜尋查詢 ===
    search_queries = []
    search_queries.append(f'site:linkedin.com "{name}" {company}')
    
    if job_title:
        first_title = job_title.split('\n')[0].strip()
        if first_title:
            search_queries.append(f'"{name}" "{first_title}" photo OR portrait')

    search_queries.append(f'"{name}" {company} 照片 OR headshot OR portrait')
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

                    if not results:
                        print(f"    [策略 {i}] 搜尋失敗: No results found.")
                        continue

                    for img in results:
                        image_url = img.get('image', '')
                        if image_url and image_url not in seen_urls:
                            seen_urls.add(image_url)
                            if validate_image_url(image_url):
                                all_results.append(img)

                    time.sleep(1)  # 增加間隔避免 rate limit

                except Exception as e:
                    error_msg = str(e).lower()
                    if 'no results' in error_msg or 'empty' in error_msg:
                        print(f"    [策略 {i}] 搜尋失敗: No results found.")
                    elif 'ratelimit' in error_msg or 'rate' in error_msg:
                        print(f"    [策略 {i}] 搜尋失敗: Rate limited, 等待中...")
                        time.sleep(5)
                    else:
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
        if score > -50:
            scored_results.append((score, img_result))

    if not scored_results:
        print("    [照片搜尋] 所有候選圖片評分過低")
        return result

    scored_results.sort(key=lambda x: x[0], reverse=True)

    print(f"    [照片搜尋] 候選圖片評分:")
    for i, (score, img_result) in enumerate(scored_results[:3], 1):
        url = img_result.get('image', '')[:50]
        source = img_result.get('url', '')[:30]
        print(f"      #{i} 分數:{score:3d} | {url}... | 來源:{source}...")

    for score, img_result in scored_results[:5]:
        result["candidates"].append({
            "url": img_result.get('image', ''),
            "score": score,
            "source": img_result.get('url', ''),
            "title": img_result.get('title', ''),
            "width": img_result.get('width', 0),
            "height": img_result.get('height', 0)
        })

    best_score, best_img = scored_results[0]
    best_url = best_img.get('image', '')
    result["best_score"] = best_score

    if best_score >= PHOTO_CONFIDENCE_THRESHOLD:
        result["best_url"] = best_url
        result["status"] = "待確認"
        print(f"    [照片搜尋] ✓ 分數 {best_score} >= {PHOTO_CONFIDENCE_THRESHOLD}，自動填入（待確認）")
    else:
        result["best_url"] = ""
        result["status"] = "待補充"
        print(f"    [照片搜尋] ✗ 分數 {best_score} < {PHOTO_CONFIDENCE_THRESHOLD}，需人工審核")

    return result


def extract_info_from_snippets(results: list[dict], name: str) -> dict:
    """從搜尋結果的摘要中提取資訊。"""
    extracted = {}
    linkedin = extract_linkedin_url(results)
    if linkedin:
        extracted['LinkedIn'] = linkedin
    return extracted


def build_executive_search_prompt(name: str, company: str) -> str:
    """建立 Executive Search Researcher & Private Investigator 品質的搜尋提示詞。"""
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

## 5. The "Professional Background" Summary (專業背景)
This is a ONE-PARAGRAPH executive summary of the person's expertise.
**Format (REQUIRED):**
"約 X 年在[產業1]、[產業2]、[產業3]等領域經歷，專長於[專業領域]，長期在[公司類型]擔任[職位層級]職務。"

## 6. The "Personal Traits" Analysis (個人特質) - MUST BE DETAILED
**Format (STRICTLY REQUIRED):**
Return a SINGLE STRING with numbered items:
"1.[特質名稱]\\n- [具體事蹟或描述]\\n2.[特質名稱]\\n- [具體事蹟或描述]\\n3.[特質名稱]\\n- [具體事蹟或描述]"

## 7. Independent Director Stats
- Search for "{name} 獨立董事" or "{name} 獨董 年資".
- If found, return count and tenure. Otherwise, null.

# Output Format
Return **ONLY** a raw JSON object (no markdown, no extra text):
{{
  "company_industry": "String (公司產業別)",
  "chamber_of_commerce": "String (所屬商會/協會)",
  "age": "String (e.g. '54歲') or null",
  "professional_category": "String (會計/財務類 | 法務類 | 商務/管理類 | 產業專業類 | 其他專門職業)",
  "professional_background": "String (約 X 年在[領域]經歷...)",
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

CRITICAL REMINDERS:
1. Age: Use the heuristic formula if direct age is not found.
2. Contact: Return "" if not 100% verified. Never guess.
3. All text in Traditional Chinese (繁體中文) for the final output.
4. Return ONLY the JSON object. No markdown, no explanations."""

    return prompt


def _clean_value(value) -> str:
    """清理欄位值，將 null、NaN、placeholder 等無效值轉為空字串。"""
    if value is None:
        return ""

    if isinstance(value, float):
        import math
        if math.isnan(value):
            return ""

    str_value = str(value).strip()

    if not str_value:
        return ""

    placeholder_values = [
        "null", "none", "nan", "n/a", "na", "undefined",
        "已略過", "待補充", "(待補充)", "（待補充）",
        "無", "無資料", "找不到", "未知", "不明",
        "暫無", "尚無", "缺", "空", "nil"
    ]

    if str_value.lower() in [p.lower() for p in placeholder_values]:
        return ""

    skip_prefixes = ["無法", "找不到", "查無", "尚未", "暫無法"]
    for prefix in skip_prefixes:
        if str_value.startswith(prefix):
            return ""

    return str_value


def _is_valid_age(age_str: str, professional_background: str = None) -> bool:
    """驗證年齡是否合理。"""
    if not age_str:
        return False

    age_match = re.search(r'(\d+)', str(age_str))
    if not age_match:
        return False

    age = int(age_match.group(1))

    if age < 35 or age > 85:
        return False

    if professional_background:
        years_match = re.search(r'約\s*(\d+)\s*年', professional_background)
        if years_match:
            experience_years = int(years_match.group(1))
            min_age_required = 22 + experience_years
            if age < min_age_required:
                return False

    return True


def _extract_experience_years(professional_background: str) -> int:
    """從專業背景中提取工作年資。"""
    if not professional_background:
        return 0

    years_match = re.search(r'約\s*(\d+)\s*年', professional_background)
    if years_match:
        return int(years_match.group(1))
    return 0


def _is_valid_education_entry(text: str) -> bool:
    """驗證學歷條目是否為有效格式。"""
    if not text or not isinstance(text, str):
        return False

    text = text.strip()

    if len(text) > 100:
        return False

    garbage_patterns = [
        r'\d+\s*(day|hour|minute|second)s?\s*ago',
        r'\d+\s*(天|小時|分鐘)前',
        r'·',
        r'》',
        r'《',
        r'http[s]?://',
        r'總經理',
        r'董事長',
        r'執行長',
        r'CEO',
    ]

    for pattern in garbage_patterns:
        if re.search(pattern, text, re.IGNORECASE):
            return False

    edu_keywords = [
        '大學', '學院', '研究所', '學系', '系',
        '學士', '碩士', '博士', '畢業',
        'University', 'College', 'Institute', 'School',
        'Bachelor', 'Master', 'MBA', 'EMBA', 'PhD', 'Doctor',
    ]

    has_edu_keyword = any(kw in text for kw in edu_keywords)
    if not has_edu_keyword:
        return False

    if len(text) < 5:
        return False

    return True


def process_api_response(api_data: dict) -> dict:
    """將 API 回傳的資料轉換為 Excel 欄位格式。"""
    result = {}

    professional_background = None
    if api_data.get("professional_background"):
        bg = api_data["professional_background"]
        if isinstance(bg, str) and bg.strip():
            professional_background = bg.strip()
            result["專業背景"] = professional_background

    if api_data.get("age"):
        age_str = str(api_data["age"])
        if _is_valid_age(age_str, professional_background):
            result["年齡"] = age_str

    if api_data.get("professional_category"):
        cat = api_data["professional_category"]
        if isinstance(cat, str) and cat.strip():
            valid_categories = ["會計/財務類", "法務類", "商務/管理類", "產業專業類", "其他專門職業"]
            cat_clean = cat.strip()
            if cat_clean in valid_categories:
                result["專業分類"] = cat_clean
            else:
                for valid_cat in valid_categories:
                    if valid_cat in cat_clean or cat_clean in valid_cat:
                        result["專業分類"] = valid_cat
                        break

    if api_data.get("education"):
        edu = api_data["education"]
        if isinstance(edu, list):
            valid_edu = []
            for item in edu:
                if isinstance(item, str) and _is_valid_education_entry(item):
                    valid_edu.append(item.strip())
            if valid_edu:
                result["學歷"] = "\n".join(valid_edu)
        elif isinstance(edu, str) and _is_valid_education_entry(edu):
            result["學歷"] = edu.strip()

    if api_data.get("key_experience"):
        exp = api_data["key_experience"]
        if isinstance(exp, list):
            result["主要經歷"] = "\n".join(exp)
        else:
            result["主要經歷"] = str(exp)

    if api_data.get("current_position"):
        pos = api_data["current_position"]
        if isinstance(pos, list):
            result["現職/任"] = "\n".join(pos)
        else:
            result["現職/任"] = str(pos)

    if api_data.get("personal_traits"):
        traits = api_data["personal_traits"]
        if isinstance(traits, list):
            result["個人特質"] = "\n".join(traits)
        else:
            result["個人特質"] = str(traits)

    if api_data.get("independent_director_count") is not None:
        result["現擔任獨董家數(年)"] = str(api_data["independent_director_count"])

    if api_data.get("independent_director_tenure"):
        result["擔任獨董年資(年)"] = str(api_data["independent_director_tenure"])

    email = api_data.get("email")
    if email and isinstance(email, str) and "@" in email and email.lower() not in ["", "null", "none"]:
        generic_patterns = ["info@", "contact@", "service@", "support@", "admin@", "hello@"]
        is_generic = any(pattern in email.lower() for pattern in generic_patterns)
        if not is_generic:
            result["電子郵件"] = email

    phone = api_data.get("phone")
    if phone and isinstance(phone, str) and phone.lower() not in ["", "null", "none"]:
        if re.search(r'\d{6,}', phone.replace("-", "").replace(" ", "")):
            result["公司電話"] = phone

    if api_data.get("photo_search_term"):
        result["_photo_search_term"] = api_data["photo_search_term"]

    cleaned_result = {}
    for key, value in result.items():
        if key.startswith("_"):
            cleaned_result[key] = value
            continue

        if isinstance(value, str):
            cleaned_value = _clean_value(value)
            if cleaned_value:
                cleaned_result[key] = cleaned_value
        elif value is not None:
            cleaned_result[key] = value

    return cleaned_result


def search_with_perplexity(name: str, company: str) -> dict:
    """使用 Perplexity API 進行深度搜尋。"""
    api_key = os.getenv("PERPLEXITY_API_KEY")

    if not api_key:
        print("    警告: PERPLEXITY_API_KEY 未設定")
        return {}

    prompt = build_executive_search_prompt(name, company)

    system_prompt = """You are an elite Executive Search Researcher & Private Investigator.
CRITICAL RULES:
1. Age Heuristic: Calculate from graduation year if not found directly.
2. Zero Fabrication: NEVER guess contact info.
3. Executive Tone: Use Traditional Chinese (繁體中文).
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
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.1,
                    "max_tokens": 4000
                },
                timeout=120
            )

            if response.status_code == 200:
                result = response.json()
                content = result.get("choices", [{}])[0].get("message", {}).get("content", "")

                content = content.strip()
                if content.startswith("```json"):
                    content = content[7:]
                if content.startswith("```"):
                    content = content[3:]
                if content.endswith("```"):
                    content = content[:-3]
                content = content.strip()

                json_match = re.search(r'\{[\s\S]*\}', content)
                if json_match:
                    try:
                        api_data = json.loads(json_match.group())
                        excel_data = process_api_response(api_data)
                        found_fields = [k for k, v in excel_data.items() if v and not k.startswith("_")]
                        if found_fields:
                            print(f"    → 找到 {len(found_fields)} 個欄位: {', '.join(found_fields)}")
                        return excel_data

                    except json.JSONDecodeError as e:
                        print(f"    JSON 解析錯誤 ({attempt + 1}/{max_retries}): {e}")

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
    """使用多重搜尋策略獲取主管資訊。"""
    result = {field: "" for field in missing_fields}

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

        linkedin_url = extract_linkedin_url(search_results_linkedin)
        if linkedin_url:
            print(f"    → 找到 LinkedIn: {linkedin_url[:60]}...")

    time.sleep(1)

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

    # === 搜尋策略 C: Perplexity API ===
    still_missing = [f for f in missing_fields if not result.get(f)]

    if still_missing:
        print(f"    [策略 C] Perplexity Executive Search Researcher...")
        print(f"    → 搜尋欄位: {', '.join(still_missing)}")
        perplexity_result = search_with_perplexity(name, company)

        for key, value in perplexity_result.items():
            if key.startswith("_"):
                continue
            if key in result and not result[key] and value:
                result[key] = value

    time.sleep(1)

    # === 搜尋策略 D: Python 端照片搜尋 ===
    photo_result = {"best_url": "", "best_score": 0, "status": "待補充", "candidates": []}

    if "照片" in missing_fields:
        print(f"    [策略 D] Python 端照片搜尋...")
        job_title = result.get("現職/任", "")
        photo_result = find_executive_photo_python(name, company, str(job_title) if job_title else "")

        if photo_result["best_url"]:
            result["照片"] = photo_result["best_url"]
        result["照片狀態"] = photo_result["status"]

    result["_photo_candidates"] = photo_result

    return result


def generate_photo_review_html(photo_data: dict):
    """生成照片審核 HTML 報告。"""
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
        h1 { color: #333; border-bottom: 3px solid #007bff; padding-bottom: 10px; }
        .instructions {
            background: #e7f3ff;
            border: 1px solid #b3d9ff;
            border-radius: 8px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .instructions h3 { margin-top: 0; color: #0056b3; }
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
        .candidate-info { padding: 10px; font-size: 0.85em; }
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
        }
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-weight: bold;
            margin-right: 10px;
        }
        .btn-success { background: #28a745; color: white; }
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
                <li>點擊正確的照片選擇它</li>
                <li>如果所有照片都不對，點擊「都不正確」</li>
                <li>完成後點擊「儲存選擇」按鈕</li>
                <li>將下載的 JSON 檔案放到 output/data/ 資料夾</li>
            </ol>
        </div>
        <div id="persons-container">
"""

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

        for i, candidate in enumerate(candidates):
            url = candidate.get("url", "")
            score = candidate.get("score", 0)
            source = candidate.get("source", "")

            if score >= 40:
                score_class = "score-high"
            elif score >= 20:
                score_class = "score-medium"
            else:
                score_class = "score-low"

            selected_class = "selected" if url == best_url and best_url else ""

            html_content += f"""
                <div class="candidate {selected_class}" data-url="{url}" onclick="selectCandidate(this, '{row_str}')">
                    <img src="{url}" alt="候選照片 {i+1}" onerror="this.parentElement.innerHTML='<div class=img-error>圖片載入失敗</div>'">
                    <div class="candidate-info">
                        <span class="candidate-score {score_class}">分數: {score}</span>
                    </div>
                </div>
"""

        html_content += f"""
                <div class="no-select" data-url="" onclick="selectCandidate(this, '{row_str}')">
                    <span>❌ 都不正確</span>
                </div>
            </div>
            <div class="actions">
                <label>手動輸入照片 URL：</label>
                <input type="text" class="url-input" id="url-{row_str}" placeholder="貼上正確的照片 URL...">
            </div>
        </div>
"""

    html_content += """
        </div>
        <div class="save-section">
            <button class="btn btn-success" onclick="saveSelections()">💾 儲存選擇</button>
        </div>
    </div>
    <script>
        let selections = {};
        document.querySelectorAll('.person-card').forEach(card => {
            const row = card.dataset.row;
            const selected = card.querySelector('.candidate.selected');
            if (selected) selections[row] = selected.dataset.url;
        });

        function selectCandidate(element, row) {
            const card = element.closest('.person-card');
            card.querySelectorAll('.candidate, .no-select').forEach(c => c.classList.remove('selected'));
            element.classList.add('selected');
            selections[row] = element.dataset.url;
        }

        function saveSelections() {
            const output = {};
            document.querySelectorAll('.person-card').forEach(card => {
                const row = card.dataset.row;
                const manualUrl = document.getElementById('url-' + row).value.trim();
                output[row] = {
                    selected_url: manualUrl || selections[row] || '',
                    status: (manualUrl || selections[row]) ? '已確認' : '待補充'
                };
            });
            const blob = new Blob([JSON.stringify(output, null, 2)], {type: 'application/json'});
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'photo_selections.json';
            a.click();
        }
    </script>
</body>
</html>
"""

    html_path = Path(PHOTO_REVIEW_HTML)
    html_path.parent.mkdir(parents=True, exist_ok=True)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)


def search_photos_only(rows_str: str):
    """僅搜尋照片模式。"""
    print("=" * 60)
    print("照片搜尋程序啟動 (Photos Only Mode)")
    print("=" * 60)

    # 測試網路連線
    print("\n檢查網路連線...")
    if not test_network_connection():
        print("錯誤: 網路連線異常，請檢查網路設定")
        sys.exit(1)
    print("網路連線正常")

    target_rows = parse_row_numbers(rows_str)
    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    print(f"\n目標 Excel 列號: {target_rows}")
    print(f"共 {len(target_rows)} 列待處理")

    try:
        if Path(EXCEL_OUTPUT).exists():
            df = read_excel_safe(EXCEL_OUTPUT)
            print(f"\n讀取 '{EXCEL_OUTPUT}'")
        else:
            df = read_excel_safe(EXCEL_INPUT)
            print(f"\n讀取 '{EXCEL_INPUT}'")

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

    max_excel_row = len(df) + 1
    invalid_rows = [r for r in target_rows if r > max_excel_row or r < 2]
    if invalid_rows:
        print(f"警告: 以下列號超出範圍: {invalid_rows}")
        target_rows = [r for r in target_rows if r <= max_excel_row and r >= 2]

    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    existing_photo_candidates = {}
    if Path(PHOTO_CANDIDATES_JSON).exists():
        try:
            with open(PHOTO_CANDIDATES_JSON, 'r', encoding='utf-8') as f:
                existing_photo_candidates = json.load(f)
        except:
            pass

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

        photo_result = find_executive_photo_python(name, company, str(job_title) if pd.notna(job_title) else "")

        if photo_result.get("candidates"):
            all_photo_candidates[excel_row] = {
                "name": name,
                "company": company,
                "best_url": photo_result.get("best_url", ""),
                "best_score": photo_result.get("best_score", 0),
                "status": photo_result.get("status", "待補充"),
                "candidates": photo_result.get("candidates", [])
            }

            if photo_result["best_url"]:
                df.at[pandas_idx, "照片"] = photo_result["best_url"]
                updated_count += 1
            df.at[pandas_idx, "照片狀態"] = photo_result["status"]

        time.sleep(2)

    output_path = Path(EXCEL_OUTPUT)
    output_path.parent.mkdir(parents=True, exist_ok=True)

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
        backup_path = output_path.with_name("Standard_Example_Enriched_backup.xlsx")
        try:
            df.to_excel(backup_path, index=False, engine='openpyxl')
            print(f"\n⚠️  原檔案被鎖定，已儲存到: {backup_path}")
        except Exception as e2:
            print(f"錯誤: 儲存 Excel 失敗 - {e2}")
            sys.exit(1)
    except Exception as e:
        print(f"錯誤: 儲存 Excel 失敗 - {e}")
        sys.exit(1)

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

        try:
            generate_photo_review_html(existing_photo_candidates)
            print(f"照片審核報告已生成: {PHOTO_REVIEW_HTML}")
        except Exception as e:
            print(f"警告: 生成照片審核報告失敗 - {e}")

    print(f"\n{'=' * 60}")


def enrich_data(rows_str: str, photos_only: bool = False):
    """主要資料擴充函式。"""
    if photos_only:
        search_photos_only(rows_str)
        return

    print("=" * 60)
    print("資料擴充程序啟動 (Executive Search Researcher Quality)")
    print("=" * 60)

    # 測試網路連線
    print("\n檢查網路連線...")
    if not test_network_connection():
        print("警告: 網路連線可能有問題，繼續執行但可能會失敗...")
    else:
        print("網路連線正常")

    search_client = None
    if UNIFIED_SEARCH_AVAILABLE:
        search_client = UnifiedSearchClient()
        status = search_client.get_status()
        print("\n搜尋引擎狀態:")
        print(f"  主要引擎: {status['primary_engine']}")
    else:
        print("\n搜尋引擎狀態:")
        print("  使用: DuckDuckGo (直接模式)")

    target_rows = parse_row_numbers(rows_str)
    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    print(f"\n目標 Excel 列號: {target_rows}")
    print(f"共 {len(target_rows)} 列待處理")

    try:
        if Path(EXCEL_OUTPUT).exists():
            df = read_excel_safe(EXCEL_OUTPUT)
            print(f"\n讀取 '{EXCEL_OUTPUT}'")
        else:
            df = read_excel_safe(EXCEL_INPUT)
            print(f"\n讀取 '{EXCEL_INPUT}'")

        print(f"資料結構: {len(df)} 列 x {len(df.columns)} 欄")

        for col in ENRICHABLE_COLUMNS + ["照片狀態", "專業分類"]:
            if col not in df.columns:
                df[col] = None
            df[col] = df[col].astype(object)

    except FileNotFoundError:
        print(f"錯誤: 找不到 Excel 檔案")
        sys.exit(1)
    except Exception as e:
        print(f"錯誤: 讀取 Excel 失敗 - {e}")
        sys.exit(1)

    max_excel_row = len(df) + 1
    invalid_rows = [r for r in target_rows if r > max_excel_row or r < 2]
    if invalid_rows:
        print(f"警告: 以下列號超出範圍: {invalid_rows}")
        target_rows = [r for r in target_rows if r <= max_excel_row and r >= 2]

    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    updated_count = 0
    total_fields = 0
    all_photo_candidates = {}

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

        found_data = multi_search_executive(name, company, missing_fields, search_client)

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

        fields_filled = 0
        for field, value in found_data.items():
            if field.startswith("_"):
                continue
            if field in df.columns and value:
                if field in missing_fields or field == "照片狀態":
                    df.at[pandas_idx, field] = value

                    display_value = str(value).replace('\n', ' | ')
                    if len(display_value) > 60:
                        display_value = display_value[:60] + "..."
                    print(f"  ✓ [{field}]: {display_value}")

                    if field in missing_fields:
                        updated_count += 1
                        fields_filled += 1

        print(f"\n  → 本列填入 {fields_filled}/{len(missing_fields)} 個欄位")
        time.sleep(2)

    output_path = Path(EXCEL_OUTPUT)
    output_path.parent.mkdir(parents=True, exist_ok=True)

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

    if all_photo_candidates:
        json_path = Path(PHOTO_CANDIDATES_JSON)
        try:
            existing_data = {}
            if json_path.exists():
                with open(json_path, 'r', encoding='utf-8') as f:
                    existing_data = json.load(f)

            for row, data in all_photo_candidates.items():
                existing_data[str(row)] = data

            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(existing_data, f, ensure_ascii=False, indent=2)

            print(f"\n照片候選資料已儲存: {json_path}")
        except Exception as e:
            print(f"警告: 儲存照片候選 JSON 失敗 - {e}")

        try:
            generate_photo_review_html(existing_data if existing_data else all_photo_candidates)
            print(f"照片審核報告已生成: {PHOTO_REVIEW_HTML}")
        except Exception as e:
            print(f"警告: 生成照片審核報告失敗 - {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="資料擴充腳本 (Executive Search Researcher Quality)",
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
        help="僅搜尋照片"
    )

    args = parser.parse_args()
    enrich_data(args.rows, photos_only=args.photos_only)
