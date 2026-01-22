# -*- coding: utf-8 -*-
"""
enrich_cell.py - 針對特定儲存格進行資料收集

此腳本允許使用者指定特定欄位和列號，針對性地收集資料。
支援 Excel 欄位代號 (如 H26) 或中文欄位名稱 (如 學歷:26)。

使用方式:
    python src/enrich_cell.py --field "學歷" --rows "26"
    python src/enrich_cell.py --field "H" --rows "26-30"
    python src/enrich_cell.py --cell "H26"
    python src/enrich_cell.py --cell "H26,I27,J26-J30"
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

# 載入環境變數
load_dotenv()

# === 常數定義 ===
EXCEL_INPUT = "Standard Example.xlsx"
EXCEL_OUTPUT = "output/data/Standard_Example_Enriched.xlsx"

# Excel 欄位代號對應到中文欄位名稱
# 注意：這個對應可能需要根據實際 Excel 結構調整
COLUMN_LETTER_TO_NAME = {
    "A": "姓名（中英）",
    "B": "所屬公司",
    "C": "年齡",
    "D": "照片",
    "E": "照片狀態",
    "F": "專業分類",
    "G": "專業背景",
    "H": "學歷",
    "I": "主要經歷",
    "J": "現職/任",
    "K": "個人特質",
    "L": "現擔任獨董家數(年)",
    "M": "擔任獨董年資(年)",
    "N": "電子郵件",
    "O": "公司電話"
}

# 中文欄位名稱對應到 Excel 欄位代號（反向對應）
COLUMN_NAME_TO_LETTER = {v: k for k, v in COLUMN_LETTER_TO_NAME.items()}

# 可搜尋的欄位（排除姓名、公司、照片狀態等）
SEARCHABLE_FIELDS = [
    "年齡", "專業分類", "專業背景", "學歷", "主要經歷",
    "現職/任", "個人特質", "現擔任獨董家數(年)", "擔任獨董年資(年)",
    "電子郵件", "公司電話", "照片"
]

# 欄位編號對應
FIELD_NUMBER_TO_NAME = {
    "1": "年齡",
    "2": "專業分類",
    "3": "專業背景",
    "4": "學歷",
    "5": "主要經歷",
    "6": "現職/任",
    "7": "個人特質",
    "8": "現擔任獨董家數(年)",
    "9": "擔任獨董年資(年)",
    "10": "電子郵件",
    "11": "公司電話",
    "12": "照片"
}


def parse_row_numbers(rows_str: str) -> list[int]:
    """解析列號字串，轉換為 Excel 列號列表。"""
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
            try:
                result.add(int(part))
            except ValueError:
                pass

    result.discard(1)  # 排除標題列
    return sorted(result)


def parse_cell_references(cell_str: str) -> list[tuple[str, int]]:
    """
    解析儲存格參照，支援多種格式。

    支援格式:
    - "H26" -> [("學歷", 26)]
    - "H26,I27" -> [("學歷", 26), ("主要經歷", 27)]
    - "H26-H30" -> [("學歷", 26), ("學歷", 27), ...]
    - "學歷:26" -> [("學歷", 26)]
    - "學歷:26-30" -> [("學歷", 26), ("學歷", 27), ...]

    Returns:
        list of (field_name, row_number) tuples
    """
    result = []
    cell_str = cell_str.replace('，', ',').replace('：', ':')
    parts = cell_str.replace(" ", "").split(",")

    for part in parts:
        if not part:
            continue

        # 格式 1: 中文欄位名:列號 (如 "學歷:26" 或 "學歷:26-30")
        if ":" in part:
            field_part, row_part = part.split(":", 1)
            field_name = resolve_field_name(field_part)
            if field_name:
                rows = parse_row_numbers(row_part)
                for row in rows:
                    result.append((field_name, row))
            continue

        # 格式 2: Excel 欄位代號 + 列號 (如 "H26" 或 "H26-H30")
        # 檢查是否為範圍格式 (H26-H30)
        range_match = re.match(r"^([A-Z]+)(\d+)-([A-Z]+)(\d+)$", part, re.IGNORECASE)
        if range_match:
            col1, row1, col2, row2 = range_match.groups()
            col1, col2 = col1.upper(), col2.upper()
            row1, row2 = int(row1), int(row2)

            if col1 == col2:  # 同一欄位，不同列
                field_name = COLUMN_LETTER_TO_NAME.get(col1)
                if field_name:
                    for row in range(min(row1, row2), max(row1, row2) + 1):
                        if row >= 2:  # 排除標題列
                            result.append((field_name, row))
            continue

        # 格式 3: 單一儲存格 (如 "H26")
        cell_match = re.match(r"^([A-Z]+)(\d+)$", part, re.IGNORECASE)
        if cell_match:
            col, row = cell_match.groups()
            col = col.upper()
            row = int(row)
            field_name = COLUMN_LETTER_TO_NAME.get(col)
            if field_name and row >= 2:
                result.append((field_name, row))

    return result


def resolve_field_name(field_input: str) -> str:
    """
    解析欄位輸入，返回正確的欄位名稱。

    支援:
    - 編號: "1", "2", ...
    - Excel 欄位代號: "H", "I", ...
    - 中文名稱: "學歷", "年齡", ...
    - 部分匹配: "獨董家數", "經歷", ...
    """
    field_input = field_input.strip()

    # 1. 檢查是否為編號
    if field_input in FIELD_NUMBER_TO_NAME:
        return FIELD_NUMBER_TO_NAME[field_input]

    # 2. 檢查是否為 Excel 欄位代號
    if field_input.upper() in COLUMN_LETTER_TO_NAME:
        return COLUMN_LETTER_TO_NAME[field_input.upper()]

    # 3. 檢查是否為完整欄位名稱
    if field_input in SEARCHABLE_FIELDS:
        return field_input

    # 4. 部分匹配
    for field in SEARCHABLE_FIELDS:
        if field_input in field or field in field_input:
            return field

    return None


def build_focused_search_prompt(name: str, company: str, target_field: str) -> str:
    """
    建立針對特定欄位的搜尋提示詞。
    """
    import datetime
    current_year = datetime.datetime.now().year

    # 根據目標欄位建立專屬的搜尋指令
    field_instructions = {
        "年齡": f"""
Search specifically for the person's AGE or BIRTH YEAR.

Strategy:
1. Search for graduation year -> Add 22 to estimate birth year -> Calculate age
2. Search for "born in YYYY" or "出生於"
3. Search for news mentioning age (e.g., "45歲的{name}")
4. Current year is {current_year}

Output: Return age as "XX歲" format (e.g., "55歲") or null if not found.
""",
        "學歷": f"""
Search specifically for EDUCATION background.

Strategy:
1. Search: "{name}" "{company}" 學歷 OR 畢業 OR alumni
2. Search: "{name}" LinkedIn education
3. Look for: University name, Department/Major, Degree level

Output format:
- Each degree on separate line
- Format: "學校名稱 科系 學位" (e.g., "國立台灣大學 電機系 學士")
- Return as array of strings
""",
        "專業背景": f"""
Search specifically for PROFESSIONAL BACKGROUND summary.

Output format (REQUIRED):
"約 X 年在[產業1]、[產業2]等領域經歷，專長於[專業領域]，長期在[公司類型]擔任[職位層級]職務。"

Must be a single paragraph in Traditional Chinese.
""",
        "主要經歷": f"""
Search specifically for KEY CAREER EXPERIENCE.

Strategy:
1. Search for past positions and companies
2. Look for notable achievements at each role
3. Include years/duration if available

Output format:
- Each position on separate line
- Format: "公司名稱: 職位 (成就/年份)"
- Return as array of strings
""",
        "現職/任": f"""
Search specifically for CURRENT POSITIONS.

Look for:
- Current job title and company
- Board positions
- Advisory roles
- Other concurrent positions

Output: Array of current positions.
""",
        "個人特質": f"""
Search specifically for PERSONAL TRAITS and leadership style.

Strategy:
1. Search for interviews and speeches
2. Look for media descriptions of personality
3. Find quotes from colleagues

Output format:
"1.[特質名稱]\\n- [具體描述]\\n2.[特質名稱]\\n- [具體描述]\\n3.[特質名稱]\\n- [具體描述]"

Must include 3-5 traits with specific examples.
""",
        "專業分類": """
Classify into ONE of these categories based on PRIMARY expertise:
- "會計/財務類" - 會計師、財務長、CFO
- "法務類" - 律師、法官、法務長
- "商務/管理類" - CEO、總經理、董事長
- "產業專業類" - 工程師、技術專家
- "其他專門職業" - 建築師、技師

Output: Return ONLY the category name.
""",
        "現擔任獨董家數(年)": """
Search for number of INDEPENDENT DIRECTOR positions currently held.

Search: "{name}" 獨立董事 OR 獨董

Output: Integer number or null.
""",
        "擔任獨董年資(年)": """
Search for total years of INDEPENDENT DIRECTOR experience.

Output: "X年" format or null.
""",
        "電子郵件": """
Search for VERIFIED EMAIL address.

CRITICAL: Only return if 100% verified from official source.
DO NOT guess or construct emails.
Ignore generic emails like info@, contact@, service@.

Output: Verified email string or null.
""",
        "公司電話": """
Search for VERIFIED PHONE number.

CRITICAL: Only return if 100% verified from official source.
Must be a direct/office phone, not general company switchboard.

Output: Verified phone string or null.
""",
        "照片": f"""
Suggest the best search query to find a professional photo.

Output: A search query string like "{name} {company} headshot portrait"
"""
    }

    instruction = field_instructions.get(target_field, f"Search for information about {target_field}.")

    prompt = f"""# Target Executive
Name: {name}
Company: {company}

# Task
Find ONLY the following information: {target_field}

# Instructions
{instruction}

# Output Format
Return ONLY a JSON object:
{{
  "{target_field}": <value or null>
}}

CRITICAL:
- Return null if information cannot be verified
- Use Traditional Chinese (繁體中文)
- No markdown, no explanations, ONLY the JSON object
"""
    return prompt


def search_field_with_perplexity(name: str, company: str, target_field: str) -> str:
    """
    使用 Perplexity API 針對特定欄位進行搜尋。
    """
    api_key = os.getenv("PERPLEXITY_API_KEY")

    if not api_key:
        print("    警告: PERPLEXITY_API_KEY 未設定")
        return None

    prompt = build_focused_search_prompt(name, company, target_field)

    system_prompt = """You are an elite Executive Search Researcher specializing in finding specific information about executives.
Your task is to find ONLY the requested information with 100% accuracy.
If you cannot verify the information, return null.
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
                    "max_tokens": 2000
                },
                timeout=60
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
                        data = json.loads(json_match.group())
                        value = data.get(target_field)

                        # 處理陣列類型的值
                        if isinstance(value, list):
                            value = "\n".join(str(v) for v in value if v)

                        # 處理 null 值
                        if value is None or value == "null" or value == "":
                            return None

                        return str(value)

                    except json.JSONDecodeError as e:
                        print(f"    JSON 解析錯誤 ({attempt + 1}/{max_retries}): {e}")

            elif response.status_code == 429:
                print(f"    API 請求過於頻繁，等待 10 秒...")
                time.sleep(10)
            else:
                print(f"    Perplexity API 錯誤: {response.status_code}")

        except requests.exceptions.Timeout:
            print(f"    API 請求超時 ({attempt + 1}/{max_retries})")
        except Exception as e:
            print(f"    搜尋錯誤 ({attempt + 1}/{max_retries}): {e}")

        if attempt < max_retries - 1:
            time.sleep(3)

    return None


def search_photo_with_ddg(name: str, company: str) -> dict:
    """使用 DuckDuckGo 搜尋照片。"""
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            print("    警告: ddgs 未安裝，無法搜尋照片")
            return {"best_url": "", "status": "待補充", "candidates": []}

    search_queries = [
        f'site:linkedin.com "{name}" {company}',
        f'"{name}" {company} portrait OR headshot',
        f'{name} {company} profile photo'
    ]

    all_results = []
    seen_urls = set()

    try:
        with DDGS() as ddgs:
            for query in search_queries:
                try:
                    results = list(ddgs.images(query, max_results=3))
                    for img in results:
                        url = img.get('image', '')
                        if url and url not in seen_urls:
                            seen_urls.add(url)
                            score = 50 if 'linkedin' in url.lower() else 20
                            all_results.append({
                                "url": url,
                                "score": score,
                                "source": img.get('url', '')
                            })
                    time.sleep(0.5)
                except Exception as e:
                    continue
    except Exception as e:
        print(f"    DuckDuckGo 搜尋錯誤: {e}")

    if all_results:
        all_results.sort(key=lambda x: x['score'], reverse=True)
        best = all_results[0]
        return {
            "best_url": best['url'],
            "status": "待確認" if best['score'] >= 30 else "待補充",
            "candidates": all_results[:5]
        }

    return {"best_url": "", "status": "待補充", "candidates": []}


def clean_value(value) -> str:
    """清理欄位值，將 placeholder 轉為空字串。"""
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
        "無", "無資料", "找不到", "未知", "不明"
    ]

    if str_value.lower() in [p.lower() for p in placeholder_values]:
        return ""

    return str_value


def enrich_cells(cells: list[tuple[str, int]], force: bool = False):
    """
    針對特定儲存格進行資料收集。

    Args:
        cells: list of (field_name, row_number) tuples
        force: 是否強制更新已有資料的欄位
    """
    print("=" * 60)
    print("針對特定儲存格資料收集")
    print("=" * 60)

    if not cells:
        print("錯誤: 沒有有效的目標儲存格")
        sys.exit(1)

    # 顯示目標儲存格
    print(f"\n目標儲存格 ({len(cells)} 個):")
    for field, row in cells[:10]:  # 只顯示前 10 個
        col_letter = COLUMN_NAME_TO_LETTER.get(field, "?")
        print(f"  - {col_letter}{row} ({field})")
    if len(cells) > 10:
        print(f"  ... 還有 {len(cells) - 10} 個")

    # 讀取 Excel 檔案
    try:
        if Path(EXCEL_OUTPUT).exists():
            df = pd.read_excel(EXCEL_OUTPUT)
            print(f"\n讀取 '{EXCEL_OUTPUT}'")
        else:
            df = pd.read_excel(EXCEL_INPUT)
            print(f"\n讀取 '{EXCEL_INPUT}'")
    except FileNotFoundError:
        print(f"錯誤: 找不到 Excel 檔案")
        sys.exit(1)
    except Exception as e:
        print(f"錯誤: 讀取 Excel 失敗 - {e}")
        sys.exit(1)

    # 確保所有目標欄位存在
    for field, _ in cells:
        if field not in df.columns:
            df[field] = None
        df[field] = df[field].astype(object)

    # 收集需要處理的唯一列號
    unique_rows = sorted(set(row for _, row in cells))
    max_row = len(df) + 1

    # 過濾無效列號
    valid_rows = [r for r in unique_rows if 2 <= r <= max_row]
    if len(valid_rows) < len(unique_rows):
        invalid = [r for r in unique_rows if r not in valid_rows]
        print(f"警告: 以下列號超出範圍: {invalid}")

    # 處理每個儲存格
    updated_count = 0
    photo_candidates = {}

    for field, excel_row in cells:
        if excel_row not in valid_rows:
            continue

        pandas_idx = excel_row - 2
        row_data = df.iloc[pandas_idx]

        name = row_data.get("姓名（中英）", "")
        company = row_data.get("所屬公司", "")

        if pd.isna(name) or not name:
            print(f"\n[{COLUMN_NAME_TO_LETTER.get(field, '?')}{excel_row}] 跳過 - 無姓名資料")
            continue

        # 檢查是否已有資料
        current_value = row_data.get(field)
        has_value = pd.notna(current_value) and str(current_value).strip() != ""

        if has_value and not force:
            print(f"\n[{COLUMN_NAME_TO_LETTER.get(field, '?')}{excel_row}] {name} - {field}")
            print(f"    已有資料，跳過 (使用 --force 強制更新)")
            continue

        print(f"\n[{COLUMN_NAME_TO_LETTER.get(field, '?')}{excel_row}] {name} ({company})")
        print(f"    搜尋: {field}")
        print("-" * 50)

        # 根據欄位類型選擇搜尋方式
        if field == "照片":
            result = search_photo_with_ddg(name, company)
            if result["best_url"]:
                df.at[pandas_idx, "照片"] = result["best_url"]
                if "照片狀態" in df.columns:
                    df.at[pandas_idx, "照片狀態"] = result["status"]
                print(f"    ✓ 找到照片 (狀態: {result['status']})")
                updated_count += 1

                # 儲存候選照片
                photo_candidates[str(excel_row)] = {
                    "name": name,
                    "company": company,
                    "best_url": result["best_url"],
                    "status": result["status"],
                    "candidates": result["candidates"]
                }
            else:
                print(f"    ✗ 未找到照片")
        else:
            # 使用 Perplexity API 搜尋
            value = search_field_with_perplexity(name, company, field)
            if value:
                cleaned = clean_value(value)
                if cleaned:
                    df.at[pandas_idx, field] = cleaned
                    display = cleaned.replace('\n', ' | ')
                    if len(display) > 60:
                        display = display[:60] + "..."
                    print(f"    ✓ [{field}]: {display}")
                    updated_count += 1
                else:
                    print(f"    ✗ 搜尋結果無效")
            else:
                print(f"    ✗ 未找到資料")

        # 避免 API 限制
        time.sleep(2)

    # 儲存結果
    output_path = Path(EXCEL_OUTPUT)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n{'=' * 60}")
        print(f"資料收集完成！")
        print(f"  - 目標儲存格: {len(cells)}")
        print(f"  - 成功更新: {updated_count}")
        print(f"  - 輸出檔案: {output_path}")
        print(f"{'=' * 60}")
    except PermissionError:
        backup_path = output_path.with_name("Standard_Example_Enriched_backup.xlsx")
        df.to_excel(backup_path, index=False, engine='openpyxl')
        print(f"\n⚠️  原檔案被鎖定，已儲存到: {backup_path}")
    except Exception as e:
        print(f"錯誤: 儲存 Excel 失敗 - {e}")
        sys.exit(1)

    # 儲存照片候選資料
    if photo_candidates:
        json_path = Path("output/data/photo_candidates.json")
        try:
            existing = {}
            if json_path.exists():
                with open(json_path, 'r', encoding='utf-8') as f:
                    existing = json.load(f)
            existing.update(photo_candidates)
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(existing, f, ensure_ascii=False, indent=2)
            print(f"\n照片候選資料已儲存: {json_path}")
        except Exception as e:
            print(f"警告: 儲存照片候選資料失敗 - {e}")

    return updated_count


def main():
    parser = argparse.ArgumentParser(
        description="針對特定儲存格進行資料收集",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
    # 使用 Excel 儲存格參照
    python src/enrich_cell.py --cell "H26"
    python src/enrich_cell.py --cell "H26,I27,J28"
    python src/enrich_cell.py --cell "H26-H30"

    # 使用欄位名稱 + 列號
    python src/enrich_cell.py --field "學歷" --rows "26"
    python src/enrich_cell.py --field "4" --rows "26-30"
    python src/enrich_cell.py --field "H" --rows "26,27,28"

欄位對應:
    1=年齡  2=專業分類  3=專業背景  4=學歷  5=主要經歷  6=現職/任
    7=個人特質  8=獨董家數  9=獨董年資  10=電子郵件  11=公司電話  12=照片

    H=學歷  I=主要經歷  J=現職/任  K=個人特質  ...
        """
    )

    parser.add_argument(
        "--cell",
        type=str,
        help="儲存格參照 (如 H26, H26-H30, H26,I27)"
    )
    parser.add_argument(
        "--field",
        type=str,
        help="欄位名稱/編號/代號 (如 學歷, 4, H)"
    )
    parser.add_argument(
        "--rows",
        type=str,
        help="列號 (如 26, 26-30, 26,27,28)"
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="強制更新已有資料的欄位"
    )

    args = parser.parse_args()

    # 解析目標儲存格
    cells = []

    if args.cell:
        cells = parse_cell_references(args.cell)
    elif args.field and args.rows:
        field_name = resolve_field_name(args.field)
        if not field_name:
            print(f"錯誤: 無法識別欄位 '{args.field}'")
            print("\n可用欄位:")
            for num, name in FIELD_NUMBER_TO_NAME.items():
                letter = COLUMN_NAME_TO_LETTER.get(name, "?")
                print(f"  {num}. {name} ({letter})")
            sys.exit(1)

        rows = parse_row_numbers(args.rows)
        cells = [(field_name, row) for row in rows]
    else:
        parser.print_help()
        sys.exit(1)

    if not cells:
        print("錯誤: 沒有有效的目標儲存格")
        sys.exit(1)

    enrich_cells(cells, force=args.force)


if __name__ == "__main__":
    main()
