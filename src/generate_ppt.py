"""
generate_ppt.py - PowerPoint 簡報生成腳本

根據擴充後的 Excel 資料和 CV 範本，
為指定的主管生成獨立的 PowerPoint 簡報。

功能:
1. 使用 CVTemplateEngine 處理範本
2. 粗體標題格式（如「學歷：」）
3. 多行內容正確換行顯示
4. 無資料時顯示「(待補充)」

使用方式:
    python src/generate_ppt.py --rows "2, 5-10, 15"
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
from pathlib import Path

import pandas as pd
import requests

# 導入 PPT 模組
try:
    from src.ppt import CVTemplateEngine, FIELD_CONFIG
except ImportError:
    # 直接從同級目錄導入
    from ppt import CVTemplateEngine, FIELD_CONFIG


# === 常數定義 ===
EXCEL_INPUT = "output/data/Standard_Example_Enriched.xlsx"
EXCEL_FALLBACK = "Standard Example.xlsx"
TEMPLATE_PATH = "CV_標準範本.pptx"
OUTPUT_DIR = "output/ppt"

# 照片選擇檔案可能的位置
PHOTO_SELECTION_PATHS = [
    "output/data/photo_selections.json",
    "photo_selections.json",
    Path.home() / "Downloads" / "photo_selections.json"
]


def apply_photo_selections(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """
    自動套用照片選擇到 DataFrame。

    搜尋 photo_selections.json 並套用選擇的照片 URL。

    Args:
        df: Excel DataFrame

    Returns:
        (更新後的 df, 套用的筆數)
    """
    # 尋找 photo_selections.json
    json_path = None
    for path in PHOTO_SELECTION_PATHS:
        p = Path(path)
        if p.exists():
            json_path = p
            break

    if not json_path:
        return df, 0

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            selections = json.load(f)

        if not selections:
            return df, 0

        print(f"\n找到照片選擇檔案: {json_path}")
        print(f"套用 {len(selections)} 筆照片選擇...")

        # 確保欄位為 object 類型
        if "照片" in df.columns:
            df["照片"] = df["照片"].astype(object)
        if "照片狀態" in df.columns:
            df["照片狀態"] = df["照片狀態"].astype(object)

        applied_count = 0
        for row_str, data in selections.items():
            try:
                excel_row = int(row_str)
                pandas_idx = excel_row - 2

                if pandas_idx < 0 or pandas_idx >= len(df):
                    continue

                # 取得選擇的 URL
                if isinstance(data, dict):
                    selected_url = data.get("selected_url", "")
                    status = data.get("status", "已確認")
                else:
                    selected_url = str(data) if data else ""
                    status = "已確認" if selected_url else "待補充"

                # 更新 DataFrame
                if selected_url:
                    df.at[pandas_idx, "照片"] = selected_url
                    if "照片狀態" in df.columns:
                        df.at[pandas_idx, "照片狀態"] = status
                    applied_count += 1

            except (ValueError, KeyError):
                continue

        if applied_count > 0:
            # 儲存更新後的 Excel
            excel_path = Path(EXCEL_INPUT)
            if excel_path.exists():
                df.to_excel(excel_path, index=False, engine='openpyxl')
                print(f"已套用 {applied_count} 筆照片選擇並儲存到 Excel")

        return df, applied_count

    except Exception as e:
        print(f"警告: 套用照片選擇失敗 - {e}")
        return df, 0


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


def download_image(url: str) -> io.BytesIO | None:
    """從 URL 下載圖片並返回 BytesIO 物件。"""
    if not url or pd.isna(url):
        return None

    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if not content_type.startswith("image/"):
            print(f"  警告: URL 不是圖片格式 ({content_type})")
            return None

        return io.BytesIO(response.content)

    except requests.exceptions.RequestException as e:
        print(f"  警告: 圖片下載失敗 - {e}")
        return None


def sanitize_filename(name: str) -> str:
    """清理檔案名稱，移除不合法字元。"""
    invalid_chars = r'<>:"/\|?*'
    for char in invalid_chars:
        name = name.replace(char, "")
    return name.strip()


def generate_ppt(rows_str: str):
    """主要 PowerPoint 生成函式。"""
    print("=" * 60)
    print("PowerPoint 簡報生成程序啟動")
    print("=" * 60)

    # 1. 解析列號
    target_rows = parse_row_numbers(rows_str)
    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    print(f"\n目標 Excel 列號: {target_rows}")
    print(f"共 {len(target_rows)} 份簡報待生成")

    # 2. 讀取 Excel 檔案
    excel_path = EXCEL_INPUT
    if not Path(excel_path).exists():
        print(f"注意: 找不到 '{EXCEL_INPUT}'，改用 '{EXCEL_FALLBACK}'")
        excel_path = EXCEL_FALLBACK

    try:
        df = pd.read_excel(excel_path)
        print(f"\n成功讀取 '{excel_path}'")
        print(f"資料共 {len(df)} 列")

        # 自動套用照片選擇（如果有 photo_selections.json）
        df, photo_applied = apply_photo_selections(df)

    except FileNotFoundError:
        print(f"錯誤: 找不到 Excel 檔案")
        sys.exit(1)
    except Exception as e:
        print(f"錯誤: 讀取 Excel 失敗 - {e}")
        sys.exit(1)

    # 3. 建立輸出目錄
    output_path = Path(OUTPUT_DIR)
    output_path.mkdir(parents=True, exist_ok=True)

    # 4. 驗證列號範圍
    max_excel_row = len(df) + 1
    invalid_rows = [r for r in target_rows if r > max_excel_row or r < 2]
    if invalid_rows:
        print(f"警告: 以下列號超出範圍: {invalid_rows}")
        target_rows = [r for r in target_rows if r <= max_excel_row and r >= 2]

    if not target_rows:
        print("錯誤: 沒有有效的目標列號")
        sys.exit(1)

    # 5. 生成每份簡報
    success_count = 0

    for excel_row in target_rows:
        pandas_idx = excel_row_to_pandas_index(excel_row)
        row_data = df.iloc[pandas_idx]

        name = row_data.get("姓名（中英）", "")
        company = row_data.get("所屬公司", "")

        if pd.isna(name) or not name:
            print(f"\n[列 {excel_row}] 跳過 - 無姓名資料")
            continue

        print(f"\n[列 {excel_row}] 生成中: {name} ({company})")

        # 載入範本
        engine = CVTemplateEngine()
        if not engine.load_template():
            print(f"  錯誤: 載入範本失敗")
            continue

        # 設定姓名（不含年齡，年齡在 Shape 5 Group 中單獨處理）
        engine.set_name(str(name))
        print(f"  → 設定姓名: {name}")

        # 設定年齡（Shape 5 Group 內的子形狀）
        age = row_data.get("年齡", "")
        engine.set_age(age)

        # 準備左側資料
        left_data = {
            "專業背景": row_data.get("專業背景", ""),  # 由 Perplexity API 自動生成
            "學歷": row_data.get("學歷", ""),
            "主要經歷": row_data.get("主要經歷", "")
        }
        engine.fill_left_content(left_data)

        # 統計左側多行欄位
        left_multiline = sum(1 for k, v in left_data.items() if v and '\n' in str(v))
        print(f"  → 填入左側內容: 專業背景、學歷、主要經歷")

        # 準備右側資料
        right_data = {
            "現任": row_data.get("現職/任", ""),
            "個人特質": row_data.get("個人特質", ""),
            "現擔任獨董家數": row_data.get("現擔任獨董家數(年)", 0),
            "擔任獨董年資": row_data.get("擔任獨董年資(年)", 0)
        }
        engine.fill_right_content(right_data)

        # 統計右側多行欄位
        right_multiline = sum(1 for k, v in right_data.items() if v and '\n' in str(v))
        print(f"  → 填入右側內容: 現任、個人特質、獨董資訊")

        # 多行欄位統計
        total_multiline = left_multiline + right_multiline
        if total_multiline > 0:
            print(f"  → 多行欄位: {total_multiline} 個")

        # 設定照片
        photo_url = row_data.get("照片", "")
        is_valid_url = isinstance(photo_url, str) and photo_url.lower().startswith("http")

        if is_valid_url:
            print(f"  → 下載照片中...")
            image_stream = download_image(photo_url)
            if image_stream:
                if engine.set_photo(image_stream):
                    print(f"  → 照片已更新")
                else:
                    print(f"  → 照片設定失敗，保留範本圖片")
            else:
                print(f"  → 照片下載失敗，保留範本圖片")
        else:
            print(f"  → 無有效照片 URL，保留範本圖片")

        # 取得專業分類（用於資料夾分類）
        category = row_data.get("專業分類", "")
        if pd.isna(category) or not category:
            category = "未分類"
        category = str(category).strip()

        # 驗證分類名稱，避免無效資料夾名稱
        valid_categories = ["會計/財務類", "法務類", "商務/管理類", "產業專業類", "其他專門職業"]
        if category not in valid_categories and category != "未分類":
            category = "未分類"

        # 替換斜線為底線（避免資料夾名稱問題）
        category_folder = category.replace("/", "_")
        print(f"  → 專業分類: {category}")

        # 生成檔名
        name_parts = str(name).split()
        chinese_name = name_parts[0] if name_parts else str(name)

        company_parts = str(company).split()
        company_name = company_parts[-1] if company_parts else str(company)
        if not any('\u4e00' <= c <= '\u9fff' for c in company_name):
            company_name = company_parts[0] if company_parts else str(company)

        filename = sanitize_filename(f"{chinese_name}_{company_name}_CV.pptx")

        # 建立分類資料夾路徑
        category_path = output_path / category_folder
        category_path.mkdir(parents=True, exist_ok=True)
        filepath = category_path / filename

        # 儲存簡報
        if engine.save(filepath):
            print(f"  → 已儲存: {filepath}")
            success_count += 1
        else:
            print(f"  錯誤: 儲存簡報失敗")

    # 6. 完成摘要
    print(f"\n{'=' * 60}")
    print(f"生成完成！")
    print(f"  - 目標列數: {len(target_rows)}")
    print(f"  - 成功生成: {success_count} 份")
    print(f"  - 輸出目錄: {output_path}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="PowerPoint 簡報生成腳本 - 使用 CV 範本引擎",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
    python src/generate_ppt.py --rows "2"
    python src/generate_ppt.py --rows "2, 5, 10"
    python src/generate_ppt.py --rows "2-10"

功能:
    - 粗體標題格式（學歷：、主要經歷：、現任：等）
    - 多行內容正確換行顯示
    - 無資料時顯示「(待補充)」
        """
    )
    parser.add_argument(
        "--rows",
        type=str,
        required=True,
        help="要處理的 Excel 列號"
    )

    args = parser.parse_args()
    generate_ppt(args.rows)
