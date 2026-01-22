"""
field_formatter.py - 欄位格式化模組

定義每個欄位的顯示設定，並提供格式化函式。
"""

import pandas as pd

# 欄位設定字典
FIELD_CONFIG = {
    "專業背景": {
        "label": "專業背景",
        "excel_column": "專業背景",  # 現在由 Perplexity API 自動生成
        "multiline": False,
        "empty_text": "(待補充)"
    },
    "學歷": {
        "label": "學歷",
        "excel_column": "學歷",
        "multiline": True,
        "empty_text": "(待補充)"
    },
    "主要經歷": {
        "label": "主要經歷",
        "excel_column": "主要經歷",
        "multiline": True,
        "empty_text": "(待補充)"
    },
    "現任": {
        "label": "現任",
        "excel_column": "現職/任",
        "multiline": True,
        "empty_text": "(待補充)"
    },
    "個人特質": {
        "label": "個人特質",
        "excel_column": "個人特質",
        "multiline": True,
        "empty_text": "(待補充)"
    },
    "現擔任獨董家數": {
        "label": "現擔任獨董家數",
        "excel_column": "現擔任獨董家數(年)",
        "multiline": False,
        "empty_text": "0"
    },
    "擔任獨董年資": {
        "label": "擔任獨董年資",
        "excel_column": "擔任獨董年資(年)",
        "multiline": False,
        "empty_text": "0"
    }
}


def format_field_content(field_name: str, raw_value, config: dict) -> list[dict]:
    """
    將欄位值格式化為文字片段列表。

    Args:
        field_name: 欄位名稱
        raw_value: 原始值
        config: 欄位設定字典

    Returns:
        格式化內容列表:
        [
            {"text": "學歷", "bold": True},
            {"text": "：", "bold": False},
            {"text": "台灣大學 資訊系 碩士", "bold": False}
        ]
    """
    label = config.get("label", field_name)
    empty_text = config.get("empty_text", "(待補充)")

    # 處理空值
    if raw_value is None or (isinstance(raw_value, float) and pd.isna(raw_value)):
        content_text = empty_text
    elif isinstance(raw_value, str) and raw_value.strip() == "":
        content_text = empty_text
    else:
        content_text = str(raw_value).strip()

    return [
        {"text": label, "bold": True},
        {"text": "：", "bold": False},
        {"text": content_text, "bold": False}
    ]


def get_field_value_from_data(field_name: str, data: dict) -> str:
    """
    從資料字典中取得欄位值。

    Args:
        field_name: 欄位名稱（如 "學歷"）
        data: 資料字典

    Returns:
        欄位值或空字串
    """
    config = FIELD_CONFIG.get(field_name, {})
    excel_column = config.get("excel_column")

    # 優先使用 excel_column 映射
    if excel_column and excel_column in data:
        return data[excel_column]

    # 直接使用欄位名稱
    if field_name in data:
        return data[field_name]

    return ""


def is_empty_value(value) -> bool:
    """
    檢查值是否為空。

    Args:
        value: 要檢查的值

    Returns:
        是否為空值
    """
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False
