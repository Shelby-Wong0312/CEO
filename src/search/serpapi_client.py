"""
serpapi_client.py - SerpAPI 搜尋客戶端

提供 Google 搜尋和圖片搜尋功能，包含額度管理（60次/月）。
"""

import os
import json
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv

# 載入環境變數
load_dotenv()

# 嘗試導入 serpapi
try:
    from serpapi import GoogleSearch
    SERPAPI_AVAILABLE = True
except ImportError:
    SERPAPI_AVAILABLE = False

# 額度追蹤檔案
USAGE_FILE = ".serpapi_usage.json"
MONTHLY_QUOTA = 60


class SerpAPIClient:
    """SerpAPI 搜尋客戶端，包含額度管理。"""

    def __init__(self):
        self.api_key = os.getenv("SERPAPI_API_KEY")
        self.usage_file = Path(USAGE_FILE)
        self._load_usage()

    def _load_usage(self):
        """載入使用量紀錄。"""
        if self.usage_file.exists():
            try:
                with open(self.usage_file, 'r', encoding='utf-8') as f:
                    self.usage_data = json.load(f)
            except (json.JSONDecodeError, IOError):
                self.usage_data = self._init_usage_data()
        else:
            self.usage_data = self._init_usage_data()

        # 檢查是否需要重置月度計數
        current_month = datetime.now().strftime("%Y-%m")
        if self.usage_data.get("month") != current_month:
            self.usage_data = self._init_usage_data()
            self._save_usage()

    def _init_usage_data(self) -> dict:
        """初始化使用量資料。"""
        return {
            "month": datetime.now().strftime("%Y-%m"),
            "count": 0,
            "quota": MONTHLY_QUOTA,
            "last_used": None
        }

    def _save_usage(self):
        """儲存使用量紀錄。"""
        try:
            with open(self.usage_file, 'w', encoding='utf-8') as f:
                json.dump(self.usage_data, f, ensure_ascii=False, indent=2)
        except IOError as e:
            print(f"警告: 無法儲存 SerpAPI 使用量紀錄 - {e}")

    def _increment_usage(self):
        """增加使用次數。"""
        self.usage_data["count"] += 1
        self.usage_data["last_used"] = datetime.now().isoformat()
        self._save_usage()

    def is_available(self) -> bool:
        """檢查 SerpAPI 是否可用。"""
        return SERPAPI_AVAILABLE and bool(self.api_key)

    def check_quota(self) -> bool:
        """檢查是否還有額度。"""
        if not self.is_available():
            return False
        return self.usage_data["count"] < self.usage_data["quota"]

    def get_usage_stats(self) -> dict:
        """取得使用量統計。"""
        return {
            "available": self.is_available(),
            "month": self.usage_data.get("month"),
            "used": self.usage_data.get("count", 0),
            "quota": self.usage_data.get("quota", MONTHLY_QUOTA),
            "remaining": max(0, self.usage_data.get("quota", MONTHLY_QUOTA) - self.usage_data.get("count", 0)),
            "last_used": self.usage_data.get("last_used")
        }

    def search_google(self, query: str, num_results: int = 5) -> list[dict]:
        """
        使用 SerpAPI 進行 Google 搜尋。

        Args:
            query: 搜尋關鍵字
            num_results: 回傳結果數量

        Returns:
            搜尋結果列表，每個結果包含 title, href, body
        """
        if not self.is_available():
            return []

        if not self.check_quota():
            print("    [SerpAPI] 本月額度已用完")
            return []

        try:
            params = {
                "engine": "google",
                "q": query,
                "api_key": self.api_key,
                "num": num_results,
                "hl": "zh-TW",
                "gl": "tw"
            }

            search = GoogleSearch(params)
            results = search.get_dict()

            self._increment_usage()

            # 轉換為統一格式
            formatted_results = []
            organic_results = results.get("organic_results", [])

            for item in organic_results[:num_results]:
                formatted_results.append({
                    "title": item.get("title", ""),
                    "href": item.get("link", ""),
                    "body": item.get("snippet", "")
                })

            return formatted_results

        except Exception as e:
            print(f"    [SerpAPI] 搜尋錯誤: {e}")
            return []

    def search_images(self, query: str, num_results: int = 5) -> list[str]:
        """
        使用 SerpAPI 進行 Google 圖片搜尋。

        Args:
            query: 搜尋關鍵字
            num_results: 回傳結果數量

        Returns:
            圖片 URL 列表
        """
        if not self.is_available():
            return []

        if not self.check_quota():
            print("    [SerpAPI] 本月額度已用完")
            return []

        try:
            params = {
                "engine": "google_images",
                "q": query,
                "api_key": self.api_key,
                "num": num_results,
                "hl": "zh-TW",
                "gl": "tw"
            }

            search = GoogleSearch(params)
            results = search.get_dict()

            self._increment_usage()

            # 提取圖片 URL
            image_urls = []
            images_results = results.get("images_results", [])

            for item in images_results[:num_results]:
                original_url = item.get("original", "")
                if original_url:
                    image_urls.append(original_url)

            return image_urls

        except Exception as e:
            print(f"    [SerpAPI] 圖片搜尋錯誤: {e}")
            return []
