"""
unified_search.py - 統一搜尋客戶端

整合 SerpAPI 和 DuckDuckGo，提供自動 fallback 機制。
優先使用 SerpAPI（品質較佳），額度不足時自動切換到 DuckDuckGo。
"""

from .serpapi_client import SerpAPIClient
from .ddg_client import DDGClient


class UnifiedSearchClient:
    """統一搜尋客戶端，整合 SerpAPI 和 DuckDuckGo。"""

    def __init__(self):
        self.serpapi = SerpAPIClient()
        self.ddg = DDGClient()

    def get_status(self) -> dict:
        """
        取得搜尋引擎狀態。

        Returns:
            包含各搜尋引擎狀態的字典
        """
        serpapi_stats = self.serpapi.get_usage_stats()

        return {
            "serpapi": {
                "available": serpapi_stats["available"],
                "has_quota": self.serpapi.check_quota(),
                "used": serpapi_stats["used"],
                "remaining": serpapi_stats["remaining"],
                "quota": serpapi_stats["quota"],
                "month": serpapi_stats["month"]
            },
            "duckduckgo": {
                "available": self.ddg.is_available()
            },
            "primary_engine": self._get_primary_engine()
        }

    def _get_primary_engine(self) -> str:
        """取得目前使用的主要搜尋引擎。"""
        if self.serpapi.check_quota():
            return "SerpAPI"
        elif self.ddg.is_available():
            return "DuckDuckGo"
        else:
            return "None"

    def search(self, query: str, num_results: int = 5) -> list[dict]:
        """
        執行搜尋，自動選擇最佳搜尋引擎。

        優先順序: SerpAPI (有額度時) > DuckDuckGo

        Args:
            query: 搜尋關鍵字
            num_results: 回傳結果數量

        Returns:
            搜尋結果列表，每個結果包含 title, href, body
        """
        # 優先使用 SerpAPI
        if self.serpapi.check_quota():
            results = self.serpapi.search_google(query, num_results)
            if results:
                return results

        # Fallback 到 DuckDuckGo
        if self.ddg.is_available():
            return self.ddg.search_text(query, num_results)

        return []

    def search_images(self, query: str, num_results: int = 5) -> list[str]:
        """
        執行圖片搜尋，自動選擇最佳搜尋引擎。

        優先順序: SerpAPI (有額度時) > DuckDuckGo

        Args:
            query: 搜尋關鍵字
            num_results: 回傳結果數量

        Returns:
            圖片 URL 列表
        """
        # 優先使用 SerpAPI
        if self.serpapi.check_quota():
            results = self.serpapi.search_images(query, num_results)
            if results:
                return results

        # Fallback 到 DuckDuckGo
        if self.ddg.is_available():
            return self.ddg.search_images(query, num_results)

        return []

    def search_with_engine(self, query: str, engine: str, num_results: int = 5) -> list[dict]:
        """
        使用指定的搜尋引擎執行搜尋。

        Args:
            query: 搜尋關鍵字
            engine: 搜尋引擎名稱 ("serpapi" 或 "duckduckgo")
            num_results: 回傳結果數量

        Returns:
            搜尋結果列表
        """
        if engine.lower() == "serpapi":
            return self.serpapi.search_google(query, num_results)
        elif engine.lower() in ["duckduckgo", "ddg"]:
            return self.ddg.search_text(query, num_results)
        else:
            return self.search(query, num_results)
