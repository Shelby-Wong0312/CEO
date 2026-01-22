"""
ddg_client.py - DuckDuckGo 搜尋客戶端

提供免費的網路搜尋和圖片搜尋功能，作為 SerpAPI 的 fallback。
"""

# 嘗試導入 DuckDuckGo 搜尋
try:
    from duckduckgo_search import DDGS
    DDGS_AVAILABLE = True
except ImportError:
    DDGS_AVAILABLE = False


class DDGClient:
    """DuckDuckGo 搜尋客戶端。"""

    def __init__(self):
        pass

    def is_available(self) -> bool:
        """檢查 DuckDuckGo 是否可用。"""
        return DDGS_AVAILABLE

    def search_text(self, query: str, max_results: int = 5) -> list[dict]:
        """
        使用 DuckDuckGo 進行網路搜尋。

        Args:
            query: 搜尋關鍵字
            max_results: 回傳結果數量

        Returns:
            搜尋結果列表，格式與 SerpAPI 相同 (title, href, body)
        """
        if not DDGS_AVAILABLE:
            return []

        try:
            with DDGS() as ddgs:
                results = list(ddgs.text(query, max_results=max_results, region='tw-tzh'))

                # 轉換為統一格式
                formatted_results = []
                for item in results:
                    formatted_results.append({
                        "title": item.get("title", ""),
                        "href": item.get("href", ""),
                        "body": item.get("body", "")
                    })

                return formatted_results

        except Exception as e:
            print(f"    [DuckDuckGo] 搜尋錯誤: {e}")
            return []

    def search_images(self, query: str, max_results: int = 5) -> list[str]:
        """
        使用 DuckDuckGo 進行圖片搜尋。

        Args:
            query: 搜尋關鍵字
            max_results: 回傳結果數量

        Returns:
            圖片 URL 列表
        """
        if not DDGS_AVAILABLE:
            return []

        try:
            with DDGS() as ddgs:
                results = list(ddgs.images(query, max_results=max_results))

                # 提取圖片 URL 並過濾
                image_urls = []
                for result in results:
                    image_url = result.get('image', '')
                    if image_url:
                        # 驗證 URL 是否為有效圖片格式
                        lower_url = image_url.lower()
                        if any(ext in lower_url for ext in ['.jpg', '.jpeg', '.png', '.webp']):
                            # 排除明顯的非照片來源
                            if not any(bad in lower_url for bad in ['logo', 'icon', 'banner', 'placeholder']):
                                image_urls.append(image_url)

                return image_urls

        except Exception as e:
            print(f"    [DuckDuckGo] 圖片搜尋錯誤: {e}")
            return []
