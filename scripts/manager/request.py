"""
请求管理器模块

该模块提供了 RequestManager 类，用于从全国人大法律数据库网站（flk.npc.gov.cn）
获取法律相关的数据，包括法律列表和Word文档。
所有请求都通过缓存机制进行优化，避免重复请求。
"""
import json
import logging
import urllib.request
from hashlib import sha1
from time import sleep

import requests
import certifi
from docx import Document
from manager.cache import CacheManager, CacheType

logger = logging.getLogger(__name__)

# HTTP请求头，模拟浏览器访问全国人大法律数据库网站
REQUEST_HEADER = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9,en-GB-oxendict;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
    "Content-Type": "application/json;charset=UTF-8",
    "Origin": "https://flk.npc.gov.cn",
    "Referer": "https://flk.npc.gov.cn/search",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36",
    "sec-ch-ua": '"Not(A:Brand";v="8", "Chromium";v="144", "Google Chrome";v="144"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"macOS"',
}


class RequestManager(object):
    """
    请求管理器类
    
    负责从全国人大法律数据库网站获取法律数据，包括：
    - 法律列表（分页）
    - Word格式的法律文档
    
    所有请求结果都会通过缓存机制进行存储，以提高效率并减少对服务器的请求。
    """
    
    def __init__(self) -> None:
        """
        初始化请求管理器
        
        初始化缓存管理器和默认搜索参数。
        """
        self.cache = CacheManager()  # 缓存管理器实例
        # 默认搜索参数
        self.default_search_params = {
            "searchRange": 1,
            "sxrq": [],
            "gbrq": [],
            "searchType": 2,
            "sxx": [],
            "gbrqYear": [],
            "flfgCodeId": [],
            "zdjgCodeId": [],
            "searchContent": "",
            "orderByParam": {"order": "-1", "sort": ""},
        }

    def getLawList(self, page=1, page_size=20):
        """
        获取法律列表
        
        从全国人大法律数据库API获取分页的法律列表数据。
        支持缓存机制，相同参数的请求会直接返回缓存结果。
        
        Args:
            page (int): 页码，默认为1
            page_size (int): 每页数量，默认为20
            
        Returns:
            dict: 包含法律列表的JSON数据，包括 total（总数）、rows（法律条目列表）等信息
                  每个法律条目包含 bbbs（唯一标识）、title（标题）、gbrq（公布日期）等字段
        """
        # 构建请求参数
        request_body = {
            **self.default_search_params,
            "pageNum": page,
            "pageSize": page_size,
        }

        # 使用参数的SHA1哈希值作为缓存键
        cache_key = sha1(json.dumps(request_body, sort_keys=True).encode()).hexdigest()

        # 检查缓存，如果存在则直接返回
        if cache := self.cache.get(cache_key, CacheType.WebPage, "json"):
            return cache

        # 发送POST请求获取法律列表
        response = requests.post(
            "https://flk.npc.gov.cn/law-search/search/list",
            headers=REQUEST_HEADER,
            json=request_body,
        )
        sleep(1)  # 延迟1秒，避免请求过于频繁
        logger.debug(f"requesting [{response.status_code}] page={page} pageSize={page_size}")

        # 解析JSON响应并缓存
        ret = response.json()
        rows = ret.get("rows", []) if isinstance(ret, dict) else []
        if "flxz" in extra_filters:
            rows = [row for row in rows if row.get("flxz") == extra_filters["flxz"]]
        mapped_rows = [
            {
                "id": row.get("bbbs"),
                "title": row.get("title"),
                "publish": row.get("gbrq"),
                "effect": row.get("sxrq"),
                "level": row.get("flxz"),
                "raw": row,
            }
            for row in rows
        ]
        mapped = {
            "result": {
                "data": mapped_rows,
                "total": ret.get("total", 0),
            }
        }
        self.cache.set(cache_key, CacheType.WebPage, mapped, "json")
        return mapped

    def get_download_url(self, bbbs: str, file_format: str = "docx"):
        """
        获取法律文档下载链接
        
        根据法律的唯一标识符（bbbs）获取文档的下载链接。
        支持缓存机制，相同参数的请求会直接返回缓存结果。
        
        Args:
            bbbs (str): 法律的唯一标识符（从法律列表中获取）
            file_format (str): 文件格式，默认为 "docx"
            
        Returns:
            dict: 包含下载链接的JSON数据，结构为:
                  {"code": 200, "msg": "操作成功", "data": {"url": "...", "urlIn": "..."}}
                  其中 url 为外网下载链接
        """
        cache_key = f"download_{bbbs}_{file_format}"
        
        # 检查缓存，如果存在则直接返回
        if cache := self.cache.get(cache_key, CacheType.WebPage, "json"):
            return cache
        
        logger.debug(f"getting download url for {bbbs}")
        
        # 发送GET请求获取下载链接
        download_headers = {k: v for k, v in REQUEST_HEADER.items() if k != "Content-Type"}
        ret = requests.get(
            "https://flk.npc.gov.cn/law-search/download/pc",
            headers=download_headers,
            params={"format": file_format, "bbbs": bbbs},
        )
        sleep(1)  # 延迟1秒，避免请求过于频繁
        ret = ret.json()
        # 缓存结果
        self.cache.set(cache_key, CacheType.WebPage, ret, "json")
        return ret

    def get_word(self, bbbs: str, title: str) -> Document:
        """
        获取Word格式的法律文档
        
        根据法律的唯一标识符（bbbs）获取下载链接，然后下载Word文档。
        支持缓存机制，如果文件已存在则直接读取，否则下载并缓存。
        
        Args:
            bbbs (str): 法律的唯一标识符（从法律列表中获取）
            title (str): 文档标题，用于确定缓存文件名
            
        Returns:
            Document: python-docx的Document对象，如果下载失败则返回None
        """
        file_extension = ".docx"
        
        # 检查缓存中是否已存在该文档
        ok, path = self.cache.is_exists(title, CacheType.WordDocument, file_extension)
        if not ok:
            # 获取下载链接
            download_info = self.get_download_url(bbbs, "docx")
            if download_info.get("code") != 200:
                logger.error(f"Failed to get download url for {bbbs}: {download_info.get('msg')}")
                return None
            
            download_url = download_info.get("data", {}).get("url")
            if not download_url:
                logger.error(f"No download url found for {bbbs}")
                return None

            logger.debug(f"getting law word file for {title}")

            try:
                # 下载Word文档到指定路径
                urllib.request.urlretrieve(download_url, path)
                sleep(1)  # 延迟1秒，避免请求过于频繁
            except Exception as e:
                logger.error(f"Failed to download word file: {e}")
                return None

        # 打开文件并解析为Document对象
        with open(path, "rb") as f:
            try:
                return Document(f)
            except Exception as e:
                # 如果解析失败（可能是文件损坏），返回None
                logger.error(f"Failed to parse word document: {e}")
                return None
