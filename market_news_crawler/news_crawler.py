#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import os
import platform
import re
import socket
import sys
import time
import textwrap
import urllib.parse
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup
from dateutil import parser as date_parser
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from country_config import default_country_file_paths, resolve_project_path

try:
    from deep_translator import GoogleTranslator
except ImportError:  # pragma: no cover - dependency is installed at runtime
    GoogleTranslator = None


USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/135.0.0.0 Safari/537.36"
)

TOKYO_TZ = "Asia/Tokyo"
TRANSLATION_MAX_CHARS = 600
SUMMARY_MAX_CHARS = 420
SUMMARY_MIN_PARAGRAPH_CHARS = 20
DEFAULT_COUNTRY_FILE_PATHS = default_country_file_paths("japan")
DEFAULT_SITE_CREDENTIALS_PATH = DEFAULT_COUNTRY_FILE_PATHS["site_credentials_path"]
DEFAULT_EXTRA_SOURCES_PATH = DEFAULT_COUNTRY_FILE_PATHS["extra_sources_path"]
DEFAULT_ADAPTER_CONFIGS_PATH = DEFAULT_COUNTRY_FILE_PATHS["adapter_configs_path"]
TRANSLATION_TARGETS = {
    "zh": "zh-CN",
    "zh-cn": "zh-CN",
    "cn": "zh-CN",
    "chinese": "zh-CN",
    "en": "en",
    "english": "en",
}
SUMMARY_SENTENCE_BREAK_RE = re.compile(r"(?<=[。！？!?\.])\s+")
SUMMARY_NOISE_PATTERNS = [
    re.compile(pattern, re.IGNORECASE)
    for pattern in [
        r"^read more$",
        r"^learn more$",
        r"^click here$",
        r"^続きを読む$",
        r"^もっと見る$",
        r"^詳細はこちら$",
        r"^関連(記事|情報).*$",
        r"^おすすめ(記事)?$",
        r"^share this.*$",
        r"^copyright .*",
        r"^all rights reserved.*",
    ]
]

TRUTHY_ENV_VALUES = {"1", "true", "yes", "on"}
FALSY_ENV_VALUES = {"0", "false", "no", "off"}

PLATFORM_LABELS = {
    "tiktok_shop": "TikTok Shop",
    "amazon_japan": "Amazon Japan",
    "rakuten_ichiba": "Rakuten Ichiba",
    "qoo10": "Qoo10",
    "temu": "TEMU",
    "shein": "SHEIN",
}

BUILTIN_PLATFORM_ALIASES = {
    "tiktok_shop": [
        "tiktok_shop",
        "tiktok shop",
        "tiktok/tiktok shop",
        "tiktok",
    ],
    "amazon_japan": [
        "amazon_japan",
        "amazon japan",
        "amazon",
    ],
    "rakuten_ichiba": [
        "rakuten_ichiba",
        "rakuten ichiba",
        "rakuten",
    ],
    "qoo10": [
        "qoo10",
    ],
    "temu": [
        "temu",
    ],
    "shein": [
        "shein",
        "she in",
    ],
}

NPS_DIMENSION_RULES = [
    ("Post-purchase service", ["refund", "return", "返品", "返金", "cancel", "cancellation", "解約", "aftersales", "アフター", "customer service", "サポート"]),
    ("Logistics", ["shipping", "delivery", "配送", "発送", "物流", "warehouse", "fulfillment", "courier", "parcel", "same-day", "翌日"]),
    ("Price", ["price", "pricing", "discount", "coupon", "voucher", "sale", "promotion", "promo", "point", "ポイント", "価格", "値下げ", "割引"]),
    ("Content", ["video", "live", "livestream", "short video", "creator", "content", "recommendation", "discovery", "動画", "ライブ", "配信", "コンテンツ"]),
    ("Seller quality", ["seller", "merchant", "shop", "partner", "marketplace seller", "shop owner", "出店", "店舗", "加盟店"]),
    ("Product quality", ["quality", "counterfeit", "authentic", "safety", "安全", "品質", "正規品", "偽物"]),
    ("Product variety", ["brand", "assortment", "catalog", "selection", "variety", "category expansion", "品揃え", "ブランド", "カテゴリ"]),
    ("Product feature", ["feature", "function", "app", "search", "payment", "checkout", "membership", "review", "rating", "ui", "ux", "security", "機能", "アプリ", "検索", "決済", "レビュー"]),
]

SOURCE_AUDIT_TARGETS = {
    "tiktok_shop": [
        ("pdf_listed", "TikTok Shop Seller Center", "https://seller.tiktokglobalshop.com/"),
        ("pdf_listed", "TikTok Newsroom (JP)", "https://newsroom.tiktok.com/ja-jp/"),
    ],
    "amazon_japan": [
        ("pdf_listed", "Amazon Seller Central JP", "https://sellercentral.amazon.co.jp/"),
        ("pdf_listed", "About Amazon Japan", "https://www.aboutamazon.jp/"),
    ],
    "rakuten_ichiba": [
        ("pdf_listed", "RMS Login", "https://glogin.rms.rakuten.co.jp/"),
        ("pdf_listed", "Rakuten Seller Manual", "https://navi-manual.faq.rakuten.net/"),
        ("crawl_fallback", "Rakuten Corporate Press", "https://corp.rakuten.co.jp/news/press/"),
    ],
    "qoo10": [
        ("pdf_listed", "QSM Seller Manager", "https://qsm.qoo10.jp/"),
        ("pdf_listed", "Qoo10 Consumer Help", "https://www.qoo10.jp/gmkt.inc/CS/NHelpHome.aspx"),
        ("crawl_fallback", "Qoo10 University Feed", "https://article-university.qoo10.jp/feed"),
    ],
    "temu": [
        ("pdf_listed", "Sell on TEMU", "https://www.temu.com/sell-on-temu.html"),
        ("pdf_listed", "TEMU Official", "https://www.temu.com/"),
        ("crawl_fallback", "TEMU Announcements", "https://www.temu.com/about_temu_home.html"),
    ],
    "shein": [
        ("pdf_listed", "SHEIN Seller Center", "https://seller.shein.com/"),
        ("crawl_fallback", "SHEIN Group Newsroom", "https://www.sheingroup.com/newsroom"),
    ],
}


@dataclass
class SourceAuditRecord:
    platform: str
    platform_label: str
    source_role: str
    source_name: str
    source_url: str
    status: str
    http_status: int | None
    final_url: str | None
    note: str


@dataclass
class NewsItem:
    platform: str
    platform_label: str
    source_name: str
    source_url: str
    article_url: str
    title: str
    published_at: str
    category: str | None
    summary: str | None
    nps_dimension_guess: str | None
    source_country_guess: str
    verification_status: str
    verification_checked_at: str
    verification_final_url: str | None


@dataclass
class ExtraSourceEntry:
    platform: str
    side: str
    source_url: str


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="抓取日本电商平台官方新闻，并输出可核验的真实链接结果。",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=textwrap.dedent(
            """\
            示例:
              python3 news_crawler.py
              python3 news_crawler.py --days 14
              python3 news_crawler.py --start-date 2026-04-01 --end-date 2026-04-14
              python3 news_crawler.py --platform qoo10 --platform temu
            """
        ),
    )
    parser.add_argument("--days", type=int, default=7, help="向前回溯多少天，默认 7 天。")
    parser.add_argument("--start-date", help="起始日期，格式 YYYY-MM-DD。")
    parser.add_argument("--end-date", help="结束日期，格式 YYYY-MM-DD。默认是当前时间。")
    parser.add_argument(
        "--platform",
        dest="platforms",
        action="append",
        help="限制抓取的平台，可重复传入；支持内置平台 ID、展示名，或国家子目录 extra_sources.json 中的自定义平台名。",
    )
    parser.add_argument("--timezone", default=TOKYO_TZ, help=f"日期范围比较所用时区，默认 {TOKYO_TZ}。")
    parser.add_argument("--output-dir", default="outputs", help="输出目录，默认 ./outputs")
    parser.add_argument("--translate-to", default="zh-CN", help="翻译输出语言，默认 zh-CN，可选 zh-CN / en")
    parser.add_argument("--extra-sources", default=DEFAULT_EXTRA_SOURCES_PATH, help="额外来源配置 JSON，默认使用国家子目录文件")
    parser.add_argument("--adapter-configs", default=DEFAULT_ADAPTER_CONFIGS_PATH, help="站点适配配置 JSON，默认使用国家子目录文件")
    parser.add_argument("--site-credentials", default=DEFAULT_SITE_CREDENTIALS_PATH, help="站点凭据配置 JSON，默认使用国家子目录文件")
    return parser.parse_args(argv)


def build_session() -> requests.Session:
    session = requests.Session()
    if should_ignore_system_proxy_by_default():
        session.trust_env = False
    apply_explicit_or_detected_proxy(session)
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ja,en;q=0.9,zh-CN;q=0.8"})
    retry = Retry(
        total=2,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session._credential_login_markers = set()
    return session


def apply_explicit_or_detected_proxy(session: requests.Session) -> None:
    explicit_proxy = os.getenv("MARKET_NEWS_PROXY", "").strip()
    if explicit_proxy:
        session.proxies.update({"http": explicit_proxy, "https": explicit_proxy})
        return

    explicit_http_proxy = os.getenv("MARKET_NEWS_HTTP_PROXY", "").strip()
    explicit_https_proxy = os.getenv("MARKET_NEWS_HTTPS_PROXY", "").strip()
    if explicit_http_proxy or explicit_https_proxy:
        if explicit_http_proxy:
            session.proxies["http"] = explicit_http_proxy
        if explicit_https_proxy:
            session.proxies["https"] = explicit_https_proxy
        return

    env_proxy = os.getenv("ALL_PROXY", "").strip()
    env_http_proxy = os.getenv("HTTP_PROXY", "").strip()
    env_https_proxy = os.getenv("HTTPS_PROXY", "").strip()
    if env_proxy or env_http_proxy or env_https_proxy:
        if env_proxy:
            normalized_proxy = normalize_proxy_url(env_proxy)
            session.proxies.update({"http": normalized_proxy, "https": normalized_proxy})
        if env_http_proxy:
            session.proxies["http"] = normalize_proxy_url(env_http_proxy)
        if env_https_proxy:
            session.proxies["https"] = normalize_proxy_url(env_https_proxy)
        session.trust_env = False
        return

    if os.name != "nt":
        return
    if session.trust_env is False:
        return
    registry_proxies = windows_system_proxy_settings()
    if registry_proxies:
        session.proxies.update(registry_proxies)
        session.trust_env = False
        return
    if os.getenv("MARKET_NEWS_AUTO_LOCAL_PROXY", "").strip().lower() in FALSY_ENV_VALUES:
        return

    for port in (7890, 7897, 10809):
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.15):
                proxy = f"http://127.0.0.1:{port}"
                session.proxies.update({"http": proxy, "https": proxy})
                session.trust_env = False
                return
        except OSError:
            continue


def normalize_proxy_url(value: str) -> str:
    proxy = value.strip()
    if not proxy:
        return ""
    if "://" not in proxy:
        proxy = f"http://{proxy}"
    return proxy


def windows_system_proxy_settings() -> dict[str, str]:
    if os.name != "nt":
        return {}
    try:
        import winreg
    except ImportError:
        return {}

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Internet Settings") as key:
            proxy_enable = int(winreg.QueryValueEx(key, "ProxyEnable")[0] or 0)
            proxy_server = str(winreg.QueryValueEx(key, "ProxyServer")[0] or "").strip()
    except OSError:
        return {}

    if not proxy_enable or not proxy_server:
        return {}

    proxies: dict[str, str] = {}
    if "=" not in proxy_server:
        proxy = normalize_proxy_url(proxy_server)
        return {"http": proxy, "https": proxy} if proxy else {}

    for item in proxy_server.split(";"):
        if "=" not in item:
            continue
        scheme, proxy_value = item.split("=", 1)
        scheme = scheme.strip().lower()
        proxy = normalize_proxy_url(proxy_value)
        if scheme in {"http", "https"} and proxy:
            proxies[scheme] = proxy
    if "http" in proxies and "https" not in proxies:
        proxies["https"] = proxies["http"]
    return proxies


def should_ignore_system_proxy_by_default() -> bool:
    env_value = os.getenv("MARKET_NEWS_IGNORE_SYSTEM_PROXY", "").strip().lower()
    if env_value in TRUTHY_ENV_VALUES:
        return True
    if env_value in FALSY_ENV_VALUES:
        return False
    return False


def request_error_chain_text(exc: BaseException) -> str:
    parts: list[str] = []
    current: BaseException | None = exc
    seen_ids: set[int] = set()
    while current is not None and id(current) not in seen_ids:
        seen_ids.add(id(current))
        message = str(current or "").strip()
        if message:
            parts.append(f"{type(current).__name__}: {message}")
        next_exc = getattr(current, "__cause__", None) or getattr(current, "__context__", None)
        current = next_exc if isinstance(next_exc, BaseException) else None
    return " | ".join(parts).lower()


def should_retry_without_proxy(exc: requests.exceptions.RequestException) -> bool:
    error_text = request_error_chain_text(exc)
    if isinstance(exc, requests.exceptions.ProxyError):
        return True
    if isinstance(exc, requests.exceptions.SSLError) and (
        "eof occurred in violation of protocol" in error_text
        or "wrong version number" in error_text
        or "tlsv1 alert" in error_text
    ):
        return True
    if isinstance(exc, requests.exceptions.ConnectionError) and "proxy" in error_text:
        return True
    if isinstance(exc, (requests.exceptions.ConnectTimeout, requests.exceptions.ReadTimeout)) and "proxy" in error_text:
        return True
    return "unable to connect to proxy" in error_text or "proxyerror" in error_text


def clone_session_without_env_proxy(session: requests.Session) -> requests.Session:
    direct_session = requests.Session()
    direct_session.trust_env = False
    direct_session.headers.update(dict(session.headers))
    for prefix, adapter in session.adapters.items():
        direct_session.mount(prefix, adapter)
    markers = getattr(session, "_credential_login_markers", None)
    if isinstance(markers, set):
        direct_session._credential_login_markers = set(markers)
    return direct_session


def now_iso(tz_name: str) -> str:
    return datetime.now(ZoneInfo(tz_name)).isoformat()


def parse_date_range(args: argparse.Namespace) -> tuple[datetime, datetime]:
    tz = ZoneInfo(args.timezone)
    end = (
        datetime.now(tz)
        if not args.end_date
        else datetime.combine(date_parser.parse(args.end_date).date(), datetime.max.time(), tz)
    )
    start = (
        end - timedelta(days=args.days)
        if not args.start_date
        else datetime.combine(date_parser.parse(args.start_date).date(), datetime.min.time(), tz)
    )
    if start > end:
        raise SystemExit("start-date 不能晚于 end-date。")
    return start, end


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    return " ".join(value.replace("\u3000", " ").replace("\xa0", " ").split())


def html_to_text(value: str | None) -> str:
    if not value:
        return ""
    normalized = clean_text(value)
    # Some feed summaries are plain URLs rather than HTML snippets.
    # Parsing those with BeautifulSoup triggers MarkupResemblesLocatorWarning.
    if normalized.startswith(("http://", "https://")) and "<" not in normalized and ">" not in normalized:
        return normalized
    soup = BeautifulSoup(value, "lxml")
    return clean_text(soup.get_text(" ", strip=True))


def parse_dt(value: str | None, default_tz: str = TOKYO_TZ) -> datetime | None:
    if not value:
        return None
    normalized = clean_text(value)
    if "—" in normalized:
        normalized = normalized.split("—", 1)[1].strip()
    normalized = re.sub(r"（[^）]*）", "", normalized)
    normalized = re.sub(r"\([^)]*\)", "", normalized)
    normalized = normalized.replace("年", "-").replace("月", "-").replace("日", "")
    normalized = normalized.replace("/", "-").replace(".", "-")
    normalized = re.sub(r"\s+", " ", normalized).strip(" -")
    try:
        dt = date_parser.parse(normalized)
    except (ValueError, TypeError, OverflowError):
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=ZoneInfo(default_tz))
    return dt


def dt_in_range(dt: datetime | None, start: datetime, end: datetime) -> bool:
    if dt is None:
        return False
    return start <= dt.astimezone(start.tzinfo) <= end


def abs_url(base_url: str, maybe_relative: str) -> str:
    return urllib.parse.urljoin(base_url, maybe_relative)


def guess_country(url: str) -> str:
    lowered = url.lower()
    if ".jp" in lowered or "ja-jp" in lowered or "aboutamazon.jp" in lowered:
        return "JP"
    return "Global"


def source_site(url: str) -> str:
    return urllib.parse.urlparse(url).netloc


def read_json_file(path: str, default: Any) -> Any:
    file_path = resolve_project_path(path)
    if not file_path.exists():
        return default
    return json.loads(file_path.read_text(encoding="utf-8"))


def load_site_credentials(path: str = DEFAULT_SITE_CREDENTIALS_PATH) -> dict[str, dict[str, Any]]:
    payload = read_json_file(path, {})
    if not isinstance(payload, dict):
        return {}
    normalized: dict[str, dict[str, Any]] = {}
    for domain, config in payload.items():
        if isinstance(config, dict):
            normalized[clean_text(domain).lower()] = config
    return normalized


def load_extra_source_entries(path: str = DEFAULT_EXTRA_SOURCES_PATH) -> list[ExtraSourceEntry]:
    payload = read_json_file(path, [])
    if isinstance(payload, dict):
        payload = payload.get("sources", [])
    if not isinstance(payload, list):
        return []

    entries: list[ExtraSourceEntry] = []
    seen: set[tuple[str, str, str]] = set()
    for row in payload:
        if not isinstance(row, dict) or row.get("active") is False:
            continue
        platform = clean_text(row.get("platform"))
        side = clean_text(row.get("side"))
        source_url = clean_text(row.get("source_url"))
        if not platform or side not in {"media", "buyer", "seller"} or not source_url:
            continue
        key = (platform, side, source_url)
        if key in seen:
            continue
        seen.add(key)
        entries.append(ExtraSourceEntry(platform=platform, side=side, source_url=source_url))
    return entries


def builtin_platform_lookup() -> dict[str, str]:
    mapping: dict[str, str] = {}
    for platform_id, platform_label in PLATFORM_LABELS.items():
        mapping[clean_text(platform_id).lower()] = platform_id
        mapping[clean_text(platform_label).lower()] = platform_id
        for alias in BUILTIN_PLATFORM_ALIASES.get(platform_id, []):
            mapping[clean_text(alias).lower()] = platform_id
    return mapping


def list_available_platform_labels(extra_sources_path: str = DEFAULT_EXTRA_SOURCES_PATH) -> list[str]:
    labels = list(PLATFORM_LABELS.values())
    builtin_lookup = builtin_platform_lookup()
    for entry in load_extra_source_entries(extra_sources_path):
        if builtin_lookup.get(clean_text(entry.platform).lower()):
            continue
        if entry.platform not in labels:
            labels.append(entry.platform)
    return labels


def resolve_requested_platforms(
    requested_platforms: list[str] | None,
    extra_source_entries: list[ExtraSourceEntry],
) -> tuple[list[str], dict[str, list[ExtraSourceEntry]], list[str]]:
    builtin_lookup = builtin_platform_lookup()
    extra_platform_lookup: dict[str, str] = {}
    extra_grouped: dict[str, list[ExtraSourceEntry]] = {}
    for entry in extra_source_entries:
        normalized_name = clean_text(entry.platform).lower()
        if builtin_lookup.get(normalized_name):
            continue
        extra_platform_lookup.setdefault(normalized_name, entry.platform)
        extra_grouped.setdefault(entry.platform, []).append(entry)

    requested = [clean_text(item) for item in (requested_platforms or []) if clean_text(item)]
    if not requested:
        return list(PLATFORM_LABELS.keys()), extra_grouped, []

    builtin_platforms: list[str] = []
    custom_platforms: dict[str, list[ExtraSourceEntry]] = {}
    unknown_platforms: list[str] = []

    for item in requested:
        normalized = item.lower()
        builtin_platform = builtin_lookup.get(normalized)
        if builtin_platform:
            if builtin_platform not in builtin_platforms:
                builtin_platforms.append(builtin_platform)
            continue

        custom_platform = extra_platform_lookup.get(normalized)
        if custom_platform:
            custom_platforms.setdefault(custom_platform, extra_grouped.get(custom_platform, []))
            continue

        unknown_platforms.append(item)

    return builtin_platforms, custom_platforms, unknown_platforms


def emit_progress(
    callback,
    *,
    stage: str,
    total_sites: int,
    completed_sites: int,
    active_sites: int,
    current_site: str | None = None,
    message: str | None = None,
) -> None:
    if callback is None:
        return
    callback(
        {
            "stage": stage,
            "total_sites": total_sites,
            "completed_sites": completed_sites,
            "active_sites": active_sites,
            "current_site": current_site or "",
            "message": message or "",
        }
    )


def find_site_credential(url: str, site_credentials: dict[str, dict[str, Any]]) -> tuple[str, dict[str, Any]] | None:
    domain = source_site(url).lower()
    if domain in site_credentials:
        return domain, site_credentials[domain]
    for candidate, config in site_credentials.items():
        normalized_candidate = clean_text(candidate).lower()
        if domain.endswith(f".{normalized_candidate}"):
            return normalized_candidate, config
    return None


def parse_cookie_header(cookie_text: str) -> dict[str, str]:
    cookies: dict[str, str] = {}
    for part in cookie_text.split(";"):
        if "=" not in part:
            continue
        name, value = part.split("=", 1)
        name = clean_text(name)
        if not name:
            continue
        cookies[name] = value.strip()
    return cookies


def apply_site_credentials(session: requests.Session, url: str, site_credentials: dict[str, dict[str, Any]]) -> bool:
    match = find_site_credential(url, site_credentials)
    if not match:
        return False

    domain, config = match
    if config.get("enabled") is False:
        return False

    headers = config.get("headers")
    if isinstance(headers, dict):
        session.headers.update({clean_text(str(k)): clean_text(str(v)) for k, v in headers.items() if clean_text(str(k))})

    cookies = config.get("cookies")
    if isinstance(cookies, dict):
        for name, value in cookies.items():
            cookie_name = clean_text(str(name))
            if not cookie_name:
                continue
            session.cookies.set(cookie_name, str(value), domain=domain)

    cookie_header = clean_text(config.get("cookie_header"))
    if cookie_header:
        for name, value in parse_cookie_header(cookie_header).items():
            session.cookies.set(name, value, domain=domain)

    auth_type = clean_text(config.get("auth_type")).lower()
    if auth_type == "basic" and clean_text(config.get("username")) and clean_text(config.get("password")):
        session.auth = (clean_text(config.get("username")), clean_text(config.get("password")))

    login_url = clean_text(config.get("login_url"))
    username = clean_text(config.get("username"))
    password = clean_text(config.get("password"))
    username_field = clean_text(config.get("username_field")) or "username"
    password_field = clean_text(config.get("password_field")) or "password"
    extra_form_fields = config.get("extra_form_fields")
    marker = f"{domain}|{login_url}|{username}"
    should_submit_form = (
        auth_type == "form"
        and login_url
        and username
        and password
        and marker not in session._credential_login_markers
    )
    if should_submit_form:
        payload = {username_field: username, password_field: password}
        if isinstance(extra_form_fields, dict):
            payload.update({clean_text(str(k)): str(v) for k, v in extra_form_fields.items() if clean_text(str(k))})
        try:
            session.get(login_url, timeout=30, allow_redirects=True)
            session.post(login_url, data=payload, timeout=30, allow_redirects=True)
            session._credential_login_markers.add(marker)
        except requests.RequestException:
            pass

    return True


def safe_slug(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug or "unknown"


def normalize_translation_target(target: str) -> str:
    normalized = clean_text(target).lower()
    return TRANSLATION_TARGETS.get(normalized, target)


def mostly_ascii(text: str) -> bool:
    text = clean_text(text)
    if not text:
        return True
    ascii_chars = sum(1 for ch in text if ord(ch) < 128)
    return ascii_chars / len(text) >= 0.85


def contains_japanese_kana(text: str) -> bool:
    text = clean_text(text)
    return any("\u3040" <= ch <= "\u30ff" for ch in text)


def mostly_chinese(text: str) -> bool:
    text = clean_text(text)
    if not text:
        return False
    if contains_japanese_kana(text):
        return False
    chinese_chars = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
    return chinese_chars / len(text) >= 0.3


def contains_chinese_chars(text: str) -> bool:
    text = clean_text(text)
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def trim_summary_text(text: str | None, max_chars: int = SUMMARY_MAX_CHARS) -> str | None:
    normalized = html_to_text(text)
    if not normalized:
        return None
    normalized = re.sub(r"\s+", " ", normalized).strip(" -|")
    for pattern in SUMMARY_NOISE_PATTERNS:
        if pattern.match(normalized):
            return None
    if len(normalized) <= max_chars:
        return normalized
    shortened = normalized[:max_chars].rstrip(" ,;:/-|")
    last_break = max(
        shortened.rfind("。"),
        shortened.rfind("！"),
        shortened.rfind("？"),
        shortened.rfind("."),
        shortened.rfind("!"),
        shortened.rfind("?"),
    )
    if last_break >= max_chars // 2:
        shortened = shortened[: last_break + 1]
    return shortened.strip() or None


def merge_summary_candidates(*values: str | None, max_chars: int = SUMMARY_MAX_CHARS) -> str | None:
    seen: list[str] = []
    for value in values:
        normalized = trim_summary_text(value, max_chars=max_chars)
        if not normalized or normalized in seen:
            continue
        seen.append(normalized)
    if not seen:
        return None
    if len(seen) == 1:
        return seen[0]
    merged = " ".join(seen)
    return trim_summary_text(merged, max_chars=max_chars)


class TextTranslator:
    def __init__(self, target_language: str) -> None:
        self.target_language = normalize_translation_target(target_language)
        self.enabled = GoogleTranslator is not None

    def should_skip(self, text: str) -> bool:
        if not text:
            return True
        if self.target_language == "en" and mostly_ascii(text):
            return True
        if self.target_language == "zh-CN" and mostly_chinese(text):
            return True
        return False

    def _translate_once(self, text: str) -> str | None:
        for attempt in range(3):
            translator = GoogleTranslator(source="auto", target=self.target_language)
            try:
                translated = clean_text(translator.translate(text))
                if translated:
                    return translated
                if attempt < 2:
                    time.sleep(0.6 * (attempt + 1))
                    continue
            except Exception:
                if attempt < 2:
                    time.sleep(0.6 * (attempt + 1))
                    continue
            break
        return None

    def _chunk_text(self, text: str, max_chars: int = 220) -> list[str]:
        normalized = clean_text(text)
        if len(normalized) <= max_chars:
            return [normalized]

        sentence_parts = [part.strip() for part in SUMMARY_SENTENCE_BREAK_RE.split(normalized) if clean_text(part)]
        if not sentence_parts:
            sentence_parts = [normalized]

        chunks: list[str] = []
        current = ""
        for part in sentence_parts:
            candidate = f"{current} {part}".strip() if current else part
            if current and len(candidate) > max_chars:
                chunks.append(current)
                current = part
                continue
            if len(part) > max_chars:
                if current:
                    chunks.append(current)
                    current = ""
                for index in range(0, len(part), max_chars):
                    piece = part[index:index + max_chars].strip()
                    if piece:
                        chunks.append(piece)
                continue
            current = candidate
        if current:
            chunks.append(current)
        return chunks or [normalized]

    def _translate_in_chunks(self, text: str) -> str | None:
        chunks = self._chunk_text(text)
        if len(chunks) <= 1:
            return None
        translated_chunks: list[str] = []
        changed = False
        for chunk in chunks:
            translated = clean_text(self._translate_once(chunk) or "")
            if not translated:
                return None
            if translated != chunk:
                changed = True
            translated_chunks.append(translated)
        if not changed:
            return None
        return clean_text(" ".join(translated_chunks))

    @lru_cache(maxsize=2048)
    def _translate_cached(self, text: str) -> str:
        if not self.enabled or self.should_skip(text):
            return text
        translated = clean_text(self._translate_once(text) or "")
        if translated and translated != text:
            return translated
        fallback = self._translate_in_chunks(text)
        if fallback and fallback != text:
            return fallback
        return text

    def translate(self, text: str | None) -> str | None:
        normalized = clean_text(text)
        if not normalized:
            return None
        return self._translate_cached(normalized)

    def needs_retry(self, original: str, translated: str | None) -> bool:
        normalized_original = clean_text(original)
        normalized_translated = clean_text(translated)
        if not normalized_original or self.should_skip(normalized_original):
            return False
        return not normalized_translated or normalized_translated == normalized_original


def add_translation_fields(
    rows: list[dict[str, Any]],
    translator: TextTranslator,
    progress_callback=None,
    extra_text_fields: list[str] | None = None,
) -> list[dict[str, Any]]:
    def prepare_summary(row: dict[str, Any]) -> str | None:
        normalized = merge_summary_candidates(
            row.get("summary"),
            row.get("body_excerpt"),
            max_chars=TRANSLATION_MAX_CHARS,
        )
        if not normalized:
            return None
        if len(normalized) <= TRANSLATION_MAX_CHARS:
            return normalized
        return normalized[:TRANSLATION_MAX_CHARS].rstrip() + "..."

    def split_extra_text(value: str | None) -> list[str]:
        normalized = clean_text(value)
        if not normalized:
            return []
        return [clean_text(part) for part in normalized.split(" | ") if clean_text(part)]

    def should_translate_extra_text(text: str) -> bool:
        normalized = clean_text(text)
        if not normalized:
            return False
        if translator.target_language == "zh-CN" and contains_chinese_chars(normalized):
            return False
        if translator.target_language == "en" and mostly_ascii(normalized):
            return False
        return True

    title_texts = {clean_text(row.get("title")) for row in rows if clean_text(row.get("title"))}
    summary_texts = {prepare_summary(row) for row in rows if prepare_summary(row)}
    extra_fields = [field for field in (extra_text_fields or []) if clean_text(field)]
    extra_texts = {
        part
        for row in rows
        for field in extra_fields
        for part in split_extra_text(row.get(field))
        if should_translate_extra_text(part)
    }
    translation_cache: dict[str, str | None] = {}
    translation_items = sorted(title_texts | summary_texts | extra_texts)
    total_items = len(translation_items)
    completed_items = 0
    retry_count = 0
    apply_count = 0
    progress_step = max(1, total_items // 20) if total_items else 1

    def emit_translation_progress(*, phase: str, last_text: str = '') -> None:
        if progress_callback is None or total_items <= 0:
            return
        if (
            completed_items <= 3
            or completed_items == total_items
            or completed_items % progress_step == 0
            or phase in {'retry', 'apply_done'}
        ):
            progress_callback(
                {
                    'phase': phase,
                    'completed': completed_items,
                    'total': total_items,
                    'retry_count': retry_count,
                    'applied_rows': apply_count,
                    'row_total': len(rows),
                    'last_text': clean_text(last_text)[:120],
                }
            )

    def translate_one(text: str) -> tuple[str, str | None]:
        return text, translator.translate(text)

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(translate_one, text): text for text in translation_items}
        for future in as_completed(futures):
            original, translated = future.result()
            translation_cache[original] = translated
            completed_items += 1
            emit_translation_progress(phase='translate', last_text=original)

    # Retry sequentially for texts that fell back to the original text during the concurrent pass.
    for original_text, translated_text in list(translation_cache.items()):
        if translator.needs_retry(original_text, translated_text):
            retry_count += 1
            translation_cache[original_text] = translator.translate(original_text)
            emit_translation_progress(phase='retry', last_text=original_text)

    translated_rows: list[dict[str, Any]] = []
    for row in rows:
        enriched = dict(row)
        enriched["translation_target"] = translator.target_language
        title = clean_text(row.get("title"))
        summary = prepare_summary(row)
        if summary:
            enriched["summary"] = summary
        enriched["title_translated"] = translation_cache.get(title) if title else None
        enriched["summary_translated"] = translation_cache.get(summary) if summary else None
        for field in extra_fields:
            field_text = clean_text(row.get(field))
            field_parts = split_extra_text(field_text)
            if field_parts:
                enriched[field] = " | ".join(
                    (translation_cache.get(part) or part) if should_translate_extra_text(part) else part
                    for part in field_parts
                )
        translated_rows.append(enriched)
        apply_count += 1
    emit_translation_progress(phase='apply_done')
    return translated_rows


def infer_nps_dimension(title: str, summary: str | None) -> str | None:
    haystack = f"{title} {summary or ''}".lower()
    for dimension, keywords in NPS_DIMENSION_RULES:
        if any(keyword.lower() in haystack for keyword in keywords):
            return dimension
    return None


def extract_meta_description(soup: BeautifulSoup) -> str:
    selectors = [
        ('meta[name="description"]', "content"),
        ('meta[property="og:description"]', "content"),
        ('meta[name="twitter:description"]', "content"),
    ]
    for selector, attr in selectors:
        node = soup.select_one(selector)
        if node and node.get(attr):
            summary = trim_summary_text(node[attr])
            if summary:
                return summary
    paragraph_nodes = soup.select("article p, main p, .article p, .content p, p, article li, main li")
    snippets: list[str] = []
    for node in paragraph_nodes:
        text = trim_summary_text(node.get_text(" ", strip=True))
        if not text or len(text) < SUMMARY_MIN_PARAGRAPH_CHARS:
            continue
        snippets.append(text)
        merged = merge_summary_candidates(*snippets)
        if merged and len(merged) >= min(180, SUMMARY_MAX_CHARS):
            return merged
    merged = merge_summary_candidates(*snippets)
    if merged:
        return merged
    return ""


def detect_source_status(url: str, status_code: int | None, final_url: str | None, text: str) -> tuple[str, str]:
    lowered_url = (final_url or url).lower()
    lowered_text = text.lower()
    seller_domains = [
        "seller.tiktokglobalshop.com",
        "sellercentral.amazon.co.jp",
        "qsm.qoo10.jp",
        "seller.shein.com",
    ]
    if status_code is None:
        return "error", "request failed"
    if status_code >= 400:
        return "http_error", f"http {status_code}"
    if "523 error" in lowered_text or "error occurred" in lowered_text and "qoo10" in lowered_url:
        return "blocked", "page returns an error body despite HTTP 200"
    if any(token in lowered_url for token in ["login", "signin", "sign-in", "auth"]):
        return "login_required", "final url indicates login flow"
    if any(domain in lowered_url for domain in seller_domains) and any(token in lowered_text for token in ["login", "signin", "ログイン"]):
        return "login_required", "seller-side page content indicates authentication"
    if any(
        token in lowered_text
        for token in [
            "type=\"password\"",
            "seller central",
            "seller center",
            "returnurl=",
        ]
    ):
        return "login_required", "page content indicates authentication"
    return "public", "page is publicly reachable"


def fetch(session: requests.Session, url: str) -> requests.Response:
    try:
        response = session.get(url, timeout=30, allow_redirects=True)
    except requests.exceptions.RequestException as exc:
        if getattr(session, "trust_env", True) is False or not should_retry_without_proxy(exc):
            raise
        direct_session = clone_session_without_env_proxy(session)
        response = direct_session.get(url, timeout=30, allow_redirects=True)
    if not response.encoding or response.encoding.lower() == "iso-8859-1":
        response.encoding = response.apparent_encoding or response.encoding
    return response


def audit_sources(session: requests.Session, platforms: list[str]) -> list[SourceAuditRecord]:
    records: list[SourceAuditRecord] = []
    for platform in platforms:
        for source_role, source_name, source_url in SOURCE_AUDIT_TARGETS[platform]:
            try:
                response = fetch(session, source_url)
                status, note = detect_source_status(source_url, response.status_code, response.url, response.text[:50000])
                records.append(
                    SourceAuditRecord(
                        platform=platform,
                        platform_label=PLATFORM_LABELS[platform],
                        source_role=source_role,
                        source_name=source_name,
                        source_url=source_url,
                        status=status,
                        http_status=response.status_code,
                        final_url=response.url,
                        note=note,
                    )
                )
            except requests.RequestException as exc:
                records.append(
                    SourceAuditRecord(
                        platform=platform,
                        platform_label=PLATFORM_LABELS[platform],
                        source_role=source_role,
                        source_name=source_name,
                        source_url=source_url,
                        status="error",
                        http_status=None,
                        final_url=None,
                        note=f"{type(exc).__name__}: {exc}",
                    )
                )
    return records


def make_custom_platform_id(platform_label: str) -> str:
    return f"custom_{safe_slug(platform_label)}"


def make_verified_candidate(
    *,
    platform_id: str,
    platform_label: str,
    source_name: str,
    source_url: str,
    article_url: str,
    title: str,
    published_at: str,
    category: str | None = None,
    summary: str | None = None,
) -> dict[str, Any]:
    normalized_title = clean_text(title)
    normalized_summary = clean_text(summary) or None
    return {
        "platform": platform_id,
        "platform_label": platform_label,
        "brand": platform_label,
        "source_name": source_name,
        "source_url": source_url,
        "source_site": source_site(source_url),
        "article_url": article_url,
        "title": normalized_title,
        "published_at": published_at,
        "category": clean_text(category) or None,
        "summary": normalized_summary,
        "nps_dimension_guess": infer_nps_dimension(normalized_title, normalized_summary),
        "source_country_guess": guess_country(article_url),
        "verification_status": "pending",
        "verification_checked_at": "",
        "verification_final_url": None,
    }


def crawl_custom_platform_source(
    entry: ExtraSourceEntry,
    start: datetime,
    end: datetime,
    adapter_configs_path: str,
    site_credentials: dict[str, dict[str, Any]],
) -> tuple[SourceAuditRecord, list[dict[str, Any]]]:
    import xlsx_source_test as xst

    platform_id = make_custom_platform_id(entry.platform)
    session = build_session()
    apply_site_credentials(session, entry.source_url, site_credentials)

    try:
        response = fetch(session, entry.source_url)
    except requests.RequestException as exc:
        return (
            SourceAuditRecord(
                platform=platform_id,
                platform_label=entry.platform,
                source_role=f"custom_{entry.side}",
                source_name=source_site(entry.source_url),
                source_url=entry.source_url,
                status="error",
                http_status=None,
                final_url=None,
                note=f"{type(exc).__name__}: {exc}",
            ),
            [],
        )

    status, note = detect_source_status(entry.source_url, response.status_code, response.url, response.text[:50000])
    if status != "public":
        return (
            SourceAuditRecord(
                platform=platform_id,
                platform_label=entry.platform,
                source_role=f"custom_{entry.side}",
                source_name=source_site(entry.source_url),
                source_url=entry.source_url,
                status=status,
                http_status=response.status_code,
                final_url=response.url,
                note=note,
            ),
            [],
        )

    adapter_configs = xst.load_adapter_configs(adapter_configs_path)
    source_entry = xst.SourceEntry(platform=entry.platform, side=entry.side, source_url=entry.source_url)
    raw_articles: list[dict[str, Any]] = []

    try:
        explicit_run = xst.run_explicit_media_adapter(source_entry, session, start, end, adapter_configs)
    except Exception as exc:
        explicit_run = None
        note = f"explicit adapter error: {type(exc).__name__}: {exc}"

    if explicit_run is not None and explicit_run.status == "ok" and explicit_run.articles:
        raw_articles = explicit_run.articles
        response_url = explicit_run.final_url or response.url
    else:
        response_url = response.url
        custom_articles = xst.custom_site_articles(response.url, session, start, end)
        if custom_articles is not None:
            raw_articles = custom_articles
        else:
            soup = BeautifulSoup(response.text, "lxml")
            for feed_url in xst.discover_feeds(soup, response.url):
                raw_articles.extend(xst.parse_feed_articles(feed_url, session, start, end))
                if raw_articles:
                    break
            if not raw_articles:
                for candidate_url in xst.discover_candidate_links(soup, response.url, 10):
                    article_meta = xst.extract_article_metadata(candidate_url, session, start, end)
                    if article_meta:
                        raw_articles.append(article_meta)

    deduped_articles: list[dict[str, Any]] = []
    seen_urls: set[str] = set()
    for article in raw_articles:
        article_url = clean_text(article.get("article_url"))
        title = clean_text(article.get("title"))
        published_at = clean_text(article.get("published_at"))
        if not article_url or not title or not published_at or article_url in seen_urls:
            continue
        seen_urls.add(article_url)
        deduped_articles.append(
            make_verified_candidate(
                platform_id=platform_id,
                platform_label=entry.platform,
                source_name=source_site(entry.source_url),
                source_url=response_url,
                article_url=article_url,
                title=title,
                published_at=published_at,
                category=article.get("category"),
                summary=article.get("summary"),
            )
        )

    audit_status = "ok" if deduped_articles else "public_no_recent_articles"
    audit_note = note if deduped_articles else "public page fetched but no recent articles found"
    return (
        SourceAuditRecord(
            platform=platform_id,
            platform_label=entry.platform,
            source_role=f"custom_{entry.side}",
            source_name=source_site(entry.source_url),
            source_url=entry.source_url,
            status=audit_status,
            http_status=response.status_code,
            final_url=response_url,
            note=audit_note,
        ),
        deduped_articles,
    )


def verify_article(session: requests.Session, item: dict[str, Any], checked_at: str) -> dict[str, Any] | None:
    try:
        response = fetch(session, item["article_url"])
    except requests.RequestException:
        return None

    status, _ = detect_source_status(item["article_url"], response.status_code, response.url, response.text[:5000])
    if status != "public":
        return None

    soup = BeautifulSoup(response.text, "lxml")
    summary = item.get("summary") or extract_meta_description(soup)
    item["summary"] = clean_text(summary) or None
    item["verification_status"] = "verified"
    item["verification_checked_at"] = checked_at
    item["verification_final_url"] = response.url
    if not item.get("nps_dimension_guess"):
        item["nps_dimension_guess"] = infer_nps_dimension(item["title"], item.get("summary"))
    return item


def crawl_tiktok_newsroom(session: requests.Session, start: datetime, end: datetime) -> list[dict[str, Any]]:
    response = fetch(session, "https://newsroom.tiktok.com/?lang=ja-JP")
    soup = BeautifulSoup(response.text, "lxml")
    for script in soup.find_all("script", {"type": "application/json"}):
        payload = script.get_text(strip=True)
        if not payload.startswith("%7B%22url%22"):
            continue
        data = json.loads(urllib.parse.unquote(payload))
        article = data["state"]["loaderData"]["routes/_app._index"]["mainArticle"]
        published_dt = parse_dt(article.get("publishedDate"), default_tz="UTC")
        if not dt_in_range(published_dt, start, end):
            return []
        article_url = f"https://newsroom.tiktok.com/{article['id']}?lang=ja-JP"
        return [
            {
                "platform": "tiktok_shop",
                "platform_label": PLATFORM_LABELS["tiktok_shop"],
                "brand": PLATFORM_LABELS["tiktok_shop"],
                "source_name": "TikTok Newsroom (JP)",
                "source_url": "https://newsroom.tiktok.com/?lang=ja-JP",
                "source_site": source_site("https://newsroom.tiktok.com/?lang=ja-JP"),
                "article_url": article_url,
                "title": clean_text(article.get("title")),
                "published_at": published_dt.isoformat(),
                "category": "Newsroom",
                "summary": clean_text(article.get("content", ""))[:280] or None,
                "nps_dimension_guess": infer_nps_dimension(article.get("title", ""), article.get("content", "")),
                "source_country_guess": "JP",
                "verification_status": "pending",
                "verification_checked_at": "",
                "verification_final_url": None,
            }
        ]
    return []


def crawl_amazon_news(session: requests.Session, start: datetime, end: datetime) -> list[dict[str, Any]]:
    response = fetch(session, "https://www.aboutamazon.jp/news")
    soup = BeautifulSoup(response.text, "lxml")
    articles: list[dict[str, Any]] = []
    seen: set[str] = set()
    for card in soup.select(".promo-card-v2"):
        title_node = card.select_one(".promo-card-v2__title a")
        date_node = card.select_one(".card-meta__published")
        category_node = card.select_one(".card-meta__category a")
        if not title_node or not date_node or not title_node.get("href"):
            continue
        article_url = abs_url(response.url, title_node["href"])
        if article_url in seen:
            continue
        seen.add(article_url)
        published_dt = parse_dt(clean_text(date_node.get_text(" ", strip=True)))
        if not dt_in_range(published_dt, start, end):
            continue
        title = clean_text(title_node.get_text(" ", strip=True))
        category = clean_text(category_node.get_text(" ", strip=True)) or None
        articles.append(
            {
                "platform": "amazon_japan",
                "platform_label": PLATFORM_LABELS["amazon_japan"],
                "brand": PLATFORM_LABELS["amazon_japan"],
                "source_name": "About Amazon Japan",
                "source_url": response.url,
                "source_site": source_site(response.url),
                "article_url": article_url,
                "title": title,
                "published_at": published_dt.isoformat(),
                "category": category,
                "summary": category,
                "nps_dimension_guess": infer_nps_dimension(title, category),
                "source_country_guess": "JP",
                "verification_status": "pending",
                "verification_checked_at": "",
                "verification_final_url": None,
            }
        )
    return articles


def crawl_rakuten_press(session: requests.Session, start: datetime, end: datetime) -> list[dict[str, Any]]:
    response = fetch(session, "https://corp.rakuten.co.jp/news/press/")
    soup = BeautifulSoup(response.text, "lxml")
    articles: list[dict[str, Any]] = []
    seen: set[str] = set()
    pattern = re.compile(r"/news/press/(\d{4})/(\d{2})(\d{2})_\d+\.html$")
    for anchor in soup.find_all("a", href=True):
        href = abs_url(response.url, anchor["href"])
        match = pattern.search(urllib.parse.urlparse(href).path)
        if not match or href in seen:
            continue
        seen.add(href)
        title = clean_text(anchor.get_text(" ", strip=True))
        if not title:
            continue
        year, month, day = match.groups()
        published_dt = datetime(int(year), int(month), int(day), 0, 0, tzinfo=ZoneInfo(TOKYO_TZ))
        if not dt_in_range(published_dt, start, end):
            continue
        articles.append(
            {
                "platform": "rakuten_ichiba",
                "platform_label": PLATFORM_LABELS["rakuten_ichiba"],
                "brand": PLATFORM_LABELS["rakuten_ichiba"],
                "source_name": "Rakuten Corporate Press",
                "source_url": response.url,
                "source_site": source_site(response.url),
                "article_url": href,
                "title": title,
                "published_at": published_dt.isoformat(),
                "category": "Press Release",
                "summary": None,
                "nps_dimension_guess": infer_nps_dimension(title, None),
                "source_country_guess": "JP",
                "verification_status": "pending",
                "verification_checked_at": "",
                "verification_final_url": None,
            }
        )
    return articles


def crawl_qoo10_feed(session: requests.Session, start: datetime, end: datetime) -> list[dict[str, Any]]:
    response = fetch(session, "https://article-university.qoo10.jp/feed")
    root = ET.fromstring(response.text)
    ns = {"atom": "http://www.w3.org/2005/Atom"}
    articles: list[dict[str, Any]] = []
    for entry in root.findall("atom:entry", ns):
        title = clean_text(entry.findtext("atom:title", default="", namespaces=ns))
        link_node = entry.find("atom:link", ns)
        if link_node is None or not link_node.get("href"):
            continue
        article_url = link_node.get("href")
        published_dt = parse_dt(entry.findtext("atom:published", default="", namespaces=ns))
        if published_dt is None or published_dt > datetime.now(published_dt.tzinfo):
            continue
        if not dt_in_range(published_dt, start, end):
            continue
        summary = html_to_text(entry.findtext("atom:summary", default="", namespaces=ns))
        articles.append(
            {
                "platform": "qoo10",
                "platform_label": PLATFORM_LABELS["qoo10"],
                "brand": PLATFORM_LABELS["qoo10"],
                "source_name": "Qoo10 University Feed",
                "source_url": response.url,
                "source_site": source_site(response.url),
                "article_url": article_url,
                "title": title,
                "published_at": published_dt.isoformat(),
                "category": "Seller University",
                "summary": summary or None,
                "nps_dimension_guess": infer_nps_dimension(title, summary),
                "source_country_guess": "JP",
                "verification_status": "pending",
                "verification_checked_at": "",
                "verification_final_url": None,
            }
        )
    return articles


def extract_temu_raw_data(html: str) -> dict[str, Any] | None:
    soup = BeautifulSoup(html, "lxml")
    for script in soup.find_all("script"):
        content = script.get_text()
        if "window.rawData=" not in content or ";document.dispatchEvent" not in content:
            continue
        start = content.index("window.rawData=") + len("window.rawData=")
        end = content.index(";document.dispatchEvent")
        return json.loads(content[start:end])
    return None


def crawl_temu_announcements(session: requests.Session, start: datetime, end: datetime) -> list[dict[str, Any]]:
    response = fetch(session, "https://www.temu.com/about_temu_home.html")
    raw_data = extract_temu_raw_data(response.text)
    if not raw_data:
        return []

    articles: list[dict[str, Any]] = []
    for entry in raw_data["store"].get("articleLists", []):
        article_url = abs_url(response.url, entry.get("customLink") or entry.get("seoLink") or entry.get("link", ""))
        if not article_url:
            continue
        try:
            detail_response = fetch(session, article_url)
        except requests.RequestException:
            continue
        detail_raw = extract_temu_raw_data(detail_response.text)
        if not detail_raw:
            continue
        detail = detail_raw["store"].get("detail", {})
        published_dt = parse_dt(detail.get("showTime"))
        if not dt_in_range(published_dt, start, end):
            continue
        title = clean_text(entry.get("title"))
        summary = clean_text(entry.get("briefText")) or None
        articles.append(
            {
                "platform": "temu",
                "platform_label": PLATFORM_LABELS["temu"],
                "brand": PLATFORM_LABELS["temu"],
                "source_name": "TEMU Announcements",
                "source_url": response.url,
                "source_site": source_site(response.url),
                "article_url": article_url,
                "title": title,
                "published_at": published_dt.isoformat(),
                "category": clean_text(entry.get("currentCategoryName")) or None,
                "summary": summary,
                "nps_dimension_guess": infer_nps_dimension(title, summary),
                "source_country_guess": guess_country(response.url),
                "verification_status": "pending",
                "verification_checked_at": "",
                "verification_final_url": None,
            }
        )
    return articles


def crawl_shein_newsroom(session: requests.Session, start: datetime, end: datetime) -> list[dict[str, Any]]:
    response = fetch(session, "https://www.sheingroup.com/newsroom")
    soup = BeautifulSoup(response.text, "lxml")
    articles: list[dict[str, Any]] = []
    seen: set[str] = set()
    for article in soup.find_all("article"):
        title_node = article.select_one("a.h4")
        meta_node = article.find(string=re.compile(r"—"))
        if not title_node or not title_node.get("href") or not meta_node:
            continue
        article_url = abs_url(response.url, title_node["href"])
        if article_url in seen:
            continue
        seen.add(article_url)
        meta_text = clean_text(str(meta_node))
        if "—" not in meta_text:
            continue
        category, date_text = [clean_text(part) for part in meta_text.split("—", 1)]
        published_dt = parse_dt(date_text)
        if not dt_in_range(published_dt, start, end):
            continue
        title = clean_text(title_node.get_text(" ", strip=True))
        articles.append(
            {
                "platform": "shein",
                "platform_label": PLATFORM_LABELS["shein"],
                "brand": PLATFORM_LABELS["shein"],
                "source_name": "SHEIN Group Newsroom",
                "source_url": response.url,
                "source_site": source_site(response.url),
                "article_url": article_url,
                "title": title,
                "published_at": published_dt.isoformat(),
                "category": category,
                "summary": None,
                "nps_dimension_guess": infer_nps_dimension(title, category),
                "source_country_guess": guess_country(response.url),
                "verification_status": "pending",
                "verification_checked_at": "",
                "verification_final_url": None,
            }
        )
    return articles


CRAWLERS = {
    "tiktok_shop": crawl_tiktok_newsroom,
    "amazon_japan": crawl_amazon_news,
    "rakuten_ichiba": crawl_rakuten_press,
    "qoo10": crawl_qoo10_feed,
    "temu": crawl_temu_announcements,
    "shein": crawl_shein_newsroom,
}


def dedupe_articles(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    deduped: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for item in items:
        key = (item["platform"], item["article_url"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)
    return deduped


def write_json(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def write_grouped_outputs(base_dir: Path, rows: list[dict[str, Any]]) -> None:
    by_brand_dir = base_dir / "by_brand"
    by_source_dir = base_dir / "by_source_site"
    by_brand_dir.mkdir(exist_ok=True)
    by_source_dir.mkdir(exist_ok=True)

    brand_groups: dict[str, list[dict[str, Any]]] = {}
    source_groups: dict[str, list[dict[str, Any]]] = {}

    for row in rows:
        brand_groups.setdefault(row["platform"], []).append(row)
        source_groups.setdefault(row["source_site"], []).append(row)

    for platform, items in brand_groups.items():
        brand_file_base = by_brand_dir / platform
        write_json(brand_file_base.with_suffix(".json"), items)
        write_csv(brand_file_base.with_suffix(".csv"), items)

    for site, items in source_groups.items():
        site_file_base = by_source_dir / safe_slug(site)
        write_json(site_file_base.with_suffix(".json"), items)
        write_csv(site_file_base.with_suffix(".csv"), items)


def main(argv: list[str] | None = None, progress_callback=None) -> int:
    args = parse_args(argv)
    start, end = parse_date_range(args)
    session = build_session()
    translator = TextTranslator(args.translate_to)
    extra_source_entries = load_extra_source_entries(args.extra_sources)
    builtin_platforms, custom_platforms, unknown_platforms = resolve_requested_platforms(args.platforms, extra_source_entries)
    site_credentials = load_site_credentials(args.site_credentials)
    checked_at = now_iso(args.timezone)
    total_sites = len(builtin_platforms) + sum(len(entries) for entries in custom_platforms.values())
    completed_sites = 0

    emit_progress(
        progress_callback,
        stage="setup",
        total_sites=total_sites,
        completed_sites=completed_sites,
        active_sites=0,
        current_site="",
        message="已完成参数解析，准备开始抓取网站",
    )

    audit_records = audit_sources(session, builtin_platforms)

    candidates: list[dict[str, Any]] = []
    for platform in builtin_platforms:
        crawler = CRAWLERS[platform]
        platform_label = PLATFORM_LABELS[platform]
        emit_progress(
            progress_callback,
            stage="crawl_site",
            total_sites=total_sites,
            completed_sites=completed_sites,
            active_sites=1,
            current_site=platform_label,
            message=f"正在抓取 {platform_label}",
        )
        try:
            candidates.extend(crawler(session, start, end))
        except Exception as exc:  # pragma: no cover - keeps the run resilient across site changes
            audit_records.append(
                SourceAuditRecord(
                    platform=platform,
                    platform_label=PLATFORM_LABELS[platform],
                    source_role="crawler",
                    source_name=crawler.__name__,
                    source_url="-",
                    status="error",
                    http_status=None,
                    final_url=None,
                    note=f"{type(exc).__name__}: {exc}",
                )
            )
        completed_sites += 1
        emit_progress(
            progress_callback,
            stage="crawl_site",
            total_sites=total_sites,
            completed_sites=completed_sites,
            active_sites=0,
            current_site=platform_label,
            message=f"已完成 {platform_label}，正在切换下一个网站",
        )

    for platform_label, entries in custom_platforms.items():
        if not entries:
            audit_records.append(
                SourceAuditRecord(
                    platform=make_custom_platform_id(platform_label),
                    platform_label=platform_label,
                    source_role="custom_media",
                    source_name=platform_label,
                    source_url="-",
                    status="error",
                    http_status=None,
                    final_url=None,
                    note="custom platform has no active source entries in the country extra_sources.json",
                )
            )
            continue
        for entry in entries:
            current_site_label = entry.source_url or platform_label
            emit_progress(
                progress_callback,
                stage="crawl_site",
                total_sites=total_sites,
                completed_sites=completed_sites,
                active_sites=1,
                current_site=current_site_label,
                message=f"正在抓取 {current_site_label}",
            )
            audit_record, custom_candidates = crawl_custom_platform_source(
                entry,
                start,
                end,
                args.adapter_configs,
                site_credentials,
            )
            audit_records.append(audit_record)
            candidates.extend(custom_candidates)
            completed_sites += 1
            emit_progress(
                progress_callback,
                stage="crawl_site",
                total_sites=total_sites,
                completed_sites=completed_sites,
                active_sites=0,
                current_site=current_site_label,
                message=f"已完成 {current_site_label}",
            )

    for raw_name in unknown_platforms:
        audit_records.append(
            SourceAuditRecord(
                platform=make_custom_platform_id(raw_name),
                platform_label=raw_name,
                source_role="custom_unknown",
                source_name=raw_name,
                source_url="-",
                status="error",
                http_status=None,
                final_url=None,
                note="unknown platform; add active source entries to the country extra_sources.json or use a built-in platform name",
            )
        )

    emit_progress(
        progress_callback,
        stage="verify_articles",
        total_sites=total_sites,
        completed_sites=completed_sites,
        active_sites=0,
        current_site="文章验证与输出整理",
        message="网站抓取已完成，正在校验文章链接并整理输出",
    )

    verified_items: list[dict[str, Any]] = []
    for item in dedupe_articles(candidates):
        verified = verify_article(session, item, checked_at)
        if verified:
            verified_items.append(verified)

    verified_items.sort(key=lambda row: row["published_at"], reverse=True)
    verified_items = add_translation_fields(verified_items, translator)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path(args.output_dir) / f"run_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    metadata = {
        "generated_at": checked_at,
        "timezone": args.timezone,
        "translation_target": translator.target_language,
        "range_start": start.isoformat(),
        "range_end": end.isoformat(),
        "mode": "raw_recent_news_no_content_filter",
        "platforms": [PLATFORM_LABELS[item] for item in builtin_platforms] + list(custom_platforms.keys()),
        "unknown_platforms": unknown_platforms,
        "candidate_count": len(candidates),
        "verified_count": len(verified_items),
    }

    write_json(output_dir / "metadata.json", metadata)
    write_json(output_dir / "source_audit.json", [asdict(record) for record in audit_records])
    write_json(output_dir / "articles.json", verified_items)
    write_csv(output_dir / "articles.csv", verified_items)
    write_grouped_outputs(output_dir, verified_items)

    emit_progress(
        progress_callback,
        stage="done",
        total_sites=total_sites,
        completed_sites=completed_sites,
        active_sites=0,
        current_site="",
        message="所有网站抓取与输出写入已完成",
    )

    print(f"输出目录: {output_dir}")
    print(f"候选新闻数: {len(candidates)}")
    print(f"已验证新闻数: {len(verified_items)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
