#!/usr/bin/env python3
from __future__ import annotations

import argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from bs4 import BeautifulSoup
from country_config import default_country_file_paths, resolve_project_path

from news_crawler import TOKYO_TZ, abs_url, build_session, clean_text, detect_source_status, fetch
from xlsx_source_test import (
    DEFAULT_ADAPTER_CONFIGS_PATH,
    DEFAULT_EXTRA_SOURCES_PATH,
    canonicalize_source_url,
    normalize_url,
    post_json_request_with_proxy_fallback,
    source_domain,
)


DEFAULT_QUERY_LOCALE = "jp"
GENERAL_MEDIA_PLATFORM = "General Media"
SEARCH_FIELD_NAMES = {"q", "query", "keyword", "keywords", "keys", "term", "p", "search_word"}
FEED_CONTENT_TYPES = {"application/rss+xml", "application/atom+xml", "application/xml", "text/xml"}
API_HINT_TOKENS = [
    "/api/",
    "api.",
    "graphql",
    "wp-json",
    "/rest/",
    "format=json",
    "output=json",
    ".json",
]
ANTI_BOT_HINT_TOKENS = [
    "captcha",
    "verify you are human",
    "access denied",
    "attention required",
    "cf-browser-verification",
    "cloudflare",
    "bot detection",
    "security check",
    "unusual traffic",
    "robot or human",
    "forbidden",
    "blocked",
]
MEDIA_PLATFORM_DISPLAY = "all brands·media"
DEFAULT_CAPABILITY_CACHE_PATH = default_country_file_paths("japan")["source_capability_cache_path"]
CAPABILITY_INSPECTION_TIMEOUT_SECONDS = 12


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="抓取网址管理系统：支持新增、停用、列表查看，并按需生成站点适配配置。")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--remove", action="store_true", help="停用一个已登记的网址")
    mode.add_argument("--reactivate", action="store_true", help="重新启用一个已停用的网址")
    mode.add_argument("--list", action="store_true", help="列出已登记的网址")
    parser.add_argument("url", nargs="?", help="要新增、停用或重新启用的抓取网址")
    parser.add_argument("--platform", help="所属品牌/平台名，例如 TikTok/TikTok Shop")
    parser.add_argument("--side", choices=["media", "buyer", "seller"], default="", help="来源侧别；新增时默认 media，删除时可选")
    parser.add_argument("--domain", default="", help="按域名过滤列表或停用目标")
    parser.add_argument("--extra-sources", default=DEFAULT_EXTRA_SOURCES_PATH, help="额外来源配置 JSON，默认使用国家子目录文件")
    parser.add_argument("--adapter-configs", default=DEFAULT_ADAPTER_CONFIGS_PATH, help="站点适配配置 JSON，默认使用国家子目录文件")
    parser.add_argument("--capability-cache", default=DEFAULT_CAPABILITY_CACHE_PATH, help="来源抓取能力缓存 JSON，默认使用国家子目录文件")
    parser.add_argument("--query-locale", choices=["jp", "global"], default=DEFAULT_QUERY_LOCALE, help="搜索词语言范围")
    parser.add_argument("--skip-api", action="store_true", help="只登记来源，不调用 API 生成适配器")
    parser.add_argument("--force-api", action="store_true", help="即使已有适配配置，也强制重新调用 API")
    parser.add_argument("--show-inactive", action="store_true", help="列表模式下显示已停用的网址")
    parser.add_argument("--api-url", default="", help="适配器 API 地址，默认读环境变量 CRAWLER_ADAPTER_API_URL")
    parser.add_argument("--api-key", default="", help="适配器 API Key，默认读环境变量 CRAWLER_ADAPTER_API_KEY")
    parser.add_argument("--api-model", default="", help="适配器模型名，默认读环境变量 CRAWLER_ADAPTER_API_MODEL")
    return parser.parse_args(argv)


def prompt_if_missing(value: str | None, prompt: str) -> str:
    normalized = clean_text(value)
    if normalized:
        return normalized
    return clean_text(input(prompt).strip())


def read_json_file(path: str, default: Any) -> Any:
    file_path = resolve_project_path(path)
    if not file_path.exists():
        return default
    return json.loads(file_path.read_text(encoding="utf-8"))


def write_json_file(path: str, payload: Any) -> None:
    file_path = resolve_project_path(path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def now_iso() -> str:
    return datetime.now().astimezone().isoformat()


def normalize_source_url(url: str) -> str:
    return canonicalize_source_url(normalize_url(url))


def normalize_registry_record(row: dict[str, Any]) -> dict[str, Any] | None:
    if not isinstance(row, dict):
        return None
    platform = clean_text(row.get("platform"))
    side = clean_text(row.get("side"))
    source_url = normalize_source_url(str(row.get("source_url") or ""))
    if side == "media" and not platform:
        platform = GENERAL_MEDIA_PLATFORM
    if not platform or side not in {"media", "buyer", "seller"} or not source_url:
        return None
    active = row.get("active")
    return {
        "platform": platform,
        "side": side,
        "source_url": source_url,
        "domain": clean_text(row.get("domain")) or source_domain(source_url),
        "active": True if active is None else bool(active),
        "created_at": clean_text(row.get("created_at")),
        "updated_at": clean_text(row.get("updated_at")),
        "deactivated_at": clean_text(row.get("deactivated_at")),
        "reactivated_at": clean_text(row.get("reactivated_at")),
    }


def display_platform(platform: str, side: str) -> str:
    if clean_text(side) == "media":
        return MEDIA_PLATFORM_DISPLAY
    return clean_text(platform)


def capability_cache_key(source_url: str) -> str:
    return normalize_source_url(source_url)


def default_capability_snapshot() -> dict[str, Any]:
    return {
        "status": "",
        "note": "",
        "http_status": None,
        "final_url": "",
        "domain": "",
        "adapter_configured": False,
        "feed_candidates": [],
        "searchable_form_count": 0,
        "search_links": [],
        "api_candidates": [],
        "anti_bot_hints": [],
        "capability_tags": ["尚未评估"],
        "completeness_risk": "medium",
        "completeness_reason": "尚未生成网站评估结果，先按中风险显示；会在新闻抓取成功完成后刷新。",
        "evaluated_at": "",
        "capability_error": "",
    }


def read_capability_cache(path: str = DEFAULT_CAPABILITY_CACHE_PATH) -> dict[str, dict[str, Any]]:
    payload = read_json_file(path, {})
    if isinstance(payload, dict) and isinstance(payload.get("entries"), dict):
        payload = payload.get("entries", {})
    if not isinstance(payload, dict):
        return {}
    cache: dict[str, dict[str, Any]] = {}
    for key, value in payload.items():
        if isinstance(key, str) and isinstance(value, dict):
            cache[key] = value
    return cache


def write_capability_cache(entries: dict[str, dict[str, Any]], path: str = DEFAULT_CAPABILITY_CACHE_PATH) -> None:
    write_json_file(
        path,
        {
            "version": 1,
            "updated_at": now_iso(),
            "entries": entries,
        },
    )


def infer_capability_risk(
    record: dict[str, Any],
    snapshot: dict[str, Any],
    *,
    adapter_configs_path: str = DEFAULT_ADAPTER_CONFIGS_PATH,
) -> tuple[str, str]:
    status = clean_text(snapshot.get("status")).lower()
    note = clean_text(snapshot.get("note")).lower()
    capability_error = clean_text(snapshot.get("capability_error")).lower()
    adapter_configured = bool(snapshot.get("adapter_configured"))
    feed_count = len(snapshot.get("feed_candidates") or [])
    search_count = int(snapshot.get("searchable_form_count") or 0) + len(snapshot.get("search_links") or [])
    api_count = len(snapshot.get("api_candidates") or [])
    anti_bot_count = len(snapshot.get("anti_bot_hints") or [])
    http_status = int(snapshot.get("http_status") or 0)

    domain = clean_text(snapshot.get("domain")) or clean_text(record.get("domain"))
    if domain and not adapter_configured:
        adapter_configured = adapter_exists_for_domain(adapter_configs_path, domain)

    if status in {"blocked", "login_required"}:
        return "high", "站点当前对脚本访问不友好，现有抓取结果较可能不完整。"
    if adapter_configured or feed_count > 0 or search_count > 0:
        return "low", "站点存在较稳定的公开抓取入口，当前抓取完整性相对更高。"
    if api_count > 0:
        return "medium", "站点疑似存在 JSON/API 接口，但当前未必已接入，可能仍有遗漏。"
    if anti_bot_count > 0:
        return "high", "站点存在反爬或风控迹象，现有抓取结果较可能不完整。"
    if http_status >= 400 or "ssl" in capability_error or "timeout" in capability_error or "connection" in capability_error:
        return "medium", "站点评估请求不稳定或返回异常，暂按中风险显示，后续抓取会继续刷新。"
    if "error" in note or "forbidden" in note or "blocked" in note:
        return "high", "站点访问反馈异常，现有抓取结果较可能不完整。"
    return "medium", "站点可访问，但没有发现明显 RSS、公开搜索或稳定 API 入口，可能只能抓到部分新闻。"


def normalize_capability_snapshot(
    record: dict[str, Any],
    snapshot: dict[str, Any] | None,
    *,
    adapter_configs_path: str = DEFAULT_ADAPTER_CONFIGS_PATH,
) -> dict[str, Any]:
    base = default_capability_snapshot()
    if isinstance(snapshot, dict):
        base.update(snapshot)
    inferred_risk, inferred_reason = infer_capability_risk(record, base, adapter_configs_path=adapter_configs_path)
    risk = clean_text(base.get("completeness_risk")).lower()
    if risk not in {"low", "medium", "high"}:
        base["completeness_risk"] = inferred_risk
    if not clean_text(base.get("completeness_reason")) or clean_text(base.get("completeness_risk")).lower() == "unknown":
        base["completeness_reason"] = inferred_reason
    elif clean_text(base.get("completeness_reason")) == default_capability_snapshot()["completeness_reason"]:
        base["completeness_reason"] = inferred_reason
    return base


def merge_capability_snapshot(
    record: dict[str, Any],
    snapshot: dict[str, Any] | None,
    *,
    adapter_configs_path: str = DEFAULT_ADAPTER_CONFIGS_PATH,
) -> dict[str, Any]:
    merged = dict(record)
    base = normalize_capability_snapshot(record, snapshot, adapter_configs_path=adapter_configs_path)
    merged.update(
        {
            "capability_tags": base.get("capability_tags") or ["尚未评估"],
            "completeness_risk": base.get("completeness_risk") or "medium",
            "completeness_reason": base.get("completeness_reason") or default_capability_snapshot()["completeness_reason"],
            "access_status": base.get("status") or "",
            "http_status": base.get("http_status"),
            "final_url": base.get("final_url") or "",
            "adapter_configured": bool(base.get("adapter_configured")),
            "feed_candidates": base.get("feed_candidates") or [],
            "searchable_form_count": int(base.get("searchable_form_count") or 0),
            "search_links": base.get("search_links") or [],
            "api_candidates": base.get("api_candidates") or [],
            "anti_bot_hints": base.get("anti_bot_hints") or [],
            "access_note": base.get("note") or "",
            "capability_error": base.get("capability_error") or "",
            "capability_evaluated_at": base.get("evaluated_at") or "",
        }
    )
    return merged


def load_source_registry(path: str) -> list[dict[str, Any]]:
    payload = read_json_file(path, [])
    if isinstance(payload, dict):
        payload = payload.get("sources", [])
    if not isinstance(payload, list):
        return []
    records: list[dict[str, Any]] = []
    for row in payload:
        normalized = normalize_registry_record(row)
        if normalized is not None:
            records.append(normalized)
    return records


def save_source_registry(path: str, records: list[dict[str, Any]]) -> None:
    write_json_file(path, records)


def find_exact_record(records: list[dict[str, Any]], platform: str, side: str, source_url: str) -> dict[str, Any] | None:
    for record in records:
        if (
            record["platform"] == platform
            and record["side"] == side
            and record["source_url"] == source_url
        ):
            return record
    return None


def add_or_reactivate_source(path: str, platform: str, side: str, source_url: str) -> tuple[str, dict[str, Any]]:
    platform = clean_text(platform)
    side = clean_text(side) or "media"
    if side == "media" and not platform:
        platform = GENERAL_MEDIA_PLATFORM
    records = load_source_registry(path)
    record = find_exact_record(records, platform, side, source_url)
    current_time = now_iso()
    if record is not None:
        if record.get("active", True):
            return "already_active", record
        record["active"] = True
        record["updated_at"] = current_time
        record["reactivated_at"] = current_time
        record["deactivated_at"] = ""
        save_source_registry(path, records)
        return "reactivated", record

    record = {
        "platform": platform,
        "side": side,
        "source_url": source_url,
        "domain": source_domain(source_url),
        "active": True,
        "created_at": current_time,
        "updated_at": current_time,
        "deactivated_at": "",
        "reactivated_at": "",
    }
    records.append(record)
    save_source_registry(path, records)
    return "added", record


def deactivate_sources(
    path: str,
    source_url: str | None = None,
    platform: str | None = None,
    side: str | None = None,
    domain: str | None = None,
) -> list[dict[str, Any]]:
    records = load_source_registry(path)
    current_time = now_iso()
    changed: list[dict[str, Any]] = []
    for record in records:
        if source_url and record["source_url"] != source_url:
            continue
        if platform and record["platform"] != platform:
            continue
        if side and record["side"] != side:
            continue
        if domain and record["domain"] != domain:
            continue
        if not record.get("active", True):
            continue
        record["active"] = False
        record["updated_at"] = current_time
        record["deactivated_at"] = current_time
        changed.append(record)
    if changed:
        save_source_registry(path, records)
    return changed


def reactivate_sources(
    path: str,
    source_url: str | None = None,
    platform: str | None = None,
    side: str | None = None,
    domain: str | None = None,
) -> list[dict[str, Any]]:
    records = load_source_registry(path)
    current_time = now_iso()
    changed: list[dict[str, Any]] = []
    for record in records:
        if source_url and record["source_url"] != source_url:
            continue
        if platform and record["platform"] != platform:
            continue
        if side and record["side"] != side:
            continue
        if domain and record["domain"] != domain:
            continue
        if record.get("active", True):
            continue
        record["active"] = True
        record["updated_at"] = current_time
        record["reactivated_at"] = current_time
        record["deactivated_at"] = ""
        changed.append(record)
    if changed:
        save_source_registry(path, records)
    return changed


def list_sources(
    path: str,
    show_inactive: bool,
    *,
    source_url: str | None = None,
    platform: str | None = None,
    side: str | None = None,
    domain: str | None = None,
) -> list[dict[str, Any]]:
    records = load_source_registry(path)
    if not show_inactive:
        records = [record for record in records if record.get("active", True)]
    if source_url:
        records = [record for record in records if record["source_url"] == source_url]
    if platform:
        records = [record for record in records if record["platform"] == platform]
    if side:
        records = [record for record in records if record["side"] == side]
    if domain:
        records = [record for record in records if record["domain"] == domain]
    records.sort(key=lambda row: (row["platform"], row["side"], row["source_url"]))
    return records


def adapter_exists_for_domain(path: str, domain: str) -> bool:
    payload = read_json_file(path, {})
    if not isinstance(payload, dict):
        return False
    normalized_domain = clean_text(domain).lower()
    if normalized_domain in payload:
        return True
    return any(normalized_domain.endswith(f".{clean_text(key).lower()}") for key in payload)


def summarize_form(form, base_url: str) -> dict[str, Any]:
    fields: list[dict[str, str]] = []
    query_field = ""
    for inp in form.find_all(["input", "textarea", "select"]):
        field_name = clean_text(inp.get("name"))
        field_type = clean_text(inp.get("type") or inp.name or "")
        fields.append({"name": field_name, "type": field_type})
        if field_type == "search" or field_name.lower() in SEARCH_FIELD_NAMES:
            query_field = field_name
    return {
        "action": abs_url(base_url, form.get("action") or ""),
        "method": clean_text(form.get("method") or "get").lower(),
        "query_field": query_field,
        "fields": fields[:12],
    }


def discover_site_context(url: str, query_locale: str) -> dict[str, Any]:
    session = build_session()
    response = fetch(session, url)
    soup = BeautifulSoup(response.text, "lxml")

    forms = [summarize_form(form, response.url) for form in soup.find_all("form")[:12]]
    search_links: list[str] = []
    for anchor in soup.find_all("a", href=True):
        href = abs_url(response.url, anchor["href"])
        lowered = href.lower()
        if "search" in lowered and href not in search_links:
            search_links.append(href)
        if len(search_links) >= 10:
            break

    sample_query = "楽天市場" if query_locale == "jp" else "Rakuten"
    candidate_urls: list[str] = []
    for form in forms:
        action = clean_text(form.get("action"))
        query_field = clean_text(form.get("query_field"))
        method = clean_text(form.get("method"))
        if method == "get" and action and query_field:
            encoded = requests.utils.quote(sample_query, safe="")
            separator = "&" if "?" in action else "?"
            candidate = f"{action}{separator}{query_field}={encoded}"
            if candidate not in candidate_urls:
                candidate_urls.append(candidate)
    for href in search_links:
        if href not in candidate_urls:
            candidate_urls.append(href)
    candidate_urls = candidate_urls[:5]

    samples: list[dict[str, Any]] = []
    for candidate_url in candidate_urls:
        try:
            candidate_response = fetch(session, candidate_url)
        except Exception as exc:
            samples.append(
                {
                    "url": candidate_url,
                    "status_code": None,
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )
            continue
        candidate_soup = BeautifulSoup(candidate_response.text, "lxml")
        anchors: list[dict[str, str]] = []
        for anchor in candidate_soup.find_all("a", href=True):
            text = clean_text(anchor.get_text(" ", strip=True))
            href = abs_url(candidate_response.url, anchor["href"])
            if len(text) < 10:
                continue
            anchors.append({"href": href, "text": text[:200]})
            if len(anchors) >= 8:
                break
        samples.append(
            {
                "url": candidate_response.url,
                "status_code": candidate_response.status_code,
                "title": clean_text(candidate_soup.title.get_text(" ", strip=True)) if candidate_soup.title else "",
                "forms": [summarize_form(form, candidate_response.url) for form in candidate_soup.find_all("form")[:6]],
                "anchors": anchors,
                "html_excerpt": candidate_response.text[:6000],
            }
        )

    return {
        "source_url": response.url,
        "domain": source_domain(response.url),
        "homepage_title": clean_text(soup.title.get_text(" ", strip=True)) if soup.title else "",
        "homepage_forms": forms,
        "homepage_search_links": search_links,
        "sample_query": sample_query,
        "candidate_search_samples": samples,
    }


def build_adapter_prompt(context: dict[str, Any]) -> list[dict[str, str]]:
    schema = {
        "note": "api-generated adapter",
        "query_locale": "jp",
        "search_terms": [],
        "search_url_template": "https://example.com/search?q={query}&page={page}",
        "page_mode": "single",
        "offset_step": 10,
        "max_pages": 1,
        "item_selector": "article",
        "link_selector": "a[href]",
        "link_attr": "href",
        "link_pattern": "/article/",
        "title_selector": "h2 a",
        "title_attr": "",
        "date_selector": "time",
        "date_attr": "datetime",
        "summary_selector": "p",
        "summary_attr": "",
        "allow_search_only": True,
        "source_discovery": "configured_search",
    }
    return [
        {
            "role": "system",
            "content": (
                "You are a site-adapter generator for a news crawler. "
                "Return only one JSON object matching the requested schema. "
                "Prefer stable CSS selectors and a direct public search URL. "
                "If the site is blocked or no public search page exists, still return a JSON object with "
                "`note` explaining the limitation and leave selectors empty."
            ),
        },
        {
            "role": "user",
            "content": (
                "Generate one adapter config JSON for this news site.\n"
                f"Target domain: {context['domain']}\n"
                f"Schema example: {json.dumps(schema, ensure_ascii=False)}\n"
                "Rules:\n"
                "- `search_url_template` must include `{query}` and may include `{page}` or `{offset}`.\n"
                "- Use `page_mode` = `single`, `page`, or `offset`.\n"
                "- If a search result page shows date in text only, use `date_attr` as empty string.\n"
                "- If detail pages are publicly readable, keep `allow_search_only` as false unless necessary.\n"
                "- Return only JSON, no markdown.\n"
                f"Observed context:\n{json.dumps(context, ensure_ascii=False)}"
            ),
        },
    ]


def call_adapter_api(messages: list[dict[str, str]], api_url: str, api_key: str, api_model: str) -> dict[str, Any]:
    if not api_url or not api_key or not api_model:
        raise RuntimeError("adapter API credentials are incomplete")
    payload = {
        "model": api_model,
        "response_format": {"type": "json_object"},
        "temperature": 0.1,
        "messages": messages,
    }
    response = post_json_request_with_proxy_fallback(
        api_url,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json_payload=payload,
        timeout=120,
    )
    response.raise_for_status()
    body = response.json()
    content = body["choices"][0]["message"]["content"]
    if isinstance(content, list):
        content = "".join(part.get("text", "") for part in content if isinstance(part, dict))
    data = json.loads(content)
    if not isinstance(data, dict):
        raise RuntimeError("adapter API did not return a JSON object")
    return data


def validate_adapter_config(config: dict[str, Any]) -> dict[str, Any]:
    validated = dict(config)
    validated["note"] = clean_text(validated.get("note")) or "api-generated adapter"
    validated["query_locale"] = clean_text(validated.get("query_locale")) or DEFAULT_QUERY_LOCALE
    validated["search_url_template"] = clean_text(validated.get("search_url_template"))
    validated["page_mode"] = clean_text(validated.get("page_mode")) or "single"
    validated["item_selector"] = clean_text(validated.get("item_selector"))
    validated["link_selector"] = clean_text(validated.get("link_selector"))
    validated["link_attr"] = clean_text(validated.get("link_attr")) or "href"
    validated["link_pattern"] = clean_text(validated.get("link_pattern"))
    validated["title_selector"] = clean_text(validated.get("title_selector"))
    validated["title_attr"] = clean_text(validated.get("title_attr"))
    validated["date_selector"] = clean_text(validated.get("date_selector"))
    validated["date_attr"] = clean_text(validated.get("date_attr"))
    validated["summary_selector"] = clean_text(validated.get("summary_selector"))
    validated["summary_attr"] = clean_text(validated.get("summary_attr"))
    validated["source_discovery"] = clean_text(validated.get("source_discovery")) or "configured_search"
    validated["max_pages"] = int(validated.get("max_pages", 1) or 1)
    validated["offset_step"] = int(validated.get("offset_step", 10) or 10)
    validated["allow_search_only"] = bool(validated.get("allow_search_only", False))
    search_terms = validated.get("search_terms")
    if not isinstance(search_terms, list):
        validated["search_terms"] = []
    else:
        validated["search_terms"] = [clean_text(str(item)) for item in search_terms if clean_text(str(item))]
    if "{query}" not in validated["search_url_template"]:
        raise RuntimeError("generated adapter is missing `{query}` in search_url_template")
    return validated


def save_adapter_config(path: str, domain: str, config: dict[str, Any]) -> None:
    payload = read_json_file(path, {})
    if not isinstance(payload, dict):
        payload = {}
    payload[domain] = config
    write_json_file(path, payload)


def resolve_api_credentials(args: argparse.Namespace) -> tuple[str, str, str]:
    api_url = clean_text(args.api_url) or clean_text(os.environ.get("CRAWLER_ADAPTER_API_URL"))
    api_key = clean_text(args.api_key) or clean_text(os.environ.get("CRAWLER_ADAPTER_API_KEY"))
    api_model = clean_text(args.api_model) or clean_text(os.environ.get("CRAWLER_ADAPTER_API_MODEL"))
    return api_url, api_key, api_model


def print_source_records(records: list[dict[str, Any]]) -> None:
    if not records:
        print("当前没有符合条件的网址记录。")
        return
    for record in records:
        status = "active" if record.get("active", True) else "inactive"
        print(f"[{status}] {display_platform(record['platform'], record['side'])} | {record['side']} | {record['source_url']}")
        print(f"  domain: {record['domain']}")
        if record.get("created_at"):
            print(f"  created_at: {record['created_at']}")
        if record.get("reactivated_at"):
            print(f"  reactivated_at: {record['reactivated_at']}")
        if record.get("deactivated_at"):
            print(f"  deactivated_at: {record['deactivated_at']}")


def print_cached_capability_report(
    record: dict[str, Any],
    snapshot: dict[str, Any] | None,
    *,
    adapter_configs_path: str = DEFAULT_ADAPTER_CONFIGS_PATH,
) -> None:
    merged = merge_capability_snapshot(record, snapshot, adapter_configs_path=adapter_configs_path)
    print(f"- {record['source_url']}")
    print(f"  access_status: {merged.get('access_status') or 'unknown'} | http_status: {merged.get('http_status') or 'n/a'}")
    if merged.get("final_url"):
        print(f"  final_url: {merged['final_url']}")
    print(f"  capability_tags: {' | '.join(merged.get('capability_tags') or []) or '无'}")
    print(f"  completeness_risk: {merged.get('completeness_risk') or 'unknown'}")
    print(f"  completeness_reason: {merged.get('completeness_reason') or default_capability_snapshot()['completeness_reason']}")
    if merged.get("capability_evaluated_at"):
        print(f"  evaluated_at: {merged['capability_evaluated_at']}")


def detect_feed_candidates(soup: BeautifulSoup, base_url: str) -> list[str]:
    feeds: list[str] = []
    seen: set[str] = set()

    for link in soup.find_all("link", href=True):
        link_type = clean_text(link.get("type", "")).lower()
        rels = {clean_text(rel).lower() for rel in link.get("rel", [])}
        if "alternate" in rels and link_type in FEED_CONTENT_TYPES:
            href = abs_url(base_url, link["href"])
            if href not in seen:
                seen.add(href)
                feeds.append(href)

    for anchor in soup.find_all("a", href=True):
        href = abs_url(base_url, anchor["href"])
        lowered = href.lower()
        if any(token in lowered for token in ["/feed", "/rss", "atom.xml", "/atom"]) and href not in seen:
            seen.add(href)
            feeds.append(href)
        if len(feeds) >= 8:
            break
    return feeds[:8]


def detect_search_capabilities(soup: BeautifulSoup, base_url: str) -> tuple[int, list[str]]:
    forms = [summarize_form(form, base_url) for form in soup.find_all("form")]
    searchable_forms = [form for form in forms if clean_text(form.get("query_field"))]
    search_links: list[str] = []
    for anchor in soup.find_all("a", href=True):
        href = abs_url(base_url, anchor["href"])
        lowered = href.lower()
        if any(token in lowered for token in ["search", "keyword", "query"]) and href not in search_links:
            search_links.append(href)
        if len(search_links) >= 8:
            break
    return len(searchable_forms), search_links[:8]


def detect_api_candidates(soup: BeautifulSoup, base_url: str) -> list[str]:
    endpoints: list[str] = []
    seen: set[str] = set()
    for tag_name, attr_name in [("script", "src"), ("a", "href"), ("link", "href"), ("form", "action")]:
        for node in soup.find_all(tag_name):
            raw_value = clean_text(node.get(attr_name))
            if not raw_value:
                continue
            absolute_value = abs_url(base_url, raw_value)
            lowered = absolute_value.lower()
            if any(token in lowered for token in API_HINT_TOKENS):
                if absolute_value not in seen:
                    seen.add(absolute_value)
                    endpoints.append(absolute_value)
            if len(endpoints) >= 8:
                break
        if len(endpoints) >= 8:
            break
    return endpoints[:8]


def detect_anti_bot_hints(text: str) -> list[str]:
    lowered = clean_text(text).lower()
    return [token for token in ANTI_BOT_HINT_TOKENS if token in lowered]


def inspect_source_capabilities(source_url: str, adapter_configs_path: str) -> dict[str, Any]:
    normalized_url = normalize_source_url(source_url)
    session = build_session()
    response = session.get(normalized_url, timeout=CAPABILITY_INSPECTION_TIMEOUT_SECONDS, allow_redirects=True)
    if not response.encoding or response.encoding.lower() == "iso-8859-1":
        response.encoding = response.apparent_encoding or response.encoding
    final_url = response.url
    domain = source_domain(final_url)
    status, note = detect_source_status(normalized_url, response.status_code, final_url, response.text[:50000])
    adapter_configured = adapter_exists_for_domain(adapter_configs_path, domain)
    soup = BeautifulSoup(response.text, "lxml")

    feed_candidates = detect_feed_candidates(soup, final_url)
    searchable_form_count, search_links = detect_search_capabilities(soup, final_url)
    api_candidates = detect_api_candidates(soup, final_url)
    anti_bot_hints = detect_anti_bot_hints(response.text[:50000])

    capability_tags: list[str] = []
    if adapter_configured:
        capability_tags.append("已配置适配器")
    if feed_candidates:
        capability_tags.append("发现 RSS/Atom")
    if searchable_form_count or search_links:
        capability_tags.append("发现公开搜索入口")
    if api_candidates:
        capability_tags.append("疑似 JSON/API/GraphQL")
    if anti_bot_hints:
        capability_tags.append("疑似反爬/风控")
    if status == "login_required":
        capability_tags.append("需要登录")
    elif status == "blocked":
        capability_tags.append("脚本访问受阻")
    elif status == "public" and not capability_tags:
        capability_tags.append("仅发现普通网页入口")

    if status in {"blocked", "login_required"}:
        completeness_risk = "high"
        completeness_reason = "站点当前对脚本访问不友好，现有抓取结果较可能不完整。"
    elif adapter_configured or feed_candidates or searchable_form_count or search_links:
        completeness_risk = "low"
        completeness_reason = "站点存在较稳定的公开抓取入口，当前抓取完整性相对更高。"
    elif api_candidates:
        completeness_risk = "medium"
        completeness_reason = "站点疑似存在 JSON/API 接口，但当前未必已接入，可能仍有遗漏。"
    else:
        completeness_risk = "medium"
        completeness_reason = "站点可访问，但没有发现明显 RSS/公开搜索/API 入口，可能只能抓到部分新闻。"

    return {
        "status": status,
        "note": note,
        "http_status": response.status_code,
        "final_url": final_url,
        "domain": domain,
        "adapter_configured": adapter_configured,
        "feed_candidates": feed_candidates,
        "searchable_form_count": searchable_form_count,
        "search_links": search_links,
        "api_candidates": api_candidates,
        "anti_bot_hints": anti_bot_hints,
        "capability_tags": capability_tags,
        "completeness_risk": completeness_risk,
        "completeness_reason": completeness_reason,
    }


def refresh_capability_cache_for_records(
    records: list[dict[str, Any]],
    *,
    adapter_configs_path: str = DEFAULT_ADAPTER_CONFIGS_PATH,
    cache_path: str = DEFAULT_CAPABILITY_CACHE_PATH,
) -> dict[str, dict[str, Any]]:
    if not records:
        return read_capability_cache(cache_path)

    cache = read_capability_cache(cache_path)
    url_to_record: dict[str, dict[str, Any]] = {}
    for record in records:
        source_url = clean_text(record.get("source_url"))
        if source_url:
            url_to_record.setdefault(capability_cache_key(source_url), record)

    def inspect_one(cache_key: str, record: dict[str, Any]) -> tuple[str, dict[str, Any]]:
        try:
            report = inspect_source_capabilities(record["source_url"], adapter_configs_path)
            snapshot = default_capability_snapshot()
            snapshot.update(report)
            snapshot["evaluated_at"] = now_iso()
            snapshot["capability_error"] = ""
            return cache_key, normalize_capability_snapshot(record, snapshot, adapter_configs_path=adapter_configs_path)
        except Exception as exc:
            previous_snapshot = cache.get(cache_key) if isinstance(cache.get(cache_key), dict) else {}
            snapshot = default_capability_snapshot()
            if isinstance(previous_snapshot, dict):
                snapshot.update(previous_snapshot)
            previous_tags = [str(tag) for tag in snapshot.get("capability_tags") or [] if str(tag).strip()]
            merged_tags = []
            for tag in previous_tags + ["评估失败"]:
                if tag not in merged_tags:
                    merged_tags.append(tag)
            snapshot.update(
                {
                    "capability_tags": merged_tags or ["评估失败"],
                    "capability_error": f"{type(exc).__name__}: {exc}",
                    "evaluated_at": now_iso(),
                }
            )
            if clean_text(snapshot.get("completeness_reason")) == default_capability_snapshot()["completeness_reason"] or not clean_text(snapshot.get("completeness_reason")):
                snapshot["completeness_reason"] = "网站评估执行失败，已保留上一轮结果或按兜底规则补全风险等级。"
            return cache_key, normalize_capability_snapshot(record, snapshot, adapter_configs_path=adapter_configs_path)

    worker_count = min(6, max(1, len(url_to_record)))
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        future_map = {
            executor.submit(inspect_one, cache_key, record): cache_key
            for cache_key, record in url_to_record.items()
        }
        for future in as_completed(future_map):
            cache_key, snapshot = future.result()
            cache[cache_key] = snapshot

    write_capability_cache(cache, cache_path)
    return cache


def print_source_capability_report(report: dict[str, Any]) -> None:
    print(f"  access_status: {report['status']} | http_status: {report['http_status']}")
    if report.get("final_url"):
        print(f"  final_url: {report['final_url']}")
    print(f"  capability_tags: {' | '.join(report.get('capability_tags') or []) or '无'}")
    print(f"  completeness_risk: {report['completeness_risk']}")
    print(f"  completeness_reason: {report['completeness_reason']}")
    if report.get("feed_candidates"):
        print(f"  rss_or_atom: {report['feed_candidates'][0]}")
    if report.get("searchable_form_count") or report.get("search_links"):
        print(
            f"  public_search: forms {int(report.get('searchable_form_count') or 0)}"
            f", links {len(report.get('search_links') or [])}"
        )
    if report.get("api_candidates"):
        print(f"  api_candidate: {report['api_candidates'][0]}")
    if report.get("anti_bot_hints"):
        print(f"  anti_bot_hints: {' | '.join(report['anti_bot_hints'][:4])}")
    if report.get("note"):
        print(f"  access_note: {report['note']}")


def print_list_summary(records: list[dict[str, Any]]) -> None:
    if not records:
        return
    active_count = sum(1 for record in records if record.get("active", True))
    inactive_count = len(records) - active_count
    print(f"合计 {len(records)} 条，active {active_count} 条，inactive {inactive_count} 条。")


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    url = normalize_source_url(args.url) if clean_text(args.url) else ""
    platform = clean_text(args.platform)
    side = clean_text(args.side)
    domain = clean_text(args.domain).lower()

    if args.list:
        records = list_sources(
            args.extra_sources,
            args.show_inactive,
            source_url=url or None,
            platform=platform or None,
            side=side or None,
            domain=domain or None,
        )
        print_source_records(records)
        cache = read_capability_cache(args.capability_cache)
        if records:
            print("\n来源抓取能力识别（显示上次评估结果）:")
        for record in records:
            print_cached_capability_report(
                record,
                cache.get(capability_cache_key(record["source_url"])),
                adapter_configs_path=args.adapter_configs,
            )
        print_list_summary(records)
        return 0

    if args.remove:
        changed = deactivate_sources(
            args.extra_sources,
            source_url=url or None,
            platform=platform or None,
            side=side or None,
            domain=domain or None,
        )
        if changed:
            target = url or domain or "匹配条件"
            print(f"已停用 {len(changed)} 条来源记录 -> {target}")
            print("站点适配配置会继续保留；以后重新添加相同网址时可直接复用。")
        else:
            target = url or domain or "匹配条件"
            print(f"没有找到可停用的来源记录 -> {target}")
        return 0

    if args.reactivate:
        changed = reactivate_sources(
            args.extra_sources,
            source_url=url or None,
            platform=platform or None,
            side=side or None,
            domain=domain or None,
        )
        if changed:
            target = url or domain or "匹配条件"
            print(f"已重新启用 {len(changed)} 条来源记录 -> {target}")
        else:
            target = url or domain or "匹配条件"
            print(f"没有找到可重新启用的来源记录 -> {target}")
        return 0

    url = normalize_source_url(prompt_if_missing(args.url, "请输入抓取网址: "))
    platform = clean_text(args.platform)
    if not platform and (side or "media") == "media":
        platform = GENERAL_MEDIA_PLATFORM
    else:
        platform = prompt_if_missing(platform, "请输入平台名: ")

    action, record = add_or_reactivate_source(args.extra_sources, platform, side or "media", url)
    action_labels = {
        "added": "新增",
        "reactivated": "重新启用",
        "already_active": "已存在且启用中",
    }
    print(f"来源登记: {action_labels[action]} -> {record['source_url']}")
    print("站点抓取能力评估不会实时执行；会在下次新闻抓取成功完成后自动刷新到查看列表。")

    domain = record["domain"]
    if not args.force_api and adapter_exists_for_domain(args.adapter_configs, domain):
        print(f"已存在域名适配配置，直接复用: {domain}")
        return 0

    if args.skip_api:
        print("已跳过 API 适配器生成。")
        return 0

    api_url, api_key, api_model = resolve_api_credentials(args)
    context = discover_site_context(url, args.query_locale)
    adapter = validate_adapter_config(call_adapter_api(build_adapter_prompt(context), api_url, api_key, api_model))
    save_adapter_config(args.adapter_configs, context["domain"], adapter)

    print(f"适配配置已写入: {args.adapter_configs}")
    print(f"目标域名: {context['domain']}")
    print(f"搜索模板: {adapter['search_url_template']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
