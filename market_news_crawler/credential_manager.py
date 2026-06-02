#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from country_config import resolve_project_path
from news_crawler import (
    DEFAULT_SITE_CREDENTIALS_PATH,
    clean_text,
    parse_cookie_header,
    source_site,
)


KNOWN_LOGIN_SITE_TEMPLATES = {
    "seller.tiktokglobalshop.com": {
        "platform": "TikTok/TikTok Shop",
        "side": "seller",
        "label": "TikTok Shop Seller Center",
        "source_url": "https://seller.tiktokglobalshop.com/",
        "auth_type": "form",
        "login_url": "https://seller.tiktokglobalshop.com/",
        "username_field": "username",
        "password_field": "password",
        "extra_form_fields": {},
        "headers": {},
        "cookies": {},
        "cookie_header": "",
        "username": "",
        "password": "",
        "enabled": False,
        "notes": "",
    },
    "sellercentral.amazon.co.jp": {
        "platform": "Amazon",
        "side": "seller",
        "label": "Amazon Seller Central JP",
        "source_url": "https://sellercentral.amazon.co.jp/",
        "auth_type": "form",
        "login_url": "https://sellercentral.amazon.co.jp/",
        "username_field": "email",
        "password_field": "password",
        "extra_form_fields": {},
        "headers": {},
        "cookies": {},
        "cookie_header": "",
        "username": "",
        "password": "",
        "enabled": False,
        "notes": "",
    },
    "glogin.rms.rakuten.co.jp": {
        "platform": "Rakuten Ichiba",
        "side": "seller",
        "label": "Rakuten RMS Login",
        "source_url": "https://glogin.rms.rakuten.co.jp/",
        "auth_type": "form",
        "login_url": "https://glogin.rms.rakuten.co.jp/",
        "username_field": "username",
        "password_field": "password",
        "extra_form_fields": {},
        "headers": {},
        "cookies": {},
        "cookie_header": "",
        "username": "",
        "password": "",
        "enabled": False,
        "notes": "",
    },
    "navi-manual.faq.rakuten.net": {
        "platform": "Rakuten Ichiba",
        "side": "seller",
        "label": "Rakuten Seller Manual",
        "source_url": "https://navi-manual.faq.rakuten.net/",
        "auth_type": "cookie",
        "login_url": "",
        "username_field": "username",
        "password_field": "password",
        "extra_form_fields": {},
        "headers": {},
        "cookies": {},
        "cookie_header": "",
        "username": "",
        "password": "",
        "enabled": False,
        "notes": "推荐优先录入已登录后的 cookies。",
    },
    "qsm.qoo10.jp": {
        "platform": "Qoo10",
        "side": "seller",
        "label": "Qoo10 Seller Manager",
        "source_url": "https://qsm.qoo10.jp/",
        "auth_type": "form",
        "login_url": "https://qsm.qoo10.jp/",
        "username_field": "username",
        "password_field": "password",
        "extra_form_fields": {},
        "headers": {},
        "cookies": {},
        "cookie_header": "",
        "username": "",
        "password": "",
        "enabled": False,
        "notes": "",
    },
    "seller.shein.com": {
        "platform": "Shein",
        "side": "seller",
        "label": "SHEIN Seller Center",
        "source_url": "https://seller.shein.com/",
        "auth_type": "cookie",
        "login_url": "",
        "username_field": "username",
        "password_field": "password",
        "extra_form_fields": {},
        "headers": {},
        "cookies": {},
        "cookie_header": "",
        "username": "",
        "password": "",
        "enabled": False,
        "notes": "当前更适合录入 cookies 或自定义 headers。",
    },
}


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="站点账号/凭据管理：为需要登录的站点录入账号、cookies 或 headers。")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--list", action="store_true", help="列出当前账号模板与已录入凭据")
    mode.add_argument("--clear", action="store_true", help="清空某个站点的账号与凭据")
    parser.add_argument("target", nargs="?", help="域名或网址，例如 sellercentral.amazon.co.jp")
    parser.add_argument(
        "--credentials-file",
        "--site-credentials",
        dest="credentials_file",
        default=DEFAULT_SITE_CREDENTIALS_PATH,
        help="账号配置 JSON",
    )
    parser.add_argument("--username", default="", help="登录用户名")
    parser.add_argument("--password", default="", help="登录密码")
    parser.add_argument("--auth-type", choices=["form", "basic", "cookie"], default="", help="认证方式")
    parser.add_argument("--login-url", default="", help="登录提交地址")
    parser.add_argument("--username-field", default="", help="表单用户名字段名")
    parser.add_argument("--password-field", default="", help="表单密码字段名")
    parser.add_argument("--cookie", action="append", default=[], help="单个 cookie，格式 name=value，可重复")
    parser.add_argument("--cookie-header", default="", help="整段 Cookie header")
    parser.add_argument("--header", action="append", default=[], help="自定义请求头，格式 Name: Value，可重复")
    parser.add_argument("--note", default="", help="备注")
    parser.add_argument("--enable", action="store_true", help="启用该站点凭据")
    parser.add_argument("--disable", action="store_true", help="停用该站点凭据")
    return parser.parse_args(argv)


def normalize_target(target: str) -> str:
    normalized = clean_text(target)
    if not normalized:
        return ""
    if "://" in normalized:
        return source_site(normalized).lower()
    return normalized.lower().strip("/")


def read_json_file(path: str, default: Any) -> Any:
    file_path = resolve_project_path(path)
    if not file_path.exists():
        return default
    return json.loads(file_path.read_text(encoding="utf-8"))


def write_json_file(path: str, payload: Any) -> None:
    file_path = resolve_project_path(path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_credentials(path: str) -> dict[str, dict[str, Any]]:
    payload = read_json_file(path, {})
    if not isinstance(payload, dict):
        payload = {}
    merged = {domain: dict(config) for domain, config in KNOWN_LOGIN_SITE_TEMPLATES.items()}
    for domain, config in payload.items():
        if isinstance(config, dict):
            merged[clean_text(domain).lower()] = {**merged.get(clean_text(domain).lower(), {}), **config}
    return merged


def save_credentials(path: str, payload: dict[str, dict[str, Any]]) -> None:
    write_json_file(path, payload)


def ensure_target(payload: dict[str, dict[str, Any]], domain: str) -> dict[str, Any]:
    if domain not in payload:
        payload[domain] = {
            "platform": "",
            "side": "",
            "label": domain,
            "source_url": f"https://{domain}/",
            "auth_type": "cookie",
            "login_url": "",
            "username_field": "username",
            "password_field": "password",
            "extra_form_fields": {},
            "headers": {},
            "cookies": {},
            "cookie_header": "",
            "username": "",
            "password": "",
            "enabled": False,
            "notes": "",
        }
    return payload[domain]


def parse_header_items(items: list[str]) -> dict[str, str]:
    headers: dict[str, str] = {}
    for item in items:
        text = str(item)
        if ":" not in text:
            continue
        name, value = text.split(":", 1)
        header_name = clean_text(name)
        if not header_name:
            continue
        headers[header_name] = value.strip()
    return headers


def parse_cookie_items(items: list[str]) -> dict[str, str]:
    cookies: dict[str, str] = {}
    for item in items:
        if "=" not in item:
            continue
        name, value = item.split("=", 1)
        cookie_name = clean_text(name)
        if not cookie_name:
            continue
        cookies[cookie_name] = value.strip()
    return cookies


def print_credentials(payload: dict[str, dict[str, Any]]) -> None:
    if not payload:
        print("当前没有站点账号模板。")
        return
    for domain in sorted(payload):
        config = payload[domain]
        status = "enabled" if config.get("enabled") else "disabled"
        print(f"[{status}] {domain}")
        if clean_text(config.get("platform")) or clean_text(config.get("side")):
            print(f"  platform: {clean_text(config.get('platform'))} | side: {clean_text(config.get('side'))}")
        if clean_text(config.get("label")):
            print(f"  label: {clean_text(config.get('label'))}")
        if clean_text(config.get("source_url")):
            print(f"  source_url: {clean_text(config.get('source_url'))}")
        print(f"  auth_type: {clean_text(config.get('auth_type')) or 'cookie'}")
        if clean_text(config.get("login_url")):
            print(f"  login_url: {clean_text(config.get('login_url'))}")
        if clean_text(config.get("username")):
            print(f"  username: {clean_text(config.get('username'))}")
        if config.get("cookies"):
            print(f"  cookies: {len(config.get('cookies', {}))} 项")
        if clean_text(config.get("cookie_header")):
            print("  cookie_header: 已设置")
        if config.get("headers"):
            print(f"  headers: {len(config.get('headers', {}))} 项")
        if clean_text(config.get("notes")):
            print(f"  notes: {clean_text(config.get('notes'))}")


def clear_credentials(config: dict[str, Any]) -> None:
    config["username"] = ""
    config["password"] = ""
    config["cookies"] = {}
    config["cookie_header"] = ""
    config["headers"] = {}
    config["enabled"] = False


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    payload = load_credentials(args.credentials_file)

    if args.list:
        print_credentials(payload)
        return 0

    domain = normalize_target(args.target or "")
    if not domain:
        raise SystemExit("请提供要操作的域名或网址，例如 sellercentral.amazon.co.jp")

    config = ensure_target(payload, domain)

    if args.clear:
        clear_credentials(config)
        save_credentials(args.credentials_file, payload)
        print(f"已清空站点凭据 -> {domain}")
        return 0

    if args.auth_type:
        config["auth_type"] = clean_text(args.auth_type)
    if args.login_url:
        config["login_url"] = clean_text(args.login_url)
    if args.username_field:
        config["username_field"] = clean_text(args.username_field)
    if args.password_field:
        config["password_field"] = clean_text(args.password_field)
    if args.username:
        config["username"] = clean_text(args.username)
    if args.password:
        config["password"] = args.password
    if args.cookie_header:
        config["cookie_header"] = args.cookie_header.strip()
        parsed = parse_cookie_header(config["cookie_header"])
        if parsed:
            config["cookies"] = {**config.get("cookies", {}), **parsed}
    parsed_cookies = parse_cookie_items(args.cookie)
    if parsed_cookies:
        config["cookies"] = {**config.get("cookies", {}), **parsed_cookies}
    parsed_headers = parse_header_items(args.header)
    if parsed_headers:
        config["headers"] = {**config.get("headers", {}), **parsed_headers}
    if args.note:
        config["notes"] = clean_text(args.note)
    if args.enable:
        config["enabled"] = True
    if args.disable:
        config["enabled"] = False

    save_credentials(args.credentials_file, payload)
    print(f"站点凭据已更新 -> {domain}")
    print(f"启用状态: {'enabled' if config.get('enabled') else 'disabled'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
