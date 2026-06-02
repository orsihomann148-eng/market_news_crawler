from __future__ import annotations

import base64
import binascii
import hashlib
import hmac
import json
import os
import signal
import sys
from pathlib import Path
from typing import Any

import xlsx_source_test
from country_config import (
    DEFAULT_COUNTRY_CODE,
    get_country_config,
    normalize_country_code,
    resolve_project_path,
)


BASE_DIR = Path(__file__).resolve().parent
APP_SETTINGS_PATH = BASE_DIR / 'web_app_settings.json'
SETTINGS_SECRET_PREFIX = 'enc-v1:'
SURVEY_API_SETTING_FIELDS = ('survey_api_url', 'survey_api_key', 'survey_api_model')
CONFIG_PREVIOUS_FIELD_MAP = {
    'related_news_search_keywords': 'related_news_search_keywords_previous',
    'report_search_keywords': 'report_search_keywords_previous',
    'survey_system_prompt': 'survey_system_prompt_previous',
}


def app_secret_dir() -> Path:
    if os.name == 'nt':
        root = Path(os.environ.get('APPDATA') or (Path.home() / 'AppData' / 'Roaming'))
        generic_dir = root / 'market_news_catch_auto'
        legacy_dir = root / 'japan_news_catch_auto'
        return legacy_dir if legacy_dir.exists() else generic_dir
    xdg_root = os.environ.get('XDG_CONFIG_HOME')
    if xdg_root:
        generic_dir = Path(xdg_root) / 'market_news_catch_auto'
        legacy_dir = Path(xdg_root) / 'japan_news_catch_auto'
        return legacy_dir if legacy_dir.exists() else generic_dir
    generic_dir = Path.home() / '.config' / 'market_news_catch_auto'
    legacy_dir = Path.home() / '.config' / 'japan_news_catch_auto'
    return legacy_dir if legacy_dir.exists() else generic_dir


APP_SECRET_PATH = app_secret_dir() / 'web_app_secret.key'


def normalize_country_request(value: Any) -> str:
    return normalize_country_code(str(value or '').strip())


def country_path(country_code: str, key: str) -> Path:
    return resolve_project_path(str(get_country_config(country_code)[key]), base_dir=BASE_DIR)


def output_dir_matches_country(path: Path, country_code: str) -> bool:
    if not path.is_dir():
        return False
    name = path.name
    country_config = get_country_config(country_code)
    slug = str(country_config['output_slug'])
    if name.startswith(f'{slug}_'):
        return True
    legacy_prefixes = country_config.get('legacy_output_prefixes') or []
    return any(name.startswith(str(prefix)) for prefix in legacy_prefixes)


def install_terminal_signal_guards() -> None:
    def ignore_suspend_signal(signum, frame) -> None:
        sys.stdout.write('\n已忽略 Ctrl+Z 挂起操作。请使用 Ctrl+C 停止当前 Web 服务。\n')
        sys.stdout.flush()

    for signal_name in ('SIGTSTP', 'SIGQUIT'):
        signal_value = getattr(signal, signal_name, None)
        if signal_value is None:
            continue
        try:
            signal.signal(signal_value, ignore_suspend_signal)
        except (ValueError, OSError, RuntimeError):
            continue


def ensure_secret_file_permissions(path: Path) -> None:
    if os.name == 'nt':
        return
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


def load_or_create_app_secret() -> bytes:
    if APP_SECRET_PATH.exists():
        try:
            secret = base64.urlsafe_b64decode(APP_SECRET_PATH.read_text(encoding='utf-8').strip().encode('ascii'))
            if len(secret) >= 32:
                return secret[:32]
        except (OSError, ValueError, binascii.Error):
            pass

    APP_SECRET_PATH.parent.mkdir(parents=True, exist_ok=True)
    secret = os.urandom(32)
    encoded = base64.urlsafe_b64encode(secret).decode('ascii')
    APP_SECRET_PATH.write_text(encoded, encoding='utf-8')
    ensure_secret_file_permissions(APP_SECRET_PATH)
    return secret


def derive_secret_keys(master_secret: bytes) -> tuple[bytes, bytes]:
    encryption_key = hmac.new(master_secret, b'web-app-settings-encryption', hashlib.sha256).digest()
    mac_key = hmac.new(master_secret, b'web-app-settings-mac', hashlib.sha256).digest()
    return encryption_key, mac_key


def build_keystream(encryption_key: bytes, nonce: bytes, length: int) -> bytes:
    blocks: list[bytes] = []
    counter = 0
    while sum(len(block) for block in blocks) < length:
        counter_bytes = counter.to_bytes(4, 'big')
        blocks.append(hmac.new(encryption_key, nonce + counter_bytes, hashlib.sha256).digest())
        counter += 1
    return b''.join(blocks)[:length]


def encrypt_secret_value(value: str) -> str:
    normalized = str(value or '')
    if not normalized:
        return ''
    master_secret = load_or_create_app_secret()
    encryption_key, mac_key = derive_secret_keys(master_secret)
    nonce = os.urandom(16)
    plaintext = normalized.encode('utf-8')
    keystream = build_keystream(encryption_key, nonce, len(plaintext))
    ciphertext = bytes(left ^ right for left, right in zip(plaintext, keystream))
    tag = hmac.new(mac_key, nonce + ciphertext, hashlib.sha256).digest()
    token = base64.urlsafe_b64encode(nonce + ciphertext + tag).decode('ascii')
    return f'{SETTINGS_SECRET_PREFIX}{token}'


def decrypt_secret_value(value: str) -> str:
    normalized = str(value or '')
    if not normalized:
        return ''
    if not normalized.startswith(SETTINGS_SECRET_PREFIX):
        return normalized

    encoded = normalized[len(SETTINGS_SECRET_PREFIX):]
    try:
        blob = base64.urlsafe_b64decode(encoded.encode('ascii'))
    except (ValueError, binascii.Error):
        return ''

    if len(blob) < 16 + 32:
        return ''

    nonce = blob[:16]
    tag = blob[-32:]
    ciphertext = blob[16:-32]
    master_secret = load_or_create_app_secret()
    encryption_key, mac_key = derive_secret_keys(master_secret)
    expected_tag = hmac.new(mac_key, nonce + ciphertext, hashlib.sha256).digest()
    if not hmac.compare_digest(tag, expected_tag):
        return ''

    keystream = build_keystream(encryption_key, nonce, len(ciphertext))
    plaintext = bytes(left ^ right for left, right in zip(ciphertext, keystream))
    try:
        return plaintext.decode('utf-8')
    except UnicodeDecodeError:
        return ''


def iter_news_api_setting_states(payload: dict[str, Any]):
    news_state = payload.get('news_filter_state')
    if isinstance(news_state, dict):
        yield news_state
    ai_settings = payload.get('ai_settings')
    if isinstance(ai_settings, dict):
        yield ai_settings
    countries = payload.get('countries')
    if isinstance(countries, dict):
        for country_payload in countries.values():
            if not isinstance(country_payload, dict):
                continue
            country_news_state = country_payload.get('news_filter_state')
            if isinstance(country_news_state, dict):
                yield country_news_state


def encrypt_settings_payload(payload: dict[str, Any]) -> dict[str, Any]:
    encrypted = json.loads(json.dumps(payload, ensure_ascii=False))
    for news_state in iter_news_api_setting_states(encrypted):
        news_state['survey_api_key'] = encrypt_secret_value(str(news_state.get('survey_api_key') or ''))
    source_state = encrypted.get('source_state')
    if isinstance(source_state, dict):
        source_state['api_key'] = encrypt_secret_value(str(source_state.get('api_key') or ''))
    return encrypted


def decrypt_settings_payload(payload: dict[str, Any]) -> dict[str, Any]:
    decrypted = json.loads(json.dumps(payload, ensure_ascii=False))
    for news_state in iter_news_api_setting_states(decrypted):
        news_state['survey_api_key'] = decrypt_secret_value(str(news_state.get('survey_api_key') or ''))
    source_state = decrypted.get('source_state')
    if isinstance(source_state, dict):
        source_state['api_key'] = decrypt_secret_value(str(source_state.get('api_key') or ''))
    return decrypted


def settings_payload_needs_secret_migration(payload: dict[str, Any]) -> bool:
    for news_state in iter_news_api_setting_states(payload):
        survey_api_key = str(news_state.get('survey_api_key') or '')
        if survey_api_key and not survey_api_key.startswith(SETTINGS_SECRET_PREFIX):
            return True
    source_state = payload.get('source_state')
    if isinstance(source_state, dict):
        api_key = str(source_state.get('api_key') or '')
        if api_key and not api_key.startswith(SETTINGS_SECRET_PREFIX):
            return True
    return False


def read_app_settings() -> dict[str, Any]:
    if not APP_SETTINGS_PATH.exists():
        write_app_settings({})
        return {}
    try:
        payload = json.loads(APP_SETTINGS_PATH.read_text(encoding='utf-8'))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    decrypted = decrypt_settings_payload(payload)
    if settings_payload_needs_secret_migration(payload):
        try:
            write_app_settings(decrypted)
        except Exception:
            pass
    return decrypted


def write_app_settings(payload: dict[str, Any]) -> None:
    encrypted_payload = encrypt_settings_payload(payload)
    APP_SETTINGS_PATH.write_text(json.dumps(encrypted_payload, ensure_ascii=False, indent=2), encoding='utf-8')


def read_country_settings(country_code: str = DEFAULT_COUNTRY_CODE) -> dict[str, Any]:
    normalized_country = normalize_country_request(country_code)
    payload = read_app_settings()
    countries = payload.get('countries')
    if isinstance(countries, dict):
        country_payload = countries.get(normalized_country)
        if isinstance(country_payload, dict):
            return country_payload

    if normalized_country != DEFAULT_COUNTRY_CODE:
        return {}

    legacy_payload: dict[str, Any] = {}
    if isinstance(payload.get('news_filter_state'), dict):
        legacy_payload['news_filter_state'] = payload['news_filter_state']
    if isinstance(payload.get('news_form_state'), dict):
        legacy_payload['news_form_state'] = payload['news_form_state']
    return legacy_payload


def text_config_changed(previous_value: Any, next_value: Any) -> bool:
    return str(previous_value or '').strip() != str(next_value or '').strip()


def set_country_previous_config_value(country_code: str, field: str, value: Any) -> None:
    normalized_country_code = normalize_country_request(country_code)
    text_value = str(value or '').strip()
    if not text_value:
        return

    payload = read_app_settings()
    countries = payload.get('countries')
    if not isinstance(countries, dict):
        countries = {}
    country_payload = countries.get(normalized_country_code)
    if not isinstance(country_payload, dict):
        country_payload = {}
    country_payload[field] = text_value
    countries[normalized_country_code] = country_payload
    payload['countries'] = countries
    write_app_settings(payload)


def read_country_previous_config_value(country_code: str, field: str) -> str:
    country_payload = read_country_settings(country_code)
    return str(country_payload.get(field) or '').strip()


def backup_news_filter_previous_values(
    country_payload: dict[str, Any],
    current_state: dict[str, Any],
    next_state: dict[str, Any],
) -> None:
    for state_field, previous_field in CONFIG_PREVIOUS_FIELD_MAP.items():
        current_value = str(current_state.get(state_field) or '').strip()
        next_value = str(next_state.get(state_field) or '').strip()
        if current_value and current_value != next_value:
            country_payload[previous_field] = current_value


def extract_survey_api_settings(state: dict[str, Any] | None) -> dict[str, str]:
    if not isinstance(state, dict):
        return {}
    return {
        key: str(state.get(key) or '').strip()
        for key in SURVEY_API_SETTING_FIELDS
        if str(state.get(key) or '').strip()
    }


def read_global_survey_api_settings() -> dict[str, str]:
    payload = read_app_settings()
    global_settings = extract_survey_api_settings(payload.get('ai_settings') if isinstance(payload.get('ai_settings'), dict) else None)
    if global_settings:
        return global_settings

    countries = payload.get('countries')
    if isinstance(countries, dict):
        last_country = normalize_country_request(payload.get('last_country_code') or DEFAULT_COUNTRY_CODE)
        country_order = [last_country] + [code for code in countries if code != last_country]
        for country_code in country_order:
            country_payload = countries.get(country_code)
            if not isinstance(country_payload, dict):
                continue
            country_settings = extract_survey_api_settings(country_payload.get('news_filter_state'))
            if country_settings:
                return country_settings

    return extract_survey_api_settings(payload.get('news_filter_state') if isinstance(payload.get('news_filter_state'), dict) else None)


def extract_survey_api_settings_from_form(form) -> dict[str, str]:
    api_url = xlsx_source_test.normalize_chat_completions_url(
        form.get('alias_ai_api_url') or form.get('country_ai_api_url') or form.get('survey_api_url') or ''
    )
    api_key = (form.get('alias_ai_api_key') or form.get('country_ai_api_key') or form.get('survey_api_key') or '').strip()
    api_model = (form.get('alias_ai_api_model') or form.get('country_ai_api_model') or form.get('survey_api_model') or '').strip()
    return {
        key: value
        for key, value in {
            'survey_api_url': api_url,
            'survey_api_key': api_key,
            'survey_api_model': api_model,
        }.items()
        if value
    }


def persist_global_survey_api_settings(settings: dict[str, str]) -> None:
    if not settings:
        return
    payload = read_app_settings()
    ai_settings = payload.get('ai_settings')
    if not isinstance(ai_settings, dict):
        ai_settings = {}
    for key, value in settings.items():
        if key in SURVEY_API_SETTING_FIELDS and str(value or '').strip():
            ai_settings[key] = str(value).strip()
    payload['ai_settings'] = ai_settings
    write_app_settings(payload)


def read_country_ai_prompt_setting() -> str:
    payload = read_app_settings()
    return str(payload.get('country_ai_prompt') or '').strip()


def persist_country_ai_prompt_setting(prompt: str) -> None:
    payload = read_app_settings()
    normalized_prompt = str(prompt or '').strip()
    if normalized_prompt:
        payload['country_ai_prompt'] = normalized_prompt
    else:
        payload.pop('country_ai_prompt', None)
    write_app_settings(payload)


def remove_country_settings(country_code: str) -> None:
    payload = read_app_settings()
    countries = payload.get('countries')
    if isinstance(countries, dict) and country_code in countries:
        countries.pop(country_code, None)
        payload['countries'] = countries
    if normalize_country_request(payload.get('last_country_code')) == country_code:
        payload['last_country_code'] = DEFAULT_COUNTRY_CODE
    write_app_settings(payload)


def read_last_country_code() -> str:
    payload = read_app_settings()
    return normalize_country_request(payload.get('last_country_code') or DEFAULT_COUNTRY_CODE)

