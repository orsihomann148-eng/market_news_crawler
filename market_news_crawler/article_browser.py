from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import tempfile
from datetime import datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Any

import db_store
import dedupe
import news_crawler
import runtime_paths
import xlsx_source_test
from country_config import (
    DEFAULT_COUNTRY_CODE,
    get_country_config,
    normalize_country_code,
    normalize_project_relative_path,
)


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = runtime_paths.outputs_dir()
RUNTIME_OUTPUT_DIR = Path.cwd() / 'outputs'
LEGACY_WORKSPACE_OUTPUT_DIR = BASE_DIR.parents[2] / 'outputs' if len(BASE_DIR.parents) > 2 else RUNTIME_OUTPUT_DIR
DEFAULT_ARTICLES_CSV_PATH = BASE_DIR / 'articles.csv'
ARTICLE_STAR_STORE_PATH = runtime_paths.article_star_store_path()


def normalize_country_request(value: Any) -> str:
    return normalize_country_code(str(value or '').strip())


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


def read_article_star_store() -> dict[str, dict[str, Any]]:
    try:
        db_store.migrate_star_json_if_needed(ARTICLE_STAR_STORE_PATH)
        return db_store.load_star_store()
    except Exception:
        if not ARTICLE_STAR_STORE_PATH.exists():
            return {}
        try:
            payload = json.loads(ARTICLE_STAR_STORE_PATH.read_text(encoding='utf-8'))
        except Exception:
            return {}
        if not isinstance(payload, dict):
            return {}
        return {
            str(key): value
            for key, value in payload.items()
            if isinstance(key, str) and isinstance(value, dict)
        }


def write_article_star_store(payload: dict[str, dict[str, Any]]) -> None:
    for article_id, value in payload.items():
        if isinstance(value, dict):
            db_store.set_article_star(str(article_id), value, bool(value.get('starred', True)))


def build_article_star_key(
    *,
    platform_label: str = '',
    title: str = '',
    title_original: str = '',
    source_name: str = '',
    article_url: str = '',
    source_url: str = '',
    published_at: str = '',
) -> str:
    raw = '||'.join(
        item.strip()
        for item in [
            platform_label,
            title,
            title_original,
            source_name,
            article_url,
            source_url,
            published_at,
        ]
    )
    return hashlib.sha1(raw.encode('utf-8')).hexdigest()


def normalize_briefing_sentiment(value: Any) -> str:
    return xlsx_source_test.normalize_briefing_sentiment(value)


def read_output_metadata_for_file(path: Path) -> dict[str, Any]:
    metadata_path = path.parent / 'metadata.json'
    if not metadata_path.exists():
        return {}
    try:
        payload = json.loads(metadata_path.read_text(encoding='utf-8'))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def article_csv_translation_status_label(path: Path) -> str:
    metadata = read_output_metadata_for_file(path)
    name = path.name
    if name in {'before.csv', 'before.json'}:
        status = str(metadata.get('before_file_translation_status') or '').strip()
        if status == 'raw_untranslated':
            return '原文未翻译'
        if status == 'translated':
            return '已翻译'
        translation_timing = str(metadata.get('translation_timing') or '').strip()
        if translation_timing == 'after_ai_filter':
            return '原文未翻译'
        if translation_timing == 'before_filter':
            return '已翻译'
    if name in {'after.csv', 'after.json'}:
        status = str(metadata.get('after_file_translation_status') or '').strip()
        if status == 'translated':
            return '已翻译'
        translation_timing = str(metadata.get('translation_timing') or '').strip()
        if translation_timing in {'after_ai_filter', 'before_filter'}:
            return '已翻译'
    return ''


def classify_article_csv_label(path: Path) -> str:
    name = path.name
    if name in {'before.csv', 'before.json'} or name.startswith('articles_before_filter_'):
        translation_status = article_csv_translation_status_label(path)
        return f"筛选前/{translation_status}" if translation_status else '筛选前'
    if name in {'after.csv', 'after.json'} or name.startswith('articles_after_filter_'):
        translation_status = article_csv_translation_status_label(path)
        return f"筛选后/{translation_status}" if translation_status else '筛选后'
    if name in {'sources.csv', 'sources.json'} or name.startswith('source_results_'):
        return '来源结果'
    if name.startswith('articles_'):
        return '兼容结果'
    return '新闻数据'


def build_article_csv_display_label(path: Path) -> str:
    if path.parent == BASE_DIR:
        return path.name
    return f'{path.parent.name} / {path.name}'


RUN_OUTPUT_NAME_RE = re.compile(
    r'run_(?P<run_date>\d{8})_(?P<run_time>\d{6})(?:_range_(?P<start>\d{8})_(?P<end>\d{8}))?'
)


def parse_dt_for_label(value: Any) -> datetime | None:
    raw = str(value or '').strip()
    if not raw:
        return None
    try:
        return datetime.fromisoformat(raw.replace('Z', '+00:00'))
    except ValueError:
        pass
    for fmt in ('%Y%m%d%H%M%S', '%Y%m%d', '%Y-%m-%d'):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def format_date_for_label(value: Any) -> str:
    parsed = parse_dt_for_label(value)
    return parsed.strftime('%Y-%m-%d') if parsed else ''


def format_run_time_for_label(value: Any) -> str:
    parsed = parse_dt_for_label(value)
    return parsed.strftime('%Y-%m-%d %H:%M') if parsed else ''


def parse_run_name_parts(run_id: str) -> dict[str, str]:
    match = RUN_OUTPUT_NAME_RE.search(str(run_id or ''))
    if not match:
        return {}
    run_date = match.group('run_date') or ''
    run_time = match.group('run_time') or ''
    return {
        'run_time': f'{run_date}{run_time}' if run_date and run_time else '',
        'range_start': match.group('start') or '',
        'range_end': match.group('end') or '',
    }


def metadata_from_run_row(run: dict[str, Any]) -> dict[str, Any]:
    metadata = run.get('metadata') if isinstance(run.get('metadata'), dict) else None
    if metadata is not None:
        return metadata
    try:
        payload = json.loads(str(run.get('metadata_json') or ''))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def friendly_range_label(start: Any = '', end: Any = '') -> str:
    start_label = format_date_for_label(start)
    end_label = format_date_for_label(end)
    if start_label and end_label:
        return f'{start_label} 至 {end_label}'
    return '时间范围未知'


def friendly_article_source_label(
    *,
    country_label: str = '',
    run_id: str = '',
    generated_at: Any = '',
    range_start: Any = '',
    range_end: Any = '',
    article_count: int = 0,
    is_latest: bool = False,
    csv_history: bool = False,
) -> str:
    name_parts = parse_run_name_parts(run_id)
    run_time_label = format_run_time_for_label(name_parts.get('run_time')) or format_run_time_for_label(generated_at)
    range_label = friendly_range_label(range_start or name_parts.get('range_start'), range_end or name_parts.get('range_end'))
    normalized_country = str(country_label or '').strip()
    count_label = f'{int(article_count)}条' if article_count else '0条'
    if csv_history:
        prefix = '历史结果'
        suffix = 'CSV'
    else:
        prefix = '最近一次抓取' if is_latest else (f'{run_time_label} 抓取' if run_time_label else '历史结果')
        suffix = count_label
    parts = [prefix]
    if normalized_country:
        parts.append(normalized_country)
    parts.append(range_label)
    parts.append(suffix)
    return '｜'.join(parts)


def friendly_db_article_source_label(run: dict[str, Any], *, is_latest: bool = False) -> str:
    metadata = metadata_from_run_row(run)
    country_label = str(metadata.get('country_label') or run.get('country_label') or '').strip()
    return friendly_article_source_label(
        country_label=country_label,
        run_id=str(run.get('run_id') or ''),
        generated_at=metadata.get('generated_at') or run.get('generated_at'),
        range_start=metadata.get('range_start'),
        range_end=metadata.get('range_end'),
        article_count=int(run.get('after_count') or 0),
        is_latest=is_latest,
    )


def friendly_csv_article_source_label(path: Path, *, country_code: str = DEFAULT_COUNTRY_CODE, is_latest: bool = False) -> str:
    metadata = read_output_metadata_for_file(path)
    country_label = str(metadata.get('country_label') or get_country_config(country_code).get('label') or '').strip()
    return friendly_article_source_label(
        country_label=country_label,
        run_id=path.parent.name,
        generated_at=metadata.get('generated_at'),
        range_start=metadata.get('range_start'),
        range_end=metadata.get('range_end'),
        article_count=count_csv_data_rows(path),
        is_latest=is_latest,
        csv_history=True,
    )


def is_article_browser_csv(path: Path) -> bool:
    name = path.name
    return name in {'before.csv', 'after.csv', 'articles.csv'} or name.startswith('articles_before_filter_') or name.startswith('articles_after_filter_') or name == 'articles.csv'


def count_csv_data_rows(path: Path) -> int:
    try:
        with path.open('r', encoding='utf-8-sig', newline='') as handle:
            reader = csv.reader(handle)
            next(reader, None)
            return sum(1 for row in reader if row)
    except Exception:
        return 0


def iter_output_dirs() -> list[Path]:
    dirs: list[Path] = []
    seen: set[str] = set()
    for path in [OUTPUT_DIR, RUNTIME_OUTPUT_DIR, LEGACY_WORKSPACE_OUTPUT_DIR]:
        resolved = path.resolve()
        key = str(resolved).lower() if os.name == 'nt' else str(resolved)
        if key in seen:
            continue
        seen.add(key)
        dirs.append(resolved)
    return dirs


def article_source_value(path: Path) -> str:
    resolved = path.resolve()
    if resolved.is_relative_to(BASE_DIR):
        return str(resolved.relative_to(BASE_DIR)).replace('\\', '/')
    return str(resolved)


def db_article_source_value(run_id: str, stage: str = 'after') -> str:
    normalized_stage = 'before' if stage == 'before' else 'after'
    return f'db:{run_id}:{normalized_stage}'


def parse_db_article_source(value: str | None) -> tuple[str, str] | None:
    raw = str(value or '').strip()
    if not raw.startswith('db:'):
        return None
    parts = raw.split(':', 2)
    if len(parts) != 3 or not parts[1].strip():
        return None
    stage = 'before' if parts[2].strip() == 'before' else 'after'
    return parts[1].strip(), stage


def is_db_article_source(value: str | None) -> bool:
    return parse_db_article_source(value) is not None


def resolve_output_dir_for_run(value: str | None) -> str:
    return str(runtime_paths.runtime_output_dir_for_user_value(value).resolve())


def list_recent_output_dirs(limit: int = 10, *, country_code: str = DEFAULT_COUNTRY_CODE) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen_paths: set[str] = set()
    for output_dir in iter_output_dirs():
        if not output_dir.exists():
            continue
        for path in output_dir.iterdir():
            resolved_key = str(path.resolve()).lower() if os.name == 'nt' else str(path.resolve())
            if resolved_key in seen_paths or not output_dir_matches_country(path, country_code):
                continue
            seen_paths.add(resolved_key)
            files = sorted(child.name for child in path.iterdir())
            rows.append(
                {
                    'name': path.name,
                    'path': str(path),
                    'updated_at': path.stat().st_mtime,
                    'files': files,
                }
            )
    rows.sort(key=lambda row: row['updated_at'], reverse=True)
    return rows[:limit]


def list_article_csv_paths(limit: int | None = 30, *, country_code: str = DEFAULT_COUNTRY_CODE) -> list[Path]:
    rows: list[Path] = []
    seen_paths: set[str] = set()

    output_paths: list[Path] = []
    for output_dir in iter_output_dirs():
        if not output_dir.exists():
            continue
        output_paths.extend(
            path
            for path in output_dir.glob('**/*.csv')
            if path.exists()
            and path.is_file()
            and is_article_browser_csv(path)
            and output_dir_matches_country(path.parent, country_code)
        )
    output_paths.sort(key=lambda item: item.stat().st_mtime, reverse=True)

    candidate_paths: list[Path] = list(output_paths)
    if DEFAULT_ARTICLES_CSV_PATH.exists() and DEFAULT_ARTICLES_CSV_PATH.is_file():
        candidate_paths.append(DEFAULT_ARTICLES_CSV_PATH)

    for path in candidate_paths:
        resolved = str(path.resolve())
        if resolved in seen_paths:
            continue
        is_legacy_output_copy = (
            path.parent != BASE_DIR
            and path.name.startswith('articles_')
            and not path.name.startswith('articles_before_filter_')
            and not path.name.startswith('articles_after_filter_')
            and path.name != 'articles.csv'
        )
        if is_legacy_output_copy:
            continue
        seen_paths.add(resolved)
        rows.append(path)
        if limit is not None and len(rows) >= limit:
            break

    return rows


def list_article_csv_files(limit: int = 30, *, country_code: str = DEFAULT_COUNTRY_CODE) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    remaining_csv_limit = limit

    try:
        db_runs = db_store.list_runs(country_code, limit=limit)
    except Exception:
        db_runs = []
    for run in db_runs:
        run_id = str(run.get('run_id') or '').strip()
        if not run_id:
            continue
        for stage, label_stage, count_key in [
            ('after', '筛选后/SQLite', 'after_count'),
            ('before', '筛选前/SQLite', 'before_count'),
        ]:
            article_count = int(run.get(count_key) or 0)
            if article_count <= 0:
                continue
            rows.append(
                {
                    'value': db_article_source_value(run_id, stage),
                    'label': f'{run_id} / {stage}.db · {label_stage} · {article_count}条',
                }
            )
            if remaining_csv_limit is not None:
                remaining_csv_limit = max(0, remaining_csv_limit - 1)
        if limit is not None and len(rows) >= limit:
            return rows[:limit]

    if remaining_csv_limit == 0:
        return rows[:limit] if limit is not None else rows

    for path in list_article_csv_paths(remaining_csv_limit, country_code=country_code):
        relative_label = article_source_value(path)
        display_label = build_article_csv_display_label(path)
        article_type_label = classify_article_csv_label(path)
        article_count = count_csv_data_rows(path)
        rows.append(
            {
                'value': relative_label,
                'label': f'{display_label} · {article_type_label} · {article_count}条',
            }
        )

    return rows[:limit] if limit is not None else rows


def list_user_article_sources(limit: int = 30, *, country_code: str = DEFAULT_COUNTRY_CODE) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    remaining_csv_limit = limit

    try:
        db_runs = db_store.list_runs(country_code, limit=limit)
    except Exception:
        db_runs = []
    for run in db_runs:
        run_id = str(run.get('run_id') or '').strip()
        article_count = int(run.get('after_count') or 0)
        if not run_id or article_count <= 0:
            continue
        rows.append(
            {
                'value': db_article_source_value(run_id, 'after'),
                'label': friendly_db_article_source_label(run, is_latest=not rows),
            }
        )
        if remaining_csv_limit is not None:
            remaining_csv_limit = max(0, remaining_csv_limit - 1)
        if limit is not None and len(rows) >= limit:
            return rows[:limit]

    if remaining_csv_limit == 0:
        return rows[:limit] if limit is not None else rows

    csv_paths = [
        path
        for path in list_article_csv_paths(None, country_code=country_code)
        if path.name == 'after.csv' or path.name.startswith('articles_after_filter_') or path.name == 'articles.csv'
    ]
    for path in csv_paths:
        rows.append(
            {
                'value': article_source_value(path),
                'label': friendly_csv_article_source_label(path, country_code=country_code, is_latest=not rows),
            }
        )
        if limit is not None and len(rows) >= limit:
            return rows[:limit]

    return rows[:limit] if limit is not None else rows


def resolve_article_csv_path(raw_value: str | None, *, country_code: str = DEFAULT_COUNTRY_CODE) -> Path:
    article_sources = [
        source
        for source in list_article_csv_files(country_code=country_code)
        if not is_db_article_source(source.get('value') or '')
    ]
    default_path = DEFAULT_ARTICLES_CSV_PATH if country_code == DEFAULT_COUNTRY_CODE and DEFAULT_ARTICLES_CSV_PATH.exists() else None
    selected_value = (raw_value or '').strip().replace('\\', '/')
    if selected_value:
        selected_path = Path(selected_value)
        candidate = selected_path.resolve() if selected_path.is_absolute() else (BASE_DIR / selected_value).resolve()
        if candidate.exists() and candidate.is_file():
            return candidate
    if default_path is not None:
        return default_path
    if article_sources:
        source_path = Path(article_sources[0]['value'])
        return source_path.resolve() if source_path.is_absolute() else (BASE_DIR / source_path).resolve()
    return DEFAULT_ARTICLES_CSV_PATH


@lru_cache(maxsize=256)
def translate_article_summary_fallback(summary_text: str, translation_target: str = '') -> str:
    normalized_summary = str(summary_text or '').strip()
    normalized_target = str(translation_target or '').strip()
    if not normalized_summary or not normalized_target:
        return normalized_summary
    try:
        translator = news_crawler.TextTranslator(normalized_target)
        translated = translator.translate(normalized_summary)
    except Exception:
        translated = None
    return str(translated or normalized_summary).strip()


@lru_cache(maxsize=512)
def translate_article_title_for_display(
    title_original: str,
    title_translated: str = '',
    translation_target: str = '',
) -> str:
    normalized_original = str(title_original or '').strip()
    normalized_translated = str(title_translated or '').strip()
    if normalized_translated and news_crawler.contains_chinese_chars(normalized_translated):
        return normalized_translated

    if not normalized_original:
        return ''

    normalized_target = 'zh-CN'
    if str(translation_target or '').strip() == 'zh-CN' and normalized_translated:
        candidate = normalized_translated
    else:
        candidate = normalized_original

    if news_crawler.contains_chinese_chars(candidate):
        return candidate

    try:
        translator = news_crawler.TextTranslator(normalized_target)
        translated = str(translator.translate(candidate) or '').strip()
    except Exception:
        translated = ''

    if translated and translated != normalized_original and news_crawler.contains_chinese_chars(translated):
        return translated
    return ''


def normalize_article_card_summary(
    *,
    title: str = '',
    title_original: str = '',
    summary: str = '',
    body_excerpt: str = '',
    translation_target: str = '',
) -> str:
    normalized_title = str(title or '').strip()
    normalized_title_original = str(title_original or '').strip()
    normalized_summary = str(summary or '').strip()
    normalized_body_excerpt = str(body_excerpt or '').strip()

    def normalize_text_token(value: str) -> str:
        text = news_crawler.clean_text(str(value or ''))
        if not text:
            return ''
        text = re.sub(r'\s*[|｜\-–—:：_]\s*', ' ', text.lower())
        text = re.sub(r'[^0-9a-z一-鿿ぁ-んァ-ヶ]+', ' ', text)
        return re.sub(r'\s+', ' ', text).strip()

    def canonicalize(value: str) -> str:
        text = news_crawler.clean_text(str(value or ''))
        if not text:
            return ''
        stripped = dedupe.strip_title_source_suffix(text)
        normalized = dedupe.normalize_article_title_for_dedupe(stripped or text)
        if normalized:
            return normalized
        return normalize_text_token(stripped or text)

    def canonicalize_full(value: str) -> str:
        text = news_crawler.clean_text(str(value or ''))
        normalized = dedupe.normalize_article_title_for_dedupe(text)
        if normalized:
            return normalized
        return normalize_text_token(text)

    if normalized_body_excerpt:
        normalized_body_excerpt = translate_article_summary_fallback(normalized_body_excerpt, translation_target)

    title_candidates: set[str] = set()
    for value in [normalized_title, normalized_title_original]:
        for candidate in [canonicalize(value), canonicalize_full(value)]:
            if candidate:
                title_candidates.add(candidate)
    normalized_summary_key = canonicalize(normalized_summary)
    normalized_body_excerpt_key = canonicalize(normalized_body_excerpt)
    normalized_summary_full_key = canonicalize_full(normalized_summary)
    normalized_body_excerpt_full_key = canonicalize_full(normalized_body_excerpt)

    if normalized_summary and normalized_summary_key not in title_candidates and normalized_summary_full_key not in title_candidates:
        return normalized_summary
    if normalized_body_excerpt and normalized_body_excerpt_key not in title_candidates and normalized_body_excerpt_full_key not in title_candidates:
        return normalized_body_excerpt
    return normalized_summary or normalized_body_excerpt


def format_article_published_date(value: Any) -> str:
    raw = str(value or '').strip()
    if not raw:
        return ''

    date_match = re.match(r'^(\d{4}-\d{2}-\d{2})(?:[T\s].*)?$', raw)
    if date_match:
        return date_match.group(1)

    normalized = raw.replace('Z', '+00:00') if raw.endswith('Z') else raw
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return raw
    return parsed.date().isoformat()


@lru_cache(maxsize=256)
def fetch_article_summary_fallback(
    article_url: str,
    published_at: str = '',
    title_hint: str = '',
    translation_target: str = '',
) -> str:
    normalized_url = str(article_url or '').strip()
    if not normalized_url:
        return ''

    published_dt = news_crawler.parse_dt(published_at) if published_at else None
    if published_dt is None:
        start = news_crawler.parse_dt('2000-01-01')
        end = news_crawler.parse_dt('2100-12-31')
    else:
        start = published_dt - timedelta(days=1)
        end = published_dt + timedelta(days=1)

    if start is None or end is None:
        return ''

    session = news_crawler.build_session()
    try:
        meta = xlsx_source_test.extract_article_metadata(
            normalized_url,
            session,
            start,
            end,
            published_at_hint=published_dt,
            title_hint=title_hint,
            source_discovery='article_browser_fallback',
        )
    except Exception:
        return ''

    if not isinstance(meta, dict):
        return ''
    return normalize_article_card_summary(
        title=str(meta.get('title') or title_hint or '').strip(),
        title_original=str(title_hint or '').strip(),
        summary=str(meta.get('summary') or '').strip(),
        body_excerpt=str(meta.get('body_excerpt') or '').strip(),
        translation_target=translation_target,
    )


def load_article_rows(
    csv_path: Path,
    *,
    selected_platform: str = '',
    keyword: str = '',
    star_filter: str = '',
    limit: int | None = 50,
    hydrate_remote_summary: bool = False,
) -> tuple[list[dict[str, Any]], list[str], int, int, int]:
    if not csv_path.exists():
        return [], [], 0, 0, 0

    def merge_pipe_text(left: str, right: str) -> str:
        values: list[str] = []
        for raw in [left, right]:
            for item in re.split(r'\s*\|\s*', str(raw or '').strip()):
                normalized = item.strip()
                if normalized and normalized not in values:
                    values.append(normalized)
        return ' | '.join(values)

    def merge_brand_text(left: str, right: str) -> str:
        values: list[str] = []
        for raw in [left, right]:
            for item in [part.strip() for part in str(raw or '').split(',')]:
                if item and item not in values:
                    values.append(item)
        return ', '.join(values)

    rows: list[dict[str, Any]] = []
    grouped_rows: dict[str, dict[str, Any]] = {}
    platform_values: set[str] = set()
    keyword_lower = keyword.strip().lower()
    file_total_rows = 0
    starred_ids = set(read_article_star_store().keys())
    file_starred_rows = 0

    with csv_path.open('r', encoding='utf-8-sig', newline='') as handle:
        reader = csv.DictReader(handle)
        for raw_row in reader:
            if not raw_row:
                continue
            file_total_rows += 1
            matched_brands = xlsx_source_test.normalize_string_list(raw_row.get('matched_brands'))
            if matched_brands:
                platform_label = dedupe.format_brand_labels(matched_brands)
                for brand in matched_brands:
                    platform_values.add(brand)
            else:
                platform_label = (raw_row.get('platform_label') or raw_row.get('brand') or raw_row.get('platform') or '').strip()
            if platform_label:
                platform_values.add(platform_label)

            title_original = (raw_row.get('title') or '').strip()
            title_translated = (raw_row.get('title_translated') or '').strip()
            title = (title_translated or title_original or '').strip()
            summary = (raw_row.get('summary_translated') or raw_row.get('summary') or '').strip()
            raw_summary_fallback = (raw_row.get('body_excerpt') or '').strip()
            source_name = (raw_row.get('source_name') or '').strip()
            article_url = (raw_row.get('verification_final_url') or raw_row.get('article_url') or '').strip()
            source_url = (raw_row.get('source_url') or '').strip()
            published_at = (raw_row.get('published_at') or '').strip()
            published_at_display = format_article_published_date(published_at)
            translation_target = (raw_row.get('translation_target') or '').strip()
            survey_dimensions = (raw_row.get('survey_dimensions') or '').strip()
            survey_question_ids = (raw_row.get('survey_question_ids') or '').strip()
            survey_indicator_examples = (raw_row.get('survey_indicator_examples') or '').strip()
            survey_ai_reason_raw = (raw_row.get('survey_ai_reason_raw') or '').strip()
            survey_ai_reason_translated = (raw_row.get('survey_ai_reason_translated') or '').strip()
            survey_filter_confidence = (raw_row.get('survey_filter_confidence') or '').strip().lower()
            if survey_filter_confidence not in {'high', 'medium', 'low'}:
                survey_filter_confidence = ''
            volume_fill = str(raw_row.get('volume_fill') or '').strip().lower() in {'1', 'true', 'yes', 'y'}
            raw_briefing_sentiment = (raw_row.get('briefing_sentiment') or '').strip()
            briefing_sentiment = normalize_briefing_sentiment(raw_briefing_sentiment) if raw_briefing_sentiment else ''
            briefing_sentiment_reason = (raw_row.get('briefing_sentiment_reason') or '').strip()
            industry_trend_flag = xlsx_source_test.normalize_industry_trend_flag(raw_row.get('industry_trend_flag'))
            industry_trend_category = xlsx_source_test.normalize_industry_trend_category(raw_row.get('industry_trend_category')) if industry_trend_flag else ''
            industry_trend_impact = xlsx_source_test.normalize_industry_trend_impact(raw_row.get('industry_trend_impact')) if industry_trend_flag else ''
            industry_trend_reason = (raw_row.get('industry_trend_reason') or '').strip()
            manual_added = str(raw_row.get('manual_added') or '').strip().lower() in {'1', 'true', 'yes', 'y'}
            manual_article_id = (raw_row.get('manual_article_id') or '').strip()
            title_display_zh = translate_article_title_for_display(
                title_original,
                title_translated,
                translation_target,
            )
            summary = normalize_article_card_summary(
                title=title,
                title_original=title_original,
                summary=summary,
                body_excerpt=raw_summary_fallback,
                translation_target=translation_target,
            )
            if hydrate_remote_summary and not summary and article_url:
                summary = fetch_article_summary_fallback(article_url, published_at, title_original or title, translation_target)
            article_id = build_article_star_key(
                platform_label=platform_label,
                title=title,
                title_original=title_original,
                source_name=source_name,
                article_url=article_url,
                source_url=source_url,
                published_at=published_at,
            )
            is_starred = article_id in starred_ids
            if is_starred:
                file_starred_rows += 1

            haystack = ' '.join(
                item for item in [platform_label, title, title_original, title_display_zh, summary, source_name, published_at, survey_dimensions, survey_question_ids, briefing_sentiment, industry_trend_category, industry_trend_impact, industry_trend_reason] if item
            ).lower()
            if keyword_lower and keyword_lower not in haystack:
                continue

            item = {
                'article_id': article_id,
                'starred': is_starred,
                'platform_label': platform_label,
                'title': title,
                'title_original': title_original,
                'title_display_zh': title_display_zh,
                'summary': summary,
                'source_name': source_name,
                'source_url': source_url,
                'article_url': article_url,
                'published_at': published_at,
                'published_at_display': published_at_display,
                'category': (raw_row.get('category') or '').strip(),
                'verification_status': (raw_row.get('verification_status') or '').strip(),
                'survey_dimensions': survey_dimensions,
                'survey_question_ids': survey_question_ids,
                'survey_indicator_examples': survey_indicator_examples,
                'survey_ai_reason_raw': survey_ai_reason_raw,
                'survey_ai_reason_translated': survey_ai_reason_translated,
                'survey_filter_confidence': survey_filter_confidence,
                'volume_fill': volume_fill,
                'briefing_sentiment': briefing_sentiment,
                'briefing_sentiment_reason': briefing_sentiment_reason,
                'industry_trend_flag': industry_trend_flag,
                'industry_trend_category': industry_trend_category,
                'industry_trend_impact': industry_trend_impact,
                'industry_trend_reason': industry_trend_reason,
                'manual_added': manual_added,
                'manual_article_id': manual_article_id,
            }
            group_key = article_url or '||'.join([published_at, title_original or title, source_url])
            existing = grouped_rows.get(group_key)
            if existing is None:
                grouped_rows[group_key] = item
                continue
            existing['platform_label'] = merge_brand_text(existing.get('platform_label', ''), item['platform_label'])
            existing['starred'] = bool(existing.get('starred')) or bool(item['starred'])
            if len(item['summary']) > len(existing.get('summary', '')):
                existing['summary'] = item['summary']
            if len(item['title']) > len(existing.get('title', '')):
                existing['title'] = item['title']
            if len(item['title_original']) > len(existing.get('title_original', '')):
                existing['title_original'] = item['title_original']
            if len(item['title_display_zh']) > len(existing.get('title_display_zh', '')):
                existing['title_display_zh'] = item['title_display_zh']
            existing['source_name'] = existing.get('source_name') or item['source_name']
            existing['source_url'] = existing.get('source_url') or item['source_url']
            existing['article_url'] = existing.get('article_url') or item['article_url']
            existing['published_at'] = existing.get('published_at') or item['published_at']
            existing['published_at_display'] = existing.get('published_at_display') or item['published_at_display']
            existing['category'] = merge_pipe_text(existing.get('category', ''), item['category'])
            existing['verification_status'] = merge_pipe_text(existing.get('verification_status', ''), item['verification_status'])
            existing['survey_dimensions'] = merge_pipe_text(existing.get('survey_dimensions', ''), item['survey_dimensions'])
            existing['survey_question_ids'] = merge_pipe_text(existing.get('survey_question_ids', ''), item['survey_question_ids'])
            existing['survey_indicator_examples'] = merge_pipe_text(existing.get('survey_indicator_examples', ''), item['survey_indicator_examples'])
            existing['survey_ai_reason_raw'] = merge_pipe_text(existing.get('survey_ai_reason_raw', ''), item['survey_ai_reason_raw'])
            existing['survey_ai_reason_translated'] = merge_pipe_text(existing.get('survey_ai_reason_translated', ''), item['survey_ai_reason_translated'])
            if not existing.get('survey_filter_confidence') and item.get('survey_filter_confidence'):
                existing['survey_filter_confidence'] = item['survey_filter_confidence']
            existing['volume_fill'] = bool(existing.get('volume_fill')) or bool(item.get('volume_fill'))
            existing['manual_added'] = bool(existing.get('manual_added')) or bool(item.get('manual_added'))
            existing['manual_article_id'] = existing.get('manual_article_id') or item.get('manual_article_id')
            if (not existing.get('briefing_sentiment') or existing.get('briefing_sentiment') == 'Neutral') and item.get('briefing_sentiment') and item.get('briefing_sentiment') != 'Neutral':
                existing['briefing_sentiment'] = item['briefing_sentiment']
            existing['briefing_sentiment_reason'] = merge_pipe_text(existing.get('briefing_sentiment_reason', ''), item['briefing_sentiment_reason'])
            existing['industry_trend_flag'] = bool(existing.get('industry_trend_flag')) or bool(item.get('industry_trend_flag'))
            existing['industry_trend_category'] = merge_pipe_text(existing.get('industry_trend_category', ''), item['industry_trend_category'])
            if (not existing.get('industry_trend_impact') or existing.get('industry_trend_impact') == 'Neutral') and item.get('industry_trend_impact') and item.get('industry_trend_impact') != 'Neutral':
                existing['industry_trend_impact'] = item['industry_trend_impact']
            existing['industry_trend_reason'] = merge_pipe_text(existing.get('industry_trend_reason', ''), item['industry_trend_reason'])
            existing['article_id'] = build_article_star_key(
                platform_label=existing.get('platform_label', ''),
                title=existing.get('title', ''),
                title_original=existing.get('title_original', ''),
                source_name=existing.get('source_name', ''),
                article_url=existing.get('article_url', ''),
                source_url=existing.get('source_url', ''),
                published_at=existing.get('published_at', ''),
            )

    rows = list(grouped_rows.values())
    platform_values = set()
    for row in rows:
        label = str(row.get('platform_label') or '').strip()
        if not label:
            continue
        platform_values.add(label)
        for brand in [item.strip() for item in label.split(',') if item.strip()]:
            platform_values.add(brand)
    if selected_platform:
        rows = [
            row for row in rows
            if selected_platform in [item.strip() for item in str(row.get('platform_label') or '').split(',') if item.strip()]
        ]
    if star_filter == 'starred':
        rows = [row for row in rows if row.get('starred')]
    elif star_filter == 'unstarred':
        rows = [row for row in rows if not row.get('starred')]
    rows.sort(key=lambda row: row['published_at'], reverse=True)
    matching_rows = len(rows)
    if limit is None:
        return rows, sorted(platform_values), file_total_rows, matching_rows, file_starred_rows
    return rows[:limit], sorted(platform_values), file_total_rows, matching_rows, file_starred_rows


def load_article_rows_from_db_source(
    source_value: str,
    *,
    selected_platform: str = '',
    keyword: str = '',
    star_filter: str = '',
    limit: int | None = 50,
    hydrate_remote_summary: bool = False,
) -> tuple[list[dict[str, Any]], list[str], int, int, int]:
    parsed = parse_db_article_source(source_value)
    if parsed is None:
        return [], [], 0, 0, 0
    run_id, stage = parsed
    raw_rows = db_store.load_articles(run_id, stage)
    if stage == 'after':
        raw_rows = raw_rows + db_store.load_manual_articles(run_id, stage)
    if not raw_rows:
        return [], [], 0, 0, 0

    fieldnames: list[str] = []
    for row in raw_rows:
        for key in row:
            if key not in fieldnames:
                fieldnames.append(key)
    if not fieldnames:
        return [], [], 0, 0, 0

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile('w', newline='', encoding='utf-8-sig', suffix='.csv', delete=False) as handle:
            temp_path = Path(handle.name)
            writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(raw_rows)
        return load_article_rows(
            temp_path,
            selected_platform=selected_platform,
            keyword=keyword,
            star_filter=star_filter,
            limit=limit,
            hydrate_remote_summary=hydrate_remote_summary,
        )
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass


def article_export_row_score(row: dict[str, Any]) -> tuple[int, int, int, int, int]:
    summary = str(row.get('summary') or '')
    return (
        1 if summary else 0,
        1 if row.get('survey_dimensions') else 0,
        1 if row.get('source_name') else 0,
        len(summary),
        len(str(row.get('title') or '')),
    )


def format_starred_updated_at(value: Any) -> str:
    try:
        timestamp = int(value)
    except (TypeError, ValueError):
        return ''
    try:
        return datetime.fromtimestamp(timestamp).astimezone().isoformat()
    except Exception:
        return str(timestamp)


def build_starred_export_rows(
    *,
    csv_path: Path | None = None,
    article_source: str = '',
    selected_platform: str = '',
    keyword: str = '',
    country_code: str = DEFAULT_COUNTRY_CODE,
) -> list[dict[str, Any]]:
    starred_store = read_article_star_store()
    if not starred_store:
        return []

    if article_source and is_db_article_source(article_source):
        scoped_rows, _, _, _, _ = load_article_rows_from_db_source(
            article_source,
            selected_platform=selected_platform,
            keyword=keyword,
            star_filter='starred',
            limit=None,
        )
        source_label = article_source
    elif csv_path is not None:
        scoped_rows, _, _, _, _ = load_article_rows(
            csv_path,
            selected_platform=selected_platform,
            keyword=keyword,
            star_filter='starred',
            limit=None,
        )
        source_label = build_article_csv_display_label(csv_path) if csv_path.exists() else ''
        export_rows = [
            {
                'article_id': row['article_id'],
                'starred_updated_at': format_starred_updated_at((starred_store.get(row['article_id']) or {}).get('updated_at')),
                'platform_label': row.get('platform_label') or '',
                'title': row.get('title') or '',
                'title_original': row.get('title_original') or '',
                'summary': row.get('summary') or '',
                'source_name': row.get('source_name') or '',
                'source_url': row.get('source_url') or '',
                'article_url': row.get('article_url') or '',
                'published_at': row.get('published_at') or '',
                'category': row.get('category') or '',
                'survey_dimensions': row.get('survey_dimensions') or '',
                'survey_question_ids': row.get('survey_question_ids') or '',
                'survey_indicator_examples': row.get('survey_indicator_examples') or '',
                'survey_ai_reason_raw': row.get('survey_ai_reason_raw') or '',
                'survey_ai_reason_translated': row.get('survey_ai_reason_translated') or '',
                'survey_filter_confidence': row.get('survey_filter_confidence') or '',
                'volume_fill': 'true' if row.get('volume_fill') else '',
                'briefing_sentiment': row.get('briefing_sentiment') or '',
                'briefing_sentiment_reason': row.get('briefing_sentiment_reason') or '',
                'source_file': source_label,
            }
            for row in scoped_rows
        ]
        export_rows.sort(
            key=lambda row: (
                str(row.get('starred_updated_at') or ''),
                str(row.get('published_at') or ''),
                str(row.get('title') or ''),
            ),
            reverse=True,
        )
        return export_rows

    export_rows_by_id: dict[str, dict[str, Any]] = {}
    for csv_path in list_article_csv_paths(limit=None, country_code=country_code):
        csv_label = build_article_csv_display_label(csv_path)
        try:
            handle = csv_path.open('r', encoding='utf-8-sig', newline='')
        except Exception:
            continue

        with handle:
            reader = csv.DictReader(handle)
            for raw_row in reader:
                if not raw_row:
                    continue

                platform_label = (raw_row.get('platform_label') or raw_row.get('brand') or raw_row.get('platform') or '').strip()
                title = (raw_row.get('title_translated') or raw_row.get('title') or '').strip()
                title_original = (raw_row.get('title') or '').strip()
                source_name = (raw_row.get('source_name') or '').strip()
                article_url = (raw_row.get('verification_final_url') or raw_row.get('article_url') or '').strip()
                source_url = (raw_row.get('source_url') or '').strip()
                published_at = (raw_row.get('published_at') or '').strip()
                article_id = build_article_star_key(
                    platform_label=platform_label,
                    title=title,
                    title_original=title_original,
                    source_name=source_name,
                    article_url=article_url,
                    source_url=source_url,
                    published_at=published_at,
                )
                star_meta = starred_store.get(article_id)
                if not star_meta:
                    continue

                translation_target = (raw_row.get('translation_target') or '').strip()
                summary = (raw_row.get('summary_translated') or raw_row.get('summary') or '').strip()
                raw_summary_fallback = (raw_row.get('body_excerpt') or '').strip()
                if not summary and raw_summary_fallback:
                    summary = translate_article_summary_fallback(raw_summary_fallback, translation_target)
                if not summary and article_url:
                    summary = fetch_article_summary_fallback(article_url, published_at, title_original or title, translation_target)

                export_row = {
                    'article_id': article_id,
                    'starred_updated_at': format_starred_updated_at(star_meta.get('updated_at')),
                    'platform_label': platform_label,
                    'title': title,
                    'title_original': title_original,
                    'summary': summary,
                    'source_name': source_name,
                    'source_url': source_url,
                    'article_url': article_url,
                    'published_at': published_at,
                    'category': (raw_row.get('category') or '').strip(),
                    'survey_dimensions': (raw_row.get('survey_dimensions') or '').strip(),
                    'survey_question_ids': (raw_row.get('survey_question_ids') or '').strip(),
                    'survey_indicator_examples': (raw_row.get('survey_indicator_examples') or '').strip(),
                    'survey_ai_reason_raw': (raw_row.get('survey_ai_reason_raw') or '').strip(),
                    'survey_ai_reason_translated': (raw_row.get('survey_ai_reason_translated') or '').strip(),
                    'briefing_sentiment': normalize_briefing_sentiment(raw_row.get('briefing_sentiment')),
                    'briefing_sentiment_reason': (raw_row.get('briefing_sentiment_reason') or '').strip(),
                    'industry_trend_flag': 'true' if xlsx_source_test.normalize_industry_trend_flag(raw_row.get('industry_trend_flag')) else '',
                    'industry_trend_category': (raw_row.get('industry_trend_category') or '').strip(),
                    'industry_trend_impact': xlsx_source_test.normalize_industry_trend_impact(raw_row.get('industry_trend_impact')) if raw_row.get('industry_trend_impact') else '',
                    'industry_trend_reason': (raw_row.get('industry_trend_reason') or '').strip(),
                    'source_file': csv_label,
                }
                existing = export_rows_by_id.get(article_id)
                if existing is None or article_export_row_score(export_row) > article_export_row_score(existing):
                    export_rows_by_id[article_id] = export_row

    for article_id, star_meta in starred_store.items():
        if article_id in export_rows_by_id:
            continue
        export_rows_by_id[article_id] = {
            'article_id': article_id,
            'starred_updated_at': format_starred_updated_at(star_meta.get('updated_at')),
            'platform_label': '',
            'title': str(star_meta.get('title') or ''),
            'title_original': '',
            'summary': '',
            'source_name': '',
            'source_url': '',
            'article_url': str(star_meta.get('article_url') or ''),
            'published_at': '',
            'category': '',
            'survey_dimensions': '',
            'survey_question_ids': '',
            'survey_indicator_examples': '',
            'survey_ai_reason_raw': '',
            'survey_ai_reason_translated': '',
            'briefing_sentiment': '',
            'briefing_sentiment_reason': '',
            'industry_trend_flag': '',
            'industry_trend_category': '',
            'industry_trend_impact': '',
            'industry_trend_reason': '',
            'source_file': '',
        }

    export_rows = list(export_rows_by_id.values())
    export_rows.sort(
        key=lambda row: (
            str(row.get('starred_updated_at') or ''),
            str(row.get('published_at') or ''),
            str(row.get('title') or ''),
        ),
        reverse=True,
    )
    return export_rows

