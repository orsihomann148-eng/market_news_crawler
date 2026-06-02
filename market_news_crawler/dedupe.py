from __future__ import annotations

import difflib
import json
import re
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Any, Callable

from news_crawler import clean_text, trim_summary_text


AI_DEDUPE_DATE_WINDOW_DAYS = 14


def merge_string_values(values: list[str]) -> str:
    ordered: list[str] = []
    for value in values:
        normalized = clean_text(value)
        if normalized and normalized not in ordered:
            ordered.append(normalized)
    return " | ".join(ordered)


def merge_list_values(values: list[Any]) -> list[str]:
    ordered: list[str] = []
    for value in values:
        if isinstance(value, list):
            for item in value:
                normalized = clean_text(item)
                if normalized and normalized not in ordered:
                    ordered.append(normalized)
        else:
            normalized = clean_text(value)
            if normalized and normalized not in ordered:
                ordered.append(normalized)
    return ordered


def article_brand_labels(row: dict[str, Any]) -> list[str]:
    matched = merge_list_values([row.get('matched_brands', [])])
    if matched:
        return matched
    fallback = [
        clean_text(row.get('platform_label') or ''),
        clean_text(row.get('platform') or ''),
        clean_text(row.get('source_platform') or ''),
    ]
    return [item for item in fallback if item]


def format_brand_labels(labels: list[str]) -> str:
    ordered = [clean_text(item) for item in labels if clean_text(item)]
    deduped: list[str] = []
    for item in ordered:
        if item not in deduped:
            deduped.append(item)
    return ', '.join(sorted(deduped, key=lambda item: item.casefold()))


def merge_article_group(items: list[dict[str, Any]]) -> dict[str, Any]:
    ranked_items = sorted(
        items,
        key=lambda item: (
            len(clean_text(item.get('summary_translated') or item.get('summary') or '')),
            len(clean_text(item.get('title_translated') or item.get('title') or '')),
            clean_text(item.get('published_at') or ''),
        ),
        reverse=True,
    )
    base = dict(ranked_items[0])
    base['side'] = merge_string_values([item.get('side', '') for item in items])
    base['source_platform'] = merge_string_values([item.get('source_platform', '') for item in items])
    base['source_url'] = merge_string_values([item.get('source_url', '') for item in items])
    base['source_final_url'] = merge_string_values([item.get('source_final_url', '') for item in items])
    base['source_site'] = merge_string_values([item.get('source_site', '') for item in items])
    base['matched_brands'] = merge_list_values([item.get('matched_brands', []) for item in items])
    if base['matched_brands']:
        base['platform_label'] = format_brand_labels(base['matched_brands'])
        base['platform'] = base['platform_label']
    else:
        base['platform_label'] = clean_text(base.get('platform_label') or base.get('platform') or '')
    base['survey_dimensions'] = merge_string_values([item.get('survey_dimensions', '') for item in items])
    base['survey_question_ids'] = merge_string_values([item.get('survey_question_ids', '') for item in items])
    base['survey_indicator_examples'] = merge_string_values([item.get('survey_indicator_examples', '') for item in items])
    base['survey_ai_reason_raw'] = merge_string_values([item.get('survey_ai_reason_raw', '') for item in items])
    base['survey_ai_reason_translated'] = merge_string_values([item.get('survey_ai_reason_translated', '') for item in items])
    return base


def normalize_article_title_for_dedupe(value: str) -> str:
    normalized = clean_text(value).lower()
    normalized = re.sub(r'\s*[|〜~\-–—:：]\s*', ' ', normalized)
    normalized = re.sub(r'[^0-9a-z\u4e00-\u9fff\u3040-\u30ff]+', ' ', normalized)
    normalized = re.sub(r'\s+', ' ', normalized).strip()
    return normalized


def strip_title_source_suffix(value: str) -> str:
    text = clean_text(value)
    if not text:
        return ''
    text = re.sub(r'\s+(?:-|\||–|—)\s+[^-–—|]{2,80}$', '', text).strip() or text
    text = re.sub(r'\s*[（(][^（）()]{2,40}[）)]\s*$', '', text).strip()
    text = re.split(r'\s+[|｜]\s+', text, maxsplit=1)[0].strip()
    text = re.split(r'\s+[-–—]\s+', text, maxsplit=1)[0].strip()
    text = re.sub(r'\s*(?:-|\||｜)\s*(?:Yahoo!?ニュース|雅虎!?消息|Yahoo!? News).*$', '', text, flags=re.IGNORECASE).strip()
    return text


def normalize_article_url_for_dedupe(value: str) -> str:
    normalized = clean_text(value)
    if not normalized:
        return ''
    try:
        parsed = urllib.parse.urlsplit(normalized)
    except Exception:
        return normalized.rstrip('/')
    if not parsed.scheme or not parsed.netloc:
        return normalized.rstrip('/')
    query_items = []
    for key, raw_value in urllib.parse.parse_qsl(parsed.query, keep_blank_values=True):
        lowered_key = key.lower()
        if lowered_key.startswith('utm_') or lowered_key in {'fbclid', 'gclid', 'yclid', 'mc_cid', 'mc_eid'}:
            continue
        query_items.append((key, raw_value))
    query = urllib.parse.urlencode(sorted(query_items), doseq=True)
    path = parsed.path or '/'
    if path != '/':
        path = path.rstrip('/')
    return urllib.parse.urlunsplit(
        (
            parsed.scheme.lower(),
            parsed.netloc.lower(),
            path,
            query,
            '',
        )
    )


def source_key_for_initial_dedupe(row: dict[str, Any]) -> str:
    for key in ['source_site', 'source_url', 'source_final_url']:
        value = clean_text(row.get(key) or '')
        if value:
            if key in {'source_url', 'source_final_url'}:
                return normalize_article_url_for_dedupe(value) or value
            return value.lower()
    return ''


def article_duplicate_cluster_key(row: dict[str, Any]) -> tuple[str, str, str] | None:
    platform = format_brand_labels(article_brand_labels(row))
    published_at = clean_text(row.get('published_at') or '')
    published_date = published_at[:10] if len(published_at) >= 10 else published_at
    title = clean_text(row.get('title_translated') or row.get('title') or '')
    normalized_title = normalize_article_title_for_dedupe(strip_title_source_suffix(title) or title)
    if not platform or not published_date or len(normalized_title) < 12:
        return None
    return platform, published_date, normalized_title


def initial_dedupe_exact_key(row: dict[str, Any]) -> tuple[str, ...]:
    article_url = normalize_article_url_for_dedupe(clean_text(row.get('article_url') or ''))
    if article_url:
        return ('url', article_url)

    platform = format_brand_labels(article_brand_labels(row)) or clean_text(row.get('platform_label') or row.get('platform') or '')
    published_at = clean_text(row.get('published_at') or '')
    published_date = published_at[:10] if len(published_at) >= 10 else published_at
    title = clean_text(row.get('title_translated') or row.get('title') or '')
    normalized_title = normalize_article_title_for_dedupe(strip_title_source_suffix(title) or title)
    source_key = source_key_for_initial_dedupe(row)
    if platform and published_date and normalized_title:
        return ('fallback', platform, published_date, normalized_title, source_key)
    return (
        'weak',
        platform,
        published_date,
        normalized_title,
        source_key,
        clean_text(row.get('summary') or row.get('summary_translated') or '')[:120],
    )


def add_initial_dedupe_stats_group(
    stats: dict[str, Any] | None,
    *,
    reason: str,
    items: list[dict[str, Any]],
) -> None:
    if stats is None or len(items) <= 1:
        return
    removed = len(items) - 1
    stats['removed_count'] = int(stats.get('removed_count', 0) or 0) + removed
    removed_by_reason = stats.setdefault('removed_by_reason', {})
    removed_by_reason[reason] = int(removed_by_reason.get(reason, 0) or 0) + removed
    samples = stats.setdefault('removed_samples', [])
    if len(samples) >= 30:
        return
    samples.append(
        {
            'reason': reason,
            'group_size': len(items),
            'titles': [clean_text(item.get('title') or item.get('title_translated') or '') for item in items[:5]],
            'article_urls': [clean_text(item.get('article_url') or '') for item in items[:5]],
            'source_sites': [clean_text(item.get('source_site') or '') for item in items[:5]],
            'published_at': [clean_text(item.get('published_at') or '') for item in items[:5]],
        }
    )


def dedupe_articles(rows: list[dict[str, Any]], stats: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    if stats is not None:
        stats.clear()
        stats.update(
            {
                'input_count': len(rows),
                'missing_url_count': sum(1 for row in rows if not clean_text(row.get('article_url') or '')),
                'group_count': 0,
                'removed_count': 0,
                'removed_by_reason': {},
                'removed_samples': [],
            }
        )

    exact_grouped: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for row in rows:
        key = initial_dedupe_exact_key(row)
        exact_grouped.setdefault(key, []).append(row)

    if stats is not None:
        stats['group_count'] = len(exact_grouped)

    exact_deduped = []
    for key, items in exact_grouped.items():
        if key and key[0] == 'url':
            add_initial_dedupe_stats_group(stats, reason='url_duplicate', items=items)
        elif key and key[0] == 'fallback':
            add_initial_dedupe_stats_group(stats, reason='same_source_title_duplicate', items=items)
        elif key and key[0] == 'weak':
            add_initial_dedupe_stats_group(stats, reason='weak_key_duplicate', items=items)
        exact_deduped.append(merge_article_group(items))

    clustered_groups: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    passthrough_rows: list[dict[str, Any]] = []
    for row in exact_deduped:
        cluster_key = article_duplicate_cluster_key(row)
        if cluster_key is None:
            passthrough_rows.append(row)
            continue
        clustered_groups.setdefault(cluster_key, []).append(row)

    deduped = list(passthrough_rows)
    for items in clustered_groups.values():
        add_initial_dedupe_stats_group(stats, reason='same_day_exact_title_duplicate', items=items)
        deduped.append(merge_article_group(items))

    deduped.sort(key=lambda row: (row['platform'], row['source_site'], row['published_at']), reverse=True)
    if stats is not None:
        stats['output_count'] = len(deduped)
        stats['removed_count'] = len(rows) - len(deduped)
    return deduped


def published_date_key(row: dict[str, Any]) -> str:
    published_at = clean_text(row.get('published_at') or '')
    return published_at[:10] if len(published_at) >= 10 else published_at


def title_token_set_for_dedupe(value: str) -> set[str]:
    normalized = normalize_article_title_for_dedupe(value)
    return {token for token in normalized.split() if len(token) >= 2}


def summary_token_set_for_dedupe(value: str) -> set[str]:
    normalized = normalize_article_title_for_dedupe(value)
    return {
        token
        for token in normalized.split()
        if len(token) >= 3 and token not in {'news', 'report', 'update', 'media'}
    }


def dedupe_brand_keys(row: dict[str, Any]) -> set[str]:
    keys: set[str] = set()
    for raw_value in [
        row.get('platform'),
        row.get('platform_label'),
        row.get('source_platform'),
        *(row.get('matched_brands') or []),
    ]:
        normalized = clean_text(raw_value).lower()
        if normalized:
            keys.add(normalized)
    return keys


def event_text_for_dedupe(row: dict[str, Any]) -> str:
    parts = [
        clean_text(row.get('title_translated') or row.get('title') or ''),
        clean_text(row.get('title') or ''),
        clean_text(row.get('summary_translated') or row.get('summary') or ''),
        clean_text(row.get('summary') or ''),
    ]
    return ' '.join(part for part in parts if part)


def dates_are_close_for_dedupe(left_date: str, right_date: str, max_days: int = AI_DEDUPE_DATE_WINDOW_DAYS) -> bool:
    if not left_date or not right_date:
        return False
    if left_date == right_date:
        return True
    try:
        left_dt = datetime.fromisoformat(left_date[:10])
        right_dt = datetime.fromisoformat(right_date[:10])
    except ValueError:
        return False
    return abs((left_dt - right_dt).days) <= max_days


def title_candidates_for_dedupe(row: dict[str, Any]) -> list[str]:
    candidates = [
        clean_text(row.get('title_translated') or ''),
        clean_text(row.get('title') or ''),
    ]
    return [candidate for candidate in candidates if candidate]


def significant_number_tokens_for_dedupe(value: str) -> set[str]:
    tokens: set[str] = set()
    for raw_number in re.findall(r'\d[\d.,\s]*\d|\d+', clean_text(value)):
        normalized = re.sub(r'\D+', '', raw_number)
        if len(normalized) < 3:
            continue
        if normalized.startswith(('19', '20')) and len(normalized) == 4:
            continue
        tokens.add(normalized)
    return tokens


def event_phrase_keys_for_dedupe(value: str) -> set[str]:
    normalized = normalize_article_title_for_dedupe(value)
    phrase_groups = {
        'cash_payment': [
            'cash payment',
            'cash payments',
            'pagamenti in contanti',
            'pagamento in contanti',
            'paga in contanti',
            'contanti',
            '现金支付',
        ],
        'payment_points': [
            'punti vendita',
            'punti fisici',
            'sales points',
            'physical points',
            'payment points',
            '实体点',
            '销售点',
            '支付点',
            '付款点',
        ],
        'promotion_event': [
            'prime day',
            'black friday',
            'cyber monday',
            'spring sale',
            'offerte di primavera',
            '促销',
            '折扣',
            '优惠',
        ],
        'investigation_or_penalty': [
            'investigation',
            'indagine',
            'fine',
            'penalty',
            'sanzione',
            '调查',
            '罚款',
            '处罚',
        ],
    }
    keys: set[str] = set()
    for key, phrases in phrase_groups.items():
        for phrase in phrases:
            if normalize_article_title_for_dedupe(phrase) in normalized:
                keys.add(key)
                break
    return keys


def event_markers_look_like_duplicates(row_a: dict[str, Any], row_b: dict[str, Any]) -> bool:
    event_a = event_text_for_dedupe(row_a)
    event_b = event_text_for_dedupe(row_b)
    shared_numbers = significant_number_tokens_for_dedupe(event_a) & significant_number_tokens_for_dedupe(event_b)
    if not shared_numbers:
        return False
    shared_phrases = event_phrase_keys_for_dedupe(event_a) & event_phrase_keys_for_dedupe(event_b)
    if shared_phrases:
        return True
    tokens_a = summary_token_set_for_dedupe(event_a)
    tokens_b = summary_token_set_for_dedupe(event_b)
    return len(tokens_a & tokens_b) >= 4


def titles_look_like_duplicates(title_a: str, title_b: str) -> bool:
    normalized_a = normalize_article_title_for_dedupe(title_a)
    normalized_b = normalize_article_title_for_dedupe(title_b)
    primary_a = normalize_article_title_for_dedupe(strip_title_source_suffix(title_a))
    primary_b = normalize_article_title_for_dedupe(strip_title_source_suffix(title_b))
    if not normalized_a or not normalized_b:
        return False
    if normalized_a == normalized_b:
        return True
    if primary_a and primary_b and primary_a == primary_b:
        return True
    if min(len(normalized_a), len(normalized_b)) >= 18 and (normalized_a in normalized_b or normalized_b in normalized_a):
        return True
    if primary_a and primary_b and min(len(primary_a), len(primary_b)) >= 16 and (primary_a in primary_b or primary_b in primary_a):
        return True
    ratio = difflib.SequenceMatcher(None, normalized_a, normalized_b).ratio()
    primary_ratio = difflib.SequenceMatcher(None, primary_a or normalized_a, primary_b or normalized_b).ratio()
    tokens_a = title_token_set_for_dedupe(normalized_a)
    tokens_b = title_token_set_for_dedupe(normalized_b)
    if not tokens_a or not tokens_b:
        return max(ratio, primary_ratio) >= 0.82
    intersection = len(tokens_a & tokens_b)
    union = len(tokens_a | tokens_b)
    jaccard = (intersection / union) if union else 0.0
    if max(ratio, primary_ratio) >= 0.86:
        return True
    if primary_ratio >= 0.76:
        return True
    if max(ratio, primary_ratio) >= 0.74 and jaccard >= 0.55:
        return True
    if jaccard >= 0.82:
        return True
    return False


def event_texts_look_like_duplicates(row_a: dict[str, Any], row_b: dict[str, Any]) -> bool:
    title_a = clean_text(row_a.get('title_translated') or row_a.get('title') or '')
    title_b = clean_text(row_b.get('title_translated') or row_b.get('title') or '')
    if titles_look_like_duplicates(title_a, title_b):
        return True
    for candidate_a in title_candidates_for_dedupe(row_a):
        for candidate_b in title_candidates_for_dedupe(row_b):
            if titles_look_like_duplicates(candidate_a, candidate_b):
                return True
    if event_markers_look_like_duplicates(row_a, row_b):
        return True

    event_a = event_text_for_dedupe(row_a)
    event_b = event_text_for_dedupe(row_b)
    if not event_a or not event_b:
        return False

    normalized_a = normalize_article_title_for_dedupe(event_a)
    normalized_b = normalize_article_title_for_dedupe(event_b)
    if not normalized_a or not normalized_b:
        return False

    ratio = difflib.SequenceMatcher(None, normalized_a, normalized_b).ratio()
    title_tokens_a = title_token_set_for_dedupe(title_a)
    title_tokens_b = title_token_set_for_dedupe(title_b)
    tokens_a = summary_token_set_for_dedupe(event_a)
    tokens_b = summary_token_set_for_dedupe(event_b)
    if not tokens_a or not tokens_b:
        return ratio >= 0.84

    title_overlap = len(title_tokens_a & title_tokens_b)
    intersection = len(tokens_a & tokens_b)
    union = len(tokens_a | tokens_b)
    jaccard = (intersection / union) if union else 0.0
    if ratio >= 0.88:
        return True
    if title_overlap >= 3 and intersection >= 4:
        return True
    if title_overlap >= 2 and intersection >= 5:
        return True
    if ratio >= 0.8 and jaccard >= 0.48:
        return True
    if jaccard >= 0.68 and intersection >= 6:
        return True
    return False


def build_ai_dedupe_candidate_clusters(rows: list[dict[str, Any]]) -> list[list[int]]:
    if len(rows) <= 1:
        return []

    parent = list(range(len(rows)))

    def find(index: int) -> int:
        while parent[index] != index:
            parent[index] = parent[parent[index]]
            index = parent[index]
        return index

    def union(left: int, right: int) -> None:
        root_left = find(left)
        root_right = find(right)
        if root_left != root_right:
            parent[root_right] = root_left

    for left in range(len(rows)):
        left_row = rows[left]
        left_brands = dedupe_brand_keys(left_row)
        left_date = published_date_key(left_row)
        if not left_brands or not left_date:
            continue
        for right in range(left + 1, len(rows)):
            right_row = rows[right]
            right_brands = dedupe_brand_keys(right_row)
            right_date = published_date_key(right_row)
            if not right_brands or not right_date:
                continue
            if not (left_brands & right_brands):
                continue
            if not dates_are_close_for_dedupe(left_date, right_date):
                continue
            if event_texts_look_like_duplicates(left_row, right_row):
                union(left, right)

    grouped: dict[int, list[int]] = {}
    for index in range(len(rows)):
        root = find(index)
        grouped.setdefault(root, []).append(index)
    return [indices for indices in grouped.values() if len(indices) > 1]


def build_ai_dedupe_messages(cluster_items: list[tuple[int, dict[str, Any]]]) -> list[dict[str, str]]:
    article_payload = []
    for index, row in cluster_items:
        article_payload.append(
            {
                'article_id': str(index),
                'brand_keys': sorted(dedupe_brand_keys(row)),
                'title': clean_text(row.get('title_translated') or row.get('title') or ''),
                'title_original': clean_text(row.get('title') or ''),
                'summary': trim_summary_text(clean_text(row.get('summary_translated') or row.get('summary') or ''), 320),
                'source_site': clean_text(row.get('source_site') or ''),
                'article_url': clean_text(row.get('article_url') or ''),
                'published_at': clean_text(row.get('published_at') or ''),
            }
        )
    system_prompt = (
        "You are a news deduplication assistant. Identify duplicate reports about the same ecommerce brand "
        "and the same real-world event. Treat articles as duplicates even when they have different source sites, "
        "different publication dates within the candidate window, different wording, or different languages, "
        "if they describe the same feature launch, policy change, payment/logistics/service change, promotion, "
        "investigation, penalty, announcement, survey, or public incident. Keep one representative article, "
        "preferring the clearest title, richer summary, more authoritative source, or article closest to the "
        "original report. Do not merge articles about different events, different countries/markets, different "
        "products, different promotions, different policy details, or different financial-report moments. "
        "Return only a JSON object with fixed fields keep_article_ids and duplicate_groups."
    )
    user_payload = {
        'task': 'Decide which articles are duplicate reports of the same ecommerce-brand event. Keep only one representative article per duplicate event.',
        'articles': article_payload,
        'output_schema': {
            'keep_article_ids': ['0'],
            'duplicate_groups': [
                {
                    'canonical_article_id': '0',
                    'duplicate_article_ids': ['1'],
                    'reason': 'Both articles report the same event for the same brand; the duplicate is a syndicated or rewritten version.',
                }
            ],
        },
    }
    return [
        {'role': 'system', 'content': system_prompt},
        {'role': 'user', 'content': json.dumps(user_payload, ensure_ascii=False)},
    ]
    return [
        {
            'role': 'system',
            'content': (
                '你是一名新闻去重助手。你的任务是识别同一电商品牌在同一事件上的重复新闻，'
                '尤其是不同新闻平台、聚合站、转载站、搜索引擎来源、门户站对同一事件的转载、摘编、改写。'
                '如果两条新闻围绕同一品牌、同一事件、同一公告、同一政策变化、同一促销活动、同一调查或同一处罚，'
                '即使标题写法不同、语言不同、来源站不同，也应视为重复。'
                '如果新闻讨论的是不同事件、不同商品、不同活动、不同政策、不同处罚、不同财报节点，'
                '即使品牌相同，也绝对不能去重。'
                '请优先保留信息更完整、标题更清晰、摘要更完整、来源更权威或更接近原始报道的一条。'
                '返回 JSON 对象，字段固定为 keep_article_ids 和 duplicate_groups。'
            ),
        },
        {
            'role': 'user',
            'content': json.dumps(
                {
                    'task': '判断下面这些新闻中哪些属于同一电商品牌的同一事件重复报道，只保留代表性的一条。',
                    'articles': article_payload,
                    'output_schema': {
                        'keep_article_ids': ['0'],
                        'duplicate_groups': [
                            {
                                'canonical_article_id': '0',
                                'duplicate_article_ids': ['1'],
                                'reason': '两条新闻报道的是同一品牌的同一事件，后一条是转载/改写版本。',
                            }
                        ],
                    },
                },
                ensure_ascii=False,
            ),
        },
    ]


def normalize_string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        raw_items = value
    else:
        raw_items = re.split(r"[|,;\n]+", str(value).strip())
    ordered: list[str] = []
    for item in raw_items:
        normalized = clean_text(str(item))
        if normalized and normalized not in ordered:
            ordered.append(normalized)
    return ordered


def classify_duplicate_cluster_with_ai(
    cluster_items: list[tuple[int, dict[str, Any]]],
    api_url: str,
    api_key: str,
    api_model: str,
    call_filter_api: Callable[[list[dict[str, str]], str, str, str], dict[str, Any]],
) -> tuple[set[int], list[dict[str, Any]]]:
    payload = call_filter_api(
        build_ai_dedupe_messages(cluster_items),
        api_url,
        api_key,
        api_model,
    )
    valid_indices = {index for index, _ in cluster_items}
    keep_ids: set[int] = set()
    duplicate_groups = payload.get('duplicate_groups') if isinstance(payload, dict) else []

    for raw_id in normalize_string_list(payload.get('keep_article_ids') if isinstance(payload, dict) else []):
        try:
            index = int(raw_id)
        except ValueError:
            continue
        if index in valid_indices:
            keep_ids.add(index)

    normalized_duplicate_groups: list[dict[str, Any]] = []
    if isinstance(duplicate_groups, list):
        for group in duplicate_groups:
            if not isinstance(group, dict):
                continue
            try:
                canonical_id = int(str(group.get('canonical_article_id') or '').strip())
            except ValueError:
                continue
            if canonical_id not in valid_indices:
                continue
            duplicate_ids: list[int] = []
            for raw_id in normalize_string_list(group.get('duplicate_article_ids')):
                try:
                    duplicate_id = int(raw_id)
                except ValueError:
                    continue
                if duplicate_id in valid_indices and duplicate_id != canonical_id and duplicate_id not in duplicate_ids:
                    duplicate_ids.append(duplicate_id)
            keep_ids.add(canonical_id)
            normalized_duplicate_groups.append(
                {
                    'canonical_article_id': canonical_id,
                    'duplicate_article_ids': duplicate_ids,
                    'reason': clean_text(group.get('reason')),
                }
            )

    if not keep_ids:
        keep_ids = set(valid_indices)
    return keep_ids, normalized_duplicate_groups


def apply_final_ai_dedupe(
    rows: list[dict[str, Any]],
    *,
    api_url: str,
    api_key: str,
    api_model: str,
    call_filter_api: Callable[[list[dict[str, str]], str, str, str], dict[str, Any]],
    ai_workers: int = 3,
    progress_callback=None,
    progress_emitter: Callable[..., None] | None = None,
    total_sites: int = 0,
    completed_sites: int = 0,
    progress_start: int = 92,
    progress_cap: int = 96,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    if not rows:
        return rows, {
            'enabled': False,
            'configured': bool(api_url and api_key and api_model),
            'candidate_group_count': 0,
            'request_count': 0,
            'removed_count': 0,
            'api_error_count': 0,
            'date_window_days': AI_DEDUPE_DATE_WINDOW_DAYS,
        }
    if not api_url or not api_key or not api_model:
        return rows, {
            'enabled': False,
            'configured': False,
            'candidate_group_count': 0,
            'request_count': 0,
            'removed_count': 0,
            'api_error_count': 0,
            'date_window_days': AI_DEDUPE_DATE_WINDOW_DAYS,
        }

    candidate_clusters = build_ai_dedupe_candidate_clusters(rows)
    if not candidate_clusters:
        return rows, {
            'enabled': True,
            'configured': True,
            'candidate_group_count': 0,
            'request_count': 0,
            'removed_count': 0,
            'api_error_count': 0,
            'date_window_days': AI_DEDUPE_DATE_WINDOW_DAYS,
        }

    keep_indices = set(range(len(rows)))
    removed_count = 0
    api_error_count = 0
    completed_groups = 0
    worker_count = max(1, min(ai_workers, len(candidate_clusters), 4))

    def worker(cluster_indices: list[int]) -> tuple[list[int], set[int] | None, list[dict[str, Any]] | None, str | None]:
        cluster_items = [(index, rows[index]) for index in cluster_indices]
        try:
            keep_ids, duplicate_groups = classify_duplicate_cluster_with_ai(cluster_items, api_url, api_key, api_model, call_filter_api)
            return cluster_indices, keep_ids, duplicate_groups, None
        except Exception as exc:
            return cluster_indices, None, None, clean_text(str(exc)) or exc.__class__.__name__

    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        futures = [executor.submit(worker, cluster_indices) for cluster_indices in candidate_clusters]
        for future in as_completed(futures):
            cluster_indices, cluster_keep_ids, _, error_message = future.result()
            completed_groups += 1
            if cluster_keep_ids is None:
                api_error_count += 1
            else:
                for index in cluster_indices:
                    if index not in cluster_keep_ids and index in keep_indices:
                        keep_indices.remove(index)
                        removed_count += 1

            if progress_emitter:
                progress_percent = int(progress_start + ((progress_cap - progress_start) * completed_groups / max(1, len(candidate_clusters))))
                progress_emitter(
                    progress_callback,
                    stage='ai_dedupe',
                    total_sites=total_sites,
                    completed_sites=completed_sites,
                    active_sites=max(0, min(worker_count, len(candidate_clusters) - completed_groups)),
                    current_site=f'AI 去重（已完成 {completed_groups}/{len(candidate_clusters)} 组候选重复新闻）',
                    last_completed_site='',
                    message=(
                        f'正在使用 AI 判断高相似新闻是否重复，已完成 {completed_groups}/{len(candidate_clusters)} 组'
                        + (f'；异常 {api_error_count} 组，已保留原文' if api_error_count else '')
                    ),
                    progress_percent=min(progress_cap, progress_percent),
                )

    deduped_rows = [row for index, row in enumerate(rows) if index in keep_indices]
    return deduped_rows, {
        'enabled': True,
        'configured': True,
        'candidate_group_count': len(candidate_clusters),
        'request_count': len(candidate_clusters),
        'removed_count': removed_count,
        'api_error_count': api_error_count,
        'date_window_days': AI_DEDUPE_DATE_WINDOW_DAYS,
    }
