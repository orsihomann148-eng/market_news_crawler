from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import xlsx_source_test


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / 'outputs'
NEWS_SUMMARY_OUTPUT_DIR = OUTPUT_DIR / 'generated_news_summaries'

PLATFORM_ORDER = ['行业', 'Amazon', 'SHEIN', 'TEMU', 'TTS', 'eBay', 'IG']
COUNTRY_ZH_LABELS = {
    'japan': '日本',
    'italy': '意大利',
    'france': '法国',
    'germany': '德国',
    'spain': '西班牙',
}
SENTIMENT_ZH_LABELS = {
    'Positive': '正向',
    'Neutral': '中性',
    'Negative': '负向',
}
SUMMARY_TAG_OPTIONS = [
    '行业生态',
    '卖家生态',
    '平台生态',
    '平台竞争',
    '平台发展',
    '卖家规范',
    '卖家权益',
    '商品品类',
    '商品品质',
    '商品价格',
    '商品丰富度',
    '商品合规',
    '广告',
    '直播',
    '促销活动',
    '支付手段',
    '物流仓储',
    '运费竞争',
    '发货时效',
    '功能',
    '推荐',
    '搜索',
    '订单查询',
    '行业发展',
    '消费趋势',
    '品类趋势',
    '银发经济',
    'Z世代消费',
    '循环时尚',
    '知识产权',
    '政策监管',
    '平台政策',
    '平台合规',
    '报告',
    '品牌形象',
    '商家赋能',
]


@dataclass
class NewsSummaryResult:
    text: str
    output_path: Path
    stats: dict[str, int]


def compact_summary_text(value: Any, max_chars: int = 1000) -> str:
    text = re.sub(r'\s+', ' ', str(value or '')).strip()
    return text[:max_chars].strip()


def country_zh_label(country_code: str) -> str:
    return COUNTRY_ZH_LABELS.get(str(country_code or '').strip().lower(), str(country_code or '').strip() or '目标国家')


def normalize_sentiment(value: Any) -> str:
    return xlsx_source_test.normalize_briefing_sentiment(value)


def sentiment_zh(value: Any) -> str:
    return SENTIMENT_ZH_LABELS.get(normalize_sentiment(value), '中性')


def parse_publish_dt(value: Any) -> datetime | None:
    text = str(value or '').strip()
    if not text:
        return None
    normalized = text.replace('Z', '+00:00')
    try:
        return datetime.fromisoformat(normalized)
    except ValueError:
        pass
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d']:
        try:
            return datetime.strptime(text[:10], fmt)
        except ValueError:
            continue
    return None


def format_publish_date_zh(value: Any) -> str:
    parsed = parse_publish_dt(value)
    if parsed is None:
        return str(value or '').strip()
    return f'{parsed.year}年{parsed.month}月{parsed.day}日'


def primary_dimension(row: dict[str, Any]) -> str:
    dimensions = str(row.get('survey_dimensions') or '').strip()
    if not dimensions:
        return '品牌感知'
    first = re.split(r'\s*[|,;/]\s*', dimensions)[0].strip()
    return first or '品牌感知'


def platform_summary_label(row: dict[str, Any]) -> str:
    if xlsx_source_test.normalize_industry_trend_flag(row.get('industry_trend_flag')):
        return '行业'
    platform = str(row.get('platform_label') or '').strip()
    if not platform:
        return '行业'
    upper = platform.upper()
    if upper in {'TIKTOK SHOP', 'TTS'}:
        return 'TTS'
    if upper in {'INSTAGRAM', 'INSTAGRAM SHOPPING', 'IG'}:
        return 'IG'
    if upper == 'EBAY':
        return 'eBay'
    if upper == 'TEMU':
        return 'TEMU'
    if upper == 'SHEIN':
        return 'SHEIN'
    if upper == 'AMAZON':
        return 'Amazon'
    return platform


def infer_tags(row: dict[str, Any]) -> list[str]:
    text = ' '.join(
        str(row.get(key) or '')
        for key in [
            'title',
            'title_original',
            'title_display_zh',
            'summary',
            'survey_dimensions',
            'survey_indicator_examples',
            'industry_trend_category',
            'industry_trend_reason',
        ]
    ).lower()
    checks = [
        ('政策监管', ['regulation', 'policy', 'commission', 'eu ', 'european', '监管', '政策', '法规', '调查']),
        ('平台合规', ['compliance', 'illegal', 'safety', 'privacy', 'data', '合规', '非法', '安全', '隐私', '数据']),
        ('商品价格', ['price', 'discount', 'deal', 'coupon', 'sale', '促销', '折扣', '优惠', '价格']),
        ('支付手段', ['payment', 'cash', 'checkout', '支付', '现金']),
        ('物流仓储', ['delivery', 'logistics', 'shipping', 'warehouse', '物流', '配送', '仓储']),
        ('广告', ['ad ', 'ads', 'advertising', '广告']),
        ('直播', ['live', 'creator', 'reels', 'tiktok', '直播', '达人', '创作者']),
        ('消费趋势', ['consumer', 'generation z', 'gen z', 'z世代', '消费', '人群']),
        ('报告', ['report', 'survey', 'study', 'data', '报告', '调研', '数据']),
        ('平台竞争', ['competition', 'rival', 'market share', '竞争', '份额']),
        ('品牌形象', ['trust', 'quality', 'brand', '信任', '品牌', '形象']),
        ('商品合规', ['pfas', 'counterfeit', 'illegal products', '假货', '正品', '商品合规']),
        ('平台生态', ['marketplace', 'seller', 'merchant', 'ecosystem', '卖家', '商家', '生态']),
        ('功能', ['feature', 'tool', 'function', '功能', '工具']),
    ]
    tags: list[str] = []
    for tag, needles in checks:
        if any(needle in text for needle in needles):
            tags.append(tag)
    if xlsx_source_test.normalize_industry_trend_flag(row.get('industry_trend_flag')):
        tags.insert(0, '行业发展')
    if not tags:
        tags.append(primary_dimension(row))
    output: list[str] = []
    for tag in tags:
        normalized = tag if tag in SUMMARY_TAG_OPTIONS else tag.strip()
        if normalized and normalized not in output:
            output.append(normalized)
        if len(output) >= 4:
            break
    return output


def limit_core_claim(value: Any, max_chars: int = 50) -> str:
    text = compact_summary_text(value, 200)
    text = text.strip('。？！?；; ')
    if len(text) <= max_chars:
        return text
    return text[:max_chars].rstrip('，、；;：: ') + '…'


def article_title(row: dict[str, Any]) -> str:
    return compact_summary_text(
        row.get('title_display_zh')
        or row.get('title_translated')
        or row.get('title')
        or row.get('title_original')
        or row.get('summary')
        or '未命名新闻',
        300,
    )


def default_summary_item(row: dict[str, Any], country_code: str) -> dict[str, Any]:
    platform = platform_summary_label(row)
    country = country_zh_label(country_code)
    date_text = format_publish_date_zh(row.get('published_at')) or '近期'
    source_name = compact_summary_text(row.get('source_name'), 80) or '媒体'
    dimension = primary_dimension(row)
    tags = infer_tags(row)
    sentiment = sentiment_zh(row.get('briefing_sentiment'))
    trend_reason = compact_summary_text(row.get('industry_trend_reason'), 220)
    linkage = compact_summary_text(row.get('survey_indicator_examples'), 260)
    title = article_title(row)
    core = limit_core_claim(title)
    implication = trend_reason or linkage or f'可能影响消费者对{dimension}相关体验的判断'
    detail = f'{date_text}，{country}/{source_name}消息，{title}，{implication}。'
    return {
        'article_id': str(row.get('article_id') or ''),
        'platform_label': platform,
        'core_claim': core,
        'detail': detail,
        'sentiment': sentiment,
        'primary_metric': dimension,
        'tags': tags,
        'needs_review': True,
    }


def summary_payload_items(payload: dict[str, Any]) -> list[dict[str, Any]]:
    for key in ['items', 'summaries', 'articles', 'rows']:
        value = payload.get(key) if isinstance(payload, dict) else None
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    return []


def normalize_tags(value: Any, row: dict[str, Any]) -> list[str]:
    raw_items: list[Any]
    if isinstance(value, list):
        raw_items = value
    else:
        text = str(value or '').strip()
        raw_items = re.split(r'[|,，、]\s*', text) if text else []
    tags: list[str] = []
    for item in raw_items:
        tag = str(item or '').strip().strip('【】')
        if tag and tag not in tags:
            tags.append(tag)
        if len(tags) >= 4:
            break
    if tags:
        return tags
    return infer_tags(row)


def normalize_summary_item(item: dict[str, Any], row: dict[str, Any], country_code: str) -> dict[str, Any]:
    fallback = default_summary_item(row, country_code)
    platform = str(item.get('platform_label') or item.get('platform') or fallback['platform_label']).strip()
    if platform.startswith('【') and platform.endswith('】'):
        platform = platform.strip('【】')
    sentiment = str(item.get('sentiment') or item.get('sentiment_zh') or fallback['sentiment']).strip()
    if sentiment in {'Positive', '正面', '利好'}:
        sentiment = '正向'
    elif sentiment in {'Negative', '负面', '利空'}:
        sentiment = '负向'
    elif sentiment not in {'正向', '中性', '负向'}:
        sentiment = fallback['sentiment']
    return {
        'article_id': str(item.get('article_id') or row.get('article_id') or ''),
        'platform_label': platform or fallback['platform_label'],
        'core_claim': limit_core_claim(item.get('core_claim') or item.get('headline') or fallback['core_claim']),
        'detail': compact_summary_text(item.get('detail') or item.get('explanation') or fallback['detail'], 800),
        'sentiment': sentiment,
        'primary_metric': compact_summary_text(item.get('primary_metric') or item.get('metric') or fallback['primary_metric'], 80),
        'tags': normalize_tags(item.get('tags'), row),
        'needs_review': False,
    }


def build_news_summary_ai_messages(rows: list[dict[str, Any]], country_code: str) -> list[dict[str, str]]:
    country = country_zh_label(country_code)
    articles = []
    for index, row in enumerate(rows, start=1):
        articles.append(
            {
                'article_id': str(row.get('article_id') or index),
                'platform': row.get('platform_label') or '',
                'title_original': row.get('title_original') or row.get('title') or '',
                'title_zh': row.get('title_display_zh') or row.get('title_translated') or '',
                'summary': row.get('summary') or '',
                'article_url': row.get('article_url') or '',
                'published_at': row.get('published_at') or '',
                'source_name': row.get('source_name') or '',
                'nps_dimension': row.get('survey_dimensions') or '',
                'survey_indicator_examples': row.get('survey_indicator_examples') or '',
                'sentiment': row.get('briefing_sentiment') or '',
                'industry_trend_flag': bool(xlsx_source_test.normalize_industry_trend_flag(row.get('industry_trend_flag'))),
                'industry_trend_category': row.get('industry_trend_category') or '',
                'industry_trend_impact': row.get('industry_trend_impact') or '',
                'industry_trend_reason': row.get('industry_trend_reason') or '',
            }
        )
    schema = {
        'items': [
            {
                'article_id': 'same article_id from input',
                'platform_label': '行业 or platform label',
                'core_claim': 'Chinese core judgement, <=50 Chinese characters',
                'detail': 'Chinese paragraph with fact + implication, no URL and no tail tags',
                'sentiment': '正向|中性|负向',
                'primary_metric': 'main NPS/business metric',
                'tags': ['2-4 values from allowed_tags'],
            }
        ]
    }
    return [
        {
            'role': 'system',
            'content': (
                'You are a senior market intelligence analyst writing client-ready Chinese news summaries. '
                'Return JSON only. Do not omit any input article.'
            ),
        },
        {
            'role': 'user',
            'content': json.dumps(
                {
                    'country_code': country_code,
                    'country_zh': country,
                    'output_schema': schema,
                    'allowed_tags': SUMMARY_TAG_OPTIONS,
                    'rules': [
                        'Write in concise professional Chinese, similar to a client briefing note.',
                        'Each item must include time, country, platform, main related metric, and positive/neutral/negative direction.',
                        'core_claim must be no more than 50 Chinese characters and should state the main implication.',
                        'detail can be longer and should summarize the fact plus potential implications.',
                        'Do not include article URL, 原文链接, 媒体, or final bracket tags in detail; the system will append URL/source and tags later.',
                        'Prioritize macro policy, industry trends, consumer data, campaign results, compliance, platform functions, services, and brand-level impact.',
                        'If the article implies but does not explicitly state an impact, infer cautiously from the facts and NPS linkage.',
                        'Do not write Chinese users, 中国用户, or anything suggesting the market is China unless the input country is China.',
                        'Use platform_label 行业 for macro/industry/report/policy trend news; otherwise use the named platform.',
                    ],
                    'articles': articles,
                },
                ensure_ascii=False,
                indent=2,
            ),
        },
    ]


def summary_reference_text(row: dict[str, Any]) -> str:
    article_url = compact_summary_text(row.get('article_url'), 500)
    if article_url:
        return article_url
    source_name = compact_summary_text(row.get('source_name'), 80)
    if source_name:
        return f'来源：{source_name}'
    return ''


def format_summary_line(item: dict[str, Any], row: dict[str, Any], country_code: str) -> str:
    platform = str(item.get('platform_label') or platform_summary_label(row)).strip() or '行业'
    core = limit_core_claim(item.get('core_claim') or article_title(row))
    detail = compact_summary_text(item.get('detail'), 900)
    if not detail:
        detail = default_summary_item(row, country_code)['detail']
    detail = detail.rstrip('。')
    sentiment = str(item.get('sentiment') or sentiment_zh(row.get('briefing_sentiment'))).strip() or '中性'
    metric = compact_summary_text(item.get('primary_metric') or primary_dimension(row), 80)
    tags = normalize_tags(item.get('tags'), row)
    tag_text = ''.join(f'【{tag}】' for tag in [sentiment, metric, *tags] if tag)
    reference = summary_reference_text(row)
    reference_text = f' {reference}' if reference else ''
    return f'【{platform}】{core}。{detail}{reference_text} {tag_text}'.strip()


def summary_sort_key(row: dict[str, Any]) -> tuple[int, int, float]:
    platform = platform_summary_label(row)
    is_industry = 0 if platform == '行业' else 1
    try:
        platform_index = PLATFORM_ORDER.index(platform)
    except ValueError:
        platform_index = len(PLATFORM_ORDER)
    parsed = parse_publish_dt(row.get('published_at'))
    timestamp = parsed.timestamp() if parsed is not None else 0
    return (is_industry, platform_index, -timestamp)


def sorted_summary_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(rows, key=summary_sort_key)


def generate_news_summary(
    rows: list[dict[str, Any]],
    *,
    country_code: str,
    api_settings: dict[str, str],
    batch_size: int = 8,
    stats: dict[str, int] | None = None,
) -> NewsSummaryResult:
    if stats is not None:
        stats.setdefault('batch_count', 0)
        stats.setdefault('failed_batch_count', 0)
        stats.setdefault('failed_row_count', 0)
        stats.setdefault('ai_completed_row_count', 0)
        stats.setdefault('fallback_row_count', 0)

    ordered_rows = sorted_summary_rows(rows)
    items_by_id = {
        str(row.get('article_id') or index): default_summary_item(row, country_code)
        for index, row in enumerate(ordered_rows, start=1)
    }
    row_by_id = {
        str(row.get('article_id') or index): row
        for index, row in enumerate(ordered_rows, start=1)
    }

    for offset in range(0, len(ordered_rows), batch_size):
        batch = ordered_rows[offset:offset + batch_size]
        if stats is not None:
            stats['batch_count'] += 1
        try:
            payload = xlsx_source_test.call_survey_filter_api(
                build_news_summary_ai_messages(batch, country_code),
                api_settings['survey_api_url'],
                api_settings['survey_api_key'],
                api_settings['survey_api_model'],
            )
            seen_ids: set[str] = set()
            for item in summary_payload_items(payload):
                article_id = str(item.get('article_id') or '').strip()
                if article_id and article_id in row_by_id:
                    items_by_id[article_id] = normalize_summary_item(item, row_by_id[article_id], country_code)
                    seen_ids.add(article_id)
                    if stats is not None:
                        stats['ai_completed_row_count'] += 1
            missing = [
                str(row.get('article_id') or index)
                for index, row in enumerate(batch, start=offset + 1)
                if str(row.get('article_id') or index) not in seen_ids
            ]
            if stats is not None and missing:
                stats['fallback_row_count'] += len(missing)
        except Exception:
            if stats is not None:
                stats['failed_batch_count'] += 1
                stats['failed_row_count'] += len(batch)
                stats['fallback_row_count'] += len(batch)

    lines = []
    for index, row in enumerate(ordered_rows, start=1):
        article_id = str(row.get('article_id') or index)
        lines.append(format_summary_line(items_by_id.get(article_id) or default_summary_item(row, country_code), row, country_code))
    summary_text = '\n\n'.join(lines).strip()
    NEWS_SUMMARY_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"news_summary_{country_code}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    output_path = NEWS_SUMMARY_OUTPUT_DIR / filename
    output_path.write_text(summary_text, encoding='utf-8-sig')
    return NewsSummaryResult(summary_text, output_path, stats or {})
