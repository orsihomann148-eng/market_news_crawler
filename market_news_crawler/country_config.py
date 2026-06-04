from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import runtime_paths

DEFAULT_COUNTRY_CODE = "japan"
PROJECT_ROOT = Path(__file__).resolve().parent
CUSTOM_COUNTRY_CONFIG_PATH = runtime_paths.custom_country_config_path()
runtime_paths.ensure_from_template(CUSTOM_COUNTRY_CONFIG_PATH, PROJECT_ROOT / "country_configs_custom.json")
COUNTRY_DATA_DIRNAME = "country_data"


def _normalize_country_token(value: str | None) -> str:
    normalized = str(value or "").strip().lower()
    normalized = re.sub(r"[^a-z0-9_-]+", "_", normalized)
    normalized = re.sub(r"_+", "_", normalized).strip("_-")
    return normalized


def default_country_data_dir(country_code: str) -> str:
    normalized_code = _normalize_country_token(country_code)
    return f"{COUNTRY_DATA_DIRNAME}/{normalized_code}" if normalized_code else COUNTRY_DATA_DIRNAME


def default_country_file_paths(country_code: str) -> dict[str, str]:
    base_dir = default_country_data_dir(country_code)
    return {
        "xlsx_path": f"{base_dir}/source_survey.xlsx",
        "extra_sources_path": f"{base_dir}/extra_sources.json",
        "adapter_configs_path": f"{base_dir}/site_adapter_configs.json",
        "site_credentials_path": f"{base_dir}/site_credentials.json",
        "source_capability_cache_path": f"{base_dir}/source_capability_cache.json",
    }


def normalize_project_relative_path(value: str | None) -> str:
    raw_value = str(value or "").strip()
    if not raw_value:
        return ""
    normalized = raw_value.replace("\\", "/")
    drive_prefixed = bool(re.match(r"^[a-zA-Z]:/", normalized))
    unc_prefixed = normalized.startswith("//")
    if not drive_prefixed and not unc_prefixed:
        normalized = normalized.lstrip("/")
    normalized = re.sub(r"/+", "/", normalized)
    return normalized


def resolve_project_path(value: str | None, *, base_dir: Path | None = None) -> Path:
    normalized = normalize_project_relative_path(value)
    anchor = base_dir or PROJECT_ROOT
    if not normalized:
        return anchor
    if re.match(r"^[a-zA-Z]:/", normalized) or normalized.startswith("//"):
        return Path(normalized)
    candidate = Path(normalized)
    if candidate.is_absolute():
        return candidate
    return (anchor / candidate).resolve()


def legacy_country_file_names(country_code: str) -> dict[str, str]:
    normalized_code = _normalize_country_token(country_code)
    if normalized_code == DEFAULT_COUNTRY_CODE:
        return {
            "xlsx_path": "source_survey.xlsx",
            "extra_sources_path": "extra_sources.json",
            "adapter_configs_path": "site_adapter_configs.json",
            "site_credentials_path": "site_credentials.json",
            "source_capability_cache_path": "source_capability_cache.json",
        }
    if normalized_code == "france":
        return {
            "xlsx_path": "source_survey_france.xlsx",
            "extra_sources_path": "extra_sources_france.json",
            "adapter_configs_path": "site_adapter_configs_france.json",
            "site_credentials_path": "site_credentials_france.json",
            "source_capability_cache_path": "source_capability_cache_france.json",
        }
    return {
        "xlsx_path": f"source_survey_{normalized_code}.xlsx",
        "extra_sources_path": f"extra_sources_{normalized_code}.json",
        "adapter_configs_path": f"site_adapter_configs_{normalized_code}.json",
        "site_credentials_path": f"site_credentials_{normalized_code}.json",
        "source_capability_cache_path": f"source_capability_cache_{normalized_code}.json",
    }

JAPAN_PROMO_SEARCH_QUERY_BLOCKS = [
    (
        "promo_core",
        "(sale OR discount OR coupon OR campaign OR promotion OR promo OR deal OR deals "
        "OR offer OR offers OR セール OR キャンペーン OR 特典 OR 割引 OR 値下げ "
        "OR クーポン OR フェア OR 期間限定 OR 优惠 OR 折扣 OR 促销 OR 优惠券)",
    ),
    (
        "points_rewards",
        "(points OR \"point up\" OR \"point back\" OR cashback OR reward OR rewards "
        "OR ポイント OR ポイントアップ OR ポイント還元 OR 還元)",
    ),
    (
        "price_event",
        "(\"flash sale\" OR voucher OR bundle OR \"special offer\" OR \"limited time\" "
        "OR \"price cut\" OR markdown OR 特価 OR 値下げ OR 期間限定)",
    ),
    (
        "shopping_festivals",
        "(\"double 11\" OR \"double 12\" OR 11.11 OR 12.12 OR \"black friday\" OR \"cyber monday\" "
        "OR \"prime day\" OR \"shopping festival\" OR \"shopping season\" OR 福袋 OR 初売り OR 新春 "
        "OR ゴールデンウィーク OR GW OR 新生活 OR 夏休み OR 冬休み OR 双11 OR 双12 OR 双十二 OR 618)",
    ),
    (
        "seasonal_holidays",
        "(christmas OR valentine OR \"white day\" OR \"mothers day\" OR \"father's day\" "
        "OR gift OR gifting OR holiday OR seasonal OR 母の日 OR 父の日 OR クリスマス OR バレンタイン "
        "OR ホワイトデー OR 敬老の日 OR お盆 OR 年末年始)",
    ),
]

FRANCE_PROMO_SEARCH_QUERY_BLOCKS = [
    (
        "promo_core",
        "(promotion OR promo OR remise OR reduction OR reductions OR rabais OR soldes "
        "OR coupon OR \"code promo\" OR offre OR offres OR reduction de prix OR discount "
        "OR sale OR deal OR deals OR campaign OR cashback)",
    ),
    (
        "points_rewards",
        "(points OR fidelite OR fidélité OR recompense OR récompense OR reward OR rewards "
        "OR cashback OR cagnotte)",
    ),
    (
        "price_event",
        "(\"vente flash\" OR voucher OR bundle OR \"special offer\" OR \"limited time\" "
        "OR \"price cut\" OR markdown OR \"offre limitee\" OR \"offre limitée\" OR \"prix en baisse\")",
    ),
    (
        "shopping_festivals",
        "(\"black friday\" OR \"cyber monday\" OR \"prime day\" OR soldes OR \"soldes d'hiver\" "
        "OR \"soldes d'ete\" OR \"soldes d'été\" OR \"shopping festival\" OR \"shopping season\")",
    ),
    (
        "seasonal_holidays",
        "(noel OR Noël OR \"saint valentin\" OR \"fete des meres\" OR \"fête des mères\" "
        "OR \"fete des peres\" OR \"fête des pères\" OR rentree OR rentrée OR christmas "
        "OR valentine OR \"mothers day\" OR \"father's day\" OR holiday OR seasonal)",
    ),
]

JAPAN_REPORT_QUERY_BLOCKS = [
    (
        "ranking_report",
        "(ranking OR rankings OR rank OR benchmark OR report OR survey OR study OR index OR data OR stats "
        "OR market share OR leaderboard OR classement OR rapport OR donnees OR données OR 排名 OR 榜单 "
        "OR 报告 OR 数据 OR 调查 OR 指数 OR 市占率 OR 统计 OR ランキング OR レポート OR 調査 OR データ OR 指標)",
    ),
]

FRANCE_REPORT_QUERY_BLOCKS = [
    (
        "ranking_report",
        "(ranking OR rankings OR rank OR benchmark OR report OR survey OR study OR index OR data OR stats "
        "OR market share OR classement OR rapport OR barometre OR baromètre OR etude OR étude OR donnees "
        "OR données OR comparatif OR palmares OR palmarès)",
    ),
]

COUNTRY_CONFIGS: dict[str, dict[str, Any]] = {
    "japan": {
        "code": "japan",
        "label": "日本",
        "consumer_label": "日本消费者",
        "market_label": "日本市场",
        "timezone": "Asia/Tokyo",
        "google_news_hl": "ja",
        "google_news_gl": "JP",
        "google_news_ceid": "JP:ja",
        "bing_news_market": "ja-JP",
        "market_terms": ["japan", "japanese", "日本", "日本市场", "日本市場", "日本站", "日本向け", "国内消費者", "日本消费者"],
        "market_search_block": "(Japan OR Japanese OR 日本 OR 日本市场 OR 日本市場 OR 日本站 OR 日本向け OR 国内)",
        **default_country_file_paths("japan"),
        "include_xlsx_sources": True,
        "output_slug": "japan",
        "legacy_output_prefixes": ["xlsx_sources_", "run_"],
        "app_title": "新闻资讯抓取工具",
        "promo_search_query_blocks": JAPAN_PROMO_SEARCH_QUERY_BLOCKS,
        "related_news_query_blocks": JAPAN_PROMO_SEARCH_QUERY_BLOCKS,
        "report_query_blocks": JAPAN_REPORT_QUERY_BLOCKS,
        "dimension_search_term_overrides": {},
        "platform_display_overrides": {},
        "platform_search_term_overrides": {},
        "available_platform_labels": [],
        "platform_alias_exclude_tokens": [],
        "official_sources": {
            "TikTok/TikTok Shop": {
                "label": "TikTok Newsroom JP",
                "url": "https://newsroom.tiktok.com/?lang=ja-JP",
            },
            "Amazon": {
                "label": "About Amazon Japan",
                "url": "https://www.aboutamazon.jp/news",
            },
            "Rakuten Ichiba": {
                "label": "Rakuten Corporate Press",
                "url": "https://corp.rakuten.co.jp/news/press/",
            },
        },
        "rakuten_press_path_date_pattern": r"/news/press/(\d{4})/(\d{2})(\d{2})_\d+\.html$",
        "rakuten_press_date_group_mode": "ymd_compact",
        "tiktok_newsroom_lang": "ja-JP",
    },
    "france": {
        "code": "france",
        "label": "法国",
        "consumer_label": "法国消费者",
        "market_label": "法国市场",
        "timezone": "Europe/Paris",
        "google_news_hl": "fr",
        "google_news_gl": "FR",
        "google_news_ceid": "FR:fr",
        "bing_news_market": "fr-FR",
        "market_terms": [
            "france",
            "french",
            "francais",
            "français",
            "marché français",
            "marche francais",
            "consommateurs français",
            "consommateurs francais",
            "法国",
            "法國",
            "法国市场",
            "法國市場",
            "法国站",
            "法国消费者",
        ],
        "market_search_block": (
            "(France OR French OR Francais OR Français OR \"marché français\" OR \"marche francais\" "
            "OR \"consommateurs français\" OR \"consommateurs francais\" OR 法国 OR 法國 OR 法国市场 OR 法國市場 OR 法国站)"
        ),
        **default_country_file_paths("france"),
        "include_xlsx_sources": False,
        "output_slug": "france",
        "legacy_output_prefixes": [],
        "app_title": "新闻资讯抓取工具",
        "promo_search_query_blocks": FRANCE_PROMO_SEARCH_QUERY_BLOCKS,
        "related_news_query_blocks": FRANCE_PROMO_SEARCH_QUERY_BLOCKS,
        "report_query_blocks": FRANCE_REPORT_QUERY_BLOCKS,
        "dimension_search_term_overrides": {
            "Product quality": ["qualite produit", "qualité produit", "authenticite", "authenticité", "counterfeit"],
            "Seller quality": ["qualite vendeur", "qualité vendeur", "merchant trust", "service client vendeur"],
            "Product variety": ["variete produits", "variété produits", "assortment", "selection produit"],
            "Price": ["prix", "promotion", "discount", "coupon", "shipping cost"],
            "Content": ["contenu", "video shopping", "live shopping", "shoppable content"],
            "Logistics": ["logistique", "livraison", "delivery", "shipping"],
            "Post-purchase service": ["service apres vente", "service après-vente", "refund", "return"],
            "Product feature": ["fonctionnalite plateforme", "fonctionnalité plateforme", "payment", "checkout", "search"],
        },
        "platform_display_overrides": {
            "TikTok/TikTok Shop": "TikTok Shop France",
            "Amazon": "Amazon France",
            "Rakuten Ichiba": "Rakuten France",
        },
        "platform_search_term_overrides": {
            "TikTok Shop": [
                "TikTok Shop France",
                "TikTok Shop FR",
                "TikTok Shop",
                "TikTok France",
                "TikTok FR",
                "TikTok",
            ],
            "Amazon": [
                "Amazon France",
                "Amazon FR",
                "Amazon.fr",
                "Amazon",
            ],
            "Zalando": [
                "Zalando France",
                "Zalando FR",
                "Zalando.fr",
                "Zalando",
            ],
            "SHEIN": [
                "SHEIN France",
                "SHEIN FR",
                "SHEIN",
                "Shein",
                "shein.fr",
            ],
            "Temu": [
                "Temu France",
                "Temu FR",
                "Temu",
                "TEMU",
                "temu.fr",
            ],
            "Instagram": [
                "Instagram Shopping France",
                "Instagram Shopping FR",
                "Instagram Shopping",
                "Instagram Shop France",
                "Instagram Shop",
                "Instagram France",
                "Instagram FR",
                "Instagram",
                "Instagram DTC",
                "IG",
                "INS",
            ],
            "Instagram Shopping": [
                "Instagram Shopping France",
                "Instagram Shopping FR",
                "Instagram Shopping",
                "Instagram Shop France",
                "Instagram Shop",
                "Instagram France",
                "Instagram FR",
                "Instagram",
                "Instagram DTC",
                "IG",
                "INS",
            ],
            "Rakuten Ichiba": [
                "Rakuten France",
                "Rakuten FR",
                "Rakuten France marketplace",
                "Rakuten",
            ],
            "TikTok/TikTok Shop": [
                "TikTok Shop France",
                "TikTok Shop FR",
                "TikTok Shop",
                "TikTok France",
                "TikTok FR",
                "TikTok",
            ],
            "TEMU": [
                "TEMU France",
                "TEMU FR",
                "TEMU",
                "temu.fr",
            ],
            "Shein": [
                "SHEIN France",
                "SHEIN FR",
                "SHEIN",
                "Shein",
                "shein.fr",
            ],
        },
        "default_platform_search_term_overrides": {
            "TikTok Shop": [
                "TikTok Shop France",
                "TikTok Shop FR",
                "TikTok Shop",
                "TikTok France",
                "TikTok FR",
                "TikTok",
            ],
            "Amazon": [
                "Amazon France",
                "Amazon FR",
                "Amazon.fr",
                "Amazon",
            ],
            "Zalando": [
                "Zalando France",
                "Zalando FR",
                "Zalando.fr",
                "Zalando",
            ],
            "SHEIN": [
                "SHEIN France",
                "SHEIN FR",
                "SHEIN",
                "Shein",
                "shein.fr",
            ],
            "Temu": [
                "Temu France",
                "Temu FR",
                "Temu",
                "TEMU",
                "temu.fr",
            ],
            "Instagram": [
                "Instagram Shopping France",
                "Instagram Shopping FR",
                "Instagram Shopping",
                "Instagram Shop France",
                "Instagram Shop",
                "Instagram France",
                "Instagram FR",
                "Instagram",
                "Instagram DTC",
                "IG",
                "INS",
            ],
            "Instagram Shopping": [
                "Instagram Shopping France",
                "Instagram Shopping FR",
                "Instagram Shopping",
                "Instagram Shop France",
                "Instagram Shop",
                "Instagram France",
                "Instagram FR",
                "Instagram",
                "Instagram DTC",
                "IG",
                "INS",
            ],
            "TikTok/TikTok Shop": [
                "TikTok Shop France",
                "TikTok Shop FR",
                "TikTok Shop",
                "TikTok France",
                "TikTok FR",
                "TikTok",
            ],
            "Shein": [
                "SHEIN France",
                "SHEIN FR",
                "SHEIN",
                "Shein",
                "shein.fr",
            ],
            "TEMU": [
                "TEMU France",
                "TEMU FR",
                "TEMU",
                "temu.fr",
            ],
            "Rakuten Ichiba": [
                "Rakuten France",
                "Rakuten FR",
                "Rakuten France marketplace",
                "Rakuten",
            ],
        },
        "available_platform_labels": ["TikTok Shop", "Amazon", "Zalando", "SHEIN", "Temu", "Instagram"],
        "platform_alias_exclude_tokens": [
            "japan",
            "日本",
            "jp",
            "ichiba",
            "楽天市場",
            "楽天",
            "らくてん",
            "アマゾンジャパン",
            "アマゾン",
            "ティックトック",
            "ティックトックショップ",
            "キューテン",
            "ティームー",
            "シーイン",
        ],
        "official_sources": {
            "TikTok/TikTok Shop": {
                "label": "TikTok Newsroom France",
                "url": "https://newsroom.tiktok.com/fr-fr?lang=fr",
            },
            "Amazon": {
                "label": "About Amazon France",
                "url": "https://www.aboutamazon.fr/actualites/actualites",
            },
            "Rakuten Ichiba": {
                "label": "Rakuten France Actualites",
                "url": "https://global.fr.shopping.rakuten.com/actualite/",
            },
        },
        "rakuten_press_path_date_pattern": r"/actualites/(\d{4})/(\d{2})/(\d{2})/",
        "rakuten_press_date_group_mode": "ymd",
        "tiktok_newsroom_lang": "fr",
    },
}


def normalize_new_country_code(value: str | None) -> str:
    return _normalize_country_token(value)


def load_custom_country_configs(path: Path = CUSTOM_COUNTRY_CONFIG_PATH) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    countries = payload.get("countries", payload)
    if not isinstance(countries, dict):
        return {}

    configs: dict[str, dict[str, Any]] = {}
    for raw_code, raw_config in countries.items():
        code = normalize_new_country_code(str(raw_code))
        if not code or not isinstance(raw_config, dict):
            continue
        config = dict(raw_config)
        config["code"] = code
        if code not in COUNTRY_CONFIGS and not str(config.get("label") or "").strip():
            config["label"] = code
        configs[code] = config
    return configs


def merge_country_config(base: dict[str, Any], override: dict[str, Any], code: str) -> dict[str, Any]:
    merged = dict(base)
    merged.update(override)
    merged["code"] = code
    return merged


def all_country_configs() -> dict[str, dict[str, Any]]:
    configs = {code: dict(config) for code, config in COUNTRY_CONFIGS.items()}
    for code, config in load_custom_country_configs().items():
        if code in configs:
            configs[code] = merge_country_config(configs[code], config, code)
        else:
            configs[code] = config
    return configs


def save_custom_country_config(
    country_code: str,
    config: dict[str, Any],
    path: Path = CUSTOM_COUNTRY_CONFIG_PATH,
) -> None:
    normalized_code = normalize_new_country_code(country_code)
    if not normalized_code:
        raise ValueError("country_code 不能为空，且只能包含英文字母、数字、下划线或连字符。")
    if normalized_code in COUNTRY_CONFIGS:
        raise ValueError(f"{normalized_code} 是内置国家代码，不能通过网页覆盖。")

    existing = load_custom_country_configs(path)
    next_config = dict(config)
    next_config["code"] = normalized_code
    existing[normalized_code] = next_config

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "version": 1,
                "countries": existing,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def save_country_config_patch(
    country_code: str,
    patch: dict[str, Any],
    path: Path = CUSTOM_COUNTRY_CONFIG_PATH,
) -> None:
    normalized_code = normalize_new_country_code(country_code)
    if not normalized_code:
        raise ValueError("country_code 不能为空，且只能包含英文字母、数字、下划线或连字符。")
    existing = load_custom_country_configs(path)
    next_config = dict(existing.get(normalized_code) or {})
    next_config.update(patch)
    next_config["code"] = normalized_code
    existing[normalized_code] = next_config
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "version": 1,
                "countries": existing,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def delete_custom_country_config(
    country_code: str,
    path: Path = CUSTOM_COUNTRY_CONFIG_PATH,
) -> dict[str, Any]:
    normalized_code = normalize_new_country_code(country_code)
    if not normalized_code:
        raise ValueError("country_code 不能为空，且只能包含英文字母、数字、下划线或连字符。")
    if normalized_code in COUNTRY_CONFIGS:
        raise ValueError(f"{normalized_code} 是内置国家代码，不能删除。")

    existing = load_custom_country_configs(path)
    removed = existing.pop(normalized_code, None)
    if not isinstance(removed, dict):
        raise ValueError(f"{normalized_code} 不是可删除的自定义国家。")

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "version": 1,
                "countries": existing,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return removed


def available_country_codes() -> list[str]:
    return list(all_country_configs().keys())


def normalize_country_code(value: str | None) -> str:
    normalized = normalize_new_country_code(value)
    if normalized in all_country_configs():
        return normalized
    return DEFAULT_COUNTRY_CODE


def get_country_config(value: str | None) -> dict[str, Any]:
    configs = all_country_configs()
    return configs[normalize_country_code(value)]


def get_default_country_config() -> dict[str, Any]:
    return COUNTRY_CONFIGS[DEFAULT_COUNTRY_CODE]


def country_setting(country_code: str | None, key: str, default: Any = None) -> Any:
    config = get_country_config(country_code)
    if key in config:
        return config[key]
    return get_default_country_config().get(key, default)


def country_list_setting(country_code: str | None, key: str) -> list[Any]:
    value = country_setting(country_code, key, [])
    return value if isinstance(value, list) else []


def country_dict_setting(country_code: str | None, key: str) -> dict[str, Any]:
    value = country_setting(country_code, key, {})
    return value if isinstance(value, dict) else {}


def country_options() -> list[tuple[str, str]]:
    return [(code, config["label"]) for code, config in all_country_configs().items()]


def country_file_path(base_dir: Path, country_code: str, key: str) -> Path:
    config = get_country_config(country_code)
    return base_dir / str(config[key])
