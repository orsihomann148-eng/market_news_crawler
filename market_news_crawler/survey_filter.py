from __future__ import annotations

import re

from country_config import (
    DEFAULT_COUNTRY_CODE,
    available_country_codes,
    country_dict_setting,
    country_list_setting,
    get_country_config,
)
from news_crawler import clean_text


DEFAULT_RECALL_MODE = 'balanced'


def country_consumer_research_phrase(country_code: str) -> tuple[str, str]:
    config = get_country_config(country_code)
    return config["consumer_label"], config["market_label"]


def country_market_terms(country_code: str) -> list[str]:
    config = get_country_config(country_code)
    return [clean_text(item) for item in config.get("market_terms", []) if clean_text(item)]


def country_market_search_block(country_code: str) -> str:
    return str(get_country_config(country_code).get("market_search_block") or "").strip()


def country_prompt_platform_examples(country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    examples: list[str] = []

    for item in country_list_setting(country_code, "available_platform_labels"):
        normalized = clean_text(item)
        if normalized and normalized not in examples:
            examples.append(normalized)

    for platform, terms in country_dict_setting(country_code, "platform_search_term_overrides").items():
        normalized_platform = clean_text(platform)
        if normalized_platform and normalized_platform not in examples:
            examples.append(normalized_platform)
        if isinstance(terms, list):
            for term in terms:
                normalized_term = clean_text(term)
                if normalized_term and normalized_term not in examples:
                    examples.append(normalized_term)
                if len(examples) >= 6:
                    break
        if len(examples) >= 6:
            break

    for platform in country_dict_setting(country_code, "official_sources"):
        normalized = clean_text(platform)
        if normalized and normalized not in examples:
            examples.append(normalized)
        if len(examples) >= 6:
            break

    if not examples:
        examples = ["Amazon", "TikTok Shop", "Rakuten Ichiba", "Qoo10", "TEMU", "SHEIN"]
    return "、".join(examples[:6])


def default_survey_ai_system_prompt(country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    consumer_label, market_label = country_consumer_research_phrase(country_code)
    platform_examples = country_prompt_platform_examples(country_code)
    return (
        f'你是一名电商舆情筛选助手，当前任务服务于“{consumer_label}对平台总体感受研究”。'
        f'只有当新闻会直接影响{consumer_label}对平台在{market_label}的使用体验、价格感知、服务感知或整体评价，并进而影响问卷指标作答时，才能判定为 relevant=true。'
        f'如果新闻仅针对其他国家或地区，且没有明确证据表明会影响{market_label}或{consumer_label}，应判定为 irrelevant。'
        '如果新闻只是单个商品、单次上新、单个品牌联名、局部活动、小范围产品发布，影响面较小，也应判定为 irrelevant。'
        f'如果新闻只是某一个具体商品、某一个 SKU、某一款耳机/相机/家电/服饰的特价、折扣、优惠券、好价、上架或发售信息，即使发生在 {platform_examples} 等平台，也应判定为 irrelevant。'
        f'如果新闻只是影视、动画、综艺、音乐、艺人、演出、时尚秀、品牌活动、颁奖礼、出道发布、舞台表演、娱乐圈动态，即使标题里出现 {platform_examples} 等平台词，也通常应判定为 irrelevant。'
        '如果新闻只是流媒体内容推荐、电影推荐、动画化消息、上映信息、片单推荐、内容导流，也应判定为 irrelevant，因为这类内容更接近单条内容推荐，而不是平台机制变化。'
        '如果新闻只是媒体导购、编辑推荐、好物清单、排行榜、某月推荐、购物指南、推荐 3 款商品、值得买清单、手提包/耳机/家电等商品测评导购，即使标题中带有平台促销或优惠字样，也应判定为 irrelevant。'
        '如果平台名称只是销售渠道、播放渠道、冠名方、赞助方、活动名称的一部分，而新闻主体并不是该平台自身的规则、功能、价格机制、物流、售后或治理变化，也应判定为 irrelevant。'
        '无目标品牌、无目标市场、纯宏观行业、公益新闻、品牌宣传、普通消费者难以感知的 B2B 细节，通常都应判定为 irrelevant。'
        '但如果新闻是目标品牌/平台相关的行业趋势、市场份额/渗透率、消费者使用率、行业报告、品牌整体增长/衰退、监管/合规趋势、社交电商趋势、平台生态、物流/支付/卖家生态变化，并可能影响消费者对该平台或品牌的整体认知，也可以判定为 relevant=true。'
        '只有平台规则、价格/促销、物流、售后、内容体验、搜索推荐、支付结算、正品质量、网站/APP 功能变化等消费者能直接感知的变化，才通常应判定为 relevant。'
        '如果影响是间接的、长期的、推测性的，或主要影响卖家/合作伙伴而非普通用户，请判定为 irrelevant。'
        '如果 relevant=true，请尽量给出 matched_dimensions 和 matched_question_ids；拿不准题号时，至少给出 matched_dimensions。'
        '如果是行业趋势/品牌整体影响新闻，请同时返回 industry_trend_flag=true、industry_trend_category、industry_trend_impact 和中文 industry_trend_reason。'
        '在 balanced 召回模式下，目标品牌/平台的行业趋势、市场份额、整体增长或下滑、平台生态、监管合规、消费者信任、社交电商和长期品牌感知新闻，即使影响是间接或滞后的，也可以判定为 relevant=true，并用 medium confidence 说明影响链路。'
        '如果文章带有 broad_entry=true 或 market_scope=europe_or_global_relevant，表示它是欧洲/全球平台影响候选；请不要因标题未写本国名称而直接排除，但必须判断它是否可能影响目标国家消费者对平台安全、合规、隐私、商品质量、社交电商体验或长期品牌感知的评价。影响链路成立可 relevant=true 且 confidence=medium；影响链路不成立则 irrelevant。'
        '跨品牌监管、合规、商品安全、海关、低价跨境包裹、隐私或非法商品事件可能影响新闻中出现的多个平台；如果新闻同时点名 Temu、SHEIN、AliExpress、Instagram、Meta、TikTok Shop 等目标平台，请分别判断被点名平台的品牌感知影响，而不要只归给标题最前面的品牌。'
        'TikTok Shop/TTS 的社交电商、直播购物、创作者经济、卖家生态、平台政策、支付物流、消费者保护、监管合规新闻，可能影响消费者对平台整体感知；balanced 模式下影响链路成立时可以用 medium confidence 保留。'
        'Instagram/IG 的 Meta commerce、Instagram business tools、social commerce、creator marketplace、shoppable posts、ads-commerce、brand partnership 新闻，若影响购物、品牌发现、商家或消费者购买路径，可以判定相关；纯娱乐、明星动态、普通账号内容和无商业购物语境的社媒功能仍应排除。'
        'eBay 的 eBay Live、直播购物、卖家生态、平台级促销、大促、平台政策、marketplace strategy、收购或 takeover 相关新闻，如果影响平台生态、品牌形象、消费者信任或平台长期稳定性，balanced 模式下可以用 medium confidence 保留。'
        '平台级促销、大促、直播购物活动、创作者商业合作、AI shopping agent 和品牌发现工具可以保留；但单品折扣、普通优惠码、单款商品好价、媒体导购和具体商品清单仍应排除。'
        '收购、资本市场或平台战略事件如果涉及目标平台的战略方向、稳定性、品牌形象或消费者信任，可以保留一条代表新闻；同事件多家媒体重复报道由最终 AI 去重处理。'
        '你必须对每一条 article_id 都返回一条 decision。'
        '只能从提供的 dimension 和 question_id 中选择，返回 JSON 对象，字段固定为 decisions。'
    )


def legacy_default_survey_ai_system_prompts(country_code: str = DEFAULT_COUNTRY_CODE) -> list[str]:
    consumer_label, market_label = country_consumer_research_phrase(country_code)
    return [
        (
            f'你是一名电商舆情筛选助手，当前任务服务于“{consumer_label}对平台总体感受研究”。'
            f'只有当新闻会直接影响{consumer_label}对平台在{market_label}的使用体验、价格感知、服务感知或整体评价，并进而影响问卷指标作答时，才能判定为 relevant=true。'
            f'如果新闻仅针对其他国家或地区，且没有明确证据表明会影响{market_label}或{consumer_label}，应判定为 irrelevant。'
            '如果新闻只是单个商品、单次上新、单个品牌联名、局部活动、小范围产品发布，影响面较小，也应判定为 irrelevant。'
            '如果新闻只是某一个具体商品、某一个 SKU、某一款耳机/相机/家电/服饰的特价、折扣、优惠券、好价、上架或发售信息，即使发生在 Amazon、楽天、Qoo10 等平台，也应判定为 irrelevant。'
            '如果新闻只是影视、动画、综艺、音乐、偶像、艺人、演出、时尚秀、品牌活动、颁奖礼、GirlsAward、出道发布、舞台表演、娱乐圈动态，即使标题里出现 Amazon、Rakuten、Prime Video 等词，也通常应判定为 irrelevant。'
            '如果新闻只是 Amazon Prime Video 推荐信息、电影推荐、动画化消息、上映信息、片单推荐、内容导流，也应判定为 irrelevant，因为这类内容更接近单条内容推荐，而不是平台机制变化。'
            '如果新闻只是媒体导购、编辑推荐、好物清单、排行榜、某月推荐、购物指南、推荐 3 款商品、值得买清单、手提包/耳机/家电等商品测评导购，即使标题中带有亚马逊促销、乐天优惠等字样，也应判定为 irrelevant。'
            '如果平台名称只是销售渠道、播放渠道、冠名方、赞助方、活动名称的一部分，而新闻主体并不是该平台自身的规则、功能、价格机制、物流、售后或治理变化，也应判定为 irrelevant。'
            '公司融资、组织架构、泛行业动态、公益新闻、品牌宣传、普通消费者难以感知的 B2B 细节，通常都应判定为 irrelevant。'
            '只有平台规则、价格/促销、物流、售后、内容体验、搜索推荐、支付结算、正品质量、网站/APP 功能变化等消费者能直接感知的变化，才通常应判定为 relevant。'
            '如果影响是间接的、长期的、推测性的，或主要影响卖家/合作伙伴而非普通用户，请判定为 irrelevant。'
            '如果 relevant=true，请尽量给出 matched_dimensions 和 matched_question_ids；拿不准题号时，至少给出 matched_dimensions。'
            '你必须对每一条 article_id 都返回一条 decision。'
            '只能从提供的 dimension 和 question_id 中选择，返回 JSON 对象，字段固定为 decisions。'
        ),
        (
            f'你是一名电商舆情筛选助手，当前任务服务于“{consumer_label}对平台总体感受研究”。'
            f'只有当新闻会直接影响{consumer_label}对平台在{market_label}的使用体验、价格感知、服务感知或整体评价，并进而影响问卷指标作答时，才能判定为 relevant=true。'
            f'如果新闻仅针对其他国家或地区，且没有明确证据表明会影响{market_label}或{consumer_label}，应判定为 irrelevant。'
            '如果新闻只是单个商品、单次上新、单个品牌联名、局部活动、小范围产品发布，影响面较小，也应判定为 irrelevant。'
            '如果新闻只是某一个具体商品、某一个 SKU、某一款耳机/相机/家电/服饰的特价、折扣、优惠券、好价、上架或发售信息，即使发生在 Amazon、楽天、Qoo10 等平台，也应判定为 irrelevant。'
            '公司融资、组织架构、泛行业动态、公益新闻、品牌宣传、普通消费者难以感知的 B2B 细节，通常都应判定为 irrelevant。'
            '只有平台规则、价格/促销、物流、售后、内容体验、搜索推荐、支付结算、正品质量、网站/APP 功能变化等消费者能直接感知的变化，才通常应判定为 relevant。'
            '如果影响是间接的、长期的、推测性的，或主要影响卖家/合作伙伴而非普通用户，请判定为 irrelevant。'
            '如果 relevant=true，请尽量给出 matched_dimensions 和 matched_question_ids；拿不准题号时，至少给出 matched_dimensions。'
            '你必须对每一条 article_id 都返回一条 decision。'
            '只能从提供的 dimension 和 question_id 中选择，返回 JSON 对象，字段固定为 decisions。'
        ),
    ]



def survey_system_prompt_has_current_default_markers(value: str | None) -> bool:
    normalized = clean_text(value).lower()
    industry_markers_present = any(
        marker in normalized
        for marker in [
            "industry_trend_flag",
            "industry_trend_category",
            "industry_trend_impact",
            "industry_trend_reason",
        ]
    )
    platform_edge_markers_present = all(
        marker in normalized
        for marker in [
            "tiktok shop/tts",
            "instagram/ig",
            "meta commerce",
            "ebay live",
            "takeover",
        ]
    )
    return industry_markers_present and platform_edge_markers_present


def looks_like_legacy_default_survey_system_prompt(
    value: str | None,
    country_code: str = DEFAULT_COUNTRY_CODE,
) -> bool:
    normalized = clean_text(value)
    if not normalized or survey_system_prompt_has_current_default_markers(normalized):
        return False

    for legacy_prompt in legacy_default_survey_ai_system_prompts(country_code):
        if normalized == clean_text(legacy_prompt):
            return True

    lower_normalized = normalized.lower()
    required_markers = [
        "relevant=true",
        "irrelevant",
        "matched_dimensions",
        "matched_question_ids",
        "decisions",
    ]
    if not all(marker in lower_normalized for marker in required_markers):
        return False

    default_shape_markers = ["sku", "b2b", "app", "amazon"]
    if not all(marker in lower_normalized for marker in default_shape_markers):
        return False

    default_prompt_length = len(clean_text(default_survey_ai_system_prompt(country_code)))
    length = len(normalized)
    return max(700, int(default_prompt_length * 0.55)) <= length <= int(default_prompt_length * 1.15)


def survey_system_prompt_source(value: str | None, country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    normalized = clean_text(value)
    if not normalized:
        return "system_default"
    if normalized == clean_text(default_survey_ai_system_prompt(country_code)):
        return "system_default"
    if looks_like_legacy_default_survey_system_prompt(value, country_code):
        return "system_default"
    return "custom"


def survey_system_prompt_source_label(value: str | None, country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    if survey_system_prompt_source(value, country_code) == "system_default":
        return "\u5f53\u524d\u4f7f\u7528\uff1a\u7cfb\u7edf\u9ed8\u8ba4\u63d0\u793a\u8bcd\uff08\u4f1a\u968f\u7cfb\u7edf\u5347\u7ea7\u81ea\u52a8\u66f4\u65b0\uff09"
    return "\u5f53\u524d\u4f7f\u7528\uff1a\u81ea\u5b9a\u4e49\u63d0\u793a\u8bcd\uff08\u4e0d\u4f1a\u88ab\u7cfb\u7edf\u81ea\u52a8\u8986\u76d6\uff09"

def normalize_survey_system_prompt(value: str | None, country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    normalized = clean_text(value)
    if not normalized:
        return default_survey_ai_system_prompt(country_code)
    if normalized == clean_text(default_survey_ai_system_prompt(country_code)):
        return default_survey_ai_system_prompt(country_code)
    if looks_like_legacy_default_survey_system_prompt(value, country_code):
        return default_survey_ai_system_prompt(country_code)
    legacy_hardcoded_markers = ["Amazon、楽天、Qoo10", "GirlsAward", "乐天优惠"]
    if "字段固定为 decisions" in normalized and any(marker in normalized for marker in legacy_hardcoded_markers):
        return default_survey_ai_system_prompt(country_code)
    return str(value or '')


def promo_search_query_blocks(country_code: str = DEFAULT_COUNTRY_CODE) -> list[tuple[str, str]]:
    return [
        (str(label), str(block))
        for label, block in country_list_setting(country_code, "promo_search_query_blocks")
    ]


def default_promo_search_keywords_text(country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    return '\n'.join(block for _, block in promo_search_query_blocks(country_code))


def related_news_query_blocks(country_code: str = DEFAULT_COUNTRY_CODE) -> list[tuple[str, str]]:
    return [
        (str(label), str(block))
        for label, block in country_list_setting(country_code, "related_news_query_blocks")
    ]


def recall_enhanced_related_news_query_blocks(country_code: str = DEFAULT_COUNTRY_CODE) -> list[tuple[str, str]]:
    normalized_country = clean_text(country_code).lower()
    blocks = [
        ("market_performance", '("market share" OR growth OR decline OR sales OR revenue OR GMV OR adoption OR penetration OR users OR shoppers OR "brand performance" OR "market performance")'),
        ("consumer_experience", '("customer experience" OR "consumer experience" OR "consumer trust" OR satisfaction OR complaints OR reviews OR "shopping experience" OR "user experience")'),
        ("platform_ecosystem", '("platform ecosystem" OR marketplace OR "seller ecosystem" OR "merchant ecosystem" OR "partner ecosystem" OR "creator economy" OR "retail media")'),
        ("payment_logistics", '(payment OR payments OR checkout OR cash OR "payment method" OR delivery OR shipping OR return OR refund OR logistics)'),
        ("trust_safety", '(trust OR safety OR privacy OR data OR investigation OR illegal OR counterfeit OR compliance OR PFAS OR recall)'),
        ("regulatory_market", '(regulation OR regulator OR antitrust OR lawsuit OR investigation OR compliance OR "consumer protection" OR "market authority")'),
        ("platform_features", '("new feature" OR feature OR app OR website OR algorithm OR recommendation OR personalization OR search OR review OR reviews)'),
        ("seller_policy", '(seller OR sellers OR merchant OR merchants OR commission OR fees OR policy OR rules OR marketplace)'),
        ("social_commerce", '("social commerce" OR "live shopping" OR livestream OR creator OR creators OR "creator shop" OR "Instagram Shop" OR "TikTok Shop")'),
    ]
    if normalized_country == "italy":
        blocks.extend(
            [
                ("italy_service", '(pagamento OR pagamenti OR contanti OR consegna OR spedizione OR reso OR rimborso OR assistenza OR servizio)'),
                ("italy_trust_safety", '(sicurezza OR privacy OR dati OR indagine OR illegale OR contraffatto OR conformità OR qualita OR qualità)'),
                ("italy_market_performance", '("quota di mercato" OR crescita OR vendite OR ricavi OR utenti OR acquirenti OR "comportamento dei consumatori" OR "performance del brand")'),
                ("italy_consumer_experience", '("esperienza cliente" OR "esperienza di acquisto" OR fiducia OR soddisfazione OR reclami OR recensioni OR consumatori)'),
                ("italy_regulatory_market", '(regolamento OR regolatore OR antitrust OR indagine OR inchiesta OR conformità OR "tutela dei consumatori" OR "autorità garante")'),
                ("italy_seller_policy", '(venditore OR venditori OR commercianti OR commissioni OR tariffe OR regole OR marketplace)'),
                ("italy_social_commerce", '("social commerce" OR "shopping su Instagram" OR "acquisti su Instagram" OR "creator shop" OR "live shopping")'),
            ]
        )
    return blocks


def build_related_news_query_blocks(
    country_code: str = DEFAULT_COUNTRY_CODE,
    recall_mode: str = DEFAULT_RECALL_MODE,
) -> list[tuple[str, str]]:
    blocks = list(related_news_query_blocks(country_code))
    if clean_text(recall_mode).lower() != "balanced":
        return blocks
    for label, block in recall_enhanced_related_news_query_blocks(country_code):
        if not any(clean_text(existing_block) == clean_text(block) for _, existing_block in blocks):
            blocks.append((label, block))
    return blocks


def report_query_blocks(country_code: str = DEFAULT_COUNTRY_CODE) -> list[tuple[str, str]]:
    return [
        (str(label), str(block))
        for label, block in country_list_setting(country_code, "report_query_blocks")
    ]


def default_related_news_search_keywords_text(
    country_code: str = DEFAULT_COUNTRY_CODE,
    recall_mode: str = DEFAULT_RECALL_MODE,
) -> str:
    return '\n'.join(block for _, block in build_related_news_query_blocks(country_code, recall_mode))


def default_report_search_keywords_text(country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    return '\n'.join(block for _, block in report_query_blocks(country_code))


LEGACY_COUNTRY_PROMO_KEYWORD_TEXTS: dict[str, list[str]] = {
    "es": [
        '(Oferta OR Descuento OR Cupón OR Promoción OR Rebajas OR Venta OR Cashback OR Puntos OR Lealtad OR Sale OR Deal OR Coupon)',
    ],
    "italy": [
        '(offerta OR sconto OR promozione OR coupon OR saldi OR vendita OR affare OR codici sconto OR deal OR sale)\n'
        '(punti OR fedeltà OR cashback OR premio OR ricompensa OR punti fedeltà OR loyalty OR points)',
    ],
}


def _split_keyword_lines(value: str | None) -> list[str]:
    return [
        line.strip()
        for line in str(value or '').splitlines()
        if line.strip()
    ]


def split_keyword_line_into_blocks(line: str) -> list[str]:
    cleaned = str(line or "").strip()
    if not cleaned:
        return []
    if any(separator in cleaned for separator in ["=", ":"]):
        for separator in ["=", ":"]:
            if separator in cleaned:
                possible_label, possible_block = cleaned.split(separator, 1)
                if possible_block.strip():
                    prefix = clean_text(possible_label)
                    split_blocks = split_keyword_line_into_blocks(possible_block.strip())
                    if len(split_blocks) > 1:
                        return [f"{prefix}_{index}: {block}" if prefix else block for index, block in enumerate(split_blocks, start=1)]
                    return [cleaned]
    matches = list(re.finditer(r"\([^()]+\)", cleaned))
    if len(matches) <= 1:
        return [cleaned]
    leftover = re.sub(r"\([^()]+\)", "", cleaned).strip()
    if leftover:
        return [cleaned]
    return [match.group(0).strip() for match in matches if match.group(0).strip()]


def split_keyword_text_into_blocks(value: str | None) -> list[str]:
    blocks: list[str] = []
    for line in _split_keyword_lines(value):
        blocks.extend(split_keyword_line_into_blocks(line))
    return blocks


def keyword_auto_split_count(value: str | None) -> int:
    count = 0
    for line in _split_keyword_lines(value):
        split_count = len(split_keyword_line_into_blocks(line))
        if split_count > 1:
            count += split_count - 1
    return count


def normalize_keyword_blocks_for_storage(value: str | None) -> str:
    blocks = split_keyword_text_into_blocks(value)
    return "\n".join(blocks) if blocks else str(value or "").strip()


def _matches_legacy_country_keyword_text(value: str | None, country_code: str) -> bool:
    normalized = clean_text(value)
    if not normalized:
        return False
    for candidate in LEGACY_COUNTRY_PROMO_KEYWORD_TEXTS.get(country_code, []):
        if normalized == clean_text(candidate):
            return True
    return False


def _is_default_keyword_prefix_subset(value: str | None, default_text: str) -> bool:
    current_lines = _split_keyword_lines(value)
    default_lines = _split_keyword_lines(default_text)
    if not current_lines or len(current_lines) >= len(default_lines):
        return False
    return current_lines == default_lines[:len(current_lines)]


def normalize_promo_search_keywords_text(value: str | None, country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    normalized = clean_text(value)
    country_default = default_promo_search_keywords_text(country_code)
    if not normalized:
        return country_default
    if normalized == clean_text(country_default):
        return country_default
    if _matches_legacy_country_keyword_text(value, country_code):
        return country_default
    if _is_default_keyword_prefix_subset(value, country_default):
        return country_default
    for code in available_country_codes():
        candidate = default_promo_search_keywords_text(code)
        if normalized == clean_text(candidate):
            return country_default
    return str(value or '')


def normalize_related_news_search_keywords_text(
    value: str | None,
    country_code: str = DEFAULT_COUNTRY_CODE,
    recall_mode: str = DEFAULT_RECALL_MODE,
) -> str:
    normalized = clean_text(value)
    country_default = default_related_news_search_keywords_text(country_code, recall_mode)
    if not normalized:
        return country_default
    if normalized == clean_text(country_default):
        return country_default
    if _matches_legacy_country_keyword_text(value, country_code):
        return country_default
    if _is_default_keyword_prefix_subset(value, country_default):
        return country_default
    for code in available_country_codes():
        candidate = default_related_news_search_keywords_text(code)
        if normalized == clean_text(candidate):
            return country_default
    return str(value or '')


def normalize_report_search_keywords_text(value: str | None, country_code: str = DEFAULT_COUNTRY_CODE) -> str:
    normalized = clean_text(value)
    country_default = default_report_search_keywords_text(country_code)
    if not normalized:
        return country_default
    if normalized == clean_text(country_default):
        return country_default
    for code in available_country_codes():
        candidate = default_report_search_keywords_text(code)
        if normalized == clean_text(candidate):
            return country_default
    return str(value or '')

