import json
import sys
import unittest
from pathlib import Path
from typing import Optional


PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

from news_crawler import add_translation_fields  # noqa: E402
from xlsx_source_test import (  # noqa: E402
    SurveyIndicator,
    apply_industry_trend_fields,
    build_survey_ai_batch_messages,
    enrich_row_with_survey_match,
    infer_industry_trend_from_article,
)


class FakeTranslator:
    target_language = "zh-CN"

    def __init__(self, mapping: dict[str, str]) -> None:
        self.mapping = mapping

    def translate(self, text: Optional[str]) -> Optional[str]:
        return self.mapping.get(text or "", text)

    def needs_retry(self, original: str, translated: Optional[str]) -> bool:
        return False


class SurveyReasonHandlingTest(unittest.TestCase):
    def test_prompt_requires_simplified_chinese_reason(self) -> None:
        indicator = SurveyIndicator(
            dimension="Price",
            question_id="B3_8",
            prompt_en="Promotions and discounts are attractive",
            prompt_zh="\u5e73\u53f0\u7684\u4fc3\u9500\u548c\u6298\u6263\u5bf9\u6211\u5f88\u6709\u5438\u5f15\u529b",
        )

        messages = build_survey_ai_batch_messages(
            [(0, {"title": "Amazon Spring Deals are coming"})],
            [indicator],
            None,
        )
        payload = json.loads(messages[1]["content"])

        self.assertTrue(
            any("Simplified Chinese" in rule for rule in payload["strict_rules"]),
            payload["strict_rules"],
        )

    def test_ai_reason_is_split_and_translated_without_touching_chinese_question(self) -> None:
        indicator = SurveyIndicator(
            dimension="Price",
            question_id="B3_8",
            prompt_en="Promotions and discounts are attractive",
            prompt_zh="[平台]的促销和折扣对我很有吸引力",
        )
        reason = "News announces Amazon Spring Offers, which can affect consumer perception of discounts."
        reason_zh = "新闻宣布亚马逊春季优惠，这会影响消费者对折扣的看法。"

        enriched = enrich_row_with_survey_match(
            {},
            matched_dimensions=["Price"],
            matched_question_ids=["B3_8"],
            grouped_indicators={"Price": [indicator]},
            question_lookup={"b3_8": indicator},
            explanation=reason,
            method="ai_batch",
        )

        self.assertEqual(enriched["survey_ai_reason_raw"], reason)
        self.assertEqual(enriched["survey_ai_reason_translated"], reason)
        self.assertIn("[平台]的促销和折扣对我很有吸引力", enriched["survey_indicator_examples"])

        translated = add_translation_fields(
            [enriched],
            FakeTranslator({reason: reason_zh}),
            extra_text_fields=["survey_indicator_examples", "survey_ai_reason_translated"],
        )[0]

        self.assertEqual(translated["survey_ai_reason_raw"], reason)
        self.assertEqual(translated["survey_ai_reason_translated"], reason_zh)
        self.assertIn("[平台]的促销和折扣对我很有吸引力", translated["survey_indicator_examples"])
        self.assertIn(reason_zh, translated["survey_indicator_examples"])
        self.assertNotIn("News announces", translated["survey_indicator_examples"])

    def test_industry_trend_prompt_and_local_inference(self) -> None:
        indicator = SurveyIndicator(
            dimension="Content",
            question_id="B4_1",
            prompt_en="The platform content is attractive",
            prompt_zh="平台内容有吸引力",
        )
        messages = build_survey_ai_batch_messages(
            [
                (
                    0,
                    {
                        "platform_label": "TTS",
                        "title": "E-commerce, in Italia quasi 1 consumatore su 5 usa TikTok Shop",
                        "summary": "A report says social commerce adoption is growing in Italy.",
                    },
                )
            ],
            [indicator],
            None,
        )
        payload = json.loads(messages[1]["content"])

        self.assertTrue(any("industry_trend_flag" in rule for rule in payload["strict_rules"]))
        self.assertTrue(any("TikTok Shop/TTS" in rule for rule in payload["strict_rules"]))
        self.assertTrue(any("Instagram/IG" in rule for rule in payload["strict_rules"]))
        self.assertTrue(any("eBay Live" in rule for rule in payload["strict_rules"]))
        self.assertTrue(any("takeover" in rule for rule in payload["strict_rules"]))
        inferred = infer_industry_trend_from_article(
            {
                "platform_label": "TTS",
                "title": "E-commerce, in Italia quasi 1 consumatore su 5 usa TikTok Shop",
                "summary": "A report says social commerce adoption is growing in Italy.",
            }
        )
        self.assertTrue(inferred["industry_trend_flag"])
        self.assertEqual(inferred["industry_trend_category"], "market_adoption")

    def test_enriched_row_keeps_ai_industry_trend_fields(self) -> None:
        indicator = SurveyIndicator(
            dimension="Quality",
            question_id="B1_3",
            prompt_en="Sellers are legitimate and trustworthy",
            prompt_zh="卖家合法可信",
        )

        enriched = enrich_row_with_survey_match(
            {
                "platform_label": "SHEIN",
                "title": "EU investigates Shein illegal products and platform compliance trend",
            },
            matched_dimensions=["Quality"],
            matched_question_ids=["B1_3"],
            grouped_indicators={"Quality": [indicator]},
            question_lookup={"b1_3": indicator},
            explanation="监管调查可能影响消费者对平台可信度的整体认知。",
            method="ai_batch",
            industry_trend_flag=True,
            industry_trend_category="regulatory_risk",
            industry_trend_impact="Negative",
            industry_trend_reason="新闻涉及监管和合规趋势，会影响平台整体信任。",
        )

        self.assertEqual(enriched["industry_trend_flag"], "true")
        self.assertEqual(enriched["industry_trend_category"], "regulatory_risk")
        self.assertEqual(enriched["industry_trend_impact"], "Negative")
        self.assertIn("监管和合规趋势", enriched["industry_trend_reason"])

        applied = apply_industry_trend_fields({"title": "Generic unrelated article"})
        self.assertEqual(applied["industry_trend_flag"], "")


if __name__ == "__main__":
    unittest.main()
