import unittest
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import sys

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import xlsx_source_test as crawler
import dedupe


class SearchRecallImprovementTests(unittest.TestCase):
    def test_run_output_dir_name_includes_execution_time_and_range(self):
        name = crawler.build_run_output_dir_name(
            "italy",
            "media_buyer_seller",
            "20260601_153000",
            date(2026, 5, 1),
            date(2026, 5, 31),
        )

        self.assertEqual(
            name,
            "italy_xlsx_sources_media_buyer_seller_run_20260601_153000_range_20260501_20260531",
        )

    def test_initial_dedupe_keeps_no_url_different_titles_same_platform(self):
        rows = [
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon launches cash payment points in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-a.it",
                "article_url": "",
            },
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon announces Prime member gifts in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-b.it",
                "article_url": "",
            },
        ]
        stats = {}

        deduped = dedupe.dedupe_articles(rows, stats=stats)

        self.assertEqual(len(deduped), 2)
        self.assertEqual(stats["removed_count"], 0)
        self.assertEqual(stats["missing_url_count"], 2)

    def test_initial_dedupe_merges_same_url(self):
        rows = [
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon cash payment in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-a.it",
                "article_url": "https://example.test/news/amazon-cash?utm_source=rss",
            },
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon cash payment in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-a.it",
                "article_url": "https://example.test/news/amazon-cash",
            },
        ]
        stats = {}

        deduped = dedupe.dedupe_articles(rows, stats=stats)

        self.assertEqual(len(deduped), 1)
        self.assertEqual(stats["removed_by_reason"].get("url_duplicate"), 1)

    def test_initial_dedupe_merges_same_day_exact_title_with_source_suffix(self):
        rows = [
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon expands cash payment points in Italy - Reuters",
                "published_at": "2026-05-18",
                "source_site": "reuters.com",
                "article_url": "",
            },
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon expands cash payment points in Italy - Il Sole 24 ORE",
                "published_at": "2026-05-18",
                "source_site": "ilsole24ore.com",
                "article_url": "",
            },
        ]
        stats = {}

        deduped = dedupe.dedupe_articles(rows, stats=stats)

        self.assertEqual(len(deduped), 1)
        self.assertEqual(stats["removed_by_reason"].get("same_day_exact_title_duplicate"), 1)

    def test_initial_dedupe_keeps_similar_but_different_events(self):
        rows = [
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon expands cash payment points in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-a.it",
                "article_url": "",
            },
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon expands Prime delivery points in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-b.it",
                "article_url": "",
            },
        ]

        self.assertEqual(len(dedupe.dedupe_articles(rows)), 2)

    def test_initial_dedupe_keeps_same_title_on_different_dates_for_ai_dedupe(self):
        rows = [
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon expands cash payment points in Italy",
                "published_at": "2026-05-18",
                "source_site": "example-a.it",
                "article_url": "",
            },
            {
                "platform": "Amazon",
                "platform_label": "Amazon",
                "title": "Amazon expands cash payment points in Italy",
                "published_at": "2026-05-14",
                "source_site": "example-b.it",
                "article_url": "",
            },
        ]

        deduped = dedupe.dedupe_articles(rows)

        self.assertEqual(len(deduped), 2)
        self.assertEqual(len(dedupe.build_ai_dedupe_candidate_clusters(deduped)), 1)

    def test_low_volume_platforms_use_dedicated_google_queries(self):
        for platform in ["TEMU", "TTS", "IG"]:
            tasks = crawler.build_promo_search_tasks(
                [platform],
                "both",
                related_news_enabled=True,
                report_ranking_enabled=False,
                country_code="italy",
                recall_mode="balanced",
            )
            self.assertEqual({task["engine"] for task in tasks}, {"google"})
            self.assertIn("platform_dedicated_google", {task["query_strategy"] for task in tasks})
            self.assertIn("platform_broad_eu_global", {task["query_strategy"] for task in tasks})
            if platform in {"TEMU", "TTS", "IG"}:
                self.assertIn("platform_broad_short", {task["query_strategy"] for task in tasks})

    def test_low_volume_broad_queries_are_balanced_only(self):
        balanced_tasks = crawler.build_promo_search_tasks(
            ["TEMU", "IG"],
            "both",
            related_news_enabled=True,
            report_ranking_enabled=False,
            country_code="italy",
            recall_mode="balanced",
        )
        strict_tasks = crawler.build_promo_search_tasks(
            ["TEMU", "IG"],
            "both",
            related_news_enabled=True,
            report_ranking_enabled=False,
            country_code="italy",
            recall_mode="strict",
        )

        self.assertIn("platform_broad_eu_global", {task["query_strategy"] for task in balanced_tasks})
        self.assertIn("platform_broad_short", {task["query_strategy"] for task in balanced_tasks})
        self.assertNotIn("platform_broad_eu_global", {task["query_strategy"] for task in strict_tasks})
        self.assertNotIn("platform_broad_short", {task["query_strategy"] for task in strict_tasks})
        self.assertTrue(
            all(task.get("market_scope") == crawler.BROAD_ENTRY_MARKET_SCOPE for task in balanced_tasks if task["query_strategy"] in {"platform_broad_eu_global", "platform_broad_short"})
        )
        self.assertTrue(any(task.get("broad_entry_allowed") for task in balanced_tasks if task["query_strategy"] == "platform_dedicated_google"))
        self.assertFalse(any(task.get("broad_entry_allowed") for task in strict_tasks))

    def test_balanced_related_keywords_include_market_and_consumer_blocks(self):
        blocks = crawler.normalize_related_news_search_keyword_blocks("", "italy", recall_mode="balanced")
        combined = "\n".join(block for _, block in blocks)

        self.assertIn("market share", combined)
        self.assertIn("consumer trust", combined)
        self.assertIn("quota di mercato", combined)
        self.assertIn("esperienza cliente", combined)

    def test_spain_defaults_replace_zalando_with_aliexpress(self):
        config = crawler.get_country_config("es")
        platforms = config.get("available_platform_labels") or []
        aliases = crawler.expanded_platform_search_terms("AliExpress", "es")

        self.assertIn("AliExpress", platforms)
        self.assertNotIn("Zalando", platforms)
        self.assertIn("AliExpress Espa\u00f1a", aliases)
        self.assertIn("AliExpress.es", aliases)
        self.assertIn("AliExpress Plaza", aliases)

    def test_spain_balanced_keywords_include_local_service_blocks(self):
        blocks = crawler.normalize_related_news_search_keyword_blocks("", "es", recall_mode="balanced")
        combined = "\n".join(block for _, block in blocks)

        self.assertIn("experiencia de compra", combined)
        self.assertIn("confianza del consumidor", combined)
        self.assertIn("protecci\u00f3n del consumidor", combined)
        self.assertIn("m\u00e9todo de pago", combined)
        self.assertIn("vendedores", combined)
        self.assertIn("regulaci\u00f3n", combined)
        self.assertIn("social commerce", combined)

    def test_single_line_multiple_keyword_blocks_are_split(self):
        default_text = crawler.default_related_news_search_keywords_text("italy", recall_mode="balanced")
        compact_text = " ".join(default_text.splitlines())

        default_blocks = crawler.normalize_related_news_search_keyword_blocks(default_text, "italy", recall_mode="balanced")
        compact_blocks = crawler.normalize_related_news_search_keyword_blocks(compact_text, "italy", recall_mode="balanced")

        self.assertEqual(len(compact_blocks), len(default_blocks))
        self.assertGreaterEqual(crawler.keyword_auto_split_count(compact_text), 1)

    def test_balanced_appends_system_blocks_to_custom_keywords(self):
        custom_text = "(coupon OR discount)\n(payment OR delivery)"

        balanced_blocks = crawler.normalize_related_news_search_keyword_blocks(custom_text, "italy", recall_mode="balanced")
        strict_blocks = crawler.normalize_related_news_search_keyword_blocks(custom_text, "italy", recall_mode="strict")
        balanced_combined = "\n".join(block for _, block in balanced_blocks)

        self.assertEqual(len(strict_blocks), 2)
        self.assertGreater(len(balanced_blocks), len(strict_blocks))
        self.assertIn("market share", balanced_combined)

    def test_keyword_blocks_for_storage_splits_ai_single_line_output(self):
        compact_text = "(coupon OR discount) (payment OR delivery) (trust OR privacy)"

        stored_text = crawler.normalize_keyword_blocks_for_storage(compact_text)

        self.assertEqual(stored_text.count("\n"), 2)
        self.assertIn("(payment OR delivery)", stored_text.splitlines())

    def test_dedicated_google_queries_include_local_low_volume_terms(self):
        temu_blocks = "\n".join(block for _, block in crawler.dedicated_google_platform_blocks("TEMU", None, "italy"))
        tts_blocks = "\n".join(block for _, block in crawler.dedicated_google_platform_blocks("TTS", None, "italy"))
        ig_blocks = "\n".join(block for _, block in crawler.dedicated_google_platform_blocks("IG", None, "italy"))

        self.assertIn("Temu consumatori italiani", temu_blocks)
        self.assertIn("creator economy", tts_blocks)
        self.assertIn("Instagram per aziende", ig_blocks)
        self.assertIn("negozi Instagram", ig_blocks)
        self.assertIn("Meta commerce", ig_blocks)

        temu_broad_blocks = "\n".join(block for _, block in crawler.broad_eu_global_platform_blocks("TEMU", None, "italy"))
        ig_broad_blocks = "\n".join(block for _, block in crawler.broad_eu_global_platform_blocks("IG", None, "italy"))
        self.assertIn("European Commission", temu_broad_blocks)
        self.assertIn("low value parcel", temu_broad_blocks)
        self.assertIn("Meta", ig_broad_blocks)
        self.assertIn("creator economy", ig_broad_blocks)

        temu_short_blocks = "\n".join(block for _, block in crawler.broad_short_platform_blocks("TEMU", None, "italy"))
        ig_short_blocks = "\n".join(block for _, block in crawler.broad_short_platform_blocks("IG", None, "italy"))
        tts_short_blocks = "\n".join(block for _, block in crawler.broad_short_platform_blocks("TTS", None, "italy"))
        self.assertIn("European Commission", temu_short_blocks)
        self.assertIn("product safety", temu_short_blocks)
        self.assertIn("Digital Services Act", temu_short_blocks)
        self.assertIn("Instagram checkout", ig_short_blocks)
        self.assertIn("Meta commerce", ig_short_blocks)
        self.assertIn("Instagram ads commerce", ig_short_blocks)
        self.assertIn("social commerce Instagram", ig_short_blocks)
        self.assertIn("Instagram shoppable posts", ig_short_blocks)
        self.assertIn("TikTok Shop Italy", tts_short_blocks)
        self.assertIn("TikTok Shop live shopping", tts_short_blocks)
        self.assertIn("TikTok Shop consumer protection", tts_short_blocks)

    def test_tts_and_ig_alias_health_contains_full_names(self):
        tts_terms = crawler.expanded_platform_search_terms("TTS", "italy")
        ig_terms = crawler.expanded_platform_search_terms("IG", "italy")
        summary = crawler.platform_alias_effective_summary(["TTS", "IG"], "italy")

        self.assertIn("TikTok Shop", tts_terms)
        self.assertIn("TikTok", tts_terms)
        self.assertIn("Instagram", ig_terms)
        self.assertIn("Instagram Shop", ig_terms)
        self.assertIn("Instagram Shopping", ig_terms)
        self.assertTrue(summary["TTS"]["has_full_name"])
        self.assertTrue(summary["IG"]["has_full_name"])
        self.assertFalse(summary["TTS"]["abbreviation_only"])
        self.assertFalse(summary["IG"]["abbreviation_only"])

    def test_instagram_weak_commerce_context_is_allowed(self):
        weak_article = {
            "title": "Meta commerce tools expand Instagram business and creator economy options for brands in Italy",
            "summary": "",
        }
        noise_article = {
            "title": "Famous singer posts concert photos on Instagram",
            "summary": "",
        }
        self.assertIn(crawler.instagram_commerce_context_strength(weak_article, "italy"), {"strong", "weak"})
        self.assertTrue(crawler.article_matches_instagram_commerce_context(weak_article, "italy"))
        self.assertFalse(crawler.article_matches_instagram_commerce_context(noise_article, "italy"))
        self.assertEqual(crawler.instagram_non_commerce_skip_reason(noise_article, "italy"), "instagram_entertainment_noise")

    def test_broad_entry_context_rules_for_temu_and_ig(self):
        temu_eu_article = {
            "title": "European Commission investigates Temu over consumer protection and product safety",
            "summary": "",
        }
        temu_unrelated_article = {
            "title": "Temu launches a local coupon campaign in Brazil",
            "summary": "",
        }
        ig_broad_article = {
            "title": "Meta expands Instagram creator economy and shoppable posts tools for retailers",
            "summary": "",
        }
        ig_noise_article = {
            "title": "Celebrity shares vacation photos on Instagram",
            "summary": "",
        }
        tts_broad_article = {
            "title": "TikTok Shop live shopping tools help sellers and merchants in Europe",
            "summary": "",
        }
        tts_noise_article = {
            "title": "TikTok creator posts dance video from a music concert",
            "summary": "",
        }

        self.assertTrue(crawler.article_matches_temu_broad_context(temu_eu_article, "italy"))
        self.assertFalse(crawler.article_matches_temu_broad_context(temu_unrelated_article, "italy"))
        self.assertTrue(crawler.article_matches_instagram_broad_context(ig_broad_article, "italy"))
        self.assertFalse(crawler.article_matches_instagram_broad_context(ig_noise_article, "italy"))
        self.assertTrue(crawler.article_matches_tiktok_shop_broad_context(tts_broad_article, "italy"))
        self.assertFalse(crawler.article_matches_tiktok_shop_broad_context(tts_noise_article, "italy"))

    def test_ig_meta_commerce_can_match_platform_for_dedicated_task(self):
        article = {
            "title": "Meta commerce tools add Instagram shopping workflows for retailers",
            "summary": "",
        }
        task = {"platform_label": "IG", "source_platform": "", "match_platform_labels": ["IG"]}

        self.assertEqual(crawler.matched_platform_labels_for_task(article, task, "italy"), ["IG"])

    def test_multi_platform_regulatory_event_is_temu_broad_entry(self):
        article = {
            "title": "Shein and Temu face EU investigation over illegal products and consumer protection",
            "summary": "Chinese marketplaces are under scrutiny in Europe.",
        }
        task = {"platform_label": "TEMU", "source_platform": "", "query_strategy": "platform_broad_short"}

        self.assertTrue(crawler.article_matches_multi_platform_regulatory_event(article, "italy"))
        self.assertEqual(crawler.broad_entry_reason_for_article(article, task, "italy"), "multi_platform_regulatory_or_market_event")

    def test_ai_exclusion_reason_category_identifies_single_product_guides(self):
        row = {
            "title": "Best Amazon coupon codes for headphones this week",
            "summary": "A shopping guide for single products.",
        }
        decision = crawler.SurveyAIFilterDecision(
            relevant=False,
            matched_dimensions=[],
            matched_question_ids=[],
            reason="single product shopping guide",
            confidence="medium",
        )

        self.assertEqual(crawler.categorize_ai_exclusion_reason(row, decision), "single_product_or_coupon")

    def test_ai_exclusion_reason_category_identifies_creator_and_strategy_edges(self):
        creator_decision = crawler.SurveyAIFilterDecision(
            relevant=False,
            matched_dimensions=[],
            matched_question_ids=[],
            reason="creator economy brand partnership has weak NPS linkage",
            confidence="medium",
        )
        capital_decision = crawler.SurveyAIFilterDecision(
            relevant=False,
            matched_dimensions=[],
            matched_question_ids=[],
            reason="capital market acquisition event has weak consumer trust linkage",
            confidence="medium",
        )
        ecosystem_decision = crawler.SurveyAIFilterDecision(
            relevant=False,
            matched_dimensions=[],
            matched_question_ids=[],
            reason="eBay Live platform ecosystem relevance is indirect",
            confidence="medium",
        )

        self.assertEqual(
            crawler.categorize_ai_exclusion_reason({"title": "Instagram brand partnership tools for Creator Economy"}, creator_decision),
            "creator_commerce_weak_link",
        )
        self.assertEqual(
            crawler.categorize_ai_exclusion_reason({"title": "GameStop proposes takeover of eBay"}, capital_decision),
            "capital_market_event_weak_link",
        )
        self.assertEqual(
            crawler.categorize_ai_exclusion_reason({"title": "Behind the scenes of eBay Live"}, ecosystem_decision),
            "platform_ecosystem_weak_link",
        )

    def test_brand_stage_total_summary_keeps_all_selected_brands(self):
        weekly_rows = [
            {"brand": "Amazon", "week_start": "2026-04-20", "raw_count": 10, "initial_dedupe_count": 8, "survey_filter_count": 4, "final_count": 2},
            {"brand": "Amazon", "week_start": "2026-04-27", "raw_count": 6, "initial_dedupe_count": 5, "survey_filter_count": 2, "final_count": 1},
            {"brand": "SHEIN", "week_start": "2026-04-20", "raw_count": 0, "initial_dedupe_count": 0, "survey_filter_count": 0, "final_count": 0},
            {"brand": "TEMU", "week_start": "2026-04-20", "raw_count": 1, "initial_dedupe_count": 1, "survey_filter_count": 0, "final_count": 0},
            {"brand": "IG", "week_start": "2026-04-20", "raw_count": 5, "initial_dedupe_count": 5, "survey_filter_count": 0, "final_count": 0},
        ]

        summary = crawler.build_brand_stage_total_summary(
            ["Amazon", "SHEIN", "TEMU", "TTS", "eBay", "IG"],
            weekly_rows,
        )

        self.assertEqual([row["brand"] for row in summary], ["Amazon", "SHEIN", "TEMU", "TTS", "eBay", "IG"])
        amazon = next(row for row in summary if row["brand"] == "Amazon")
        temu = next(row for row in summary if row["brand"] == "TEMU")
        ig = next(row for row in summary if row["brand"] == "IG")
        self.assertEqual(amazon["raw_count"], 16)
        self.assertEqual(amazon["final_count"], 3)
        self.assertIn("搜索入池阶段", temu["diagnosis"])
        self.assertIn("AI 指标筛选阶段", ig["diagnosis"])

    def test_configured_adapter_parses_h3_and_text_date(self):
        html = """
        <html><body>
          <article>
            <time>Jul 05, 2023</time>
            <h3><a href="https://example.test/news/amazon-impact">Amazon Impact Report</a></h3>
            <p>Amazon report summary</p>
          </article>
        </body></html>
        """
        config = {
            "search_url_template": "https://example.test/?s={query}",
            "item_selector": "article",
            "link_selector": "h3 a[href]",
            "link_attr": "href",
            "title_selector": "h3 a",
            "date_selector": "time",
            "date_attr": "",
            "summary_selector": "p",
            "allow_search_only": True,
        }
        entry = crawler.SourceEntry("General Media", "media", "https://example.test/")
        diagnostics = {}

        def fake_fetch(_session, url):
            if "?s=" in url:
                return SimpleNamespace(status_code=200, url=url, text=html)
            raise RuntimeError("skip article hydration")

        with patch.object(crawler, "fetch", side_effect=fake_fetch):
            rows = crawler.collect_configured_adapter_articles(
                entry,
                object(),
                datetime.fromisoformat("2023-01-01T00:00:00+00:00"),
                datetime.fromisoformat("2023-12-31T23:59:59+00:00"),
                config,
                ["Amazon"],
                "italy",
                diagnostics=diagnostics,
            )

        self.assertEqual(len(rows), 1)
        self.assertEqual(diagnostics["selector_match_count"], 1)
        self.assertEqual(diagnostics["parsed_date_count"], 1)
        self.assertEqual(rows[0]["title"], "Amazon Impact Report")


if __name__ == "__main__":
    unittest.main()
