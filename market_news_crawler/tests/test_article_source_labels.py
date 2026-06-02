import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

import article_browser  # noqa: E402
import db_store  # noqa: E402


class ArticleSourceLabelTest(unittest.TestCase):
    def test_user_sources_show_friendly_after_only_label_and_keep_value(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "test.db"
            original_db_path = db_store.DEFAULT_DB_PATH
            try:
                db_store.DEFAULT_DB_PATH = db_path
                run_id = db_store.save_run(
                    Path(temp_dir)
                    / "italy_xlsx_sources_media_buyer_seller_run_20260601_122554_range_20260502_20260601",
                    {
                        "country_code": "italy",
                        "country_label": "意大利",
                        "generated_at": "2026-06-01T06:18:39+02:00",
                        "range_start": "2026-05-02T00:00:00+02:00",
                        "range_end": "2026-06-01T23:59:59+02:00",
                    },
                    [{"title": "Before article", "platform_label": "Amazon"}],
                    [{"title": "After article", "platform_label": "Amazon"}],
                    db_path=db_path,
                )

                sources = article_browser.list_user_article_sources(country_code="italy")

                self.assertGreaterEqual(len(sources), 1)
                self.assertEqual(sources[0]["value"], f"db:{run_id}:after")
                self.assertIn("最近一次抓取", sources[0]["label"])
                self.assertIn("意大利", sources[0]["label"])
                self.assertIn("2026-05-02 至 2026-06-01", sources[0]["label"])
                self.assertIn("1条", sources[0]["label"])
                self.assertNotIn("xlsx_sources", sources[0]["label"])
                self.assertNotIn(":before", sources[0]["value"])
            finally:
                db_store.DEFAULT_DB_PATH = original_db_path

    def test_friendly_label_handles_old_run_without_range(self) -> None:
        label = article_browser.friendly_article_source_label(
            country_label="意大利",
            run_id="italy_xlsx_sources_media_buyer_seller_run_20260520_143052",
            article_count=5,
        )

        self.assertIn("2026-05-20 14:30 抓取", label)
        self.assertIn("时间范围未知", label)
        self.assertIn("5条", label)


if __name__ == "__main__":
    unittest.main()
