import json
import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

import db_store  # noqa: E402


class DbStoreTest(unittest.TestCase):
    def test_init_db_creates_expected_tables(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "test.db"
            db_store.init_db(db_path)

            tables = set(db_store.list_table_names(db_path))

            self.assertIn("runs", tables)
            self.assertIn("articles", tables)
            self.assertIn("article_stars", tables)
            self.assertIn("manual_articles", tables)
            self.assertIn("schema_migrations", tables)

    def test_save_run_is_idempotent_and_loads_before_after(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "test.db"
            output_dir = Path(temp_dir) / "italy_run_1"
            metadata = {
                "country_code": "italy",
                "country_label": "意大利",
                "generated_at": "2026-05-19T00:00:00+00:00",
            }
            before_rows = [
                {"title": "Before article", "platform_label": "Amazon", "published_at": "2026-05-18"},
                {"title": "Before article 2", "platform_label": "IG", "published_at": "2026-05-17"},
            ]
            after_rows = [
                {"title": "After article", "platform_label": "Amazon", "published_at": "2026-05-18"},
            ]

            run_id = db_store.save_run(output_dir, metadata, before_rows, after_rows, db_path=db_path)
            db_store.save_run(output_dir, metadata, before_rows, after_rows, db_path=db_path)

            runs = db_store.list_runs("italy", db_path=db_path)
            self.assertEqual(len(runs), 1)
            self.assertEqual(runs[0]["run_id"], run_id)
            self.assertEqual(runs[0]["before_count"], 2)
            self.assertEqual(runs[0]["after_count"], 1)
            self.assertEqual([row["title"] for row in db_store.load_articles(run_id, "before", db_path=db_path)], ["Before article", "Before article 2"])
            self.assertEqual([row["title"] for row in db_store.load_articles(run_id, "after", db_path=db_path)], ["After article"])

    def test_manual_articles_can_be_added_updated_and_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "test.db"
            output_dir = Path(temp_dir) / "italy_run_manual"
            run_id = db_store.save_run(
                output_dir,
                {"country_code": "italy", "country_label": "Italy"},
                [],
                [{"title": "Crawler article", "platform_label": "Amazon", "published_at": "2026-05-18"}],
                db_path=db_path,
            )

            article = {
                "title": "Manual article",
                "platform_label": "SHEIN",
                "article_url": "https://example.com/manual",
                "published_at": "2026-05-19",
            }
            manual_id, updated = db_store.save_manual_article(run_id, "italy", article, db_path=db_path)
            self.assertFalse(updated)
            self.assertEqual([row["title"] for row in db_store.load_manual_articles(run_id, db_path=db_path)], ["Manual article"])

            article["title"] = "Manual article updated"
            same_id, updated = db_store.save_manual_article(run_id, "italy", article, db_path=db_path)
            self.assertEqual(same_id, manual_id)
            self.assertTrue(updated)
            self.assertEqual([row["title"] for row in db_store.load_manual_articles(run_id, db_path=db_path)], ["Manual article updated"])

            runs = db_store.list_runs("italy", db_path=db_path)
            self.assertEqual(runs[0]["after_count"], 2)
            self.assertTrue(db_store.set_manual_article_enabled(manual_id, False, db_path=db_path))
            self.assertEqual(db_store.load_manual_articles(run_id, db_path=db_path), [])

    def test_star_json_migrates_once_and_new_writes_use_sqlite(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            db_path = Path(temp_dir) / "test.db"
            star_json = Path(temp_dir) / "article_star_store.json"
            star_json.write_text(
                json.dumps(
                    {
                        "old-star": {
                            "starred": True,
                            "updated_at": 123,
                            "title": "Old starred article",
                            "article_url": "https://example.com/old",
                        }
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            migrated_count = db_store.migrate_star_json_if_needed(star_json, db_path=db_path)
            self.assertEqual(migrated_count, 1)
            self.assertIn("old-star", db_store.load_starred_article_ids(db_path=db_path))

            db_store.set_article_star(
                "new-star",
                {"title": "New starred article", "article_url": "https://example.com/new"},
                True,
                db_path=db_path,
            )
            db_store.set_article_star("old-star", {}, False, db_path=db_path)

            store = db_store.load_star_store(db_path=db_path)
            self.assertNotIn("old-star", store)
            self.assertEqual(store["new-star"]["title"], "New starred article")
            self.assertEqual(db_store.migrate_star_json_if_needed(star_json, db_path=db_path), 0)
            self.assertNotIn("old-star", db_store.load_star_store(db_path=db_path))


if __name__ == "__main__":
    unittest.main()
