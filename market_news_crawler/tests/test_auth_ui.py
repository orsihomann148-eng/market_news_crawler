import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock


PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

import web_app  # noqa: E402


class AuthUiTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self._original_settings_path = web_app.APP_SETTINGS_PATH
        web_app.APP_SETTINGS_PATH = Path(self._tmpdir.name) / "settings.json"
        self.addCleanup(lambda: setattr(web_app, "APP_SETTINGS_PATH", self._original_settings_path))
        web_app.app.config["TESTING"] = True
        self.addCleanup(lambda: web_app.app.config.update(TESTING=False))
        web_app.WEB_VIEW_STATE.clear()
        web_app.WEB_VIEW_STATE.update(
            {
                "result": None,
                "active_tab": "home",
                "country_code": web_app.DEFAULT_COUNTRY_CODE,
                "news_form_state": None,
                "news_platforms_text": None,
                "news_sides": web_app.DEFAULT_NEWS_SIDES.copy(),
                "news_filter_state": None,
                "source_state": None,
            }
        )

    def login_client(self, client) -> None:
        web_app.save_admin_password("password123")
        with client.session_transaction() as session:
            session[web_app.AUTH_SESSION_KEY] = True

    def save_api_settings(self) -> None:
        web_app.persist_global_survey_api_settings(
            {
                "survey_api_url": "https://api.example.com/chat/completions",
                "survey_api_key": "secret-key",
                "survey_api_model": "test-model",
            }
        )

    def sample_rows(self) -> list[dict[str, str]]:
        return [
            {
                "article_id": "article-1",
                "platform_label": "Amazon",
                "title_original": "Amazon expands cash payment points in Italy",
                "title": "Amazon expands cash payment points in Italy",
                "title_display_zh": "Amazon 扩大意大利现金支付点",
                "survey_dimensions": "Features",
                "survey_indicator_examples": "支付方式更丰富，可能提升便利性感知。",
                "briefing_sentiment": "Positive",
                "published_at_display": "2026-05-18",
                "article_url": "https://example.com/1",
            },
            {
                "article_id": "article-2",
                "platform_label": "IG",
                "title_original": "Instagram tests new shopping tools",
                "title": "Instagram tests new shopping tools",
                "title_display_zh": "Instagram 测试新的购物工具",
                "survey_dimensions": "Features",
                "survey_indicator_examples": "购物工具升级可能影响内容转化。",
                "briefing_sentiment": "Neutral",
                "published_at_display": "2026-05-17",
                "article_url": "https://example.com/2",
            },
        ]

    def test_first_visit_redirects_to_setup_and_api_returns_401(self) -> None:
        client = web_app.app.test_client()
        response = client.get("/")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/setup-admin", response.headers.get("Location", ""))

        api_response = client.get("/api/jobs/missing")
        self.assertEqual(api_response.status_code, 401)
        self.assertEqual(api_response.get_json()["error"], "auth_not_configured")

    def test_setup_login_user_and_developer_pages(self) -> None:
        client = web_app.app.test_client()
        setup_response = client.post(
            "/setup-admin",
            data={"password": "password123", "confirm_password": "password123", "next": "/"},
            follow_redirects=True,
        )
        self.assertEqual(setup_response.status_code, 200)
        html = setup_response.get_data(as_text=True)
        self.assertIn("普通用户模式", html)
        self.assertIn("AI 配置", html)
        self.assertIn("新闻抓取", html)
        self.assertIn("新闻查看", html)
        self.assertIn("添加新闻", html)
        self.assertIn("导出材料", html)
        self.assertIn("请先填写 AI API 配置", html)
        self.assertNotIn("国家管理", html)
        self.assertNotIn("来源管理", html)

        developer_response = client.get("/developer")
        self.assertEqual(developer_response.status_code, 200)
        self.assertIn("普通用户模式", developer_response.get_data(as_text=True))

        client.get("/logout")
        blocked_response = client.get("/")
        self.assertEqual(blocked_response.status_code, 302)
        self.assertIn("/login", blocked_response.headers.get("Location", ""))

    def test_user_page_shows_api_form_when_not_configured(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)

        response = client.get("/?country_code=italy")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('id="user-survey-api-url"', html)
        self.assertIn('id="user-survey-api-key"', html)
        self.assertIn('id="user-survey-api-model"', html)
        self.assertIn('id="user-survey-api-test"', html)
        self.assertIn('id="user-survey-api-save"', html)
        self.assertIn('id="user-run-button" disabled', html)

    def test_user_page_shows_compact_api_status_when_configured(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)
        self.save_api_settings()

        response = client.get("/?country_code=italy")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("AI API 已配置：test-model", html)
        self.assertIn("重新配置", html)
        self.assertNotIn('value="secret-key"', html)
        self.assertIn('id="user-run-button"', html)

    def test_news_form_state_can_save_api_from_user_flow(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)

        response = client.post(
            "/api/news-form-state",
            data={
                "country_code": "italy",
                "date_mode": "days",
                "days": "30",
                "translate_to": "zh-CN",
                "output_dir": "outputs",
                "recall_mode": "balanced",
                "survey_filter_mode": "ai",
                "related_news_search_enabled": "1",
                "report_ranking_search_enabled": "1",
                "promo_search_engine": "both",
                "news_sides": ["media", "buyer", "seller"],
                "news_builtin_platforms": ["Amazon", "TTS"],
                "survey_api_url": "https://api.example.com/chat/completions",
                "survey_api_key": "saved-key",
                "survey_api_model": "saved-model",
            },
        )
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload["ok"])

        saved = web_app.read_global_survey_api_settings()
        self.assertEqual(saved["survey_api_url"], "https://api.example.com/chat/completions")
        self.assertEqual(saved["survey_api_key"], "saved-key")
        self.assertEqual(saved["survey_api_model"], "saved-model")

    def test_user_platform_labels_are_full_names_but_values_stay_internal(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)

        response = client.get("/?country_code=italy")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('value="TTS"', html)
        self.assertIn(">TikTok Shop</span>", html)
        self.assertIn('value="IG"', html)
        self.assertIn(">Instagram</span>", html)

    def test_developer_platform_labels_are_full_names_but_values_stay_internal(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)

        response = client.get("/developer?country_code=italy&tab=news")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('value="TTS"', html)
        self.assertIn("TikTok Shop", html)
        self.assertIn('value="IG"', html)
        self.assertIn("Instagram", html)

    def test_platform_display_label_handles_multi_brand_text(self) -> None:
        self.assertEqual(web_app.user_platform_display_label("TTS | IG"), "TikTok Shop | Instagram")
        self.assertEqual(web_app.user_platform_display_label("TTS, IG"), "TikTok Shop, Instagram")
        self.assertEqual(web_app.user_platform_display_label("TTS/IG"), "TikTok Shop/Instagram")

    def test_manual_article_suggest_requires_db_after_and_url(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)

        csv_response = client.post(
            "/api/articles/manual/suggest",
            data={
                "country_code": "italy",
                "article_source": "outputs/example/after.csv",
                "article_url": "https://example.com/a",
            },
        )
        self.assertEqual(csv_response.status_code, 400)

        missing_url_response = client.post(
            "/api/articles/manual/suggest",
            data={"country_code": "italy", "article_source": "db:test_run:after", "article_url": ""},
        )
        self.assertEqual(missing_url_response.status_code, 400)

    def test_manual_article_suggest_returns_fields_without_saving(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)
        self.save_api_settings()

        with mock.patch.object(
            web_app,
            "fetch_manual_article_url_context",
            return_value={"title": "Temu safety news"},
        ), mock.patch.object(
            web_app.xlsx_source_test,
            "call_survey_filter_api",
            return_value={
                "title": "Temu faces EU product safety checks",
                "title_translated": "Temu 面临欧盟商品安全检查",
                "article_url": "https://example.com/temu",
                "published_at": "2026-05-20T12:00:00+00:00",
                "source_name": "Example News",
                "platform_label": "TEMU",
                "summary": "EU checks may affect platform trust.",
                "survey_dimensions": "Quality",
                "survey_question_ids": "B1_3",
                "survey_indicator_examples": "商品安全检查可能影响消费者对平台合规的信任。",
                "briefing_sentiment": "Negative",
                "industry_trend_flag": True,
                "industry_trend_category": "regulatory_risk",
                "industry_trend_impact": "Negative",
                "industry_trend_reason": "欧盟监管会影响跨境平台形象。",
            },
        ):
            response = client.post(
                "/api/articles/manual/suggest",
                data={
                    "country_code": "italy",
                    "article_source": "db:test_run:after",
                    "article_url": "https://example.com/temu",
                },
            )

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["fields"]["platform_label"], "TEMU")
        self.assertEqual(payload["fields"]["published_at"], "2026-05-20")
        self.assertEqual(payload["dataset"]["titleTranslated"], "Temu 面临欧盟商品安全检查")

    def test_selected_article_ids_filter_briefing_table_generation(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)
        self.save_api_settings()
        captured = {}

        def fake_build_article_briefing_table(rows, country_code, api_settings, ai_stats):
            captured["rows"] = rows
            ai_stats.update({"batch_count": 0, "failed_batch_count": 0, "failed_row_count": 0, "ai_completed_row_count": 0})
            output_path = Path(self._tmpdir.name) / "briefing.xlsx"
            output_path.write_bytes(b"test")
            return output_path

        with mock.patch.object(
            web_app,
            "load_article_rows_from_db_source",
            return_value=(self.sample_rows(), [], 2, 2, 0),
        ), mock.patch.object(
            web_app,
            "build_article_briefing_table",
            side_effect=fake_build_article_briefing_table,
        ):
            response = client.post(
                "/api/articles/briefing-table",
                data={
                    "country_code": "italy",
                    "article_source": "db:test_run:after",
                    "article_selection_mode": "selected",
                    "selected_article_ids": ["article-2"],
                },
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual([row["article_id"] for row in captured["rows"]], ["article-2"])

    def test_selected_article_ids_filter_news_summary_generation(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)
        self.save_api_settings()
        captured = {}

        def fake_generate_news_summary(rows, country_code, api_settings, stats):
            captured["rows"] = rows
            stats.update({"failed_row_count": 0})
            return SimpleNamespace(
                text="summary",
                output_path=Path(self._tmpdir.name) / "summary.txt",
                stats={"failed_row_count": 0},
            )

        with mock.patch.object(
            web_app,
            "load_article_rows_from_db_source",
            return_value=(self.sample_rows(), [], 2, 2, 0),
        ), mock.patch.object(
            web_app.news_summary_utils,
            "generate_news_summary",
            side_effect=fake_generate_news_summary,
        ):
            response = client.post(
                "/api/articles/news-summary",
                data={
                    "country_code": "italy",
                    "article_source": "db:test_run:after",
                    "article_selection_mode": "selected",
                    "selected_article_ids": ["article-1"],
                },
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual([row["article_id"] for row in captured["rows"]], ["article-1"])

    def test_selected_article_generation_requires_at_least_one_article(self) -> None:
        client = web_app.app.test_client()
        self.login_client(client)
        self.save_api_settings()

        with mock.patch.object(
            web_app,
            "load_article_rows_from_db_source",
            return_value=(self.sample_rows(), [], 2, 2, 0),
        ):
            response = client.post(
                "/api/articles/news-summary",
                data={
                    "country_code": "italy",
                    "article_source": "db:test_run:after",
                    "article_selection_mode": "selected",
                },
            )

        self.assertEqual(response.status_code, 400)
        self.assertIn("至少选择", response.get_json()["error"])

    def test_user_template_is_utf8_without_bom_and_has_selection_controls(self) -> None:
        template_path = PROJECT_DIR / "templates" / "user_app.html"
        raw = template_path.read_bytes()
        self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
        text = raw.decode("utf-8")
        self.assertIn("普通用户模式", text)
        self.assertIn("添加新闻到当前批次", text)
        self.assertIn("全选", text)
        self.assertIn("取消全选", text)
        self.assertIn("selected_article_ids", text)
        self.assertIn("user-manual-details", text)
        self.assertIn("chip.sentiment-positive", text)
        self.assertIn('class="js-user-article-select"', text)
        for marker in ("鏅", "閫", "瀵", "鍥", "绛", "锛"):
            self.assertNotIn(marker, text)


if __name__ == "__main__":
    unittest.main()
