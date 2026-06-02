import sys
import tempfile
import unittest
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

import news_summary  # noqa: E402


class NewsSummaryTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmpdir.cleanup)
        self._original_output_dir = news_summary.NEWS_SUMMARY_OUTPUT_DIR
        news_summary.NEWS_SUMMARY_OUTPUT_DIR = Path(self._tmpdir.name)
        self.addCleanup(lambda: setattr(news_summary, 'NEWS_SUMMARY_OUTPUT_DIR', self._original_output_dir))

    def test_ai_summary_is_sorted_and_saved_as_txt_with_url_reference(self) -> None:
        rows = [
            {
                'article_id': 'platform-1',
                'platform_label': 'Amazon',
                'title_display_zh': 'Amazon 推出新的现金支付点',
                'published_at': '2026-05-18T16:44:40+00:00',
                'source_name': 'Test Media',
                'article_url': 'https://example.com/amazon-cash',
                'survey_dimensions': 'Features',
                'briefing_sentiment': 'Positive',
            },
            {
                'article_id': 'industry-1',
                'platform_label': 'TEMU',
                'title_display_zh': '欧盟推进电商平台监管',
                'published_at': '2026-05-20',
                'source_name': 'Policy Daily',
                'article_url': 'https://example.com/eu-regulation',
                'survey_dimensions': 'Quality | Customer / post-purchase service',
                'briefing_sentiment': 'Neutral',
                'industry_trend_flag': 'true',
            },
        ]

        def fake_call_api(messages, api_url, api_key, api_model):
            user_payload = messages[1]['content']
            self.assertIn('"article_url": "https://example.com/amazon-cash"', user_payload)
            return {
                'items': [
                    {
                        'article_id': 'platform-1',
                        'platform_label': 'Amazon',
                        'core_claim': '意大利现金支付便利性提升',
                        'detail': '2026年5月18日，意大利/Test Media消息，Amazon 扩展现金支付点，可能提升支付便利性与下单转化。',
                        'sentiment': '正向',
                        'primary_metric': 'Features',
                        'tags': ['支付手段', '功能'],
                    },
                    {
                        'article_id': 'industry-1',
                        'platform_label': '行业',
                        'core_claim': '欧盟监管强化平台合规压力',
                        'detail': '2026年5月20日，意大利/Policy Daily消息，欧盟推进电商平台监管，可能影响 Temu 等平台合规成本与消费者信任。',
                        'sentiment': '中性',
                        'primary_metric': 'Quality',
                        'tags': ['政策监管', '平台合规'],
                    },
                ]
            }

        original_call_api = news_summary.xlsx_source_test.call_survey_filter_api
        news_summary.xlsx_source_test.call_survey_filter_api = fake_call_api
        self.addCleanup(lambda: setattr(news_summary.xlsx_source_test, 'call_survey_filter_api', original_call_api))

        stats: dict[str, int] = {}
        result = news_summary.generate_news_summary(
            rows,
            country_code='italy',
            api_settings={
                'survey_api_url': 'https://example.test/v1/chat/completions',
                'survey_api_key': 'key',
                'survey_api_model': 'model',
            },
            stats=stats,
        )

        lines = [line for line in result.text.splitlines() if line.strip()]
        self.assertTrue(lines[0].startswith('【行业】'))
        self.assertIn('2026年5月20日', lines[0])
        self.assertIn('https://example.com/eu-regulation 【中性】【Quality】【政策监管】【平台合规】', lines[0])
        self.assertNotIn('媒体【', lines[0])
        self.assertTrue(lines[1].startswith('【Amazon】'))
        self.assertIn('https://example.com/amazon-cash 【正向】【Features】【支付手段】【功能】', lines[1])
        self.assertNotIn('媒体【', lines[1])
        self.assertEqual(result.output_path.read_text(encoding='utf-8-sig'), result.text)
        self.assertEqual(stats['failed_row_count'], 0)
        self.assertEqual(stats['ai_completed_row_count'], 2)

    def test_failed_ai_batch_uses_fallback_summary_with_source_reference(self) -> None:
        rows = [
            {
                'article_id': 'row-1',
                'platform_label': 'IG',
                'title_display_zh': 'Instagram 增加购物广告工具',
                'published_at': '2026-05-18',
                'source_name': 'Media',
                'survey_dimensions': 'Content',
                'briefing_sentiment': 'Neutral',
            }
        ]

        def failing_call_api(messages, api_url, api_key, api_model):
            raise RuntimeError('mock ai failed')

        original_call_api = news_summary.xlsx_source_test.call_survey_filter_api
        news_summary.xlsx_source_test.call_survey_filter_api = failing_call_api
        self.addCleanup(lambda: setattr(news_summary.xlsx_source_test, 'call_survey_filter_api', original_call_api))

        stats: dict[str, int] = {}
        result = news_summary.generate_news_summary(
            rows,
            country_code='italy',
            api_settings={
                'survey_api_url': 'https://example.test/v1/chat/completions',
                'survey_api_key': 'key',
                'survey_api_model': 'model',
            },
            stats=stats,
        )

        self.assertIn('【IG】', result.text)
        self.assertIn('2026年5月18日', result.text)
        self.assertIn('来源：Media 【中性】【Content】', result.text)
        self.assertNotIn('媒体【', result.text)
        self.assertEqual(stats['failed_batch_count'], 1)
        self.assertEqual(stats['failed_row_count'], 1)
        self.assertEqual(stats['fallback_row_count'], 1)

    def test_summary_without_url_or_source_omits_media_placeholder(self) -> None:
        row = {
            'article_id': 'row-2',
            'platform_label': 'SHEIN',
            'title_display_zh': 'SHEIN 产品合规受到关注',
            'published_at': '2026-05-19',
            'survey_dimensions': 'Quality',
            'briefing_sentiment': 'Negative',
        }
        item = news_summary.default_summary_item(row, 'italy')
        line = news_summary.format_summary_line(item, row, 'italy')

        self.assertIn('【SHEIN】', line)
        self.assertIn('【负向】【Quality】', line)
        self.assertNotIn('媒体【', line)
        self.assertNotIn('来源：', line)

    def test_core_claim_is_limited_to_50_chars(self) -> None:
        long_text = '这是一个非常长的新闻核心判断句用于测试摘要标题不会超过五十个汉字并且会自动截断保留可读性'
        self.assertLessEqual(len(news_summary.limit_core_claim(long_text)), 51)


if __name__ == '__main__':
    unittest.main()
