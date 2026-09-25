import datetime as dt
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts import parse_marketing as marketing
from scripts.parse_marketing import parse_workbook


JULY = Path(r"C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_20_июл_2026_г_26_июл_2026.xlsx")
AUGUST = Path(r"C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_17_авг_2026_г_23_авг_2026.xlsx")

HEADERS = [
    "Дата начала отчетности", "Окончание отчетности", "Название кампании",
    "Показ кампании", "Настройка атрибуции", "Результат",
    "Индикатор результата", "Охват", "Частота", "Цена за результаты",
    "Бюджет группы объявлений", "Тип бюджета группы объявлений",
    "Потраченная сумма (USD)", "Конец", "Показы",
    "CPM (цена за 1\u00a0000 показов) (USD)", "Клики по ссылке", "shop_clicks",
    "CPC (цена за клик по ссылке) (USD)",
    "CTR (отношение кликов к показам)", "Клики (все)", "CTR (все)",
    "CPC (все) (USD)", "Просмотры целевой страницы",
    "Цена за просмотр целевой страницы (USD)", "Результаты (начальные)",
    "Индикатор результатов (начальных)",
]


def campaign(name="Campaign", **values):
    data = {
        "Дата начала отчетности": "2026-07-20",
        "Окончание отчетности": "2026-07-26",
        "Название кампании": name,
        "Результат": 3,
        "Индикатор результата": "actions:leadgen.other",
        "Потраченная сумма (USD)": 10.5,
        "Показы": 100,
        "Клики по ссылке": 8,
    }
    data.update(values)
    return data


class MarketingReportTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)

    def workbook(self, rows, headers=HEADERS, name="meta.xlsx"):
        path = Path(self.temp_dir.name) / name
        book = Workbook()
        sheet = book.active
        sheet.append(headers)
        for row in rows:
            sheet.append([row.get(header) for header in headers])
        book.save(path)
        return path

    def test_result_type_mapping_retains_raw_values(self):
        cases = {
            "actions:leadgen.other": "lead",
            "actions:click_to_call_native_call_placed": "call",
            "actions:link_click": "link_click",
            "actions:omni_landing_page_view": "landing_page_view",
            "profile_visit_view": "profile_visit",
            "actions:unknown_future_metric": "other",
        }
        rows = [campaign(f"Campaign {i}", **{
            "Индикатор результата": raw
        }) for i, raw in enumerate(cases)]
        export = parse_workbook(self.workbook(rows))
        self.assertEqual(
            [(row["result_type_raw"], row["result_category"]) for row in export["campaigns"]],
            list(cases.items()),
        )

    def test_missing_required_spend_header_names_both_supported_labels(self):
        headers = [header for header in HEADERS if header != "Потраченная сумма (USD)"]
        with self.assertRaises(ValueError) as error:
            parse_workbook(self.workbook([campaign()], headers))
        self.assertIn("Потраченная сумма (USD)", str(error.exception))
        self.assertIn("Сумма затрат (USD)", str(error.exception))

    def test_summa_zatrat_spend_alias_parses(self):
        headers = ["Сумма затрат (USD)" if header == "Потраченная сумма (USD)" else header
                   for header in HEADERS]
        row = campaign("Alias", **{"Потраченная сумма (USD)": "12,75"})
        row["Сумма затрат (USD)"] = row.pop("Потраченная сумма (USD)")
        export = parse_workbook(self.workbook([row], headers))
        self.assertEqual(export["campaigns"][0]["spend"], 12.75)
        self.assertEqual(export["totals"]["spend"], 12.75)

    def test_normalized_headers_numeric_cells_and_optional_meta_fields(self):
        headers = list(reversed(HEADERS))
        headers[headers.index("Потраченная сумма (USD)")] = "  ПОТРАЧЕННАЯ   СУММА (USD) "
        row = campaign(
            "  July lead  ",
            **{
                "Потраченная сумма (USD)": "1 234,50",
                "Показы": "2\u00a0000",
                "Клики по ссылке": None,
                "Охват": 1500,
                "Частота": 1.333,
                "Просмотры целевой страницы": 7,
            },
        )
        renamed = {("  ПОТРАЧЕННАЯ   СУММА (USD) " if key == "Потраченная сумма (USD)" else key): value for key, value in row.items()}
        export = parse_workbook(self.workbook([renamed], headers))
        item = export["campaigns"][0]
        self.assertEqual((export["start"], export["end"], export["month"]),
                         ("2026-07-20", "2026-07-26", "2026-07"))
        self.assertEqual(item["name"], "July lead")
        self.assertEqual(item["spend"], 1234.5)
        self.assertEqual(item["impressions"], 2000)
        self.assertIsNone(item["link_clicks"])
        self.assertEqual(item["reach"], 1500)
        self.assertEqual(item["frequency"], 1.333)
        self.assertEqual(item["landing_page_views"], 7)

    def test_optional_excel_date_is_json_serializable(self):
        row = campaign("Timed", **{"Конец": dt.datetime(2026, 7, 26)})
        export = parse_workbook(self.workbook([row]))
        self.assertEqual(export["campaigns"][0]["ends_at"], "2026-07-26T00:00:00")
        json.dumps(export)

    def test_blank_campaign_row_is_totals_and_zero_spend_campaign_is_retained(self):
        total = campaign(None, **{
            "Результат": None, "Индикатор результата": None,
            "Потраченная сумма (USD)": 12, "Показы": 400,
            "Клики по ссылке": 9, "Охват": 300,
        })
        zero = campaign("Paused campaign", **{
            "Результат": 0, "Потраченная сумма (USD)": 0,
            "Показы": 0, "Клики по ссылке": 0,
        })
        paid = campaign("Paid campaign", **{
            "Результат": 2, "Потраченная сумма (USD)": 12,
            "Показы": 400, "Клики по ссылке": 9,
        })
        export = parse_workbook(self.workbook([total, zero, paid]))
        self.assertEqual([item["name"] for item in export["campaigns"]],
                         ["Paused campaign", "Paid campaign"])
        self.assertEqual(export["campaigns"][0]["spend"], 0)
        self.assertEqual(export["totals"]["spend"], 12)
        self.assertEqual(export["totals"]["impressions"], 400)
        self.assertEqual(export["totals"]["link_clicks"], 9)
        self.assertEqual(export["totals"]["reach"], 300)
        self.assertEqual(export["totals"]["leads"], 2)
        self.assertEqual(export["totals"]["lead_spend"], 12)

    def test_no_total_row_does_not_sum_campaign_reach(self):
        first = campaign("A", **{"Охват": 100, "Результат": 1})
        second = campaign("B", **{"Охват": 80, "Результат": 2})
        export = parse_workbook(self.workbook([first, second]))
        self.assertIsNone(export["totals"]["reach"])
        self.assertEqual(export["totals"]["leads"], 3)

    def test_safe_div_returns_none_for_zero_or_missing_denominator(self):
        self.assertIsNone(marketing.safe_div(10, 0))
        self.assertIsNone(marketing.safe_div(10, None))
        self.assertEqual(marketing.safe_div(10, 2), 5)

    def test_duplicate_period_with_same_sorted_campaigns_is_deduplicated(self):
        first = campaign("A", **{"Результат": 1})
        second = campaign("B", **{"Результат": 2})
        a = self.workbook([first, second], name="first.xlsx")
        b = self.workbook([second, first], name="second.xlsx")
        report = marketing.parse_workbooks([a, b])
        self.assertEqual(len(report["exports"]), 1)
        self.assertEqual(report["exports"][0]["source"], "first.xlsx")
        self.assertEqual(report["sources"], ["first.xlsx", "second.xlsx"])
        self.assertTrue(any("duplicate" in issue.lower() for issue in report["issues"]))

    def test_conflicting_period_keeps_first_export_and_reports_conflict(self):
        a = self.workbook([campaign("A")], name="first.xlsx")
        b = self.workbook([campaign("B")], name="second.xlsx")
        report = marketing.parse_workbooks([a, b])
        self.assertEqual(len(report["exports"]), 1)
        self.assertEqual(report["exports"][0]["campaigns"][0]["name"], "A")
        self.assertTrue(any("conflict" in issue.lower() for issue in report["issues"]))

    def test_cross_month_export_remains_for_weekly_view(self):
        row = campaign("Bridge", **{
            "Дата начала отчетности": "2026-07-30",
            "Окончание отчетности": "2026-08-02",
        })
        report = marketing.parse_workbooks([self.workbook([row])])
        self.assertEqual(len(report["exports"]), 1)
        export = report["exports"][0]
        self.assertEqual(export["id"], "2026-07-30_2026-08-02")
        self.assertIsNone(export["month"])
        self.assertFalse(export["month_eligible"])
        self.assertTrue(any("cross-month" in issue.lower() for issue in report["issues"]))

    def test_cli_writes_json_for_multiple_workbooks(self):
        first = self.workbook([campaign("July")], name="july.xlsx")
        august = campaign("August", **{
            "Дата начала отчетности": "2026-08-17",
            "Окончание отчетности": "2026-08-23",
        })
        second = self.workbook([august], name="august.xlsx")
        output = Path(self.temp_dir.name) / "marketing.json"
        result = subprocess.run(
            [sys.executable, str(Path(__file__).parents[1] / "scripts" / "parse_marketing.py"),
             str(first), str(second), "-o", str(output)],
            capture_output=True, text=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        data = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(len(data["exports"]), 2)
        self.assertEqual(data["sources"], ["july.xlsx", "august.xlsx"])
        self.assertIn("generated_at", data)


class RealMetaWorkbookTest(unittest.TestCase):
    @unittest.skipUnless(JULY.exists(), "July Meta workbook not available")
    def test_july_weekly_extract(self):
        export = parse_workbook(JULY)
        totals = export["totals"]
        self.assertAlmostEqual(totals["spend"], 162.05)
        self.assertEqual(totals["impressions"], 73683)
        self.assertEqual(totals["link_clicks"], 3423)
        self.assertEqual(totals["leads"], 120)
        self.assertAlmostEqual(totals["lead_spend"], 103.67)
        self.assertNotIn("results", totals)

    @unittest.skipUnless(AUGUST.exists(), "August Meta workbook not available")
    def test_august_weekly_extract(self):
        export = parse_workbook(AUGUST)
        totals = export["totals"]
        self.assertAlmostEqual(totals["spend"], 175.63)
        self.assertEqual(totals["impressions"], 261445)
        self.assertEqual(totals["link_clicks"], 6982)
        self.assertEqual(totals["leads"], 35)
        self.assertAlmostEqual(totals["lead_spend"], 21.28)
        self.assertNotIn("results", totals)


if __name__ == "__main__":
    unittest.main()
