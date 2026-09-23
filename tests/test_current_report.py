import json
import os
import subprocess
import tempfile
import unittest
from collections import defaultdict
from pathlib import Path


ROOT = Path(__file__).parents[1]
WORKBOOK = Path(r"C:\Users\User\Downloads\Отчеты (3).xlsx")


class CurrentReportTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.data_path = Path(cls.temp_dir.name) / "deals.json"
        subprocess.run(
            ["python", "scripts/parse_report.py", str(WORKBOOK), "-o", str(cls.data_path)],
            cwd=ROOT,
            check=True,
            env={**os.environ, "PYTHONUTF8": "1"},
        )
        cls.payload = json.loads(cls.data_path.read_text(encoding="utf-8"))

    @classmethod
    def tearDownClass(cls):
        cls.temp_dir.cleanup()

    def test_monthly_totals_match_the_approved_report(self):
        expected = {
            "2026-05": (26, 95, 127_732, 8_152),
            "2026-06": (27, 85, 93_949, 5_417),
            "2026-07": (47, 154, 192_751, 12_901),
            "2026-08": (39, 127, 166_788, 13_750),
        }
        actual = defaultdict(lambda: [0, 0, 0, 0])
        for deal in self.payload["deals"]:
            row = actual[deal["period"]]
            row[0] += 1
            row[1] += deal["pax"]
            row[2] += deal["gross"]
            row[3] += deal["profit"] or 0
        self.assertEqual({key: tuple(value) for key, value in actual.items()}, expected)

    def test_overall_totals_and_marketing_are_current(self):
        deals = self.payload["deals"]
        self.assertEqual(len(deals), 139)
        self.assertEqual(sum(deal["pax"] for deal in deals), 461)
        self.assertEqual(sum(deal["gross"] for deal in deals), 581_220)
        self.assertEqual(sum((deal["profit"] or 0) for deal in deals), 40_220)
        self.assertEqual(self.payload["ad_spend"]["2026-08"], 1_277)

    def test_august_manager_payouts_use_the_reported_ad_deductions(self):
        deductions = self.payload["manager_ad_deductions"]["2026-08"]
        self.assertEqual(
            deductions,
            {"Сарвиноз": 159.6, "Муслимжон": 159.6, "Азиза": 159.6, "Шахноза": 159.6},
        )
        august_profit = defaultdict(float)
        for deal in self.payload["deals"]:
            if deal["period"] == "2026-08":
                august_profit[deal["manager"]] += deal["profit"] or 0
        payout = sum(value * self.payload["commission_rate"] for value in august_profit.values())
        payout -= sum(deductions.values())
        self.assertAlmostEqual(payout, 6_236.6)

    def test_dashboard_declares_utf8_before_visible_content(self):
        template = (ROOT / "prototype" / "dashboard.tpl.html").read_text(encoding="utf-8")
        self.assertTrue(template.lstrip().lower().startswith("<!doctype html>"))
        self.assertLess(template.lower().index('<meta charset="utf-8">'), template.index("<title>"))

    def test_average_check_is_sales_per_tourist(self):
        expected = {
            "2026-05": 1_345,
            "2026-06": 1_105,
            "2026-07": 1_252,
            "2026-08": 1_313,
        }
        for period, wanted in expected.items():
            deals = [deal for deal in self.payload["deals"] if deal["period"] == period]
            sales = sum(deal["gross"] for deal in deals)
            tourists = sum(deal["pax"] for deal in deals)
            self.assertEqual(round(sales / tourists), wanted)

    def test_dashboard_aggregation_divides_sales_by_tourists(self):
        template = (ROOT / "prototype" / "dashboard.tpl.html").read_text(encoding="utf-8")
        start = template.index("function agg(list)")
        end = template.index("\n// Доля расхода", start)
        function_source = template[start:end]
        script = function_source + "\nconsole.log(JSON.stringify(agg([{gross:300,profit:30,pax:3,client_debt:0,operator_debt:0,gaps:[]}])));"
        completed = subprocess.run(
            ["node", "-e", script], capture_output=True, text=True, check=True, encoding="utf-8"
        )
        result = json.loads(completed.stdout)
        self.assertEqual(result["avg"], 100)


if __name__ == "__main__":
    unittest.main()
