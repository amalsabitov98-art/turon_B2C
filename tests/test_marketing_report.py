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

    def test_missing_headers_report_all_logical_fields_including_spend_aliases(self):
        missing = {"Название кампании", "Показы", "Потраченная сумма (USD)"}
        headers = [header for header in HEADERS if header not in missing]

        with self.assertRaises(ValueError) as error:
            parse_workbook(self.workbook([campaign()], headers))

        message = str(error.exception)
        for header in (*missing, "Сумма затрат (USD)"):
            with self.subTest(header=header):
                self.assertIn(header, message)

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

    def test_partial_overlap_keeps_first_export_and_names_both_ranges(self):
        first = self.workbook([campaign("First")], name="first.xlsx")
        overlapping = campaign("Overlap", **{
            "Дата начала отчетности": "2026-07-26",
            "Окончание отчетности": "2026-08-01",
        })
        second = self.workbook([overlapping], name="overlap.xlsx")

        report = marketing.parse_workbooks([first, second])

        self.assertEqual([item["source"] for item in report["exports"]], ["first.xlsx"])
        self.assertTrue(any("2026-07-20_2026-07-26" in issue
                            and "2026-07-26_2026-08-01" in issue
                            and "overlap" in issue.lower() for issue in report["issues"]))

    def test_adjacent_ranges_are_both_accepted(self):
        first = self.workbook([campaign("First")], name="first.xlsx")
        adjacent = campaign("Adjacent", **{
            "Дата начала отчетности": "2026-07-27",
            "Окончание отчетности": "2026-08-02",
        })
        second = self.workbook([adjacent], name="adjacent.xlsx")

        report = marketing.parse_workbooks([first, second])

        self.assertEqual([item["source"] for item in report["exports"]],
                         ["first.xlsx", "adjacent.xlsx"])
        self.assertFalse(any("overlap" in issue.lower() for issue in report["issues"]))

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


class MarketingAggregationTest(unittest.TestCase):
    def model(self, expression):
        template = (Path(__file__).parents[1] / "prototype" / "dashboard.tpl.html").read_text(encoding="utf-8")
        start = template.find("/* ---------- marketing model ---------- */")
        end = template.find("/* ---------- подсказка ---------- */", start)
        block = template[start:end] if start >= 0 and end > start else ""
        script = block + "\nprocess.stdout.write(JSON.stringify(" + expression + "));"
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def test_aggregation_sums_weighted_metrics_and_keeps_result_categories_separate(self):
        rows = [
            {"spend": 10, "impressions": 1000, "link_clicks": 100,
             "results": 20, "result_category": "lead"},
            {"spend": 5, "impressions": 500, "link_clicks": 50,
             "results": 300, "result_category": "profile_visit"},
        ]
        result = self.model(f"aggregateMarketing({json.dumps(rows)})")
        for key, expected in {
            "spend": 15, "impressions": 1500, "cpm": 10,
            "linkClicks": 150, "cpc": 0.1, "linkCtr": 10,
            "leads": 20, "leadSpend": 10, "cpl": 0.5,
        }.items():
            self.assertAlmostEqual(result[key], expected, msg=key)
        self.assertEqual(result["resultCategories"]["profile_visit"],
                         {"results": 300, "spend": 5, "campaigns": 1})
        self.assertNotIn("totalResults", result)
        self.assertNotIn("results", result)

    def test_unknown_raw_result_types_have_independent_counts_and_spend(self):
        rows = [
            {"spend": 12, "results": 3, "result_category": "other",
             "result_type_raw": "actions:video_view"},
            {"spend": 8, "results": 2, "result_category": "other",
             "result_type_raw": "actions:video_view"},
            {"spend": 15, "results": 5, "result_category": "other",
             "result_type_raw": "actions:post_engagement"},
        ]

        groups = self.model(f"aggregateMarketing({json.dumps(rows)})")["resultCategories"]

        self.assertEqual(groups["other:actions:video_view"],
                         {"results": 5, "spend": 20, "campaigns": 2})
        self.assertEqual(groups["other:actions:post_engagement"],
                         {"results": 5, "spend": 15, "campaigns": 1})
        self.assertNotIn("other", groups)

    def test_reach_requires_a_single_export_level_total(self):
        rows = [{"spend": 3, "reach": 100, "result_category": "lead"},
                {"spend": 4, "reach": 80, "result_category": "call"}]
        source = json.dumps(rows)
        self.assertIsNone(self.model(f"aggregateMarketing({source})")["reach"])
        self.assertEqual(self.model(
            f"aggregateMarketing({source}, [{{totals: {{reach: 145}}}}])"
        )["reach"], 145)
        self.assertIsNone(self.model(
            f"aggregateMarketing({source}, [{{totals: {{reach: 145}}}}, {{totals: {{reach: 60}}}}])"
        )["reach"])

    def test_zero_denominators_produce_null_rates(self):
        result = self.model("aggregateMarketing([{spend: 10, impressions: 0, "
                            "link_clicks: 0, results: 0, result_category: 'lead'}])")
        self.assertEqual(result["spend"], 10)
        for key in ("cpm", "cpc", "linkCtr", "cpl"):
            self.assertIsNone(result[key], key)

    def test_scope_uses_exact_week_and_excludes_cross_month_exports_from_month(self):
        exports = [
            {"id": "jul", "start": "2026-07-20", "end": "2026-07-26", "month": "2026-07",
             "month_eligible": True, "campaigns": [{"spend": 10}, {"spend": 0}]},
            {"id": "bridge", "start": "2026-07-30", "end": "2026-08-02", "month": None,
             "month_eligible": False, "campaigns": [{"spend": 5}]},
        ]
        data = json.dumps(exports)
        self.assertEqual(self.model(f"marketingScope({data}, {{mode:'month', period:'2026-07', hideZero:true}})"),
                         {"exports": [exports[0]], "campaigns": exports[0]["campaigns"],
                          "filteredCampaigns": [{"spend": 10}]})
        self.assertEqual(self.model(f"marketingScope({data}, {{mode:'week', period:'bridge', hideZero:false}})"),
                         {"exports": [exports[1]], "campaigns": [{"spend": 5}],
                          "filteredCampaigns": [{"spend": 5}]})
        self.assertEqual(self.model(f"marketingPeriods({data}, 'month')"), ["2026-07"])
        self.assertEqual(self.model(f"marketingPeriods({data}, 'week')"), ["jul", "bridge"])

    def test_zero_spend_rows_remain_for_audit_but_not_aggregate_metrics(self):
        exports = [{"id": "week", "start": "2026-07-20", "campaigns": [
            {"spend": 10, "impressions": 100, "link_clicks": 10,
             "result_category": "lead", "results": 2},
            {"spend": 0, "impressions": 100, "link_clicks": 40,
             "result_category": "lead", "results": 3},
            {"spend": None, "impressions": 500, "link_clicks": 100,
             "result_category": "profile_visit", "results": 7},
        ]}]
        scoped = self.model("(() => { const s = marketingScope(" + json.dumps(exports) +
                            ", {mode:'week', period:'week', hideZero:true}); "
                            "return {metrics: aggregateMarketing(s.campaigns, s.exports), "
                            "audit: s.campaigns, shown: s.filteredCampaigns}; })()")
        self.assertEqual(len(scoped["audit"]), 3)
        self.assertEqual(len(scoped["shown"]), 1)
        metrics = scoped["metrics"]
        self.assertEqual(metrics["spend"], 10)
        self.assertEqual(metrics["impressions"], 100)
        self.assertEqual(metrics["linkClicks"], 10)
        self.assertEqual(metrics["leads"], 2)
        self.assertEqual(metrics["leadSpend"], 10)
        self.assertEqual(metrics["cpm"], 100)
        self.assertEqual(metrics["cpc"], 1)
        self.assertEqual(metrics["linkCtr"], 10)
        self.assertEqual(metrics["cpl"], 5)
        self.assertEqual(metrics["resultCategories"],
                         {"lead": {"results": 2, "spend": 10, "campaigns": 1}})

    def test_month_coverage_unions_days_and_marks_partial_july_august(self):
        exports = [
            {"start": "2026-07-20", "end": "2026-07-26", "month": "2026-07", "month_eligible": True},
            {"start": "2026-08-17", "end": "2026-08-23", "month": "2026-08", "month_eligible": True},
            {"start": "2026-07-24", "end": "2026-07-28", "month": "2026-07", "month_eligible": True},
        ]
        data = json.dumps(exports)
        self.assertEqual(self.model(f"monthCoverage({data}, '2026-07')"),
                         {"complete": False, "coveredDays": 9, "totalDays": 31})
        self.assertEqual(self.model(f"monthCoverage({data}, '2026-08')"),
                         {"complete": False, "coveredDays": 7, "totalDays": 31})
        full = [{"start": "2026-02-01", "end": "2026-02-14", "month": "2026-02", "month_eligible": True},
                {"start": "2026-02-14", "end": "2026-02-28", "month": "2026-02", "month_eligible": True}]
        self.assertEqual(self.model(f"monthCoverage({json.dumps(full)}, '2026-02')"),
                         {"complete": True, "coveredDays": 28, "totalDays": 28})

    def test_comparison_uses_previous_week_and_only_supported_metric_deltas(self):
        exports = [
            {"id": "first", "start": "2026-07-20", "end": "2026-07-26", "month": "2026-07",
             "month_eligible": True, "campaigns": [
                 {"spend": 10, "impressions": 1000, "link_clicks": 100, "results": 20, "result_category": "lead"},
                 {"spend": 5, "impressions": 500, "link_clicks": 50, "results": 300, "result_category": "profile_visit"}]},
            {"id": "second", "start": "2026-08-17", "end": "2026-08-23", "month": "2026-08",
             "month_eligible": True, "campaigns": [
                 {"spend": 20, "impressions": 2000, "link_clicks": 200, "results": 40, "result_category": "lead"},
                 {"spend": 10, "impressions": 1000, "link_clicks": 100, "results": 600, "result_category": "profile_visit"}]},
        ]
        data = json.dumps(exports)
        self.assertEqual(self.model(f"marketingComparison({data}, {{mode:'week', period:'first', hideZero:false}})"),
                         {key: None for key in ("spend", "impressions", "cpm", "linkClicks", "cpc",
                                                 "linkCtr", "leads", "leadSpend", "cpl")})
        self.assertEqual(self.model(f"marketingComparison({data}, {{mode:'week', period:'second', hideZero:false}})"),
                         {key: 100 for key in ("spend", "impressions", "leads", "leadSpend", "linkClicks")}
                         | {key: 0 for key in ("cpm", "cpc", "linkCtr", "cpl")})

    def test_comparison_returns_null_for_zero_previous_denominator(self):
        exports = [
            {"id": "first", "start": "2026-07-20", "end": "2026-07-26", "campaigns": [
                {"spend": 0, "impressions": 0, "link_clicks": 0, "results": 0, "result_category": "lead"}]},
            {"id": "second", "start": "2026-08-17", "end": "2026-08-23", "campaigns": [
                {"spend": 10, "impressions": 100, "link_clicks": 10, "results": 2, "result_category": "lead"}]},
        ]
        self.assertTrue(all(value is None for value in self.model(
            f"marketingComparison({json.dumps(exports)}, {{mode:'week', period:'second', hideZero:false}})"
        ).values()))


class MarketingTemplateTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.template = (Path(__file__).parents[1] / "prototype" / "dashboard.tpl.html").read_text(encoding="utf-8")

    def test_section_and_marketing_mounts_exist(self):
        for element_id in ("sectionSeg", "salesView", "marketingView", "marketingModeSeg",
                           "marketingPeriodSeg", "marketingKpis", "marketingResults",
                           "marketingCampaigns"):
            with self.subTest(element_id=element_id):
                self.assertTrue(f'id="{element_id}"' in self.template,
                                f'missing #{element_id}')
        render = self.template.split("function render() {", 1)[1].split("\n}", 1)[0]
        self.assertIn("renderSectionNavigation()", render)
        self.assertIn("renderMarketing()", render)

    def test_hidden_controls_stay_hidden_with_author_display_styles(self):
        self.assertIn('[hidden] { display:none!important; }', self.template)

    def marketing_ui(self, payload, state=None):
        model_start = self.template.index("/* ---------- marketing model ---------- */")
        model_end = self.template.index("/* ---------- подсказка ---------- */", model_start)
        ui_start = self.template.find("/* ---------- marketing interface ---------- */")
        ui_end = self.template.find("/* ---------- контекст ---------- */", ui_start)
        ui = self.template[ui_start:ui_end] if ui_start >= 0 and ui_end > ui_start else ""
        script = f"""
const DATA = {json.dumps({'marketing': payload}, ensure_ascii=False)};
const S = {json.dumps(state or {'role': 'rop', 'section': 'marketing', 'currency': 'USD', 'marketing': {'mode': 'week', 'period': 'week', 'hideZero': True}}, ensure_ascii=False)};
const elements = Object.fromEntries(['sectionControl','sectionSeg','salesView','marketingView',
  'perControl','currencyControl','rateBox','maskControl','marketingModeSeg',
  'marketingPeriodSeg','marketingKpis','marketingResults','marketingCampaigns',
  'marketingCoverage','marketingIssues','marketingHideZero'].map(id => [id,
    {{innerHTML:'', hidden:false, style:{{}}, checked:false, onclick:null, onchange:null}}]));
const document = {{getElementById:id => elements[id]}};
const esc = x => String(x == null ? '' : x).replace(/[&<>\"]/g, c =>
  ({{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}}[c]));
const money = v => v == null ? '—' : '$' + Number(v).toFixed(2);
const pct = v => v == null ? '—' : Number(v).toFixed(1) + '%';
const dmy = value => value;
const nf = new Intl.NumberFormat('ru-RU');
function seg(el, items, cur, cb) {{
  el.items = items; el.current = cur; el.choose = cb;
  el.innerHTML = items.map(i => `<button aria-pressed="${{i.v === cur}}">${{esc(i.t)}}</button>`).join('');
}}
{self.template[model_start:model_end]}
{ui}
renderSectionNavigation();
renderMarketing();
process.stdout.write(JSON.stringify({{state:S, elements}}));
"""
        result = subprocess.run(["node", "-e", script], capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(result.returncode, 0, result.stderr)
        return json.loads(result.stdout)

    def test_manager_cannot_retain_marketing_view(self):
        result = self.marketing_ui({'exports': [], 'issues': []},
                                   {'role': 'Manager', 'section': 'marketing', 'currency': 'USD',
                                    'marketing': {'mode': 'week', 'period': None, 'hideZero': True}})
        self.assertEqual(result['state']['section'], 'sales')
        self.assertTrue(result['elements']['sectionControl']['hidden'])
        self.assertTrue(result['elements']['marketingView']['hidden'])

    def test_marketing_keeps_currency_control_available(self):
        result = self.marketing_ui({'exports': [], 'issues': []})['elements']
        self.assertFalse(result['currencyControl']['hidden'])
        self.assertFalse(result['rateBox']['hidden'])

    def test_empty_marketing_and_issues_are_visible_separately(self):
        result = self.marketing_ui({'exports': [], 'issues': ['Bad Meta row']})['elements']
        self.assertIn('Нет рекламных выгрузок', result['marketingResults']['innerHTML'])
        self.assertIn('Bad Meta row', result['marketingIssues']['innerHTML'])

    def test_partial_month_and_separate_results_and_zero_spend_filter(self):
        rows = [
            {'name': 'Lead', 'status': 'active', 'spend': 10, 'impressions': 1000,
             'link_clicks': 20, 'results': 2, 'result_category': 'lead',
             'result_type_raw': 'actions:leadgen.other'},
            {'name': 'Call', 'status': 'paused', 'spend': 5, 'impressions': 500,
             'link_clicks': 10, 'results': 3, 'result_category': 'call',
             'result_type_raw': 'actions:click_to_call_native_call_placed'},
            {'name': 'Zero', 'status': 'paused', 'spend': 0, 'impressions': 300,
             'link_clicks': 1, 'results': 8, 'result_category': 'lead',
             'result_type_raw': 'actions:leadgen.other'},
        ]
        export = {'id': 'week', 'start': '2026-07-20', 'end': '2026-07-26',
                  'month': '2026-07', 'month_eligible': True, 'campaigns': rows}
        state = {'role': 'rop', 'section': 'marketing', 'currency': 'USD',
                 'marketing': {'mode': 'month', 'period': '2026-07', 'hideZero': True}}
        result = self.marketing_ui({'exports': [export], 'issues': []}, state)['elements']
        self.assertIn('Неполный месяц', result['marketingCoverage']['innerHTML'])
        self.assertIn('Лиды', result['marketingResults']['innerHTML'])
        self.assertIn('Звонки', result['marketingResults']['innerHTML'])
        self.assertNotIn('Посещения профиля', result['marketingResults']['innerHTML'])
        self.assertIn('$15,00', result['marketingKpis']['innerHTML'])
        self.assertIn('Lead', result['marketingCampaigns']['innerHTML'])
        self.assertIn('Call', result['marketingCampaigns']['innerHTML'])
        self.assertNotIn('Zero', result['marketingCampaigns']['innerHTML'])

    def test_unknown_result_types_render_as_separate_cards(self):
        rows = [
            {'name': 'Video', 'spend': 12, 'results': 3,
             'result_category': 'other', 'result_type_raw': 'actions:video_view'},
            {'name': 'Engagement', 'spend': 20, 'results': 4,
             'result_category': 'other', 'result_type_raw': 'actions:post_engagement'},
        ]
        export = {'id': 'week', 'start': '2026-07-20', 'end': '2026-07-26',
                  'month': '2026-07', 'month_eligible': True, 'campaigns': rows}

        html = self.marketing_ui({'exports': [export], 'issues': []})['elements']['marketingResults']['innerHTML']
        cards = html.split('<div class="card kpi">')[1:]

        self.assertEqual(len(cards), 2)
        self.assertTrue(any('actions:video_view' in card and '>3</div>' in card
                            and '$12,00' in card and '$4,00' in card for card in cards))
        self.assertTrue(any('actions:post_engagement' in card and '>4</div>' in card
                            and '$20,00' in card and '$5,00' in card for card in cards))

    def test_non_lead_results_show_count_spend_and_cost_or_dash(self):
        rows = [
            {'name': 'Call', 'spend': 15, 'results': 3,
             'result_category': 'call', 'result_type_raw': 'actions:click_to_call_native_call_placed'},
            {'name': 'Visit', 'spend': 6, 'results': 0,
             'result_category': 'profile_visit', 'result_type_raw': 'profile_visit_view'},
            {'name': 'Click', 'spend': 4, 'results': None,
             'result_category': 'link_click', 'result_type_raw': 'actions:link_click'},
        ]
        export = {'id': 'week', 'start': '2026-07-20', 'end': '2026-07-26',
                  'month': '2026-07', 'month_eligible': True, 'campaigns': rows}

        html = self.marketing_ui({'exports': [export], 'issues': []})['elements']['marketingResults']['innerHTML']
        cards = html.split('<div class="card kpi">')[1:]
        by_label = {card.split('<div class="k-l">', 1)[1].split('</div>', 1)[0]: card
                    for card in cards}

        self.assertIn('>3</div>', by_label['Звонки'])
        self.assertIn('$15,00', by_label['Звонки'])
        self.assertIn('Стоимость результата: $5,00', by_label['Звонки'])
        self.assertIn('>0</div>', by_label['Посещения профиля'])
        self.assertIn('$6,00', by_label['Посещения профиля'])
        self.assertIn('Стоимость результата: —', by_label['Посещения профиля'])
        self.assertIn('Стоимость результата: —', by_label['Клики по ссылке'])

    def test_comparison_context_is_visible_next_to_marketing_kpis(self):
        section = self.template.split('<section id="marketingView"', 1)[1].split('</section>', 1)[0]
        note = section.index('Состав и цели кампаний могут отличаться между периодами')
        self.assertLess(section.index('id="marketingCoverage"'), note)
        self.assertLess(note, section.index('id="marketingKpis"'))


class DashboardBuildTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.directory = Path(self.temp_dir.name)
        self.sales = self.directory / "deals.json"
        self.marketing = self.directory / "marketing.json"
        self.template = self.directory / "dashboard.tpl.html"
        self.output = self.directory / "index.html"
        self.template.write_text("const DATA = /*__DATA__*/null;\n", encoding="utf-8")
        self.sales_payload = {
            "deals": [{"gross": 300, "profit": 30, "pax": 3}],
            "ad_spend": {"2026-08": 1277},
        }
        self.sales.write_text(json.dumps(self.sales_payload), encoding="utf-8")

    def build(self):
        return subprocess.run(
            [sys.executable, str(Path(__file__).parents[1] / "scripts" / "build_dashboard.py"),
             "--sales", str(self.sales), "--marketing", str(self.marketing),
             "--template", str(self.template), "--out", str(self.output)],
            cwd=self.directory, capture_output=True, text=True,
        )

    def embedded_payload(self):
        html = self.output.read_text(encoding="utf-8")
        return json.loads(html.split("const DATA = ", 1)[1].split(";", 1)[0])

    def test_build_embeds_exact_marketing_object_without_changing_sales(self):
        marketing_payload = {
            "generated_at": "2026-09-25T08:30:00+00:00",
            "sources": ["july.xlsx"],
            "exports": [{
                "id": "2026-07-20_2026-07-26", "signature": "abc123",
                "start": "2026-07-20", "end": "2026-07-26", "month": "2026-07",
                "month_eligible": True, "source": "july.xlsx",
                "totals": {"spend": 162.05, "leads": 120},
                "campaigns": [{"name": "Lead campaign", "spend": 103.67}],
            }],
            "issues": ["Weekly extract only"],
        }
        self.marketing.write_text(json.dumps(marketing_payload), encoding="utf-8")

        result = self.build()

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            self.embedded_payload(), {**self.sales_payload, "marketing": marketing_payload}
        )

    def test_build_without_marketing_file_embeds_empty_marketing_data(self):
        result = self.build()

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(
            self.embedded_payload(),
            {**self.sales_payload, "marketing": {"exports": [], "issues": []}},
        )

    def test_build_escapes_closing_script_in_campaign_name(self):
        self.template = Path(__file__).parents[1] / "prototype" / "dashboard.tpl.html"
        campaign_name = "City </script><script>alert(1)</script> sale"
        marketing_payload = {"exports": [{"campaigns": [{"name": campaign_name}]}], "issues": []}
        self.marketing.write_text(json.dumps(marketing_payload), encoding="utf-8")

        result = self.build()

        self.assertEqual(result.returncode, 0, result.stderr)
        html = self.output.read_text(encoding="utf-8")
        serialized = html.split("const DATA = ", 1)[1].split(";", 1)[0]
        self.assertNotIn("</script>", serialized.lower())
        self.assertEqual(json.loads(serialized)["marketing"], marketing_payload)


if __name__ == "__main__":
    unittest.main()
