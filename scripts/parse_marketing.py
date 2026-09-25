"""Normalize Meta Ads campaign exports for the static dashboard."""

import argparse
import datetime as dt
import hashlib
import json
import re
import unicodedata
from pathlib import Path

import openpyxl


REQUIRED_HEADERS = (
    "Дата начала отчетности", "Окончание отчетности", "Название кампании",
    "Результат", "Индикатор результата",
    "Показы", "Клики по ссылке",
)

SPEND_HEADERS = ("Потраченная сумма (USD)", "Сумма затрат (USD)")

RESULT_CATEGORIES = {
    "actions:leadgen.other": "lead",
    "actions:click_to_call_native_call_placed": "call",
    "actions:link_click": "link_click",
    "actions:omni_landing_page_view": "landing_page_view",
    "profile_visit_view": "profile_visit",
}

NUMERIC_FIELDS = {
    "results": "Результат",
    "spend": "Потраченная сумма (USD)",
    "impressions": "Показы",
    "link_clicks": "Клики по ссылке",
    "reach": "Охват",
    "frequency": "Частота",
    "cost_per_result": "Цена за результаты",
    "cpm": "CPM (цена за 1 000 показов) (USD)",
    "shop_clicks": "shop_clicks",
    "cpc_link": "CPC (цена за клик по ссылке) (USD)",
    "ctr_link": "CTR (отношение кликов к показам)",
    "all_clicks": "Клики (все)",
    "ctr_all": "CTR (все)",
    "cpc_all": "CPC (все) (USD)",
    "landing_page_views": "Просмотры целевой страницы",
    "cost_per_landing_page_view": "Цена за просмотр целевой страницы (USD)",
    "initial_results": "Результаты (начальные)",
}

TEXT_FIELDS = {
    "status": "Показ кампании",
    "attribution": "Настройка атрибуции",
    "budget_type": "Тип бюджета группы объявлений",
    "ends_at": "Конец",
    "initial_result_type_raw": "Индикатор результатов (начальных)",
}


def normalize_header(value):
    text = unicodedata.normalize("NFKC", str(value or "")).replace("ё", "е")
    return re.sub(r"\s+", " ", text).strip().casefold()


def parse_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(text) if "." in text else int(text)
    except ValueError:
        return None


def date_iso(value):
    if isinstance(value, dt.datetime):
        return value.date().isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    try:
        return dt.date.fromisoformat(str(value).strip()).isoformat()
    except ValueError as error:
        raise ValueError(f"Invalid Meta report date: {value!r}") from error


def json_cell(value):
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    return value


def sum_available(values):
    present = [value for value in values if value is not None]
    return sum(present) if present else None


def safe_div(numerator, denominator):
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def content_signature(campaigns):
    normalized = [json.dumps(row, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                  for row in campaigns]
    canonical = json.dumps(sorted(normalized), ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def parse_workbook(path):
    book = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = book.active
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows)
        columns = {normalize_header(value): index for index, value in enumerate(headers)}
        missing = [header for header in REQUIRED_HEADERS
                   if normalize_header(header) not in columns]
        spend_column = next((columns[normalize_header(header)] for header in SPEND_HEADERS
                             if normalize_header(header) in columns), None)
        if spend_column is None:
            missing.append(" / ".join(SPEND_HEADERS))
        if missing:
            raise ValueError(f"Missing required Meta headers: {', '.join(missing)}")

        def cell(row, header):
            index = (spend_column if header == SPEND_HEADERS[0]
                     else columns.get(normalize_header(header)))
            return row[index] if index is not None and index < len(row) else None

        campaigns = []
        total_row = None
        period = None
        for row in rows:
            if not any(value is not None for value in row):
                continue
            current_period = (date_iso(cell(row, "Дата начала отчетности")),
                              date_iso(cell(row, "Окончание отчетности")))
            if period is None:
                period = current_period
            elif period != current_period:
                raise ValueError(f"Mixed Meta report periods in {path}: {period} and {current_period}")
            name = cell(row, "Название кампании")
            if not name or not str(name).strip():
                total_row = row
                continue
            raw = cell(row, "Индикатор результата")
            raw = str(raw).strip() if raw is not None else None
            item = {
                "name": str(name).strip(),
                "result_type_raw": raw,
                "result_category": RESULT_CATEGORIES.get(raw.casefold() if raw else "", "other"),
            }
            item.update({field: parse_number(cell(row, header))
                         for field, header in NUMERIC_FIELDS.items()})
            item.update({field: json_cell(cell(row, header))
                         for field, header in TEXT_FIELDS.items()})
            budget = cell(row, "Бюджет группы объявлений")
            item["budget"] = parse_number(budget)
            item["budget_raw"] = str(budget).strip() if budget is not None else None
            campaigns.append(item)
        if period is None:
            raise ValueError(f"No Meta report rows in {path}")

        start, end = period
        month_eligible = start[:7] == end[:7]
        totals = {}
        for field in ("spend", "impressions", "link_clicks", "all_clicks"):
            header = NUMERIC_FIELDS[field]
            stated = parse_number(cell(total_row, header)) if total_row else None
            totals[field] = stated if stated is not None else sum_available(
                campaign[field] for campaign in campaigns)
        totals["reach"] = parse_number(cell(total_row, "Охват")) if total_row else None
        for category, count_name, spend_name in (
            ("lead", "leads", "lead_spend"),
            ("call", "calls", "call_spend"),
            ("link_click", "link_click_results", "link_click_result_spend"),
            ("landing_page_view", "landing_page_view_results", "landing_page_view_spend"),
            ("profile_visit", "profile_visits", "profile_visit_spend"),
        ):
            matching = [item for item in campaigns if item["result_category"] == category]
            totals[count_name] = sum(item["results"] or 0 for item in matching)
            totals[spend_name] = sum(item["spend"] or 0 for item in matching)
        return {
            "id": f"{start}_{end}",
            "signature": content_signature(campaigns),
            "start": start,
            "end": end,
            "month": start[:7] if month_eligible else None,
            "month_eligible": month_eligible,
            "source": Path(path).name,
            "totals": totals,
            "campaigns": campaigns,
        }
    finally:
        book.close()


def parse_workbooks(paths: list[Path]) -> dict:
    exports = []
    seen = {}
    issues = []
    sources = []
    for path in paths:
        path = Path(path)
        sources.append(path.name)
        export = parse_workbook(path)
        existing = seen.get(export["id"])
        if existing:
            if existing["signature"] == export["signature"]:
                issues.append(f"Duplicate Meta export for {export['id']}: {path.name} matches {existing['source']}")
            else:
                issues.append(f"Conflict for Meta period {export['id']}: {path.name} differs from {existing['source']}; keeping first")
            continue
        overlapping = next((accepted for accepted in exports
                            if export["start"] <= accepted["end"]
                            and accepted["start"] <= export["end"]), None)
        if overlapping:
            issues.append(f"Overlapping Meta periods {export['id']} ({path.name}) and "
                          f"{overlapping['id']} ({overlapping['source']}); keeping first")
            continue
        seen[export["id"]] = export
        exports.append(export)
        if not export["month_eligible"]:
            issues.append(f"Cross-month Meta export {export['id']} ({path.name}) is available for weekly mode only")
    return {
        "generated_at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "sources": sources,
        "exports": exports,
        "issues": issues,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbooks", nargs="+", type=Path, help="Meta Ads .xlsx campaign exports")
    parser.add_argument("-o", "--out", type=Path, help="Write normalized JSON to this path")
    args = parser.parse_args(argv)
    payload = parse_workbooks(args.workbooks)
    encoded = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    if args.out:
        args.out.write_text(encoded, encoding="utf-8")
    else:
        print(encoded, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
