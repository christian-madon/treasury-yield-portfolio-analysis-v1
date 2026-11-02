"""Compare Treasury yields and reprice fixed October 27 cash flows in two yield scenarios."""

import argparse
import csv
import math
import os
import pprint
import subprocess
import sys
import tempfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


BEFORE_DATE = date(2026, 10, 27)
AFTER_DATE = date(2026, 10, 28)
VALUATION_DATE = date(2026, 10, 27)
SCENARIO_YIELD_OBSERVATION_DATE = date(2026, 10, 28)
MATURITY_COLUMNS = ["2 Yr", "5 Yr", "10 Yr", "30 Yr"]
INPUT_PATH = Path(__file__).resolve().parent / "data" / "raw" / "treasury_yields_2026.csv"
PORTFOLIO_PATH = Path(__file__).resolve().parent / "data" / "raw" / "synthetic_portfolio.xlsx"


def read_yields(input_path):
    """Find the requested dates, check the input, and return their percent yields."""
    matches = {BEFORE_DATE: [], AFTER_DATE: []}

    # Input: DictReader uses the CSV's actual column headings as dictionary keys.
    with input_path.open("r", encoding="utf-8", newline="") as source:
        reader = csv.DictReader(source, strict=True)
        headers = reader.fieldnames or []
        for header in ["Date"] + MATURITY_COLUMNS:
            count = headers.count(header)
            if count != 1:
                raise ValueError(f"Required header {header!r} must appear once; found {count}.")

        # Date selection: parse each date; row order does not determine before/after.
        for row in reader:
            if None in row or None in row.values():
                raise ValueError(f"CSV line {reader.line_num}: field count differs from headers.")
            try:
                observed_date = datetime.strptime(row["Date"], "%m/%d/%Y").date()
            except ValueError:
                raise ValueError(
                    f"CSV line {reader.line_num}: invalid Date {row['Date']!r}; expected MM/DD/YYYY."
                ) from None
            if observed_date in matches:
                matches[observed_date].append(row)

    # Numeric conversion: Decimal keeps the input decimals without rounding them.
    observations = {}
    for requested_date in (BEFORE_DATE, AFTER_DATE):
        rows = matches[requested_date]
        if len(rows) != 1:
            raise ValueError(
                f"Expected exactly one record for {requested_date}; found {len(rows)}."
            )
        observations[requested_date] = {}
        for maturity in MATURITY_COLUMNS:
            raw_yield = rows[0][maturity]
            try:
                yield_pct = Decimal(raw_yield)
            except InvalidOperation:
                raise ValueError(
                    f"{requested_date}, {maturity}: yield {raw_yield!r} is not numeric."
                ) from None
            if not yield_pct.is_finite():
                raise ValueError(f"{requested_date}, {maturity}: yield must be finite.")
            observations[requested_date][maturity] = yield_pct
    return observations


def print_comparison(observations):
    """Calculate signed basis-point changes and format the completed comparison."""
    # Display: dates and headers make the yield and change units explicit.
    print(f"Treasury par yields: {BEFORE_DATE} (before) to {AFTER_DATE} (after)")
    print(f"{'Maturity':<10} {'Before yield (%)':>18} {'After yield (%)':>18} {'Change (bp)':>14}")
    print("-" * 63)
    for maturity in MATURITY_COLUMNS:
        before_yield_pct = observations[BEFORE_DATE][maturity]
        after_yield_pct = observations[AFTER_DATE][maturity]
        # Calculation: one percentage point equals 100 basis points.
        change_bp = (after_yield_pct - before_yield_pct) * 100
        print(
            f"{maturity:<10} {before_yield_pct:>18.2f} "
            f"{after_yield_pct:>18.2f} {change_bp:>+14.2f}"
        )


def load_portfolio(input_path):
    """Keep each populated Holdings row as a dictionary of its original values."""
    expected_headers = (
        "bond_id", "face_value_usd", "coupon_rate_pct", "issue_date",
        "maturity_date", "as_of_date", "currency", "coupon_payments_per_year",
    )
    numeric_headers = ("face_value_usd", "coupon_rate_pct", "coupon_payments_per_year")
    date_headers = ("issue_date", "maturity_date", "as_of_date")
    portfolio = []
    bond_ids = set()
    workbook = load_workbook(input_path, read_only=True, data_only=False)
    try:
        if "Holdings" not in workbook.sheetnames:
            raise ValueError("Workbook is missing the Holdings worksheet.")
        worksheet = workbook["Holdings"]
        # Read actual cells instead of truncating to cached worksheet dimensions.
        worksheet.reset_dimensions()
        rows = worksheet.iter_rows(values_only=False)
        header_cells = next(rows, ())
        if len(header_cells) != len(expected_headers):
            column = min(len(header_cells), len(expected_headers)) + 1
            raise ValueError(
                f"Holdings row 1, column {column}: expected exactly eight headings; "
                f"found {len(header_cells)}."
            )
        headers = []
        for column, cell in enumerate(header_cells, start=1):
            header = cell.value
            location = f"Holdings row 1, column {column}"
            if cell.data_type in ("f", "e"):
                raise ValueError(f"{location}: formula or Excel error in heading.")
            if not isinstance(header, str) or not header.strip():
                raise ValueError(f"{location}: heading must be nonblank text.")
            if header not in expected_headers or header in headers:
                raise ValueError(f"{location}: unexpected or duplicate heading {header!r}.")
            headers.append(header)

        for row_number, cells in enumerate(rows, start=2):
            if all(cell.value is None for cell in cells):
                continue
            if len(cells) > len(headers):
                raise ValueError(
                    f"Holdings row {row_number}, column {len(headers) + 1}: "
                    "unexpected cell outside the eight headings."
                )
            holding = {}
            for column, header in enumerate(headers, start=1):
                location = f"Holdings row {row_number}, column {column} ({header})"
                if column > len(cells):
                    raise ValueError(f"{location}: missing cell.")
                cell = cells[column - 1]
                value = cell.value
                if cell.data_type in ("f", "e"):
                    raise ValueError(f"{location}: formula or Excel error is not allowed.")
                if value is None or value == "":
                    raise ValueError(f"{location}: missing value.")
                if header in numeric_headers:
                    if (isinstance(value, bool) or not isinstance(value, (int, float))
                            or not math.isfinite(value)):
                        raise ValueError(f"{location}: expected a finite number.")
                elif header in date_headers:
                    if not isinstance(value, (date, datetime)):
                        raise ValueError(f"{location}: expected an Excel date/datetime.")
                elif not isinstance(value, str) or not value.strip():
                    raise ValueError(f"{location}: expected nonblank text.")
                # Assign in worksheet order; do not rescale coupons or convert dates.
                holding[header] = value
            bond_id = holding["bond_id"]
            if bond_id in bond_ids:
                column = headers.index("bond_id") + 1
                raise ValueError(
                    f"Holdings row {row_number}, column {column} (bond_id): "
                    f"duplicate identifier {bond_id!r}."
                )
            bond_ids.add(bond_id)
            portfolio.append(holding)
        if len(portfolio) != 3:
            raise ValueError(f"Holdings: expected three populated holdings; found {len(portfolio)}.")
        return portfolio
    finally:
        workbook.close()


def calendar_date(value, label):
    """Use a date's calendar component without changing the original object."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise ValueError(f"{label}: expected a date/datetime; found {value!r}.")


def finite_decimal(value, label):
    """Convert original numeric text representation to a finite Decimal."""
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ValueError(f"{label}: expected a finite numeric value; found {value!r}.")
    number = Decimal(str(value))
    if not number.is_finite():
        raise ValueError(f"{label}: value must be finite.")
    return number


def prepare_valuation_portfolio(portfolio, observations):
    """Copy holdings and add inputs for the supported regular semiannual model."""
    benchmark_by_years = {2: "2 Yr", 5: "5 Yr", 10: "10 Yr"}
    if VALUATION_DATE not in observations:
        raise ValueError(f"Missing Treasury observation for {VALUATION_DATE}.")
    valuation_yields = observations[VALUATION_DATE]
    valuation_portfolio = []
    for holding in portfolio:
        bond = holding.copy()
        bond_id = bond.get("bond_id")
        if not isinstance(bond_id, str) or not bond_id.strip():
            raise ValueError("Valuation holding is missing a valid bond_id.")
        if bond.get("currency") != "USD":
            raise ValueError(f"{bond_id}: valuation supports only USD bonds.")
        frequency = finite_decimal(
            bond.get("coupon_payments_per_year"), f"{bond_id} coupon_payments_per_year"
        )
        if frequency != Decimal("2"):
            raise ValueError(f"{bond_id}: valuation requires semiannual frequency 2.")
        for field in ("issue_date", "as_of_date", "maturity_date"):
            bond[field] = calendar_date(bond.get(field), f"{bond_id} {field}")
        for field in ("issue_date", "as_of_date"):
            if bond[field] != VALUATION_DATE:
                raise ValueError(f"{bond_id} {field}: must equal {VALUATION_DATE}.")
        maturity = bond["maturity_date"]
        term_years = maturity.year - VALUATION_DATE.year
        if ((maturity.month, maturity.day) != (VALUATION_DATE.month, VALUATION_DATE.day)
                or term_years not in benchmark_by_years):
            raise ValueError(
                f"{bond_id} maturity_date: require the same month/day exactly "
                "2, 5 or 10 years after the valuation date."
            )
        principal = finite_decimal(bond.get("face_value_usd"), f"{bond_id} face_value_usd")
        if principal <= Decimal("0"):
            raise ValueError(f"{bond_id} face_value_usd: principal must be positive.")
        coupon_rate_pct = finite_decimal(
            bond.get("coupon_rate_pct"), f"{bond_id} coupon_rate_pct"
        )
        benchmark = benchmark_by_years[term_years]
        annual_yield_pct = finite_decimal(
            valuation_yields.get(benchmark), f"{bond_id} {VALUATION_DATE} {benchmark} yield"
        )
        semiannual_discount_rate = annual_yield_pct / Decimal("100") / Decimal("2")
        if Decimal("1") + semiannual_discount_rate <= Decimal("0"):
            raise ValueError(f"{bond_id}: 1 + semiannual discount rate must be positive.")

        bond["valuation_date"] = VALUATION_DATE
        bond["semiannual_coupon_rate_decimal"] = coupon_rate_pct / Decimal("100") / Decimal("2")
        bond["coupon_payment_usd"] = principal * bond["semiannual_coupon_rate_decimal"]
        # Exact supported calendar years contain two complete coupon periods each.
        bond["number_of_payments"] = term_years * int(frequency)
        bond["maturity_repayment_usd"] = principal
        bond["benchmark_maturity"] = benchmark
        bond["annual_treasury_yield_pct"] = annual_yield_pct
        bond["semiannual_discount_rate_decimal"] = semiannual_discount_rate
        valuation_portfolio.append(bond)
    return valuation_portfolio


def prepare_yield_scenario(valuation_portfolio, observations, yield_observation_date):
    """Copy baseline bonds and change only the yield inputs and their date metadata."""
    if yield_observation_date not in observations:
        raise ValueError(f"Missing Treasury observation for {yield_observation_date}.")
    scenario_yields = observations[yield_observation_date]
    scenario_portfolio = []
    for baseline_bond in valuation_portfolio:
        bond = baseline_bond.copy()
        benchmark = bond["benchmark_maturity"]
        if benchmark not in ("2 Yr", "5 Yr", "10 Yr") or benchmark not in scenario_yields:
            raise ValueError(
                f"{bond['bond_id']}: missing or unsupported {benchmark!r} benchmark "
                f"for {yield_observation_date}."
            )
        annual_yield_pct = finite_decimal(
            scenario_yields[benchmark], f"{bond['bond_id']} {yield_observation_date} {benchmark} yield"
        )
        periodic_rate = annual_yield_pct / Decimal("100") / Decimal("2")
        if Decimal("1") + periodic_rate <= Decimal("0"):
            raise ValueError(f"{bond['bond_id']}: 1 + scenario discount rate must be positive.")
        bond["yield_observation_date"] = yield_observation_date
        bond["annual_treasury_yield_pct"] = annual_yield_pct
        bond["semiannual_discount_rate_decimal"] = periodic_rate
        scenario_portfolio.append(bond)
    return scenario_portfolio


def present_value(bond):
    """Return the USD PV of ONE validated, enriched bond dictionary; do not print."""
    C = bond["coupon_payment_usd"]
    N = bond["number_of_payments"]
    F = bond["maturity_repayment_usd"]
    r = bond["semiannual_discount_rate_decimal"]
    pv = Decimal("0")
    for k in range(1, N + 1):
        pv += C / (Decimal("1") + r) ** k
    pv += F / (Decimal("1") + r) ** N
    return pv


def calculate_valuation_results(enriched_portfolio):
    """Call the same PV function once per bond and retain its unrounded result."""
    results = []
    for bond in enriched_portfolio:
        pv_usd = present_value(bond)
        results.append({
            "bond_id": bond["bond_id"],
            "valuation_date": bond["valuation_date"],
            "coupon_payment_usd": bond["coupon_payment_usd"],
            "number_of_payments": bond["number_of_payments"],
            "benchmark_maturity": bond["benchmark_maturity"],
            "coupon_rate_pct": finite_decimal(bond["coupon_rate_pct"], "Annual coupon rate"),
            "annual_treasury_yield_pct": bond["annual_treasury_yield_pct"],
            "semiannual_discount_rate_decimal": bond["semiannual_discount_rate_decimal"],
            "present_value_usd": pv_usd,
        })
    return results


def calculate_valuation_change(before_pv_usd, after_pv_usd):
    """Return signed USD and percent changes from finite, unrounded PVs."""
    before_pv_usd = finite_decimal(before_pv_usd, "Starting present value")
    after_pv_usd = finite_decimal(after_pv_usd, "Ending present value")
    if before_pv_usd <= Decimal("0"):
        raise ValueError("Starting present value must be strictly positive.")
    dollar_change_usd = after_pv_usd - before_pv_usd
    return {
        "dollar_change_usd": dollar_change_usd,
        "percentage_change_pct": dollar_change_usd / before_pv_usd * Decimal("100"),
    }


def compare_valuation_results(
    before_results, after_results, before_yield_date, after_yield_date, valuation_anchor
):
    """Match existing PVs by ID and return new bond/portfolio change records."""
    if not isinstance(valuation_anchor, (date, datetime)):
        raise ValueError("Valuation anchor must be a date/datetime.")
    indexed_results = {}
    for label, records in (("Before", before_results), ("After", after_results)):
        if not isinstance(records, list) or not records:
            raise ValueError(f"{label} results must be a nonempty list.")
        values_by_id = {}
        for row_number, record in enumerate(records, start=1):
            if not isinstance(record, dict):
                raise ValueError(f"{label} record {row_number}: expected a dictionary.")
            bond_id = record.get("bond_id")
            if not isinstance(bond_id, str) or not bond_id.strip():
                raise ValueError(f"{label} record {row_number}: a nonblank text bond_id is required.")
            if bond_id in values_by_id:
                raise ValueError(f"{label} results: duplicate bond_id {bond_id!r}.")
            if record.get("valuation_date") != valuation_anchor:
                raise ValueError(
                    f"{label} {bond_id}: valuation_date must equal {valuation_anchor}."
                )
            values_by_id[bond_id] = finite_decimal(
                record.get("present_value_usd"), f"{label} {bond_id} present_value_usd"
            )
        indexed_results[label] = values_by_id
    before_values = indexed_results["Before"]
    after_values = indexed_results["After"]
    if before_values.keys() != after_values.keys():
        missing = sorted(before_values.keys() - after_values.keys())
        extra = sorted(after_values.keys() - before_values.keys())
        raise ValueError(f"Bond ID sets differ: missing after {missing}; extra after {extra}.")

    bond_changes = []
    before_total = Decimal("0")
    after_total = Decimal("0")
    for record in before_results:
        bond_id = record["bond_id"]
        before_pv = before_values[bond_id]
        after_pv = after_values[bond_id]
        change = calculate_valuation_change(before_pv, after_pv)
        bond_changes.append({
            "bond_id": bond_id,
            "before_pv_usd": before_pv,
            "after_pv_usd": after_pv,
            **change,
        })
        before_total += before_pv
        after_total += after_pv
    portfolio_change = {"before_pv_usd": before_total, "after_pv_usd": after_total}
    portfolio_change.update(calculate_valuation_change(before_total, after_total))
    return {
        "valuation_anchor": valuation_anchor,
        "before_yield_date": before_yield_date,
        "after_yield_date": after_yield_date,
        "bond_changes": bond_changes,
        "portfolio_change": portfolio_change,
    }


def calculate_bond_weight(bond_record, yield_date, portfolio_pv_by_yield_date):
    """Return one bond's percent weight using the portfolio PV for the same date."""
    if not isinstance(yield_date, (date, datetime)):
        raise ValueError("Yield observation date must be a date/datetime.")
    if not isinstance(bond_record, dict):
        raise ValueError("Bond weight input must be a dictionary.")
    dated_bond_pvs = bond_record.get("pv_by_yield_date")
    if not isinstance(dated_bond_pvs, dict) or yield_date not in dated_bond_pvs:
        raise ValueError(f"Bond PV is missing for yield date {yield_date}.")
    if (not isinstance(portfolio_pv_by_yield_date, dict)
            or yield_date not in portfolio_pv_by_yield_date):
        raise ValueError(f"Portfolio PV is missing for yield date {yield_date}.")
    bond_pv = finite_decimal(dated_bond_pvs[yield_date], f"Bond PV for {yield_date}")
    portfolio_pv = finite_decimal(
        portfolio_pv_by_yield_date[yield_date], f"Portfolio PV for {yield_date}"
    )
    if bond_pv < Decimal("0"):
        raise ValueError(f"Bond PV for {yield_date} must be nonnegative.")
    if portfolio_pv <= Decimal("0"):
        raise ValueError(f"Portfolio PV for {yield_date} must be strictly positive.")
    return bond_pv / portfolio_pv * Decimal("100")


def build_weight_analysis(comparison_results):
    """Copy existing dated PVs into a separate structure of percent weights."""
    if not isinstance(comparison_results, dict) or not comparison_results:
        raise ValueError("Weight analysis requires a nonempty comparison dictionary.")
    valuation_anchor = comparison_results.get("valuation_anchor")
    if not isinstance(valuation_anchor, (date, datetime)):
        raise ValueError("Weight analysis valuation anchor must be a date/datetime.")
    before_date = comparison_results.get("before_yield_date")
    after_date = comparison_results.get("after_yield_date")
    before_calendar = calendar_date(before_date, "Before yield date")
    after_calendar = calendar_date(after_date, "After yield date")
    if before_date == after_date or before_calendar == after_calendar:
        raise ValueError("Before and after yield dates must be distinct calendar dates and keys.")
    portfolio_change = comparison_results.get("portfolio_change")
    if not isinstance(portfolio_change, dict) or not portfolio_change:
        raise ValueError("Weight analysis requires portfolio_change values.")
    portfolio_pv_by_yield_date = {
        before_date: finite_decimal(portfolio_change.get("before_pv_usd"), "Before portfolio PV"),
        after_date: finite_decimal(portfolio_change.get("after_pv_usd"), "After portfolio PV"),
    }
    for yield_date, total in portfolio_pv_by_yield_date.items():
        if total <= Decimal("0"):
            raise ValueError(f"Portfolio PV for {yield_date} must be strictly positive.")
    rows = comparison_results.get("bond_changes")
    if not isinstance(rows, list) or not rows:
        raise ValueError("Weight analysis requires a nonempty bond_changes list.")

    bond_weights = []
    bond_ids = set()
    pv_sums = {before_date: Decimal("0"), after_date: Decimal("0")}
    for row_number, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            raise ValueError(f"Weight record {row_number}: expected a dictionary.")
        bond_id = row.get("bond_id")
        if not isinstance(bond_id, str) or not bond_id.strip():
            raise ValueError(f"Weight record {row_number}: a nonblank text bond_id is required.")
        if bond_id in bond_ids:
            raise ValueError(f"Weight analysis: duplicate bond_id {bond_id!r}.")
        bond_ids.add(bond_id)
        bond = {
            "bond_id": bond_id,
            "pv_by_yield_date": {
                before_date: finite_decimal(row.get("before_pv_usd"), f"{bond_id} before PV"),
                after_date: finite_decimal(row.get("after_pv_usd"), f"{bond_id} after PV"),
            },
        }
        before_weight = calculate_bond_weight(bond, before_date, portfolio_pv_by_yield_date)
        after_weight = calculate_bond_weight(bond, after_date, portfolio_pv_by_yield_date)
        bond["weight_pct_by_yield_date"] = {before_date: before_weight, after_date: after_weight}
        bond["weight_change_pp"] = after_weight - before_weight
        bond_weights.append(bond)
        for yield_date in pv_sums:
            pv_sums[yield_date] += bond["pv_by_yield_date"][yield_date]

    for yield_date, total in portfolio_pv_by_yield_date.items():
        # Tiny scale-relative tolerance permits Decimal grouping differences only.
        tolerance = max(abs(total), abs(pv_sums[yield_date])) * Decimal("1e-24")
        if abs(total - pv_sums[yield_date]) > tolerance:
            raise ValueError(f"Portfolio PV for {yield_date} does not reconcile with bond PVs.")
    return {
        "valuation_anchor": valuation_anchor,
        "before_yield_date": before_date,
        "after_yield_date": after_date,
        "portfolio_pv_by_yield_date": portfolio_pv_by_yield_date,
        "bond_weights": bond_weights,
    }


def index_risk_records(records, label):
    """Index nonempty supplied records by unique IDs without changing them."""
    if not isinstance(records, list) or not records:
        raise ValueError(f"{label}: expected a nonempty list.")
    indexed = {}
    for number, record in enumerate(records, start=1):
        if not isinstance(record, dict):
            raise ValueError(f"{label} record {number}: expected a dictionary.")
        bond_id = record.get("bond_id")
        if not isinstance(bond_id, str) or not bond_id.strip():
            raise ValueError(f"{label} record {number}: a nonblank bond_id is required.")
        if bond_id in indexed:
            raise ValueError(f"{label}: duplicate bond_id {bond_id!r}.")
        indexed[bond_id] = record
    return indexed


def risk_dated_value(values, yield_date, label):
    """Require one finite numeric value at the exact requested date key."""
    if not isinstance(values, dict) or yield_date not in values:
        raise ValueError(f"{label}: missing yield date {yield_date}.")
    return finite_decimal(values[yield_date], label)


def require_risk_match(actual, expected, label):
    """Allow only tiny Decimal rounding differences when reconciling values."""
    actual = finite_decimal(actual, label)
    expected = finite_decimal(expected, label)
    tolerance = max(abs(actual), abs(expected)) * Decimal("1e-24")
    if abs(actual - expected) > tolerance:
        raise ValueError(f"{label}: inconsistent values {actual} and {expected}.")


def duration_inputs(bond, pv_usd):
    """Validate the regular semiannual long-position cash flows and own PV."""
    if not isinstance(bond, dict):
        raise ValueError("Duration input must be one enriched bond dictionary.")
    N = bond.get("number_of_payments")
    if isinstance(N, bool) or not isinstance(N, int) or N <= 0:
        raise ValueError("Duration number_of_payments must be a positive integer.")
    m = finite_decimal(bond.get("coupon_payments_per_year"), "Duration frequency")
    if m != Decimal("2"):
        raise ValueError("Duration supports semiannual frequency 2 only.")
    C = finite_decimal(bond.get("coupon_payment_usd"), "Duration coupon payment")
    F = finite_decimal(bond.get("maturity_repayment_usd"), "Duration principal repayment")
    r = finite_decimal(bond.get("semiannual_discount_rate_decimal"), "Duration periodic rate")
    pv = finite_decimal(pv_usd, "Duration same-scenario PV")
    if C < Decimal("0") or F <= Decimal("0") or pv <= Decimal("0"):
        raise ValueError("Duration requires nonnegative coupons, positive principal and positive PV.")
    if Decimal("1") + r <= Decimal("0"):
        raise ValueError("Duration requires 1 + periodic discount rate > 0.")
    return C, N, F, r, m, pv


def calculate_modified_duration(bond, pv_usd):
    """Return yield-based modified duration in years using the bond's own PV."""
    C, N, F, r, m, pv = duration_inputs(bond, pv_usd)
    weighted_pv_years = Decimal("0")
    for k in range(1, N + 1):
        weighted_pv_years += (Decimal(k) / m) * C / (Decimal("1") + r) ** k
    weighted_pv_years += (Decimal(N) / m) * F / (Decimal("1") + r) ** N
    macaulay_years = weighted_pv_years / pv
    return macaulay_years / (Decimal("1") + r)


def calculate_dv01(modified_duration, pv_usd):
    """Return positive long-position USD sensitivity per 1 bp annual-yield move."""
    duration = finite_decimal(modified_duration, "Modified duration")
    pv = finite_decimal(pv_usd, "DV01 same-scenario PV")
    if duration < Decimal("0") or pv <= Decimal("0"):
        raise ValueError("DV01 requires nonnegative duration and strictly positive PV.")
    return duration * pv * Decimal("0.0001")


def calculate_portfolio_modified_duration(bond_metrics):
    """Weight durations by the same scenario's unrounded market values."""
    if not isinstance(bond_metrics, list) or not bond_metrics:
        raise ValueError("Portfolio duration requires nonempty bond metrics.")
    total_pv = Decimal("0")
    weighted_duration = Decimal("0")
    for metric in bond_metrics:
        if not isinstance(metric, dict):
            raise ValueError("Portfolio duration requires dictionary records.")
        pv = finite_decimal(metric.get("pv_usd"), "Portfolio bond PV")
        duration = finite_decimal(metric.get("modified_duration"), "Portfolio bond duration")
        if pv <= Decimal("0") or duration < Decimal("0"):
            raise ValueError("Portfolio duration requires positive PVs and nonnegative durations.")
        total_pv += pv
        weighted_duration += pv * duration
    return weighted_duration / total_pv


def calculate_portfolio_dv01(bond_metrics):
    """Sum full-position dollar DV01s directly; do not weight them again."""
    if not isinstance(bond_metrics, list) or not bond_metrics:
        raise ValueError("Portfolio DV01 requires nonempty bond metrics.")
    total_dv01 = Decimal("0")
    for metric in bond_metrics:
        if not isinstance(metric, dict):
            raise ValueError("Portfolio DV01 requires dictionary records.")
        dv01 = finite_decimal(metric.get("dv01_usd_per_bp"), "Portfolio bond DV01")
        if dv01 < Decimal("0"):
            raise ValueError("Long-position portfolio DV01 must be nonnegative.")
        total_dv01 += dv01
    return total_dv01


def build_risk_analyses(scenarios, weight_analysis):
    """Match existing enriched bonds/PVs/dated weights and create new risk records."""
    if not isinstance(weight_analysis, dict):
        raise ValueError("Risk analysis requires the dated weight dictionary.")
    anchor = weight_analysis.get("valuation_anchor")
    calendar_date(anchor, "Risk valuation anchor")
    dates = (weight_analysis.get("before_yield_date"), weight_analysis.get("after_yield_date"))
    calendars = [calendar_date(value, "Risk yield date") for value in dates]
    if dates[0] == dates[1] or calendars[0] == calendars[1]:
        raise ValueError("Risk yield dates must be distinct calendar dates and keys.")
    weights = index_risk_records(weight_analysis.get("bond_weights"), "Risk weights")
    if not isinstance(scenarios, list) or len(scenarios) != len(dates):
        raise ValueError("Risk analysis requires one enriched/result pair per yield date.")
    by_date = {}
    for scenario in scenarios:
        if not isinstance(scenario, tuple) or len(scenario) != 3:
            raise ValueError("Each risk scenario requires (yield_date, enriched_bonds, results).")
        yield_date, bonds, results = scenario
        calendar_date(yield_date, "Supplied risk yield date")
        if yield_date not in dates or yield_date in by_date:
            raise ValueError("Risk scenarios contain an unexpected or duplicate yield date.")
        by_date[yield_date] = (
            index_risk_records(bonds, "Enriched risk bonds"),
            index_risk_records(results, "Risk PV results"),
        )
    duration_rows = {bond_id: {"bond_id": bond_id, "modified_duration_by_yield_date": {}}
                     for bond_id in weights}
    dv01_rows = {bond_id: {"bond_id": bond_id, "dv01_usd_by_yield_date": {}}
                 for bond_id in weights}
    portfolio_durations = {}
    portfolio_dv01s = {}
    for yield_date in dates:
        bonds, results = by_date[yield_date]
        if bonds.keys() != weights.keys() or results.keys() != weights.keys():
            raise ValueError(f"Risk bond/PV/weight ID sets differ for {yield_date}.")
        total = risk_dated_value(
            weight_analysis.get("portfolio_pv_by_yield_date"), yield_date, "Risk portfolio PV"
        )
        if total <= Decimal("0"):
            raise ValueError("Risk portfolio PV must be positive.")
        metrics = []
        for bond_id, weight in weights.items():
            bond, result = bonds[bond_id], results[bond_id]
            if bond.get("valuation_date") != anchor or result.get("valuation_date") != anchor:
                raise ValueError(f"{bond_id}: incompatible risk valuation anchor.")
            if "yield_observation_date" in bond and bond["yield_observation_date"] != yield_date:
                raise ValueError(f"{bond_id}: incompatible yield observation date.")
            pv = finite_decimal(result.get("present_value_usd"), f"{bond_id} risk PV")
            require_risk_match(pv, risk_dated_value(
                weight.get("pv_by_yield_date"), yield_date, f"{bond_id} dated PV"
            ), f"{bond_id} same-scenario PV")
            annual_yield = finite_decimal(bond.get("annual_treasury_yield_pct"), "Risk annual yield")
            require_risk_match(bond.get("semiannual_discount_rate_decimal"),
                               annual_yield / Decimal("100") / Decimal("2"), "Risk yield units")
            for field in ("coupon_payment_usd", "number_of_payments", "annual_treasury_yield_pct",
                          "semiannual_discount_rate_decimal"):
                require_risk_match(bond.get(field), result.get(field), f"{bond_id} {field}")
            weight_pct = risk_dated_value(
                weight.get("weight_pct_by_yield_date"), yield_date, f"{bond_id} dated weight"
            )
            require_risk_match(weight_pct, pv / total * Decimal("100"), "Risk same-date weight")
            duration = calculate_modified_duration(bond, pv)
            dv01 = calculate_dv01(duration, pv)
            duration_rows[bond_id]["modified_duration_by_yield_date"][yield_date] = duration
            dv01_rows[bond_id]["dv01_usd_by_yield_date"][yield_date] = dv01
            metrics.append({"pv_usd": pv, "modified_duration": duration,
                            "dv01_usd_per_bp": dv01, "weight_pct": weight_pct})
        require_risk_match(sum((item["pv_usd"] for item in metrics), Decimal("0")),
                           total, "Risk portfolio PV reconciliation")
        portfolio_duration = calculate_portfolio_modified_duration(metrics)
        portfolio_dv01 = calculate_portfolio_dv01(metrics)
        weighted_from_p011 = sum((item["weight_pct"] / Decimal("100") * item["modified_duration"]
                                 for item in metrics), Decimal("0"))
        require_risk_match(portfolio_duration, weighted_from_p011, "Risk dated-weight aggregation")
        require_risk_match(portfolio_dv01, calculate_dv01(portfolio_duration, total), "Portfolio DV01 identity")
        portfolio_durations[yield_date] = portfolio_duration
        portfolio_dv01s[yield_date] = portfolio_dv01
    duration_analysis = {
        "valuation_anchor": anchor, "yield_dates": list(dates), "unit": "years",
        "convention": "Yield-based modified duration; own-scenario PVs; parallel annual-yield shift",
        "bonds": list(duration_rows.values()),
        "portfolio": {"modified_duration_by_yield_date": portfolio_durations},
    }
    dv01_analysis = {
        "valuation_anchor": anchor, "yield_dates": list(dates), "unit": "USD per bp",
        "convention": "Positive long-position DV01; +1 bp implies approximately -DV01 USD",
        "bonds": list(dv01_rows.values()),
        "portfolio": {"dv01_usd_by_yield_date": portfolio_dv01s},
    }
    return duration_analysis, dv01_analysis


def bump_bond_yield(bond, bump_bp):
    """Copy a bond and change only consistent annual-percent/periodic yield fields."""
    if not isinstance(bond, dict):
        raise ValueError("Yield bump requires one enriched bond dictionary.")
    if finite_decimal(bond.get("coupon_payments_per_year"), "Bump frequency") != Decimal("2"):
        raise ValueError("Yield bump supports semiannual frequency 2 only.")
    bump_bp = finite_decimal(bump_bp, "Annual yield bump in bp")
    annual = finite_decimal(bond.get("annual_treasury_yield_pct"), "Bump center annual yield")
    periodic = finite_decimal(bond.get("semiannual_discount_rate_decimal"), "Bump center periodic rate")
    require_risk_match(periodic, annual / Decimal("100") / Decimal("2"), "Bump center yield units")
    bumped = bond.copy()
    bumped["annual_treasury_yield_pct"] = annual + bump_bp * Decimal("0.01")
    bumped["semiannual_discount_rate_decimal"] = bumped["annual_treasury_yield_pct"] / Decimal("100") / Decimal("2")
    if Decimal("1") + bumped["semiannual_discount_rate_decimal"] <= Decimal("0"):
        raise ValueError("Bumped yield requires 1 + periodic discount rate > 0.")
    return bumped


def baseline_bump_changes(pv, pv_up, pv_down, dv01):
    """Use the unrounded baseline PV denominator for all four percentage changes."""
    pv = finite_decimal(pv, "Baseline bump PV")
    pv_up = finite_decimal(pv_up, "Up-bump PV")
    pv_down = finite_decimal(pv_down, "Down-bump PV")
    dv01 = finite_decimal(dv01, "Baseline DV01")
    if pv <= Decimal("0") or dv01 < Decimal("0"):
        raise ValueError("Baseline bump analysis requires positive PV and nonnegative DV01.")
    return {
        "baseline_pv_usd": pv, "pv_up_usd": pv_up, "pv_down_usd": pv_down,
        "full_repricing_up_pct": (pv_up - pv) / pv * Decimal("100"),
        "full_repricing_down_pct": (pv_down - pv) / pv * Decimal("100"),
        "dv01_estimate_up_pct": -dv01 / pv * Decimal("100"),
        "dv01_estimate_down_pct": dv01 / pv * Decimal("100"),
    }


def build_baseline_bump_analysis(baseline_bonds, baseline_results, dv01_analysis, weight_analysis):
    """Reprice copied baseline cash flows for parallel +/-1 bp annual-yield shocks."""
    if not isinstance(weight_analysis, dict) or not isinstance(dv01_analysis, dict):
        raise ValueError("Baseline bump analysis requires weight and DV01 dictionaries.")
    anchor = weight_analysis.get("valuation_anchor")
    calendar_date(anchor, "Baseline bump valuation anchor")
    baseline_date = weight_analysis.get("before_yield_date")
    after_date = weight_analysis.get("after_yield_date")
    baseline_calendar = calendar_date(baseline_date, "Baseline bump yield date")
    after_calendar = calendar_date(after_date, "Other scenario yield date")
    if baseline_date == after_date or baseline_calendar == after_calendar:
        raise ValueError("Baseline and other scenario require distinct yield dates and keys.")
    if dv01_analysis.get("valuation_anchor") != anchor:
        raise ValueError("Baseline DV01 and weight valuation anchors differ.")
    bonds = index_risk_records(baseline_bonds, "Baseline bump bonds")
    results = index_risk_records(baseline_results, "Baseline bump PV results")
    risk_rows = index_risk_records(dv01_analysis.get("bonds"), "Baseline DV01 records")
    weights = index_risk_records(weight_analysis.get("bond_weights"), "Baseline dated weights")
    if bonds.keys() != results.keys() or bonds.keys() != risk_rows.keys() or bonds.keys() != weights.keys():
        raise ValueError("Baseline bump bond/PV/DV01/weight ID sets differ.")
    rows = []
    total_before = Decimal("0")
    total_up = Decimal("0")
    total_down = Decimal("0")
    total_dv01 = Decimal("0")
    for bond_id in weights:
        bond, result = bonds[bond_id], results[bond_id]
        if bond.get("valuation_date") != anchor or result.get("valuation_date") != anchor:
            raise ValueError(f"{bond_id}: incompatible baseline bump valuation anchor.")
        if "yield_observation_date" in bond and bond["yield_observation_date"] != baseline_date:
            raise ValueError(f"{bond_id}: bump center must use the baseline yield observation.")
        pv = finite_decimal(result.get("present_value_usd"), "Baseline full-position PV")
        duration_inputs(bond, pv)
        require_risk_match(pv, risk_dated_value(weights[bond_id].get("pv_by_yield_date"),
                                              baseline_date, "Baseline dated PV"), "Baseline PV")
        dv01 = risk_dated_value(risk_rows[bond_id].get("dv01_usd_by_yield_date"), baseline_date, "Baseline bond DV01")
        pv_up = present_value(bump_bond_yield(bond, Decimal("1")))
        pv_down = present_value(bump_bond_yield(bond, Decimal("-1")))
        rows.append({"bond_id": bond_id, **baseline_bump_changes(pv, pv_up, pv_down, dv01)})
        total_before += pv
        total_up += pv_up
        total_down += pv_down
        total_dv01 += dv01
    portfolio_pv = risk_dated_value(weight_analysis.get("portfolio_pv_by_yield_date"), baseline_date, "Baseline portfolio PV")
    require_risk_match(total_before, portfolio_pv, "Baseline portfolio PV sum")
    portfolio_risk = dv01_analysis.get("portfolio")
    if not isinstance(portfolio_risk, dict):
        raise ValueError("Baseline bump analysis requires portfolio DV01 values.")
    portfolio_dv01 = risk_dated_value(portfolio_risk.get("dv01_usd_by_yield_date"), baseline_date, "Baseline portfolio DV01")
    require_risk_match(total_dv01, portfolio_dv01, "Baseline portfolio DV01 sum")
    return {
        "valuation_anchor": anchor, "baseline_yield_date": baseline_date,
        "bump_size_bp": Decimal("1"), "annual_decimal_bump": Decimal("0.0001"),
        "semiannual_decimal_bump": Decimal("0.00005"), "unit": "percentage change from baseline PV",
        "bonds": rows,
        "portfolio": baseline_bump_changes(portfolio_pv, total_up, total_down, portfolio_dv01),
    }


def print_baseline_bump_table(baseline_bump_analysis):
    """Show signed first-order estimates beside exact +/-1 bp baseline repricing."""
    print()
    print(f"Baseline parallel +/-1 bp annual-yield sensitivity (%): yields dated {baseline_bump_analysis['baseline_yield_date']}; "
          f"valuation anchor {baseline_bump_analysis['valuation_anchor']}")
    header = (f"{'Bond/portfolio':<16} {'+1 bp DV01 (%)':>18} {'+1 bp repriced (%)':>21} "
              f"{'-1 bp DV01 (%)':>18} {'-1 bp repriced (%)':>21}")
    print(header)
    print("-" * len(header))
    for row in baseline_bump_analysis["bonds"] + [{"bond_id": "Portfolio", **baseline_bump_analysis["portfolio"]}]:
        print(f"{row['bond_id']:<16} {row['dv01_estimate_up_pct']:>+18.6f} "
              f"{row['full_repricing_up_pct']:>+21.6f} {row['dv01_estimate_down_pct']:>+18.6f} "
              f"{row['full_repricing_down_pct']:>+21.6f}")


def build_movement_analysis(baseline_bonds, scenario_bonds, comparison_results, dv01_analysis):
    """Explain existing model changes with baseline DV01 and each actual tenor move."""
    if not isinstance(comparison_results, dict) or not isinstance(dv01_analysis, dict):
        raise ValueError("Movement analysis requires comparison and DV01 dictionaries.")
    anchor = comparison_results.get("valuation_anchor")
    calendar_date(anchor, "Movement valuation anchor")
    before_date = comparison_results.get("before_yield_date")
    after_date = comparison_results.get("after_yield_date")
    before_calendar = calendar_date(before_date, "Movement before yield date")
    after_calendar = calendar_date(after_date, "Movement after yield date")
    if before_date == after_date or before_calendar == after_calendar:
        raise ValueError("Movement yield dates must be distinct calendars and keys.")
    if dv01_analysis.get("valuation_anchor") != anchor or dv01_analysis.get("yield_dates") != [before_date, after_date]:
        raise ValueError("Movement DV01 dates/anchor differ from the comparison.")
    before = index_risk_records(baseline_bonds, "Movement baseline bonds")
    after = index_risk_records(scenario_bonds, "Movement scenario bonds")
    changes = index_risk_records(comparison_results.get("bond_changes"), "Movement full changes")
    risks = index_risk_records(dv01_analysis.get("bonds"), "Movement baseline DV01")
    if before.keys() != changes.keys() or after.keys() != changes.keys() or risks.keys() != changes.keys():
        raise ValueError("Movement baseline/scenario/comparison/DV01 bond ID sets differ.")
    portfolio_change = comparison_results.get("portfolio_change")
    if not isinstance(portfolio_change, dict):
        raise ValueError("Movement analysis requires the saved portfolio comparison.")
    portfolio_before = finite_decimal(portfolio_change.get("before_pv_usd"), "Movement starting total PV")
    portfolio_after = finite_decimal(portfolio_change.get("after_pv_usd"), "Movement ending total PV")
    if portfolio_before <= Decimal("0") or portfolio_after <= Decimal("0"):
        raise ValueError("Movement portfolio PVs must be positive.")
    tolerance = max(portfolio_before, portfolio_after) * Decimal("1e-24")
    rows = []
    total_estimate = Decimal("0")
    total_full = Decimal("0")
    total_residual = Decimal("0")
    total_before_pv = Decimal("0")
    total_after_pv = Decimal("0")
    total_baseline_dv01 = Decimal("0")
    for bond_id, change in changes.items():
        baseline, scenario = before[bond_id], after[bond_id]
        for bond, yield_date in ((baseline, before_date), (scenario, after_date)):
            if bond.get("valuation_date") != anchor:
                raise ValueError(f"{bond_id}: incompatible movement valuation anchor.")
            if "yield_observation_date" in bond and bond["yield_observation_date"] != yield_date:
                raise ValueError(f"{bond_id}: incompatible movement yield observation.")
        for field in ("benchmark_maturity", "coupon_payment_usd", "number_of_payments",
                      "maturity_repayment_usd", "coupon_payments_per_year", "maturity_date"):
            if field not in baseline or field not in scenario or baseline[field] != scenario[field]:
                raise ValueError(f"{bond_id}: missing or changed movement context {field}.")
        before_pv = finite_decimal(change.get("before_pv_usd"), f"{bond_id} movement before PV")
        after_pv = finite_decimal(change.get("after_pv_usd"), f"{bond_id} movement after PV")
        duration_inputs(baseline, before_pv)
        duration_inputs(scenario, after_pv)
        before_yield = finite_decimal(baseline.get("annual_treasury_yield_pct"), f"{bond_id} before yield")
        after_yield = finite_decimal(scenario.get("annual_treasury_yield_pct"), f"{bond_id} after yield")
        for bond, annual in ((baseline, before_yield), (scenario, after_yield)):
            require_risk_match(bond.get("semiannual_discount_rate_decimal"),
                               annual / Decimal("100") / Decimal("2"), "Movement yield units")
        dv01 = risk_dated_value(risks[bond_id].get("dv01_usd_by_yield_date"), before_date, f"{bond_id} baseline DV01")
        if dv01 < Decimal("0"):
            raise ValueError(f"{bond_id}: baseline long-position DV01 must be nonnegative.")
        full_change = finite_decimal(change.get("dollar_change_usd"), f"{bond_id} saved full change")
        if abs(full_change - (after_pv - before_pv)) > tolerance:
            raise ValueError(f"{bond_id}: saved full change does not match its PVs.")
        yield_change_bp = (after_yield - before_yield) * Decimal("100")
        estimate = -dv01 * yield_change_bp
        residual = full_change - estimate
        rows.append({
            "bond_id": bond_id, "benchmark_maturity": baseline["benchmark_maturity"],
            "before_annual_yield_pct": before_yield, "after_annual_yield_pct": after_yield,
            "baseline_dv01_usd_per_bp": dv01, "yield_change_bp": yield_change_bp,
            "estimated_change_usd": estimate, "full_repricing_change_usd": full_change,
            "residual_usd": residual,
        })
        total_estimate += estimate
        total_full += full_change
        total_residual += residual
        total_before_pv += before_pv
        total_after_pv += after_pv
        total_baseline_dv01 += dv01
    if abs(total_before_pv - portfolio_before) > tolerance or abs(total_after_pv - portfolio_after) > tolerance:
        raise ValueError("Movement bond PVs do not reconcile with their dated portfolio totals.")
    portfolio_risk = dv01_analysis.get("portfolio")
    if not isinstance(portfolio_risk, dict):
        raise ValueError("Movement analysis requires portfolio DV01 metadata.")
    baseline_portfolio_dv01 = risk_dated_value(
        portfolio_risk.get("dv01_usd_by_yield_date"), before_date, "Movement baseline portfolio DV01"
    )
    require_risk_match(total_baseline_dv01, baseline_portfolio_dv01, "Movement baseline DV01 sum")
    saved_portfolio_change = finite_decimal(portfolio_change.get("dollar_change_usd"), "Saved portfolio full change")
    if abs(total_full - saved_portfolio_change) > tolerance:
        raise ValueError("Movement full changes do not reconcile with the saved portfolio change.")
    if abs(saved_portfolio_change - (portfolio_after - portfolio_before)) > tolerance:
        raise ValueError("Movement saved portfolio change does not match its dated PVs.")
    if abs(total_residual - (total_full - total_estimate)) > tolerance:
        raise ValueError("Movement portfolio residual does not reconcile.")
    return {
        "valuation_anchor": anchor, "before_yield_date": before_date, "after_yield_date": after_date,
        "residual_convention": "Full repricing minus first-order estimate within this model",
        "reconciliation_tolerance_usd": tolerance, "bonds": rows,
        "portfolio": {"estimated_change_usd": total_estimate,
                      "full_repricing_change_usd": total_full, "residual_usd": total_residual},
    }


def print_movement_table(movement_analysis):
    """Show mixed-tenor movement and internal first-order approximation errors."""
    print()
    print(f"Observed yield-scenario movement: {movement_analysis['before_yield_date']} -> "
          f"{movement_analysis['after_yield_date']}; valuation anchor {movement_analysis['valuation_anchor']}")
    print("Residual = full repricing minus first-order estimate within this model (USD).")
    header = (f"{'Bond/portfolio':<16} {'Yield change (bp)':>19} {'DV01 estimate (USD)':>22} "
              f"{'Full repricing (USD)':>24} {'Residual (USD)':>20}")
    print(header)
    print("-" * len(header))
    for row in movement_analysis["bonds"]:
        print(f"{row['bond_id']:<16} {row['yield_change_bp']:>+19.2f} "
              f"{row['estimated_change_usd']:>+22,.2f} {row['full_repricing_change_usd']:>+24,.2f} "
              f"{row['residual_usd']:>+20,.2f}")
    total = movement_analysis["portfolio"]
    print(f"{'Portfolio':<16} {'mixed tenor moves':>19} {total['estimated_change_usd']:>+22,.2f} "
          f"{total['full_repricing_change_usd']:>+24,.2f} {total['residual_usd']:>+20,.2f}")
    print("This approximation check does not validate prices against market quotes or calculate convexity.")


def create_movement_figure(observations, movement_analysis, weight_analysis):
    """Convert to float only here, at the Matplotlib plotting boundary."""
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.figure import Figure
    from matplotlib.ticker import StrMethodFormatter

    anchor = movement_analysis["valuation_anchor"]
    dates = (movement_analysis["before_yield_date"], movement_analysis["after_yield_date"])
    if (weight_analysis.get("valuation_anchor") != anchor
            or (weight_analysis.get("before_yield_date"), weight_analysis.get("after_yield_date")) != dates):
        raise ValueError("Figure weight dates/anchor differ from movement analysis.")
    movements = index_risk_records(movement_analysis.get("bonds"), "Figure movement rows")
    weights = index_risk_records(weight_analysis.get("bond_weights"), "Figure value weights")
    if movements.keys() != weights.keys():
        raise ValueError("Figure movement/weight bond IDs differ.")
    bond_ids = list(movements)
    years = []
    for maturity in MATURITY_COLUMNS:
        if not maturity.endswith(" Yr"):
            raise ValueError("Selected figure maturities must be quoted in years.")
        years.append(float(Decimal(maturity.split()[0])))
    figure = Figure(figsize=(16, 12), facecolor="white")
    FigureCanvasAgg(figure)
    axes = figure.subplots(2, 2)
    figure.subplots_adjust(left=0.07, right=0.96, top=0.86, bottom=0.17, wspace=0.28, hspace=0.57)
    figure.suptitle("Treasury yield scenarios and the synthetic portfolio", fontsize=19, y=0.97)
    figure.text(0.5, 0.925, f"Treasury yields: {dates[0]} and {dates[1]} | Fixed valuation anchor: {anchor}",
                ha="center", fontsize=12, color="#444444")
    blue, orange = "#176B95", "#D58D32"
    for yield_date, color, marker, style in zip(dates, (blue, orange), ("o", "s"), ("-", "--")):
        if yield_date not in observations:
            raise ValueError(f"Figure is missing Treasury observations for {yield_date}.")
        yields = [float(finite_decimal(observations[yield_date].get(maturity), "Figure Treasury yield"))
                  for maturity in MATURITY_COLUMNS]
        axes[0, 0].plot(years, yields, color=color, marker=marker, linestyle=style,
                        linewidth=2, markersize=7, label=str(yield_date))
    axes[0, 0].set(title="Selected Treasury benchmark points", xlabel="Maturity (years)", ylabel="Annual par yield (%)")
    axes[0, 0].set_xticks(years)
    axes[0, 0].legend(frameon=False, fontsize=11)
    axes[0, 0].grid(axis="y", color="#DDDDDD", linewidth=0.7)
    positions = list(range(len(bond_ids)))
    estimates = [movements[bond_id]["estimated_change_usd"] for bond_id in bond_ids]
    full_changes = [movements[bond_id]["full_repricing_change_usd"] for bond_id in bond_ids]
    estimated_bars = axes[0, 1].bar([x - 0.19 for x in positions], [float(value) for value in estimates],
                                   width=0.36, color=blue, label="Baseline DV01 estimate")
    full_bars = axes[0, 1].bar([x + 0.19 for x in positions], [float(value) for value in full_changes],
                              width=0.36, color=orange, hatch="//", label="Full model repricing")
    axes[0, 1].axhline(0, color="#444444", linewidth=1)
    axes[0, 1].set(title="Signed valuation movement by bond", ylabel="Valuation change (USD)")
    axes[0, 1].set_xticks(positions, bond_ids)
    axes[0, 1].yaxis.set_major_formatter(StrMethodFormatter("{x:,.0f}"))
    axes[0, 1].legend(frameon=False, fontsize=10, loc="upper right")
    axes[0, 1].margins(y=0.17)
    axes[0, 1].bar_label(estimated_bars, labels=[f"{value:+,.2f}" for value in estimates], padding=4, fontsize=9)
    axes[0, 1].bar_label(full_bars, labels=[f"{value:+,.2f}" for value in full_changes], padding=4, fontsize=9)
    palette = (blue, orange, "#71864A", "#9C6380", "#62769A")
    bond_colors = [palette[index % len(palette)] for index in positions]
    for axis, yield_date in zip(axes[1], dates):
        total = risk_dated_value(weight_analysis.get("portfolio_pv_by_yield_date"), yield_date, "Figure portfolio PV")
        if total <= Decimal("0"):
            raise ValueError("Composition pies require positive portfolio PVs.")
        pvs = [risk_dated_value(weights[bond_id].get("pv_by_yield_date"), yield_date, "Figure bond PV")
               for bond_id in bond_ids]
        percentages = [risk_dated_value(weights[bond_id].get("weight_pct_by_yield_date"), yield_date, "Figure weight")
                       for bond_id in bond_ids]
        if any(value <= Decimal("0") for value in pvs):
            raise ValueError("Composition pie slices must be positive estimated PVs.")
        require_risk_match(sum(pvs, Decimal("0")), total, "Figure PV total")
        for pv, percentage in zip(pvs, percentages):
            require_risk_match(percentage, pv / total * Decimal("100"), "Figure same-date percentage")
        labels = [f"{bond_id}\n{percentage:.4f}%" for bond_id, percentage in zip(bond_ids, percentages)]
        axis.pie([float(value) for value in pvs], labels=labels, colors=bond_colors,
                 startangle=90, counterclock=False, radius=1, labeldistance=1.12,
                 wedgeprops={"edgecolor": "white", "linewidth": 1}, textprops={"fontsize": 11})
        axis.set_title(f"Estimated PV composition: yields {yield_date}", fontsize=12, pad=16)
        axis.legend(list(axis.patches), [f"{bond_id}: USD {pv:,.2f}" for bond_id, pv in zip(bond_ids, pvs)],
                    title=f"Portfolio PV: USD {total:,.2f}", loc="upper center", bbox_to_anchor=(0.5, -0.03),
                    frameon=False, fontsize=10, title_fontsize=11)
    for axis in axes[0]:
        axis.tick_params(labelsize=11)
        axis.title.set_fontsize(13)
        axis.xaxis.label.set_fontsize(11)
        axis.yaxis.label.set_fontsize(11)
        axis.spines[["top", "right"]].set_visible(False)
    figure.text(0.5, 0.04, "Pies show positive estimated portfolio value, with identical size/orientation; small weight shifts may look alike.\n"
                "Treasury par-yield proxy model. Estimate/repricing differences are internal approximation errors, not market-price validation.",
                ha="center", va="center", fontsize=10, color="#444444")
    return figure


def save_movement_figure(observations, movement_analysis, weight_analysis):
    """Save first in a unique project folder, using no global backend or cache settings."""
    output_root = Path(__file__).resolve().parent / "outputs" / "P013"
    output_root.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    output_directory = Path(tempfile.mkdtemp(prefix=f"{timestamp}-", dir=output_root))
    image_path = output_directory / f"p013-overview-{output_directory.name}.png"
    previous_config = os.environ.get("MPLCONFIGDIR")
    with tempfile.TemporaryDirectory(prefix=".matplotlib-", dir=output_directory) as cache:
        os.environ["MPLCONFIGDIR"] = cache
        try:
            figure = create_movement_figure(observations, movement_analysis, weight_analysis)
            figure.savefig(image_path, dpi=180, format="png", facecolor="white")
            figure.clear()
        finally:
            if previous_config is None:
                os.environ.pop("MPLCONFIGDIR", None)
            else:
                os.environ["MPLCONFIGDIR"] = previous_config
    return image_path


def print_weight_table(weight_analysis):
    """Display dated PVs and weights; round only the displayed strings."""
    before_date = weight_analysis["before_yield_date"]
    after_date = weight_analysis["after_yield_date"]
    print()
    print(
        f"Estimated market-value weights: Treasury yield dates {before_date} -> {after_date}; "
        f"valuation anchor {weight_analysis['valuation_anchor']}"
    )
    header = (
        f"{'Bond ID':<10} {'Before PV (USD)':>18} {'After PV (USD)':>18} "
        f"{'Before weight (%)':>19} {'After weight (%)':>19} {'Weight change (pp)':>20}"
    )
    print(header)
    print("-" * len(header))
    for bond in weight_analysis["bond_weights"]:
        print(
            f"{bond['bond_id']:<10} {bond['pv_by_yield_date'][before_date]:>18,.2f} "
            f"{bond['pv_by_yield_date'][after_date]:>18,.2f} "
            f"{bond['weight_pct_by_yield_date'][before_date]:>19.4f} "
            f"{bond['weight_pct_by_yield_date'][after_date]:>19.4f} "
            f"{bond['weight_change_pp']:>+20.4f}"
        )


def print_valuation_table(valuation_results, heading):
    """Display stored per-holding results; formatting never changes the inputs."""
    print()
    print(heading)
    print("Matched Treasury par yields are single-yield proxies, not observed bond prices.")
    header = (
        f"{'Bond ID':<10} {'Coupon (USD)':>16} {'Payments':>9} {'Benchmark':>10} "
        f"{'Coupon rate (%)':>17} {'Annual yield (%)':>18} "
        f"{'Semiannual r (decimal)':>24} {'PV (USD)':>20}"
    )
    print(header)
    print("-" * len(header))
    for result in valuation_results:
        print(
            f"{result['bond_id']:<10} {result['coupon_payment_usd']:>16,.2f} "
            f"{result['number_of_payments']:>9d} {result['benchmark_maturity']:>10} "
            f"{result['coupon_rate_pct']:>17.2f} "
            f"{result['annual_treasury_yield_pct']:>18.2f} "
            f"{result['semiannual_discount_rate_decimal']:>24.6f} "
            f"{result['present_value_usd']:>20,.2f}"
        )


def main():
    parser = argparse.ArgumentParser(description="Review fixed-anchor synthetic bond yield scenarios.")
    parser.add_argument("--no-show", action="store_true", help="Save the four-panel PNG without opening a viewer.")
    arguments = parser.parse_args()
    try:
        observations = read_yields(INPUT_PATH)
    except (OSError, UnicodeError, csv.Error, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    print_comparison(observations)
    try:
        portfolio = load_portfolio(PORTFOLIO_PATH)
    except (OSError, ValueError, InvalidFileException, BadZipFile) as error:
        print(f"Error loading portfolio: {error}", file=sys.stderr)
        return 1
    try:
        valuation_portfolio = prepare_valuation_portfolio(portfolio, observations)
        valuation_results = calculate_valuation_results(valuation_portfolio)
    except (ValueError, ArithmeticError) as error:
        print(f"Error preparing October 27 valuation: {error}", file=sys.stderr)
        return 1
    print_valuation_table(
        valuation_results, f"Calculated present values as of {VALUATION_DATE}: USD, full holdings"
    )
    total_pv_usd = sum(
        (result["present_value_usd"] for result in valuation_results), Decimal("0")
    )
    print(f"Total estimated synthetic portfolio PV as of {VALUATION_DATE} (USD): {total_pv_usd:,.2f}")
    try:
        scenario_portfolio = prepare_yield_scenario(
            valuation_portfolio, observations, SCENARIO_YIELD_OBSERVATION_DATE
        )
        scenario_results = calculate_valuation_results(scenario_portfolio)
    except (ValueError, ArithmeticError) as error:
        print(f"Error preparing October 28 yield scenario: {error}", file=sys.stderr)
        return 1
    print_valuation_table(
        scenario_results,
        f"October 28 yield scenario: Treasury yields dated {SCENARIO_YIELD_OBSERVATION_DATE}; "
        f"valuation anchor {VALUATION_DATE}",
    )
    scenario_total_pv_usd = sum(
        (result["present_value_usd"] for result in scenario_results), Decimal("0")
    )
    print(
        f"Total estimated synthetic portfolio PV using Treasury yields dated "
        f"{SCENARIO_YIELD_OBSERVATION_DATE}, valuation anchor {VALUATION_DATE} (USD): "
        f"{scenario_total_pv_usd:,.2f}"
    )
    try:
        comparison_results = compare_valuation_results(
            valuation_results, scenario_results,
            VALUATION_DATE, SCENARIO_YIELD_OBSERVATION_DATE, VALUATION_DATE,
        )
    except (ValueError, ArithmeticError) as error:
        print(f"Error comparing valuations: {error}", file=sys.stderr)
        return 1
    print()
    print("Signed yield-scenario valuation changes at the common valuation anchor:")
    print("percentage_change_pct is in percent; pprint shows Decimal(...) and datetime.date(...).")
    pprint.pp(comparison_results, sort_dicts=False)
    try:
        weight_analysis = build_weight_analysis(comparison_results)
    except (ValueError, ArithmeticError) as error:
        print(f"Error building weight analysis: {error}", file=sys.stderr)
        return 1
    print_weight_table(weight_analysis)
    try:
        scenarios = [
            (weight_analysis["before_yield_date"], valuation_portfolio, valuation_results),
            (weight_analysis["after_yield_date"], scenario_portfolio, scenario_results),
        ]
        duration_analysis, dv01_analysis = build_risk_analyses(scenarios, weight_analysis)
        baseline_bump_analysis = build_baseline_bump_analysis(
            valuation_portfolio, valuation_results, dv01_analysis, weight_analysis
        )
    except (ValueError, ArithmeticError) as error:
        print(f"Error calculating duration/DV01 or baseline sensitivity: {error}", file=sys.stderr)
        return 1
    print()
    print("Modified duration by Treasury yield date (years); common valuation anchor unchanged:")
    pprint.pp(duration_analysis, sort_dicts=False)
    print()
    print("Positive full-position DV01 by Treasury yield date (USD per bp):")
    pprint.pp(dv01_analysis, sort_dicts=False)
    print_baseline_bump_table(baseline_bump_analysis)
    try:
        movement_analysis = build_movement_analysis(
            valuation_portfolio, scenario_portfolio, comparison_results, dv01_analysis
        )
        print_movement_table(movement_analysis)
        image_path = save_movement_figure(observations, movement_analysis, weight_analysis)
    except (ValueError, ArithmeticError, OSError, ImportError) as error:
        print(f"Error creating movement analysis/figure: {error}", file=sys.stderr)
        return 1
    print(f"Saved P013 figure: {image_path}")
    if arguments.no_show:
        print("Figure display skipped (--no-show); PNG saved.")
    else:
        try:
            subprocess.run(["/usr/bin/open", str(image_path)], check=True)
        except (OSError, subprocess.CalledProcessError) as error:
            print(f"Error opening saved PNG in the default image viewer: {error}", file=sys.stderr)
            return 1
        print("Requested the saved PNG in the default image viewer; program finished.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
