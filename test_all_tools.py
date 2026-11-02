

"""Level 2 test: runs all 17 Excel tools on the ACTIVE dataset workbook.

Creates no workbook. Temporary changes are cleaned up and verified at the end.
Run (venv active, dataset workbook open in Excel):
    python test_all_tools.py                  # all 17
    python test_all_tools.py chart pivot      # only tools whose name contains these words
"""
import inspect
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import pandas as pd
import xlwings as xw

from excel_tools.analysis import analyze_excel_column
from excel_tools.charts import create_excel_chart
from excel_tools.cleaning import remove_excel_blank_rows, remove_excel_duplicates
from excel_tools.conditional_formatting import highlight_excel_values
from excel_tools.filter_sort import clear_excel_filter, filter_excel_data
from excel_tools.formatting import autofit_excel_columns, format_excel_cells
from excel_tools.formulas import apply_excel_formula, apply_excel_formula_range
from excel_tools.pivot_tables import create_excel_pivot_summary
from excel_tools.reader import get_excel_sheet_names, read_excel_data
from excel_tools.validator import validate_excel_data
from excel_tools.writer import append_excel_data, write_excel_data

WORKBOOK = "sales-transactions-100.xlsx"
SHEET = "sales-transactions"
EXPECTED_COLS = ["date", "region", "product", "channel", "units", "unit_price", "revenue"]
TEMP_SHEETS = ("Filtered_Data", "Pivot_Summary")
SCRATCH = "I1:N120"
S: dict = {}


class SkipTest(Exception):
    pass


# ---------------- helpers ----------------
def check(condition, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def ok(result: dict) -> dict:
    check(isinstance(result, dict) and result.get("success") is True, f"tool error: {result}")
    return result


def fails(result: dict) -> dict:
    check(isinstance(result, dict) and result.get("success") is False,
          f"fail aagirukkanum, aana: {result}")
    return result


def call(tool, **kwargs):
    kwargs.setdefault("workbook_name", S["book"].name)
    if "sheet_name" in inspect.signature(tool).parameters:
        kwargs.setdefault("sheet_name", SHEET)
    return tool(**kwargs)


def sh():
    return S["book"].sheets[SHEET]


def sheet_names() -> list[str]:
    return [s.name for s in S["book"].sheets]


def has_data(rng) -> bool:
    return any(v is not None for row in rng.options(ndim=2).value for v in row)


def same_color(color, rgb) -> bool:
    return color is not None and tuple(color) == rgb


# ---------------- preflight / cleanup / integrity ----------------
def preflight() -> str | None:
    if xw.apps.count == 0:
        return "Excel open-a illa."
    app = xw.apps.active
    book = next((b for b in app.books if b.name == WORKBOOK), None)
    if book is None:
        return f"'{WORKBOOK}' open-la illa. Open workbooks: {[b.name for b in app.books]}"
    names = [s.name for s in book.sheets]
    if SHEET not in names:
        return f"Sheet '{SHEET}' illa. Sheets: {names}"
    clash = [n for n in TEMP_SHEETS if n in names]
    if clash:
        return f"Indha sheets already irukku, test overwrite pannidum, so stop: {clash}"
    sheet = book.sheets[SHEET]
    df0 = sheet.range("A1:G101").options(pd.DataFrame, header=1, index=False).value
    if list(df0.columns) != EXPECTED_COLS or len(df0) != 100:
        return f"Dataset expected maadhiri illa: columns={list(df0.columns)}, rows={len(df0)}"
    if has_data(sheet.range("A102:G200")):
        return "Row 102-ku keezha data irukku. Test stop."
    if has_data(sheet.range(SCRATCH)):
        return f"{SCRATCH} area-la data irukku. Test stop."
    S.update(
        book=book, df0=df0,
        values0=sheet.range("A1:G101").options(ndim=2).value,
        names0=names, charts0=len(sheet.charts),
        widths0={c: sheet.range(f"{c}1").column_width for c in "ABCDEFG"},
        mess=False,
    )
    return None


def cleanup() -> None:
    book, sheet = S["book"], sh()
    app = book.app
    app.display_alerts = False
    try:
        for name in TEMP_SHEETS:
            if name in sheet_names():
                book.sheets[name].delete()
        # delete from the END so indexes do not shift
        for chart in reversed(list(sheet.charts)[S["charts0"]:]):
            chart.delete()
        sheet.range(SCRATCH).clear()
        if has_data(sheet.range("A102:G200")):
            sheet.range("102:200").delete()
        for col, width in S["widths0"].items():
            sheet.range(f"{col}1").column_width = width
    finally:
        app.display_alerts = True


def dataset_intact() -> tuple[bool, list[str]]:
    sheet, problems = sh(), []
    if sheet.range("A1:G101").options(ndim=2).value != S["values0"]:
        problems.append("A1:G101 values maariduchu")
    if sheet_names() != S["names0"]:
        problems.append(f"sheets maariduchu: {sheet_names()}")
    if len(sheet.charts) != S["charts0"]:
        problems.append("chart count maariduchu")
    if has_data(sheet.range("A102:G110")):
        problems.append("row 102-ku keezha data irukku")
    return not problems, problems


# ---------------- tests ----------------
def test_get_sheet_names():
    r = ok(call(get_excel_sheet_names))
    check(r["workbook"] == WORKBOOK and SHEET in r["sheets"], f"result: {r}")
    return f"sheets={r['sheets']}"


def test_read_data():
    r = ok(call(read_excel_data, max_rows=5))
    check(r["columns"] == EXPECTED_COLS, f"columns: {r['columns']}")
    check(r["total_rows"] == 100 and r["returned_rows"] == 5 and len(r["rows"]) == 5, f"rows: {r['total_rows']}")
    fails(call(read_excel_data, sheet_name="NoSuchSheet"))
    fails(call(read_excel_data, max_rows=0))
    return "100 rows, 7 columns, max_rows limit OK"


def test_analyze_column():
    df0 = S["df0"]
    r = ok(call(analyze_excel_column, column="revenue"))
    check(r["type"] == "numeric" and abs(r["sum"] - df0["revenue"].sum()) < 0.01, f"revenue: {r}")
    u = ok(call(analyze_excel_column, column="units"))
    check(u["min"] == df0["units"].min() and u["max"] == df0["units"].max(), f"units: {u}")
    for col, expected in (("region", 5), ("product", 8), ("channel", 3)):
        t = ok(call(analyze_excel_column, column=col))
        check(t["type"] == "text" and t["unique"] == expected, f"{col}: {t}")
    d = ok(call(analyze_excel_column, column="date"))
    fails(call(analyze_excel_column, column="nope"))
    return f"revenue sum={round(r['sum'], 2)}, rukku type={d['type']}"


def test_validate_data():
    r = ok(call(validate_excel_data))
    check(r["rows"] == 100 and r["is_clean"] is True, f"validate: {r}")
    return "100 rows, clean"


def test_filter_data():
    df0 = S["df0"]
    exp = int((df0["region"] == "EMEA").sum())
    r = ok(call(filter_excel_data, column="region", condition="equals", value="EMEA"))
    check(r["matched_rows"] == exp, f"matched={r['matched_rows']} expected={exp}")
    back = ok(call(read_excel_data, sheet_name="Filtered_Data", max_rows=1000))
    check(back["total_rows"] == exp and all(x["region"] == "EMEA" for x in back["rows"]), "Filtered_Data rows thappu")
    exp2 = int((pd.to_numeric(df0["units"]) > 30).sum())
    r2 = ok(call(filter_excel_data, column="units", condition="greater_than", value=30))
    check(r2["matched_rows"] == exp2, f"units>30: {r2['matched_rows']} expected={exp2}")
    exp3 = int(df0["product"].str.contains("usb", case=False).sum())
    r3 = ok(call(filter_excel_data, column="product", condition="contains", value="usb"))
    check(r3["matched_rows"] == exp3, f"contains usb: {r3['matched_rows']} expected={exp3}")
    none = ok(call(filter_excel_data, column="region", condition="equals", value="Mars"))
    check(none["matched_rows"] == 0, "no-match case")
    fails(call(filter_excel_data, column="region", condition="bigger", value="x"))
    fails(call(filter_excel_data, column="nope", condition="equals", value="x"))
    fails(call(filter_excel_data, column="region", condition="equals", value="EMEA", output_sheet=SHEET))
    return f"EMEA={exp}, units>30={exp2}, usb={exp3}"


def test_clear_filter():
    if "Filtered_Data" not in sheet_names():
        ok(call(filter_excel_data, column="region", condition="equals", value="EMEA"))
    r = ok(call(clear_excel_filter))
    check(r["removed"] is True and "Filtered_Data" not in sheet_names(), "Filtered_Data innum irukku")
    again = ok(call(clear_excel_filter))
    check(again["removed"] is False, "illaadha sheet-ku removed=False varanum")
    fails(call(clear_excel_filter, output_sheet=SHEET))
    check(SHEET in sheet_names(), "DATASET SHEET delete aayiduchu!")
    return "Filtered_Data removed, dataset sheet protected"


def test_pivot_summary():
    df0 = S["df0"]
    exp = df0.groupby("region")["revenue"].sum()
    r = ok(call(create_excel_pivot_summary, group_by="region", value_column="revenue", aggregation="sum"))
    back = ok(call(read_excel_data, sheet_name="Pivot_Summary", max_rows=100))
    got = {x["region"]: x["Sum of revenue"] for x in back["rows"]}
    check(set(got) == set(exp.index), f"regions: {sorted(got)}")
    check(all(abs(got[k] - exp[k]) < 0.01 for k in exp.index), f"sums: {got}")
    expc = df0.groupby("product")["revenue"].count()
    ok(call(create_excel_pivot_summary, group_by="product", value_column="revenue", aggregation="count"))
    back2 = ok(call(read_excel_data, sheet_name="Pivot_Summary", max_rows=100))
    gotc = {x["product"]: x["Count of revenue"] for x in back2["rows"]}
    check(gotc == {k: int(v) for k, v in expc.items()}, f"counts: {gotc}")
    fails(call(create_excel_pivot_summary, group_by="region", value_column="revenue", aggregation="median"))
    fails(call(create_excel_pivot_summary, group_by="region", value_column="region"))
    fails(call(create_excel_pivot_summary, group_by="region", value_column="revenue", output_sheet=SHEET))
    return f"{r['groups']} regions match pandas, product counts match"


def test_highlight_values():
    df0 = S["df0"]
    mask = pd.to_numeric(df0["units"], errors="coerce") > 35
    idx = [int(i) for i in mask[mask].index]
    cells = [sh().range((i + 2, 5)) for i in idx]
    original = [c.color for c in cells]
    try:
        r = ok(call(highlight_excel_values, column="units", condition="greater_than", value=35, color="#FFFF00"))
        check(r["highlighted_cells"] == len(idx), f"highlighted={r['highlighted_cells']} expected={len(idx)}")
        check(same_color(cells[0].color, (255, 255, 0)), f"first match color: {cells[0].color}")
        other = next(i for i in df0.index if i not in idx)
        check(not same_color(sh().range((other + 2, 5)).color, (255, 255, 0)), "non-match cell-um yellow aachu")
        empty = ok(call(highlight_excel_values, column="units", condition="is_empty"))
        check(empty["highlighted_cells"] == 0, f"is_empty: {empty}")
        fails(call(highlight_excel_values, column="units", condition="bigger", value=1))
        fails(call(highlight_excel_values, column="nope", condition="equals", value=1))
        fails(call(highlight_excel_values, column="units", condition="equals", value=1, color="yellow"))
    finally:
        for cell, color in zip(cells, original):
            cell.color = color
    return f"{len(idx)} cells highlighted (restored after)"


def test_append_data():
    sheet, values0 = sh(), S["values0"]
    vals = sheet.range("A2:G3").options(ndim=2).value
    if isinstance(values0[1][0], str):
        sheet.range("A102:A104").number_format = "@"   # dataset dates text-a irundha mattum
    rows = [vals[0], [None] * 7, vals[1]]                  # duplicate, blank row, duplicate
    r = ok(call(append_excel_data, rows=rows))
    S["mess"] = True
    check("102" in r["appended_range"] and "104" in r["appended_range"], f"range: {r['appended_range']}")
    check(sheet.range("A102").value == values0[1][0] and sheet.range("G104").value == values0[2][6], "data thappaana idathula")
    check(sheet.range("A103").value is None, "blank row blank-a illa")
    fails(call(append_excel_data, rows=[[1, 2], [3]]))
    fails(call(append_excel_data, rows=[[1] * 9]))
    v = ok(call(validate_excel_data))
    types = {i["type"] for i in v["issues"]}
    check({"duplicate_rows", "blank_rows"} <= types, f"validator issues: {types}")
    return f"added rows 102-104, validator now reports {sorted(types)}"


def test_remove_blank_rows():
    if not S["mess"]:
        raise SkipTest("append test pass aagala")
    sheet, values0 = sh(), S["values0"]
    p = ok(call(remove_excel_blank_rows))
    check(p["found"] == 1 and p["deleted"] is False and p["excel_rows"] == [103], f"preview: {p}")
    check(sheet.range("A104").value is not None, "preview-la data maariduchu!")
    r = ok(call(remove_excel_blank_rows, confirm=True))
    check(r["deleted"] and r["removed"] == 1, f"delete: {r}")
    check(r["backup"] and Path(r["backup"]).exists(), f"backup illa: {r['backup']}")
    check(sheet.range("A103").value == values0[2][0] and sheet.range("A104").value is None, "row delete sariyaa aagala")
    again = ok(call(remove_excel_blank_rows))
    check(again["found"] == 0, "innum blank rows irukku")
    return "preview OK, 1 blank row deleted, backup OK"


def test_remove_duplicates():
    if not S["mess"]:
        raise SkipTest("append test pass aagala")
    sheet, values0, df0 = sh(), S["values0"], S["df0"]
    p = ok(call(remove_excel_duplicates))
    check(p["found"] == 2 and p["deleted"] is False and p["excel_rows"] == [102, 103], f"preview: {p}")
    expected = (len(df0) + 2) - df0["region"].nunique()
    pc = ok(call(remove_excel_duplicates, columns=["region"]))
    check(pc["found"] == expected and pc["deleted"] is False, f"columns=['region']: {pc['found']} expected={expected}")
    check(sheet.range("A103").value is not None, "preview-la data maariduchu!")
    fails(call(remove_excel_duplicates, columns=["nope"]))
    fails(call(remove_excel_duplicates, keep="middle"))
    r = ok(call(remove_excel_duplicates, confirm=True))
    check(r["deleted"] and r["removed"] == 2, f"delete: {r}")
    check(r["backup"] and Path(r["backup"]).exists(), f"backup illa: {r['backup']}")
    check(sheet.range("A102").value is None and sheet.range("A101").value == values0[100][0], "row delete sariyaa aagala")
    again = ok(call(remove_excel_duplicates))
    check(again["found"] == 0, "innum duplicates irukku")
    S["mess"] = False
    return "preview OK, 2 duplicates deleted, backup OK, dataset back to 100 rows"


def test_write_data():
    sheet = sh()
    data = [["Item", "Value"], ["A", 1], ["B", 2]]
    r = ok(call(write_excel_data, data=data, start_cell="I1"))
    check(sheet.range("I1").value == "Item" and sheet.range("J3").value == 2, "Excel-la data ezhudhala")
    fails(call(write_excel_data, data=data, start_cell="I1"))
    r2 = ok(call(write_excel_data, data=data, start_cell="I1", overwrite=True))
    check(r2["backup"] and Path(r2["backup"]).exists(), f"backup illa: {r2['backup']}")
    fails(call(write_excel_data, data="abc", start_cell="I1"))
    fails(call(write_excel_data, data=[[1, 2], [3]], start_cell="L5"))
    fails(call(write_excel_data, data=[["x"]], start_cell="I1:J2"))
    return f"written {r['written_range']}, overwrite protection OK, backup OK"


def test_apply_formula():
    sheet = sh()
    expected = float(S["df0"]["revenue"].sum())
    r = ok(call(apply_excel_formula, cell="L1", formula="=SUM(G2:G101)"))
    check(abs(r["result"] - expected) < 0.01, f"Excel SUM={r['result']} pandas={expected}")
    check(sheet.range("L1").formula == "=SUM(G2:G101)", f"formula text: {sheet.range('L1').formula}")
    err = call(apply_excel_formula, cell="L2", formula="=1/0")
    check(err["success"] is False, f"=1/0 error detect aagala: {err}")
    fails(call(apply_excel_formula, cell="L1", formula="=1+1"))
    fails(call(apply_excel_formula, cell="L3", formula="SUM(1)"))
    fails(call(apply_excel_formula, cell="ZZZZ", formula="=1"))
    return f"Excel SUM = pandas sum = {round(r['result'], 2)}, error detection OK"


def test_apply_formula_range():
    sheet = sh()
    r = ok(call(apply_excel_formula_range, target_range="N2:N4", formula="=E2*F2"))
    got = sheet.range("N2:N4").options(ndim=1).value
    exp = sheet.range("G2:G4").options(ndim=1).value
    check(all(abs(a - b) < 0.01 for a, b in zip(got, exp)), f"relative fill: {got} vs revenue {exp}")
    check(r["cells_filled"] == 3 and r["error_cells"] == 0, f"result: {r}")
    fails(call(apply_excel_formula_range, target_range="N2:N4", formula="=E2*F2"))
    return f"N2:N4 = {got} (matches revenue)"


def test_format_cells():
    sheet = sh()
    ok(call(format_excel_cells, cell_range="I1:J1", bold=True, font_color="#FFFFFF", background_color="#1F4E78"))
    check(bool(sheet.range("I1").font.bold), "I1 bold aagala")
    check(same_color(sheet.range("I1").color, (31, 78, 120)), f"fill color: {sheet.range('I1').color}")
    ok(call(format_excel_cells, cell_range="J2:J3", number_format="#,##0.00", font_size=12, italic=True))
    check(sheet.range("J2").number_format == "#,##0.00", f"number_format: {sheet.range('J2').number_format}")
    fails(call(format_excel_cells, cell_range="I1"))
    fails(call(format_excel_cells, cell_range="I1", font_color="red"))
    fails(call(format_excel_cells, cell_range="I1", font_size=0))
    fails(call(format_excel_cells, cell_range="not-a-range", bold=True))
    return "bold, fill color, number_format OK"


def test_autofit_columns():
    cell = sh().range("C1")
    cell.column_width = 2
    ok(call(autofit_excel_columns, cell_range="A1:G101"))
    check(cell.column_width > 4, f"autofit aagala, width={cell.column_width}")
    ok(call(autofit_excel_columns))
    fails(call(autofit_excel_columns, cell_range="ZZZ"))
    return f"column C width 2 -> {round(cell.column_width, 1)} (restored after)"


def test_create_chart():
    sheet, before = sh(), len(sh().charts)
    r = ok(call(create_excel_chart, data_range="G1:G11", chart_type="column", title="Test Column", position_cell="I10"))
    check(len(sheet.charts) == before + 1, "column chart sheet-la illa")
    ok(call(create_excel_chart, data_range="G1:G11", chart_type="line", title="Test Line", position_cell="I28"))
    check(len(sheet.charts) == before + 2, "line chart sheet-la illa")
    fails(call(create_excel_chart, data_range="G1:G11", chart_type="donut"))
    fails(call(create_excel_chart, data_range="nope", chart_type="column"))
    fails(call(create_excel_chart, data_range="G1:G11", width=5))
    return f"column + line charts created (removed after), first name={r['name']}"


TESTS = [
    ("get_excel_sheet_names", test_get_sheet_names),
    ("read_excel_data", test_read_data),
    ("analyze_excel_column", test_analyze_column),
    ("validate_excel_data", test_validate_data),
    ("filter_excel_data", test_filter_data),
    ("clear_excel_filter", test_clear_filter),
    ("create_excel_pivot_summary", test_pivot_summary),
    ("highlight_excel_values", test_highlight_values),
    ("append_excel_data", test_append_data),
    ("remove_excel_blank_rows", test_remove_blank_rows),
    ("remove_excel_duplicates", test_remove_duplicates),
    ("write_excel_data", test_write_data),
    ("apply_excel_formula", test_apply_formula),
    ("apply_excel_formula_range", test_apply_formula_range),
    ("format_excel_cells", test_format_cells),
    ("autofit_excel_columns", test_autofit_columns),
    ("create_excel_chart", test_create_chart),
]


def run_one(func) -> tuple[str, str]:
    try:
        return "PASS", func()
    except SkipTest as exc:
        return "SKIP", str(exc)
    except AssertionError as exc:
        return "FAIL", str(exc)
    except Exception as exc:
        return "ERROR", f"{type(exc).__name__}: {exc}"


def main() -> int:
    words = [a.lower() for a in sys.argv[1:]]
    names = {n for n, _ in TESTS if not words or any(w in n for w in words)}
    if not names:
        print("Match aana tool illa. Available:", ", ".join(n for n, _ in TESTS))
        return 1
    if any(n.startswith("remove_") for n in names):
        names.add("append_excel_data")          # remove_* tests need the rows that append adds
    selected = [(n, f) for n, f in TESTS if n in names]

    print(f"xlwings {xw.__version__} | pandas {pd.__version__}")
    problem = preflight()
    if problem:
        print(f"PREFLIGHT FAILED (dataset-ai thodala): {problem}")
        return 1
    print(f"Dataset: {WORKBOOK} / {SHEET} (100 rows). Tests thodangudhu, Excel-la click pannaadhinga...\n")

    results = []
    try:
        for number, (name, func) in enumerate(selected, 1):
            status, detail = run_one(func)
            results.append((name, status, detail))
            print(f"[{number:>2}/{len(selected)}] {status:<5} {name:<28} {detail}")
    finally:
        try:
            cleanup()
        except Exception as exc:
            print(f"CLEANUP ERROR: {type(exc).__name__}: {exc}")

    intact, problems = dataset_intact()

    passed = sum(1 for r in results if r[1] == "PASS")
    print("\n" + "=" * 64)
    print(f"RESULT: {passed}/{len(results)} passed")
    for name, status, detail in results:
        if status != "PASS":
            print(f"  - {name} [{status}]: {detail}")
    if intact:
        print("DATASET CHECK: PASS (values, sheets, charts ellaam test-ku munnadi maadhiri irukku)")
    else:
        print("DATASET CHANGED:", "; ".join(problems))
    print("=" * 64)
    return 0 if passed == len(results) and intact else 1


if __name__ == "__main__":
    sys.exit(main())