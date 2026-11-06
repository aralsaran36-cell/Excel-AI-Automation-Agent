
"""Shared helpers used by every excel_tools file."""
import functools
import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

BACKUP_DIR = Path(__file__).resolve().parent.parent / "backups"
ADDRESS_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$")
EXCEL_ERRORS = {"#NAME?", "#DIV/0!", "#VALUE!", "#REF!", "#N/A", "#NUM!", "#NULL!",
                "#SPILL!", "#CALC!"}
CONDITIONS = (
    "equals", "not_equals", "greater_than", "less_than",
    "greater_or_equal", "less_or_equal", "contains", "is_empty",
)


class ExcelToolError(Exception):
    """Expected, user-facing error (bad input, missing sheet, etc.)."""


def safe_tool(func):
    """Wrap a tool so it always returns a dict instead of raising."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except ExcelToolError as exc:
            return {"success": False, "error": str(exc)}
        except Exception as exc:  # unexpected
            return {"success": False, "error": f"{type(exc).__name__}: {exc}"}
    return wrapper


# ---------- workbook / sheet access ----------
def get_workbook(workbook_name: str | None = None):
    if xw.apps.count == 0:
        raise ExcelToolError("Excel open-a illa. Excel Desktop-la oru workbook open pannunga.")
    app = xw.apps.active
    if workbook_name:
        for book in app.books:
            if book.name == workbook_name:
                return book
        open_names = [b.name for b in app.books]
        raise ExcelToolError(f"Workbook '{workbook_name}' open-la illa. Open workbooks: {open_names}")
    if app.books.count == 0:
        raise ExcelToolError("Excel-la workbook edhuvum open-a illa.")
    return app.books.active


def get_sheet(workbook_name: str | None = None, sheet_name: str | None = None):
    book = get_workbook(workbook_name)
    if sheet_name:
        names = [s.name for s in book.sheets]
        if sheet_name not in names:
            raise ExcelToolError(f"Sheet '{sheet_name}' illa. Available sheets: {names}")
        return book.sheets[sheet_name]
    return book.sheets.active


# ---------- validation ----------
def validate_address(address: str) -> str:
    if not isinstance(address, str) or not ADDRESS_RE.match(address.strip()):
        raise ExcelToolError(f"Invalid cell/range '{address}'. Example: 'A1' or 'A1:C10'.")
    return address.strip()


def hex_to_rgb(color: str) -> tuple[int, int, int]:
    value = color.strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        raise ExcelToolError(f"Invalid color '{color}'. Use hex like '#FFFF00'.")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


# ---------- reading data ----------
def load_dataframe(sheet):
    """Read the sheet's used range into a DataFrame (first row = header)."""
    used = sheet.used_range
    if used.count == 1 and used.value is None:
        raise ExcelToolError(f"Sheet '{sheet.name}' empty-a irukku.")
    df = used.options(pd.DataFrame, header=1, index=False).value
    names = []
    for i, col in enumerate(df.columns):
        if col is None or (isinstance(col, float) and pd.isna(col)):
            names.append(f"Unnamed_{i + 1}")
        elif isinstance(col, float) and col.is_integer():
            names.append(str(int(col)))
        else:
            names.append(str(col))
    df.columns = names
    return df, used


def df_to_records(df: pd.DataFrame, limit: int = 50) -> list[dict]:
    return json.loads(df.head(limit).to_json(orient="records", date_format="iso"))


# ---------- row helpers ----------
def contiguous_blocks(numbers) -> list[tuple[int, int]]:
    """[2,3,4,9] -> [(2,4),(9,9)]"""
    blocks: list[list[int]] = []
    for n in sorted(set(numbers)):
        if blocks and n == blocks[-1][1] + 1:
            blocks[-1][1] = n
        else:
            blocks.append([n, n])
    return [(a, b) for a, b in blocks]


def delete_rows(sheet, row_numbers) -> int:
    """Delete whole rows, bottom-up, so row numbers stay valid. Keeps formatting/formulas."""
    blocks = contiguous_blocks(row_numbers)
    for start, end in reversed(blocks):
        sheet.range(f"{start}:{end}").delete()
    return len(blocks)


def build_condition_mask(series: pd.Series, condition: str, value=None) -> pd.Series:
    """Shared by filter_excel_data and highlight_excel_values."""
    if condition not in CONDITIONS:
        raise ExcelToolError(f"Invalid condition '{condition}'. Use one of: {list(CONDITIONS)}")
    if condition == "is_empty":
        return series.isna()
    if value is None:
        raise ExcelToolError(f"Condition '{condition}'-ku 'value' thevai.")
    if condition == "contains":
        return series.notna() & series.astype(str).str.contains(str(value), case=False, regex=False)
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = None
    numeric = pd.to_numeric(series, errors="coerce")
    if condition in ("equals", "not_equals"):
        if number is not None:
            mask = numeric == number
        else:
            mask = series.astype(str).str.strip().str.lower() == str(value).strip().lower()
        return mask if condition == "equals" else ~mask
    if number is None:
        raise ExcelToolError(f"Condition '{condition}'-ku numeric value thevai.")
    return {
        "greater_than": numeric > number,
        "less_than": numeric < number,
        "greater_or_equal": numeric >= number,
        "less_or_equal": numeric <= number,
    }[condition]


# ---------- safety ----------
def backup_workbook(book) -> str | None:
    """Copy the saved file on disk into backups/. Returns path, or None if not saved yet."""
    try:
        path = Path(book.fullname)
    except Exception:
        return None
    if not path.exists():
        return None
    BACKUP_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = BACKUP_DIR / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, destination)
    return str(destination)


def write_dataframe_to_sheet(book, sheet_name: str, df: pd.DataFrame, allowed_prefix: str):
    """Write df into a helper sheet (created or cleared). Never touches non-helper sheets."""
    if not sheet_name.startswith(allowed_prefix) or len(sheet_name) > 31:
        raise ExcelToolError(f"Output sheet name '{allowed_prefix}'-la start aagi 31 characters-kulla irukkanum.")
    if sheet_name in [s.name for s in book.sheets]:
        out = book.sheets[sheet_name]
        out.clear_contents()
    else:
        out = book.sheets.add(name=sheet_name, after=book.sheets[len(book.sheets) - 1])
    out.range("A1").options(index=False, header=True).value = df
    return out