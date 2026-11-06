import pandas as pd
import numpy as np

from excel_tools.workbook_manager import (
    ExcelToolError,
    get_sheet,
    load_dataframe,
    safe_tool,
)

MAX_SUMMARY_COLUMNS = 30
NUMERIC_SHARE = 0.8  # a column is "numeric" when at least 80% of its non-empty cells are numbers

OPERATIONS = {
    "sum": "sum", "total": "sum",
    "average": "mean", "avg": "mean", "mean": "mean",
    "count": "count", "min": "min", "max": "max",
}


# ---------- helpers ----------

def _find_col(df, name, label: str = "Column"):
    """Case-insensitive column lookup."""
    key = str(name).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == key:
            return col
    raise ExcelToolError(f"{label} '{name}' not found. Available: {list(df.columns)}")


def _non_empty(series):
    """Series without None, NaN and blank-text cells."""
    s = series
    if s.dtype == object:
        s = s.where(~s.astype(str).str.strip().eq(""), np.nan)
    return s.dropna()


def _num(value):
    return None if pd.isna(value) else round(float(value), 4)


def _text(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _profile(series) -> dict:
    """Short profile of one column (used by both modes)."""
    total = int(len(series))
    values = _non_empty(series)
    nonempty = int(len(values))
    numeric = pd.to_numeric(values, errors="coerce")
    numeric_count = int(numeric.notna().sum())
    info = {
        "total_rows": total,
        "non_empty_count": nonempty,
        "missing_count": total - nonempty,
        "unique_count": int(values.nunique()),
    }
    if nonempty and numeric_count / nonempty >= NUMERIC_SHARE:
        clean = numeric.dropna()
        info.update({
            "data_type": "numeric",
            "numeric_count": numeric_count,
            "sum": _num(clean.sum()),
            "minimum": _num(clean.min()),
            "maximum": _num(clean.max()),
            "average": _num(clean.mean()),
            "median": _num(clean.median()),
            "standard_deviation": _num(clean.std()) if len(clean) > 1 else 0.0,
        })
        if numeric_count < nonempty:
            info["non_numeric_values"] = nonempty - numeric_count
    else:
        info["data_type"] = "categorical"
        if nonempty:
            counts = values.astype(str).value_counts()
            info["top_value"] = str(counts.index[0])
            info["top_value_count"] = int(counts.iloc[0])
    return info


# ---------- tools ----------

@safe_tool
def analyze_excel_column(column: str | None = None, workbook_name: str | None = None,
                         sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Analyze one column (count, missing, min, max, average, median, distribution, insights), or give a summary of every column when no column is given. Read-only.

    PARAMETERS:
        column: Column header name, not case-sensitive (optional, without it the whole sheet is summarized).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, column, data_type, counts, statistics, distribution, insights (or columns summary).
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "message": "No data found."}

    if column is None:
        cols = list(df.columns)[:MAX_SUMMARY_COLUMNS]
        summary = []
        for col in cols:
            item = {"column": str(col)}
            item.update(_profile(df[col]))
            item.pop("total_rows", None)
            summary.append(item)
        out = {"success": True, "sheet": sheet.name, "total_rows": int(len(df)),
               "total_columns": int(len(df.columns)), "columns": summary}
        if len(df.columns) > MAX_SUMMARY_COLUMNS:
            out["note"] = f"Only the first {MAX_SUMMARY_COLUMNS} columns are summarized."
        return out

    col = _find_col(df, column)
    result = {"success": True, "column": str(col)}
    result.update(_profile(df[col]))

    values = _non_empty(df[col])
    insights = []
    if result["data_type"] == "numeric":
        insights.append(
            f"Values range from {result['minimum']} to {result['maximum']}; "
            f"average {result['average']}, median {result['median']}.")
    else:
        result["distribution"] = {str(k): int(v) for k, v in
                                  values.astype(str).value_counts().head(20).items()}
        insights.append(f"Column has {result['unique_count']} unique categories.")
    if result["missing_count"]:
        pct = result["missing_count"] / max(result["total_rows"], 1) * 100
        insights.append(f"{result['missing_count']} values are missing ({pct:.1f}%).")
    result["insights"] = insights
    return result


@safe_tool
def calculate_excel_total(column: str, operation: str = "sum", group_by: str | None = None,
                          workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Calculate sum, average, count, min or max of a column, optionally grouped by another column. Read-only (does not write to the sheet).

    PARAMETERS:
        column: Column header name, not case-sensitive.
        operation: sum, average, count, min or max.
        group_by: Column to group by (optional).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, column, operation, result (or results per group).
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    col = _find_col(df, column)

    op = str(operation).lower().strip()
    if op not in OPERATIONS:
        raise ExcelToolError("operation must be sum, average, count, min or max.")

    numeric = pd.to_numeric(df[col], errors="coerce")
    if not numeric.notna().any() and OPERATIONS[op] != "count":
        raise ExcelToolError(f"Column '{col}' has no numeric values.")

    if group_by is not None:
        gcol = _find_col(df, group_by, "Group column")
        temp = pd.DataFrame({"group": df[gcol], "value": numeric})
        grouped = temp.groupby("group", dropna=False)["value"].agg(OPERATIONS[op])
        results = {str(k): _num(v) for k, v in grouped.items()}
        return {"success": True, "column": str(col), "operation": op,
                "group_by": str(gcol), "results": results}

    if OPERATIONS[op] == "count":
        result = int(_non_empty(df[col]).count())
    else:
        result = float(getattr(numeric, OPERATIONS[op])())
    return {"success": True, "column": str(col), "operation": op, "result": result}


@safe_tool
def identify_excel_outliers(column: str, method: str = "iqr", threshold: float = 1.5,
                            workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Find numeric outliers in a column using IQR or z-score. Read-only.

    PARAMETERS:
        column: Column header name, not case-sensitive.
        method: iqr or zscore.
        threshold: IQR multiplier (default 1.5) or z-score limit (example 3).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, column, method, total_values, outlier_count, outliers (excel_row, value), details.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    col = _find_col(df, column)

    method = str(method).lower().strip()
    if method not in {"iqr", "zscore"}:
        raise ExcelToolError("method must be iqr or zscore.")
    if threshold <= 0:
        raise ExcelToolError("threshold must be greater than zero.")

    numeric = pd.to_numeric(df[col], errors="coerce")
    clean = numeric.dropna()
    if len(clean) < 3:
        raise ExcelToolError("At least 3 numeric values are needed for outlier detection.")

    if method == "iqr":
        q1 = float(clean.quantile(0.25))
        q3 = float(clean.quantile(0.75))
        iqr = q3 - q1
        lower = q1 - threshold * iqr
        upper = q3 + threshold * iqr
        mask = (numeric < lower) | (numeric > upper)
        details = {"q1": q1, "q3": q3, "iqr": float(iqr),
                   "lower_bound": float(lower), "upper_bound": float(upper)}
    else:
        mean = float(clean.mean())
        std = float(clean.std())
        if std == 0:
            mask = pd.Series(False, index=df.index)
        else:
            mask = ((numeric - mean) / std).abs() > threshold
        details = {"mean": mean, "standard_deviation": std, "zscore_threshold": threshold}

    outliers = [{"excel_row": int(used.row + 1 + i), "value": float(numeric.loc[i])}
                for i in df.index[mask.fillna(False)]]
    return {"success": True, "column": str(col), "method": method,
            "total_values": int(len(clean)), "outlier_count": len(outliers),
            "outliers": outliers[:100], "details": details}


@safe_tool
def compare_excel_datasets(compare_sheet_name: str, key_column: str | None = None,
                           workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Compare two sheets in the same workbook: row counts and columns, and (with key_column) added, removed and changed records. Read-only.

    PARAMETERS:
        compare_sheet_name: The second sheet to compare against.
        key_column: Column that identifies a record, example ID (optional).
        workbook_name: Workbook name (optional).
        sheet_name: First sheet name (optional, default active sheet).

    RETURNS:
        success, rows, columns_only_in_sheet_1/2, keys_only_in_sheet_1/2, changed_records, counts.
    """
    if not compare_sheet_name:
        raise ExcelToolError("compare_sheet_name is required.")

    first = get_sheet(workbook_name, sheet_name)
    second = get_sheet(workbook_name, compare_sheet_name)
    if first.name == second.name:
        raise ExcelToolError("Two different sheets are needed to compare.")

    df1, _ = load_dataframe(first)
    df2, _ = load_dataframe(second)
    cols1, cols2 = list(df1.columns), list(df2.columns)
    common = [c for c in cols1 if c in cols2]

    result = {
        "success": True,
        "sheet_1": first.name, "sheet_2": second.name,
        "rows_sheet_1": int(len(df1)), "rows_sheet_2": int(len(df2)),
        "common_columns": common,
        "columns_only_in_sheet_1": [c for c in cols1 if c not in cols2],
        "columns_only_in_sheet_2": [c for c in cols2 if c not in cols1],
    }
    if key_column is None:
        return result

    k1 = _find_col(df1, key_column, "Key column")
    k2 = _find_col(df2, key_column, "Key column")
    if df1[k1].duplicated().any():
        raise ExcelToolError(f"Duplicate keys found in {first.name}.")
    if df2[k2].duplicated().any():
        raise ExcelToolError(f"Duplicate keys found in {second.name}.")

    a, b = df1.set_index(k1), df2.set_index(k2)
    keys1, keys2 = set(a.index), set(b.index)
    only_first, only_second = keys1 - keys2, keys2 - keys1
    common_keys = keys1 & keys2
    comparable = [c for c in common if c not in (k1, k2)]

    changed = []
    for key in common_keys:
        diffs = {}
        for c in comparable:
            v1, v2 = a.at[key, c], b.at[key, c]
            n1, n2 = pd.isna(v1), pd.isna(v2)
            if n1 and n2:
                continue
            if n1 != n2 or v1 != v2:
                diffs[str(c)] = {"sheet_1": None if n1 else str(v1),
                                 "sheet_2": None if n2 else str(v2)}
        if diffs:
            changed.append({"key": str(key), "differences": diffs})

    result.update({
        "key_column": str(k1),
        "keys_only_in_sheet_1": [str(k) for k in list(only_first)[:100]],
        "keys_only_in_sheet_2": [str(k) for k in list(only_second)[:100]],
        "keys_only_in_sheet_1_count": len(only_first),
        "keys_only_in_sheet_2_count": len(only_second),
        "changed_records": changed[:100],
        "changed_records_count": len(changed),
        "unchanged_records_count": len(common_keys) - len(changed),
    })
    return result


@safe_tool
def value_counts_excel_column(column: str, value: str | float | None = None, top: int = 20,
                              workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Count how many times each value appears in a column, or count one specific value (example: how many Male). Read-only.

    PARAMETERS:
        column: Column header name, not case-sensitive.
        value: Count only this value, not case-sensitive (optional).
        top: How many top values to return, 1-100 (default 20).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, column, total_non_empty, blank_count, counts (or value, count, percent).
    """
    if not 1 <= int(top) <= 100:
        raise ExcelToolError("top must be between 1 and 100.")
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    col = _find_col(df, column)

    values = _non_empty(df[col])
    total = int(len(values))
    blanks = int(len(df) - total)
    texts = values.map(_text)

    if value is not None:
        wanted = _text(value).lower()
        count = int((texts.str.lower() == wanted).sum())
        percent = round(count / total * 100, 2) if total else 0.0
        return {"success": True, "column": str(col), "value": str(value), "count": count,
                "percent": percent, "total_non_empty": total, "blank_count": blanks}

    counts = texts.value_counts()
    items = [{"value": str(k), "count": int(v),
              "percent": round(int(v) / total * 100, 2) if total else 0.0}
             for k, v in counts.head(int(top)).items()]
    out = {"success": True, "column": str(col), "unique_count": int(counts.size),
           "total_non_empty": total, "blank_count": blanks, "counts": items}
    if counts.size > int(top):
        out["note"] = f"Showing top {int(top)} of {int(counts.size)} values."
    return out