"""Knowledge tool: ask_knowledge_base (explain Excel formulas and concepts).

Stage-10 version: built-in curated entries + keyword matching. NOT RAG yet.
Stage 11 can replace _search() with real retrieval; the return format stays the same.
Read-only, never touches Excel. Returns {"success": bool, ...} (via @safe_tool).
"""
import re

from excel_tools.workbook_manager import ExcelToolError, safe_tool

MAX_RESULTS = 5

# name, category, keywords, syntax, explanation, example
ENTRIES: list[dict] = [
    {"name": "SUM", "category": "formula", "keywords": ["total", "add", "sum", "addition"],
     "syntax": "=SUM(number1, [number2], ...)",
     "explain": "Cells-oda values-ah kootti total tharum.", "example": "=SUM(G2:G101)"},
    {"name": "AVERAGE", "category": "formula", "keywords": ["average", "mean"],
     "syntax": "=AVERAGE(number1, [number2], ...)",
     "explain": "Numbers-oda average tharum. Empty/text cells ignore aagum.", "example": "=AVERAGE(G2:G101)"},
    {"name": "COUNT", "category": "formula", "keywords": ["count", "numbers", "how many"],
     "syntax": "=COUNT(value1, [value2], ...)",
     "explain": "Numbers irukkura cells ethanai nu ennum. Text-ku COUNTA use pannunga.", "example": "=COUNT(E2:E101)"},
    {"name": "COUNTIF", "category": "formula", "keywords": ["countif", "count if", "condition count", "criteria"],
     "syntax": "=COUNTIF(range, criteria)",
     "explain": "Condition match aagura cells-ah ennum.", "example": '=COUNTIF(B2:B101,"EMEA")'},
    {"name": "SUMIF", "category": "formula", "keywords": ["sumif", "sum if", "conditional sum", "criteria"],
     "syntax": "=SUMIF(range, criteria, [sum_range])",
     "explain": "Condition match aagura rows-oda values-ah mattum kootum.", "example": '=SUMIF(B2:B101,"EMEA",G2:G101)'},
    {"name": "IF", "category": "formula", "keywords": ["if", "condition", "logical", "true false"],
     "syntax": "=IF(logical_test, value_if_true, value_if_false)",
     "explain": "Condition true-na oru value, false-na vera value tharum.", "example": '=IF(E2>30,"High","Low")'},
    {"name": "VLOOKUP", "category": "formula", "keywords": ["vlookup", "lookup", "find value", "vertical"],
     "syntax": "=VLOOKUP(lookup_value, table_array, col_index, [range_lookup])",
     "explain": "Table-oda mudhal column-la value thedi, right-la irukkura column value tharum. Exact match-ku FALSE kudunga.",
     "example": "=VLOOKUP(A2,Data!A:C,3,FALSE)"},
    {"name": "XLOOKUP", "category": "formula", "keywords": ["xlookup", "lookup", "find value", "modern lookup"],
     "syntax": "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",
     "explain": "VLOOKUP-oda better version. Left/right edhu-um, default exact match.",
     "example": '=XLOOKUP(A2,Data!A:A,Data!C:C,"Not found")'},
    {"name": "INDEX MATCH", "category": "formula", "keywords": ["index", "match", "index match", "lookup", "left lookup"],
     "syntax": "=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",
     "explain": "MATCH row position kandupidikkum, INDEX andha position-la value edukkum. Left-side lookup-kum work aagum.",
     "example": "=INDEX(C2:C101,MATCH(A2,A2:A101,0))"},
    {"name": "TEXT", "category": "formula", "keywords": ["text", "format", "date format", "number format"],
     "syntax": "=TEXT(value, format_text)",
     "explain": "Value-ah text format-la maathum (date, number display).", "example": '=TEXT(A2,"dd-mmm-yyyy")'},
    {"name": "DATE", "category": "formula", "keywords": ["date", "year", "month", "day", "today"],
     "syntax": "=DATE(year, month, day)",
     "explain": "Year, month, day-la irundhu date create pannum. TODAY() inniku date tharum.", "example": "=DATE(2026,11,2)"},
    {"name": "AND / OR / NOT", "category": "formula", "keywords": ["and", "or", "not", "logical", "multiple conditions"],
     "syntax": "=AND(cond1, cond2) / =OR(cond1, cond2)",
     "explain": "Pala conditions-ah serthu check pannum. Bonnaiya IF-kulla use pannalam.",
     "example": '=IF(AND(E2>10,F2<100),"OK","No")'},
    {"name": "Calculated column", "category": "concept", "keywords": ["calculated", "new column", "derived", "fill down"],
     "syntax": "=E2*F2 (then fill down)",
     "explain": "Vera columns-la irundhu kanakkida oru puthu column. Oru cell-la formula ezhudhi keezha fill pannunga.",
     "example": "revenue = units * unit_price"},
    {"name": "Absolute reference", "category": "concept", "keywords": ["absolute", "relative", "dollar", "$", "reference", "lock"],
     "syntax": "$A$1 (fixed), A$1 (row fixed), $A1 (column fixed)",
     "explain": "Formula copy pannum bodhu reference maaradha irukka $ use pannunga.", "example": "=E2*$J$1"},
    {"name": "Pivot table", "category": "concept", "keywords": ["pivot", "summary", "group", "group by", "summarize"],
     "syntax": "Insert > PivotTable",
     "explain": "Data-va group panni sum/average/count summary kaattum. Row, column, value fields drag pannalam.",
     "example": "Region-wise total revenue"},
    {"name": "Conditional formatting", "category": "concept", "keywords": ["conditional", "highlight", "color scale", "data bars", "formatting"],
     "syntax": "Home > Conditional Formatting",
     "explain": "Cell value paathu auto-ah color/icon/bar apply aagum.", "example": "units > 30 red-la highlight"},
    {"name": "Data validation", "category": "concept", "keywords": ["validation", "dropdown", "drop down", "restrict", "input"],
     "syntax": "Data > Data Validation",
     "explain": "Cell-la enna value type pannalam nu control pannum, dropdown list-um create pannalam.", "example": "Region: EMEA, APAC, AMER"},
    {"name": "Remove duplicates", "category": "concept", "keywords": ["duplicate", "duplicates", "unique", "remove duplicates"],
     "syntax": "Data > Remove Duplicates",
     "explain": "Same rows-ah neekkum. Mudhalla backup edhukkunga, adhu undo-ku easy illa.", "example": "Order ID-la duplicates"},
    {"name": "Filter and sort", "category": "concept", "keywords": ["filter", "sort", "order", "ascending", "descending"],
     "syntax": "Data > Filter / Sort",
     "explain": "Filter rows-ah maraikkum (data maaradhu), Sort rows order-ah maathum (data position maarum).",
     "example": "Revenue high to low"},
    {"name": "Chart types", "category": "concept", "keywords": ["chart", "graph", "column", "bar", "line", "pie", "scatter", "which chart"],
     "syntax": "Insert > Charts",
     "explain": "Column/bar: category compare. Line: time trend. Pie: 5-ku kammi parts share. Scatter: rendu numbers relationship.",
     "example": "Region-wise revenue = column chart"},
]


def _tokens(text: str) -> set[str]:
    return {t for t in re.split(r"[^a-z0-9$]+", text.lower()) if t}


def _search(question: str) -> list[tuple[int, dict]]:
    q_text = question.lower()
    q_tokens = _tokens(question)
    scored = []
    for entry in ENTRIES:
        score = 0
        name = entry["name"].lower()
        if name in q_text:
            score += 10
        score += 3 * len(q_tokens & _tokens(name))
        for kw in entry["keywords"]:
            if kw in q_text:
                score += 2 if " " in kw else 1
        if score > 0:
            scored.append((score, entry))
    return sorted(scored, key=lambda s: -s[0])


@safe_tool
def ask_knowledge_base(question: str, max_results: int = 3) -> dict:
    """Explain an Excel formula or concept (SUM, SUMIF, VLOOKUP, XLOOKUP, INDEX MATCH, pivot,
    conditional formatting, ...). Returns syntax, explanation and an example. Read-only."""
    if not isinstance(question, str) or not question.strip():
        raise ExcelToolError("Question empty-a irukka koodadhu.")
    if not isinstance(max_results, int) or isinstance(max_results, bool) or not 1 <= max_results <= MAX_RESULTS:
        raise ExcelToolError(f"max_results 1 mudhal {MAX_RESULTS} varaikkum.")
    hits = _search(question.strip())[:max_results]
    if not hits:
        return {"success": False, "error": "Indha topic knowledge-la illa.",
                "available_topics": [e["name"] for e in ENTRIES]}
    return {"success": True, "question": question.strip(), "source": "built-in",
            "answers": [{"topic": e["name"], "category": e["category"], "syntax": e["syntax"],
                         "explanation": e["explain"], "example": e["example"]} for _, e in hits]}