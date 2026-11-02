"""Backend test: Orchestrator -> Agent -> Gemma -> MCP -> Excel, with real requests.

Run (Ollama running, dataset workbook open in Excel, Excel-la click pannaadhinga):
    python test_backend.py             # read-only
    python test_backend.py --write     # also tests the 'aam' confirmation path (writes only scratch cell K1)
"""
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import test_all_tools as T
from backend import orchestrator as O

MUTATING = {
    "write_excel_data", "append_excel_data", "apply_excel_formula", "apply_excel_formula_range",
    "format_excel_cells", "autofit_excel_columns", "highlight_excel_values", "filter_excel_data",
    "clear_excel_filter", "create_excel_chart", "create_excel_pivot_summary",
}
timings: list[tuple[str, float]] = []


def ask(text: str) -> dict:
    start = time.time()
    with O._request_lock:
        res = O._get_loop_thread().submit(O._global_orchestrator.route_request(text), O.REQUEST_TIMEOUT)
    elapsed = time.time() - start
    timings.append((text, elapsed))
    res["_seconds"] = round(elapsed, 1)
    return res


def tools_used(res: dict) -> list[str]:
    return [s["tool"] for s in res.get("steps", [])]


def reply_of(res: dict) -> str:
    return str(res.get("reply") or res.get("error") or "")


def norm(text: str) -> str:
    return text.replace(",", "").lower()


def pending():
    agent = O._global_orchestrator.agent
    return agent.pending if agent else None


def check(condition, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def short(res: dict) -> str:
    return f"[{res['_seconds']}s] tools={tools_used(res)} reply={reply_of(res)[:90]!r}"


# ---------------- tests ----------------
def t_greeting():
    r = ask("Hello")
    check(r.get("success") and not tools_used(r), f"greeting-ku tool call aachu: {short(r)}")
    check(reply_of(r).strip(), "reply empty")
    return short(r)


def t_sheet_names():
    r = ask("intha workbook-la enna sheets irukku?")
    check("get_excel_sheet_names" in tools_used(r), f"sariyaana tool illa: {short(r)}")
    check("sales-transactions" in reply_of(r), f"reply-la sheet peyar illa: {short(r)}")
    return short(r)

def t_revenue_total():
    r = ask("revenue column total enna?")
    check({"analyze_excel_column", "calculate_excel_total"} & set(tools_used(r)),
          f"sariyaana tool illa: {short(r)}")
    check("119231.42" in norm(reply_of(r)), f"sum thappu/illa: {short(r)}")
    return short(r)


def t_first_rows():
    r = ask("first 5 rows kaattu")
    check("read_excel_data" in tools_used(r), f"read_excel_data call aagala: {short(r)}")
    check(not set(tools_used(r)) & MUTATING, f"read request-la write tool aachu: {tools_used(r)}")
    return short(r)


def t_units_stats():
    r = ask("units column stats sollu")
    calls = [s for s in r.get("steps", []) if s["tool"] == "analyze_excel_column"]
    check(calls and calls[0]["arguments"].get("column") == "units", f"units-ku analyze call illa: {short(r)}")
    check("40" in reply_of(r), f"max=40 reply-la illa: {short(r)}")
    return short(r)


def t_followup_memory():
    r = ask("adhula max evlo?")
    check("40" in reply_of(r), f"follow-up purinjukkala: {short(r)}")
    return short(r)


def t_validate():
    r = ask("data clean-a irukka?")
    check("validate_excel_data" in tools_used(r), f"validate tool illa: {short(r)}")
    return short(r)


def t_duplicates_preview():
    r = ask("duplicate rows irukka?")
    check({"remove_excel_duplicates", "validate_excel_data"} & set(tools_used(r)), f"duplicates-ku sariyaana tool illa: {short(r)}")
    check(not any(s["result"].get("deleted") for s in r.get("steps", [])), "ROW DELETE AAGIDUCHU!")
    check(pending() is None, "dataset clean, pending irukka koodaadhu")
    return short(r)


def t_missing_column():
    r = ask("profit column total sollu")
    bad = [s for s in r.get("steps", []) if s["arguments"].get("column") == "profit" and s["result"].get("success")]
    check(not bad, f"illaadha column success aachu: {bad}")
    return "REVIEW: " + short(r) + "  <- 'profit column illa' nu solla vendum, number invent pannakoodaadhu"


def t_save_request():
    r = ask("workbook-a save pannu")
    check(not set(tools_used(r)) & MUTATING, f"save request-la write tool aachu: {tools_used(r)}")
    return "REVIEW: " + short(r) + "  <- save tool illa nu honest-a solla vendum"


def t_confirm_no():
    r = ask("A1-la 'test' nu ezhudhu")
    check(r.get("needs_confirmation") is True and pending() is not None,
          f"confirmation kekkala (tools={tools_used(r)}): {short(r)}")
    r2 = ask("illa")
    check(pending() is None, "'illa'-ku piragum pending irukku")
    check(T.sh().range("A1").value == "date", f"A1 MAARIDUCHU: {T.sh().range('A1').value!r}")
    return f"asked -> 'illa' -> cancelled, A1 safe | {short(r2)}"


def t_confirm_yes():
    k1 = T.sh().range("K1")
    check(k1.value is None, "K1 empty illa, skip")
    r1 = ask("K1-la 'x' nu ezhudhu")
    check(k1.value == "x", f"K1 'x' aagala (K1={k1.value!r}) tools={tools_used(r1)}")
    r2 = ask("K1-la 'y' nu ezhudhu")
    check(r2.get("needs_confirmation") is True and k1.value == "x",
          f"overwrite confirmation kekkala (K1={k1.value!r}): {short(r2)}")
    ask("aam")
    check(k1.value == "y", f"'aam'-ku piragum K1={k1.value!r} ('y' varanum)")
    return "x written, overwrite asked, 'aam' -> y (K1 cleaned after)"


def main() -> int:
    write_mode = "--write" in sys.argv
    problem = T.preflight()
    if problem:
        print(f"PREFLIGHT FAILED (dataset-ai thodala): {problem}")
        return 1
    tests = [t_greeting, t_sheet_names, t_revenue_total, t_first_rows, t_units_stats, t_followup_memory,
             t_validate, t_duplicates_preview, t_missing_column, t_save_request, t_confirm_no]
    if write_mode:
        tests.append(t_confirm_yes)
    print(f"Backend test: {len(tests)} requests (~{len(tests) * 15 // 60 + 1} min). Excel-la click pannaadhinga...\n")

    results = []
    try:
        for number, func in enumerate(tests, 1):
            try:
                detail = func()
                status = "REVIEW" if detail.startswith("REVIEW:") else "PASS"
                detail = detail.removeprefix("REVIEW: ")
            except AssertionError as exc:
                status, detail = "FAIL", str(exc)
            except Exception as exc:
                status, detail = "ERROR", f"{type(exc).__name__}: {exc}"
            results.append((func.__name__, status, detail))
            print(f"[{number:>2}/{len(tests)}] {status:<6} {func.__name__:<22} {detail}")
    finally:
        try:
            T.cleanup()
        except Exception as exc:
            print(f"CLEANUP ERROR: {type(exc).__name__}: {exc}")
        try:
            O._get_loop_thread().submit(O._global_orchestrator.close(), 30)
        except Exception:
            pass

    intact, problems = T.dataset_intact()
    passed = sum(1 for r in results if r[1] == "PASS")
    review = sum(1 for r in results if r[1] == "REVIEW")
    slowest = max(timings, key=lambda t: t[1]) if timings else ("", 0)
    print("\n" + "=" * 64)
    print(f"RESULT: {passed} passed, {review} need your review, {len(results) - passed - review} failed (of {len(results)})")
    print(f"Slowest request: {slowest[1]:.1f}s -> {slowest[0]!r}")
    for name, status, detail in results:
        if status in ("FAIL", "ERROR"):
            print(f"  - {name} [{status}]: {detail}")
    print("DATASET CHECK: PASS" if intact else "DATASET CHANGED: " + "; ".join(problems))
    print("=" * 64)
    return 0 if intact and passed + review == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())