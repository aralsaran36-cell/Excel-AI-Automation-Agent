"""Domain detector: guesses the dataset type from column names.

Pure Python, no Excel access. Rules only (no LLM), so results are predictable.
detect_domain(columns) -> {"domain", "score", "title", "theme"}
Test (no Excel needed): python -m dashboard.domain_detector
"""
from __future__ import annotations

import re
from typing import Any

DOMAINS: dict[str, dict[str, Any]] = {
    "Sales": {"theme": "blue", "keywords": [
        "revenue", "sales", "units", "price", "order", "product", "customer",
        "discount", "channel", "region", "quantity"]},
    "HR": {"theme": "purple", "keywords": [
        "salary", "employee", "department", "designation", "hire", "attrition",
        "experience", "gender", "manager", "staff"]},
    "Education": {"theme": "orange", "keywords": [
        "marks", "score", "subject", "student", "grade", "class", "exam",
        "attendance", "teacher", "semester"]},
    "Finance": {"theme": "green", "keywords": [
        "expense", "income", "profit", "budget", "balance", "account", "tax",
        "cost", "loss", "payment"]},
    "Inventory": {"theme": "teal", "keywords": [
        "stock", "inventory", "sku", "warehouse", "supplier", "reorder", "shipment"]},
    "Healthcare": {"theme": "teal", "keywords": [
        "patient", "diagnosis", "doctor", "hospital", "treatment", "bed", "disease"]},
}
DEFAULT = {"domain": "General", "score": 0, "title": "Data Dashboard", "theme": "slate"}


def _tokens(name: str) -> set[str]:
    return {t for t in re.split(r"[^a-z0-9]+", str(name).lower()) if t}


def _matches(tokens: set[str], keyword: str) -> bool:
    return keyword in tokens or f"{keyword}s" in tokens or keyword.rstrip("s") in tokens


def detect_domain(columns: list[str]) -> dict[str, Any]:
    """Score each domain by how many columns match its keywords. Best score wins."""
    best_name, best_score = None, 0
    for name, cfg in DOMAINS.items():
        score = sum(
            1 for col in columns
            if any(_matches(_tokens(col), kw) for kw in cfg["keywords"])
        )
        if score > best_score:
            best_name, best_score = name, score
    if best_name is None or best_score < 2:   # 1 match alone is too weak
        return dict(DEFAULT)
    return {"domain": best_name, "score": best_score,
            "title": f"{best_name} Dashboard", "theme": DOMAINS[best_name]["theme"]}


if __name__ == "__main__":
    cases = {
        "sales":     ["rukku", "region", "product", "channel", "units", "unit_price", "revenue"],
        "hr":        ["emp_id", "employee_name", "department", "salary", "hire_date", "gender"],
        "education": ["student", "subject", "marks", "grade"],
        "finance":   ["account", "income", "expense", "balance"],
        "unknown":   ["a", "b", "c"],
    }
    for label, cols in cases.items():
        r = detect_domain(cols)
        print(f"{label:<10} -> {r['domain']:<10} score={r['score']} theme={r['theme']} title={r['title']}")