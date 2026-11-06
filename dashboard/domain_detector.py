"""Domain detector: guesses the dataset type from column names.

Pure Python, no Excel access. Rules only (no LLM), so results are predictable.
detect_domain(columns) -> {"domain", "score", "title", "theme"}
Test (no Excel needed): python -m dashboard.domain_detector
"""
from __future__ import annotations

import re
from typing import Any

DOMAINS: dict[str, dict[str, Any]] = {
    "Sales": {"theme": "blue", "title": "Sales Dashboard", "keywords": [
        "revenue", "sales", "units", "price", "order", "product", "customer",
        "discount", "channel", "region", "quantity"]},
    "HR": {"theme": "purple", "title": "HR Dashboard", "keywords": [
        "salary", "employee", "department", "designation", "hire", "attrition",
        "experience", "gender", "manager", "staff"]},
    "Education": {"theme": "orange", "title": "Education Dashboard", "keywords": [
        "marks", "score", "subject", "student", "grade", "class", "exam",
        "attendance", "teacher", "semester"]},
    "Finance": {"theme": "green", "title": "Finance Dashboard", "keywords": [
        "expense", "income", "profit", "budget", "balance", "account", "tax",
        "cost", "loss", "payment"]},
    "Inventory": {"theme": "teal", "title": "Inventory Dashboard", "keywords": [
        "stock", "inventory", "sku", "warehouse", "supplier", "reorder", "shipment"]},
    "Healthcare": {"theme": "teal", "title": "Healthcare Dashboard", "keywords": [
        "patient", "diagnosis", "doctor", "hospital", "treatment", "bed", "disease"]},
    "Passengers": {"theme": "slate", "title": "Passenger Survival Dashboard", "keywords": [
        "passenger", "survived", "pclass", "sibsp", "parch", "embarked", "cabin", "ticket"]},
    "Survey": {"theme": "slate", "title": "Survey Dashboard", "keywords": [
        "survey", "response", "respondent", "rating", "feedback", "satisfaction", "question", "nps"]},
    "Customer": {"theme": "slate", "title": "Customer Dashboard", "keywords": [
        "churn", "subscription", "tenure", "segment", "contract", "loyalty", "customer", "plan"]},
    "Marketing": {"theme": "slate", "title": "Marketing Dashboard", "keywords": [
        "campaign", "impressions", "clicks", "ctr", "conversion", "ads", "lead", "channel"]},
    "Transport": {"theme": "slate", "title": "Transport Dashboard", "keywords": [
        "trip", "route", "vehicle", "driver", "distance", "ride", "fare", "pickup"]},
}
DEFAULT = {"domain": "General", "score": 0, "title": "Data Dashboard", "theme": "slate"}


def _tokens(name: str) -> set[str]:
    spaced = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(name))      # PassengerId -> Passenger Id
    return {t for t in re.split(r"[^a-z0-9]+", spaced.lower()) if t}


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
    cfg = DOMAINS[best_name]
    return {"domain": best_name, "score": best_score, "title": cfg["title"], "theme": cfg["theme"]}


if __name__ == "__main__":
    cases = {
        "sales": ["order_id", "region", "product", "channel", "units", "unit_price", "revenue"],
        "hr": ["emp_id", "employee_name", "department", "salary", "hire_date", "gender"],
        "titanic": ["PassengerId", "Survived", "Pclass", "Name", "Sex", "Age", "SibSp", "Parch", "Fare"],
        "survey": ["respondent", "rating", "feedback", "question"],
        "unknown": ["a", "b", "c"],
    }
    for label, cols in cases.items():
        r = detect_domain(cols)
        print(f"{label:<10} -> {r['domain']:<12} score={r['score']} title={r['title']}")