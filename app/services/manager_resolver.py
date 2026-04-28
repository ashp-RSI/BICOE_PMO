"""Resolve a First Line Manager's email by looking them up as an employee
in the same Head Count Report sheet.

Resolution strategy (Option a with guardrails):
  1. Normalize names (trim, collapse whitespace, lowercase) before matching.
  2. Exact match first.
  3. If multiple exact matches, narrow by Sub Practice.
  4. If no exact match, fuzzy match via rapidfuzz with a threshold.
  5. Return a structured result so the UI can show why resolution failed.
"""
import logging
import re

import pandas as pd

try:
    from rapidfuzz import process, fuzz
    _RAPIDFUZZ_AVAILABLE = True
except ImportError:
    _RAPIDFUZZ_AVAILABLE = False

logger = logging.getLogger(__name__)

# Possible resolution statuses surfaced to the UI.
STATUS_OK = "ok"                              # exact unique match
STATUS_OK_FUZZY = "ok_fuzzy"                  # fuzzy best match above threshold
STATUS_OK_DISAMBIGUATED = "ok_disambiguated"  # multiple matches, narrowed by Sub Practice
STATUS_NOT_FOUND = "not_found"                # no manager name on the row
STATUS_NO_MATCH = "no_match"                  # name didn't match any employee
STATUS_MULTIPLE_MATCHES = "multiple_matches"  # multiple matches, can't disambiguate
STATUS_NO_EMAIL = "no_email"                  # matched but Company Email blank


def _normalize(name):
    """Lowercase, trim, collapse whitespace. Returns '' for null/NaN."""
    if name is None:
        return ""
    if isinstance(name, float) and pd.isna(name):
        return ""
    return re.sub(r"\s+", " ", str(name).strip()).lower()


def resolve_manager_email(emp_row, headcount_df, fuzzy_threshold=90):
    """Resolve the manager email for the given employee row.

    Args:
        emp_row: A pandas Series (one employee row) or a dict.
        headcount_df: The full headcount DataFrame.
        fuzzy_threshold: Minimum rapidfuzz score (0-100) to accept a fuzzy match.

    Returns:
        dict with keys:
          status         — one of the STATUS_* constants above
          manager_name   — the raw 'First Line Manager' value from the row
          manager_email  — resolved email or '' if not resolved
          method         — 'exact' | 'fuzzy' | 'sub_practice_disambig' | ''
          candidates     — list of (name, email) tuples for the multi-match case
          fuzzy_score    — float 0-100 when method='fuzzy', else None
    """
    def _get(row, key):
        if isinstance(row, dict):
            return row.get(key)
        return row.get(key) if hasattr(row, "get") else None

    mgr_raw = _get(emp_row, "First Line Manager")
    mgr_norm = _normalize(mgr_raw)

    result = {
        "status": STATUS_NOT_FOUND,
        "manager_name": str(mgr_raw or "").strip(),
        "manager_email": "",
        "method": "",
        "candidates": [],
        "fuzzy_score": None,
    }

    if not mgr_norm:
        return result

    if headcount_df is None or headcount_df.empty:
        result["status"] = STATUS_NO_MATCH
        return result

    if "Emp Name" not in headcount_df.columns:
        result["status"] = STATUS_NO_MATCH
        return result

    df = headcount_df.copy()
    df["_norm_name"] = df["Emp Name"].apply(_normalize)

    exact = df[df["_norm_name"] == mgr_norm]

    if len(exact) == 1:
        match = exact.iloc[0]
        email = _clean_email(match.get("Company Email"))
        if not email:
            result["status"] = STATUS_NO_EMAIL
            result["method"] = "exact"
            return result
        result["status"] = STATUS_OK
        result["method"] = "exact"
        result["manager_email"] = email
        return result

    if len(exact) > 1:
        emp_sub_practice = _normalize(_get(emp_row, "Sub Practice"))
        result["candidates"] = [
            (str(r.get("Emp Name") or ""), _clean_email(r.get("Company Email")))
            for _, r in exact.iterrows()
        ]
        if emp_sub_practice:
            narrowed = exact[
                exact["Sub Practice"].fillna("").astype(str)
                     .str.strip().str.lower() == emp_sub_practice
            ]
            if len(narrowed) == 1:
                match = narrowed.iloc[0]
                email = _clean_email(match.get("Company Email"))
                if not email:
                    result["status"] = STATUS_NO_EMAIL
                    result["method"] = "sub_practice_disambig"
                    return result
                result["status"] = STATUS_OK_DISAMBIGUATED
                result["method"] = "sub_practice_disambig"
                result["manager_email"] = email
                return result
        result["status"] = STATUS_MULTIPLE_MATCHES
        return result

    # No exact match → fuzzy fallback
    if _RAPIDFUZZ_AVAILABLE:
        choices = df["_norm_name"].tolist()
        if choices:
            best = process.extractOne(
                mgr_norm, choices, scorer=fuzz.WRatio,
                score_cutoff=fuzzy_threshold,
            )
            if best:
                matched_name, score, idx = best
                match = df.iloc[idx]
                email = _clean_email(match.get("Company Email"))
                if not email:
                    result["status"] = STATUS_NO_EMAIL
                    result["method"] = "fuzzy"
                    result["fuzzy_score"] = round(float(score), 1)
                    return result
                result["status"] = STATUS_OK_FUZZY
                result["method"] = "fuzzy"
                result["manager_email"] = email
                result["fuzzy_score"] = round(float(score), 1)
                return result
    else:
        logger.debug("rapidfuzz not installed — skipping fuzzy fallback")

    result["status"] = STATUS_NO_MATCH
    return result


def _clean_email(value):
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s


def audit_unresolved(headcount_df, fuzzy_threshold=90):
    """Return a list of employees whose First Line Manager email could not
    be auto-resolved. Used to populate the Manager Email Audit page."""
    if headcount_df is None or headcount_df.empty:
        return []

    audit = []
    for idx, row in headcount_df.iterrows():
        if "Status" in row and str(row.get("Status", "")).strip() == "Resigned":
            continue
        if "Billable/Non Billable" in row and str(
            row.get("Billable/Non Billable", "")
        ).strip() == "Resigned":
            continue
        result = resolve_manager_email(row, headcount_df, fuzzy_threshold)
        if result["status"] not in (STATUS_OK, STATUS_OK_DISAMBIGUATED):
            audit.append({
                "row_index": int(idx),
                "emp_code": str(row.get("Emp Code") or ""),
                "emp_name": str(row.get("Emp Name") or ""),
                "sub_practice": str(row.get("Sub Practice") or ""),
                "manager_name": result["manager_name"],
                "status": result["status"],
                "fuzzy_email": result["manager_email"],
                "fuzzy_score": result["fuzzy_score"],
            })
    return audit
