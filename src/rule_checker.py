"""
Rule-based Excel part-list validator.

Expected rule workbook columns:
Rule ID, Column Letter, Header Name, Rule, Source, Rule Row in Sheet

Usage:
    import excel_tools as T
    import rule_checker as R

    T.set_file(partlist_bytes)
    R.set_rules_file(rule_workbook_bytes)
    result = R.check_all_rules("ZaragozaTH,PF,BS")

The validator reads part-list data through excel_tools._table(), so it uses
the same sheet/header/data-row detection as the existing Excel tools.
"""
from __future__ import annotations

import io
import math
import re
import threading
from collections import Counter
from typing import Any

import pandas as pd
from openpyxl import load_workbook

import src.excel_tools as T

_LOCK = threading.RLock()
_RULES_BYTES: bytes | None = None
_RULES_CACHE: list[dict[str, str]] | None = None

REQUIRED_RULE_COLUMNS = {
    "Rule ID", "Column Letter", "Header Name", "Rule"
}


def set_rules_file(file_bytes: bytes) -> None:
    """Load/update the uploaded rule workbook bytes."""
    global _RULES_BYTES, _RULES_CACHE
    with _LOCK:
        _RULES_BYTES = file_bytes
        _RULES_CACHE = None


def _load_rules() -> list[dict[str, str]]:
    global _RULES_CACHE
    with _LOCK:
        if _RULES_CACHE is not None:
            return _RULES_CACHE
        if not _RULES_BYTES:
            raise ValueError(
                "Rules workbook is not loaded. Call set_rules_file(rule_file_bytes) first."
            )

        wb = load_workbook(io.BytesIO(_RULES_BYTES), read_only=True,
                           data_only=True, keep_links=False)
        ws = wb["Rules"] if "Rules" in wb.sheetnames else wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)
        header_row = next(rows, None)
        if not header_row:
            raise ValueError("The rules workbook is empty.")

        headers = [str(x).strip() if x is not None else "" for x in header_row]
        missing = REQUIRED_RULE_COLUMNS - set(headers)
        if missing:
            raise ValueError(
                f"Rules workbook is missing required columns: {sorted(missing)}"
            )

        result = []
        for row in rows:
            record = {
                headers[i]: ("" if i >= len(row) or row[i] is None else str(row[i]).strip())
                for i in range(len(headers))
            }
            if not record.get("Rule ID") or not record.get("Rule"):
                continue
            result.append(record)
        wb.close()
        _RULES_CACHE = result
        return result


def _blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return isinstance(value, str) and value.replace("\xa0", " ").strip() == ""


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\xa0", " ").strip()


def _is_zero(value: Any) -> bool:
    if _blank(value):
        return False
    try:
        return float(_text(value).replace(",", ".")) == 0
    except (ValueError, TypeError):
        return _text(value).casefold() == "0"


def _is_one_or_more(value: Any) -> bool:
    if _blank(value):
        return False
    try:
        return float(_text(value).replace(",", ".")) >= 1
    except (ValueError, TypeError):
        return _text(value).casefold() in {"yes", "true", "x", "y"}


def _is_whole_number(value: Any) -> bool:
    if _blank(value):
        return False
    try:
        number = float(_text(value).replace(",", "."))
        return number >= 1 and number.is_integer()
    except (ValueError, TypeError):
        return False


def _cell_value(df: pd.DataFrame, row: int, col: str) -> Any:
    if col not in df.columns or row not in df.index:
        return None
    return df.at[row, col]


def _result(rule: dict[str, str], status: str, message: str,
            sheet: str, row: int | None = None, cell: str = "",
            value: Any = None) -> dict[str, Any]:
    out = {
        "rule_id": rule.get("Rule ID", ""),
        "column": rule.get("Column Letter", ""),
        "header": rule.get("Header Name", ""),
        "rule": rule.get("Rule", ""),
        "status": status,
        "message": message,
        "sheet": sheet,
    }
    if row is not None:
        out["excel_row"] = int(row)
    if cell:
        out["cell"] = cell
    if value is not None:
        out["actual_value"] = _text(value)
    return out


def check_all_rules(sheet: str = "ZaragozaTH,PF,BS") -> dict[str, Any]:
    """
    Run every rule from the uploaded rules workbook against the loaded partlist.

    Returns JSON-friendly summary and row-level findings. Statuses:
    PASS, FAIL, REVIEW. REVIEW is used for conditional rules whose applicability
    cannot safely be inferred from the provided rule wording.
    """
    rules = _load_rules()
    resolved_sheet, df, headers, header_row, first_data_row = T._table(sheet)
    findings: list[dict[str, Any]] = []
    rule_results: list[dict[str, Any]] = []

    # Rule conditions refer to fixed Excel columns, so evaluate by letters.
    # Required nonblank fields are checked for every populated part row.
    nonblank_ids = {
        "R002", "R003", "R004", "R005", "R006", "R007",
        "R010", "R011", "R012", "R013", "R014", "R015",
        "R017", "R018", "R019", "R020", "R021", "R022",
        "R023", "R024",
    }
    conditional_ids = {"R008", "R009", "R016", "R025", "R026", "R027",
                       "R028", "R029", "R030", "R031"}

    # Only data rows represented in the loaded dataframe are checked.
    for rule in rules:
        rid = rule["Rule ID"].strip()
        col = rule.get("Column Letter", "").strip().upper()
        description = rule.get("Rule", "")
        header_name = rule.get("Header Name", "")
        row_findings = []

        if not col or col not in df.columns:
            rule_results.append(_result(
                rule, "REVIEW",
                f"Configured column {col or '(blank)'} is not available in the parsed sheet.",
                resolved_sheet
            ))
            continue

        # R001: flag every occurrence of each duplicated part number.
        # Comparison ignores case and repeated/leading/trailing whitespace.
        if rid == "R001":
            def duplicate_key(value: Any) -> str:
                return re.sub(r"\s+", " ", _text(value)).strip().casefold()

            keys = df[col].map(duplicate_key)
            keys = keys[keys != ""]
            counts = keys.value_counts()
            duplicate_keys = set(counts[counts > 1].index)
            for row in keys.index:
                key = keys.loc[row]
                if key in duplicate_keys:
                    original = _cell_value(df, int(row), col)
                    row_findings.append(_result(
                        rule, "FAIL",
                        f"Part number appears {int(counts[key])} times.",
                        resolved_sheet, int(row), f"{col}{row}", original
                    ))
            status = "FAIL" if row_findings else "PASS"
            rule_results.append(_result(
                rule, status,
                f"{len(duplicate_keys)} duplicated part-number value(s); "
                f"{len(row_findings)} affected row(s).",
                resolved_sheet
            ))
            findings.extend(row_findings)
            continue

        # Required nonblank fields. R015 (AU) has the R016 SILS exemption:
        # a blank UC code is allowed when column CF says Yes.
        if rid in nonblank_ids:
            for row in df.index:
                value = _cell_value(df, int(row), col)
                if all(_blank(_cell_value(df, int(row), c)) for c in df.columns):
                    continue
                if not _blank(value):
                    continue
                if rid == "R015" and _text(_cell_value(df, int(row), "CF")).casefold() == "yes":
                    continue
                row_findings.append(_result(
                    rule, "FAIL", f"Required value is blank in {header_name}.",
                    resolved_sheet, int(row), f"{col}{row}", value
                ))
            status = "FAIL" if row_findings else "PASS"
            rule_results.append(_result(
                rule, status,
                f"{len(row_findings)} blank required cell(s) found.",
                resolved_sheet
            ))
            findings.extend(row_findings)
            continue

        # R008: if O=0, Q should be 'unique to O1H'.
        if rid == "R008":
            for row in df.index:
                if _is_zero(_cell_value(df, int(row), "O")):
                    actual = _text(_cell_value(df, int(row), "Q"))
                    if actual.casefold() != "unique to o1h":
                        row_findings.append(_result(
                            rule, "FAIL",
                            "Column O is 0, so Q should be 'unique to O1H'.",
                            resolved_sheet, int(row), f"Q{row}", actual
                        ))
            status = "FAIL" if row_findings else "PASS"
            rule_results.append(_result(rule, status,
                                        f"{len(row_findings)} violation(s) found.",
                                        resolved_sheet))
            findings.extend(row_findings)
            continue

        # R009: if P=0, Q should be 'unique to P1H'.
        if rid == "R009":
            for row in df.index:
                if _is_zero(_cell_value(df, int(row), "P")):
                    actual = _text(_cell_value(df, int(row), "Q"))
                    if actual.casefold() != "unique to p1h":
                        row_findings.append(_result(
                            rule, "FAIL",
                            "Column P is 0, so Q should be 'unique to P1H'.",
                            resolved_sheet, int(row), f"Q{row}", actual
                        ))
            status = "FAIL" if row_findings else "PASS"
            rule_results.append(_result(rule, status,
                                        f"{len(row_findings)} violation(s) found.",
                                        resolved_sheet))
            findings.extend(row_findings)
            continue

        # R016 defines the SILS exemption for R015. It does not create a second
        # cell-level error; the required-field check above reports any violation.
        if rid == "R016":
            flags = [_text(_cell_value(df, int(row), "CF")).casefold() for row in df.index
                     if not all(_blank(_cell_value(df, int(row), c)) for c in df.columns)]
            unknown = [flag for flag in flags if flag not in {"yes", "no"}]
            status = "REVIEW" if unknown else "PASS"
            message = ("SILS exemption applied using column CF; unknown flag values need review."
                       if unknown else "SILS exemption evaluated using column CF.")
            rule_results.append(_result(rule, status, message, resolved_sheet))
            continue

        # R025-R028: require a value when the corresponding unit applies.
        if rid in {"R025", "R026", "R027", "R028"}:
            flag_col = "O" if rid in {"R025", "R027"} else "P"
            for row in df.index:
                if all(_blank(_cell_value(df, int(row), c)) for c in df.columns):
                    continue
                if _is_one_or_more(_cell_value(df, int(row), flag_col)) and _blank(_cell_value(df, int(row), col)):
                    unit = "P1H" if flag_col == "O" else "O1H"
                    row_findings.append(_result(
                        rule, "FAIL", f"{unit} applies but required value is blank.",
                        resolved_sheet, int(row), f"{col}{row}", _cell_value(df, int(row), col)
                    ))
            status = "FAIL" if row_findings else "PASS"
            rule_results.append(_result(rule, status,
                                        f"{len(row_findings)} missing required value(s).",
                                        resolved_sheet))
            findings.extend(row_findings)
            continue

        # R029-R031: nonblank values must be positive whole numbers.
        if rid in {"R029", "R030", "R031"}:
            for row in df.index:
                value = _cell_value(df, int(row), col)
                if _blank(value):
                    continue
                try:
                    number = float(_text(value).replace(",", "."))
                except (TypeError, ValueError):
                    number = None
                if number is None:
                    message = "Not a number"
                elif number < 1 or not number.is_integer():
                    message = "Must be a positive whole number"
                else:
                    continue
                row_findings.append(_result(
                    rule, "FAIL", message, resolved_sheet, int(row), f"{col}{row}", value
                ))
            status = "FAIL" if row_findings else "PASS"
            rule_results.append(_result(rule, status,
                                        f"{len(row_findings)} invalid whole-number value(s).",
                                        resolved_sheet))
            findings.extend(row_findings)
            continue

        # Unknown conditional rule IDs are not marked as passed without validation.
        if rid in conditional_ids:
            rule_results.append(_result(
                rule, "REVIEW",
                "Conditional rule needs an explicit applicability definition before it can be validated automatically.",
                resolved_sheet
            ))
            continue

        # Unknown rule text should never be marked as passed without validation.
        rule_results.append(_result(
            rule, "REVIEW",
            "No automated validator is configured for this rule wording yet.",
            resolved_sheet
        ))

    counts = Counter(item["status"] for item in rule_results)
    finding_counts = Counter(item["status"] for item in findings)

    return {
        "sheet": resolved_sheet,
        "header_row": int(header_row),
        "first_data_row": int(first_data_row),
        "data_rows_checked": int(len(df)),
        "rules_loaded": len(rules),
        "rules_passed": counts.get("PASS", 0),
        "rules_failed": counts.get("FAIL", 0),
        "rules_needing_review": counts.get("REVIEW", 0),
        "row_level_failures": len(findings),
        "rule_summary": rule_results,
        "findings": findings,
        "notes": [
            "Conditional rules are marked REVIEW until their applicability logic is defined.",
            "Blank rows are not treated as part records unless another cell in that row contains data.",
            "Confirm that the target workbook's header/data row settings match excel_tools.py.",
        ],
    }
