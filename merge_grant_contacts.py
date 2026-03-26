#!/usr/bin/env python3
"""
Merge a grants workbook (file 1) with a contacts workbook (file 2) into one
row per matched Funder/Foundation pair for migration export.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

# Output column order (all included; unfilled cells are empty strings)
OUTPUT_COLUMNS = [
    "Prefix",
    "First Name",
    "Middle Name",
    "Last Name",
    "Suffix",
    "Individual/Company Type",
    "IsCompany?",
    "Company Name",
    "Department",
    "Job Title",
    "Preferred Name",
    "Salutation",
    "Deceased",
    "Deceased Date",
    "Do Not Contact",
    "Email Opt Out",
    "Email 1",
    "Email 2",
    "Email 3",
    "SMS/MMS Number",
    "SMS/MMS Consent",
    "Phone 1",
    "Phone 1 Type",
    "Phone 2",
    "Phone 2 Type",
    "Phone 3",
    "Phone 3 Type",
    "Website",
    "Fax",
    "Address Line 1",
    "Address Line 2",
    "Address Line 3",
    "Address Line 4",
    "Address Type",
    "City",
    "State/Province",
    "Territory",
    "County",
    "Country",
    "Zip",
    "Birthday",
    "Gender",
    "Login Name",
    "Login Password",
    "Note",
    "Note Title",
    "Note Type",
    "Pinned Note?",
    "Account Source",
    "Volunteer Role(s)",
    "Volunteer Group(s)",
    "Created By",
    "Created Date",
    "Last Updated By",
    "Last Updated Date",
    "Custom Fields",
    "Grant Status",
    "Grant System User",
    "Grant Name",
    "Ask Date",
    "Ask Amount",
    "Funded Date",
    "Funded Amount",
    "Close Date",
    "Grant Note",
    "Grant Campaign",
    "Grant Fund",
    "Grant Purpose",
    "Grant Remaining",
    "Grant Award Date",
    "Grant Report Date",
    "Grant Report Notes",
]


def _norm_key(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _find_col(df: pd.DataFrame, *candidates: str) -> str | None:
    """Return first column name whose normalized header matches a candidate."""
    norm_to_col = {_norm_key(c): c for c in df.columns}
    for cand in candidates:
        k = _norm_key(cand)
        if k in norm_to_col:
            return norm_to_col[k]
    return None


def _col_by_position(df: pd.DataFrame, idx: int) -> str | None:
    if idx < 0 or idx >= len(df.columns):
        return None
    return df.columns[idx]


def load_grants(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() if c is not None and str(c) != "nan" else "" for c in df.columns]
    return df


def load_contacts(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    df.columns = [str(c).strip() if c is not None and str(c) != "nan" else "" for c in df.columns]
    return df


def print_headers(grants: pd.DataFrame, contacts: pd.DataFrame) -> None:
    print("=== Grants file headers ===")
    for idx, col in enumerate(grants.columns, start=1):
        print(f"{idx:>2}. {col!r}")
    print("=== Contacts file headers ===")
    for idx, col in enumerate(contacts.columns, start=1):
        print(f"{idx:>2}. {col!r}")


def resolve_contact_columns(df: pd.DataFrame) -> dict[str, str | None]:
    """
    Map logical fields to actual column names. File 2 may have blank headers
    after Address (line 2, city, state, zip).
    """
    first = _find_col(df, "first name")
    last = _find_col(df, "last name")
    foundation = _find_col(df, "foundation")
    title = _find_col(df, "title")
    email = _find_col(df, "email")
    phone = _find_col(df, "phone")

    addr_header = _find_col(df, "address")
    addr_idx = None
    if addr_header is not None:
        addr_idx = list(df.columns).index(addr_header)

    addr1 = addr_header
    addr2 = _col_by_position(df, addr_idx + 1) if addr_idx is not None else None
    city_col = _col_by_position(df, addr_idx + 2) if addr_idx is not None else None
    state_col = _col_by_position(df, addr_idx + 3) if addr_idx is not None else None
    zip_col = _col_by_position(df, addr_idx + 4) if addr_idx is not None else None

    return {
        "first": first,
        "last": last,
        "foundation": foundation,
        "title": title,
        "email": email,
        "phone": phone,
        "addr1": addr1,
        "addr2": addr2,
        "city": city_col,
        "state": state_col,
        "zip": zip_col,
    }


def _cell(df: pd.DataFrame, row_idx: int, col: str | None) -> str:
    if col is None or col not in df.columns:
        return ""
    v = df.iloc[row_idx][col]
    if pd.isna(v):
        return ""
    return str(v).strip()


def _grant_cell(grants: pd.DataFrame, row_idx: int, *header_candidates: str) -> str:
    for h in header_candidates:
        col = _find_col(grants, h)
        if col:
            return _cell(grants, row_idx, col)
    return ""


def _format_amount(value: str) -> str:
    """Format numeric-looking values to 2 decimals; leave non-numeric as-is."""
    text = str(value).strip()
    if not text:
        return ""
    cleaned = text.replace("$", "").replace(",", "")
    try:
        num = float(cleaned)
    except ValueError:
        return text
    return f"{num:.2f}"


def _format_date_mmddyyyy(value: str) -> str:
    """Normalize a single date value to mm/dd/yyyy when parseable."""
    text = str(value).strip()
    if not text:
        return ""
    dt = pd.to_datetime(text, errors="coerce")
    if pd.isna(dt):
        return text
    return dt.strftime("%m/%d/%Y")


def _format_report_date_multiline(value: str) -> str:
    """
    Preserve report date as multiline text.
    Each line is normalized to mm/dd/yyyy only when parseable.
    """
    text = str(value).strip()
    if not text:
        return ""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    formatted = [_format_date_mmddyyyy(line) for line in lines]
    return ",\n".join(formatted)



def combine_grant_report_notes(report_notes: str, notes: str) -> str:
    """Report Notes and Notes both map into Grant Report Notes (combined if both set)."""
    parts = [x for x in (report_notes, notes) if x]
    return " | ".join(parts) if parts else ""


def merge_workbooks(grants: pd.DataFrame, contacts: pd.DataFrame) -> pd.DataFrame:
    g_funder = _find_col(grants, "funder")
    if not g_funder:
        raise ValueError("Grants file must have a 'Funder' column.")

    cmap = resolve_contact_columns(contacts)
    if not cmap["foundation"]:
        raise ValueError("Contacts file must have a 'Foundation' column.")

    rows_out: list[dict[str, str]] = []

    for gi in range(len(grants)):
        funder_raw = _cell(grants, gi, g_funder)
        funder_key = _norm_key(funder_raw)
        if not funder_key:
            continue

        report_notes_val = _grant_cell(grants, gi, "report notes")
        notes_val = _grant_cell(grants, gi, "notes")
        grant_block = {
            "Company Name": _grant_cell(grants, gi, "funder"),
            "Grant Purpose": _grant_cell(grants, gi, "purpose"),
            "Ask Amount": _format_amount(_grant_cell(grants, gi, "amount")),
            "Funded Amount": _format_amount(_grant_cell(grants, gi, "received")),
            "Grant Remaining": _format_amount(_grant_cell(grants, gi, "remaining")),
            "Grant Award Date": _format_date_mmddyyyy(_grant_cell(grants, gi, "award date")),
            "Funded Date": _format_date_mmddyyyy(_grant_cell(grants, gi, "grant start date")),
            "Close Date": _format_date_mmddyyyy(_grant_cell(grants, gi, "grant end date")),
            "Grant Report Date": _format_report_date_multiline(_grant_cell(grants, gi, "report date")),
            "Grant Report Notes": combine_grant_report_notes(report_notes_val, notes_val),
            "Grant Note": "",
        }

        matched_ci = []
        for ci in range(len(contacts)):
            found = _cell(contacts, ci, cmap["foundation"])
            if _norm_key(found) == funder_key:
                matched_ci.append(ci)

        if not matched_ci:
            row = {c: "" for c in OUTPUT_COLUMNS}
            row.update(grant_block)
            row["IsCompany?"] = "Yes"
            row["Country"] = "United States"
            rows_out.append(row)
            continue

        for ci in matched_ci:
            row = {c: "" for c in OUTPUT_COLUMNS}
            row.update(grant_block)
            row["IsCompany?"] = "Yes"
            row["Country"] = "United States"
            row["First Name"] = _cell(contacts, ci, cmap["first"])
            row["Last Name"] = _cell(contacts, ci, cmap["last"])
            row["Job Title"] = _cell(contacts, ci, cmap["title"])
            row["Email 1"] = _cell(contacts, ci, cmap["email"])
            row["Phone 1"] = _cell(contacts, ci, cmap["phone"])
            row["Address Line 1"] = _cell(contacts, ci, cmap["addr1"])
            row["Address Line 2"] = _cell(contacts, ci, cmap["addr2"])
            row["City"] = _cell(contacts, ci, cmap["city"])
            row["State/Province"] = _cell(contacts, ci, cmap["state"])
            row["Zip"] = _cell(contacts, ci, cmap["zip"])
            rows_out.append(row)

    out = pd.DataFrame(rows_out)
    for c in OUTPUT_COLUMNS:
        if c not in out.columns:
            out[c] = ""
    out = out[OUTPUT_COLUMNS]
    return out.astype(object).fillna("").map(lambda x: "" if x is None else str(x))


def main() -> int:
    p = argparse.ArgumentParser(description="Merge grants + contacts xlsx into one migration export.")
    p.add_argument("grants_xlsx", type=Path, help="Input file 1 (grants: Funder, Purpose, Amount, …)")
    p.add_argument("contacts_xlsx", type=Path, help="Input file 2 (contacts: First Name, Foundation, …)")
    p.add_argument(
        "--debug-headers",
        action="store_true",
        help="Print all detected headers from both input files before merging.",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("merged_export.xlsx"),
        help="Output xlsx path (default: merged_export.xlsx)",
    )
    args = p.parse_args()

    if not args.grants_xlsx.is_file():
        print(f"Not found: {args.grants_xlsx}", file=sys.stderr)
        return 1
    if not args.contacts_xlsx.is_file():
        print(f"Not found: {args.contacts_xlsx}", file=sys.stderr)
        return 1

    grants = load_grants(args.grants_xlsx)
    contacts = load_contacts(args.contacts_xlsx)
    if args.debug_headers:
        print_headers(grants, contacts)
    merged = merge_workbooks(grants, contacts)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    merged.to_excel(args.output, index=False, engine="openpyxl")
    print(f"Wrote {len(merged)} rows to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
