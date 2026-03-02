"""Export tutors from the Excel system into tutors.json for the website.

Usage:
  python export_tutors_json.py \
    --xlsx "Math_and_Sciences_Hub_Full_Institution_System.xlsx" \
    --out  "tutors.json" \
    --year 2026

Notes:
- Reads sheet: TUTOR_REGISTER
- Publishes only non-empty Tutor Name rows.
- Creates a verification code like: MSH-TUT-YYYY-001
"""

import argparse
import json
import re
from datetime import date

import openpyxl


def make_display_name(full_name: str) -> str:
    name = (full_name or "").strip()
    parts = [p for p in re.split(r"\s+", name) if p]
    if len(parts) >= 2:
        return f"{parts[0][0].upper()}. {parts[-1].title()}"
    return name


def split_subjects(text: str):
    if not text:
        return []
    # split on comma/semicolon/slash
    parts = re.split(r"\s*[,;/]\s*", str(text).strip())
    return [p.strip() for p in parts if p.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help="Path to the Excel workbook")
    ap.add_argument("--out", default="tutors.json", help="Output JSON file")
    ap.add_argument("--year", type=int, default=date.today().year, help="Year in verification code")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    if "TUTOR_REGISTER" not in wb.sheetnames:
        raise SystemExit("Sheet 'TUTOR_REGISTER' not found.")

    ws = wb["TUTOR_REGISTER"]
    headers = [c.value for c in ws[1]]
    col = {h: i+1 for i, h in enumerate(headers) if h}

    # Expected headers (safe lookups)
    def get(row, header):
        idx = col.get(header)
        return ws.cell(row, idx).value if idx else None

    tutors = []
    seq = 1
    for r in range(2, ws.max_row + 1):
        name = get(r, "Tutor Name")
        if not name or not str(name).strip():
            continue

        subject_spec = get(r, "Subject Specialty")
        qualification = get(r, "Highest Qualification")
        institution = get(r, "University/Institution")
        status_raw = (get(r, "Status (Active/Inactive)") or "").strip().lower()

        code = f"MSH-TUT-{args.year}-{seq:03d}"
        seq += 1

        role = f"{subject_spec} Tutor" if subject_spec else "Tutor"
        subjects = split_subjects(subject_spec)

        status = "Verified" if status_raw == "active" else "Pending"

        tutors.append({
            "code": code,
            "displayName": make_display_name(str(name)),
            "name": str(name).strip(),
            "role": str(role).strip(),
            "qualification": (str(qualification).strip() if qualification else ""),
            "institution": (str(institution).strip() if institution else ""),
            "subjects": subjects,
            "status": status,
        })

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(tutors, f, ensure_ascii=False, indent=2)

    print(f"Exported {len(tutors)} tutors -> {args.out}")


if __name__ == "__main__":
    main()
