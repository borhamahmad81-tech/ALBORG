"""
main.py - Alborg Lab Fetcher

Reads a list of patient IDs from an Excel file, pulls each patient's most
recent "All Services" lab report from the Al Borg results portal, parses
the PDF, and writes everything into one master Excel workbook.

USAGE (from source):
    python src/main.py --input patients.xlsx --output master_labs.xlsx

USAGE (compiled .exe, built by GitHub Actions - see .github/workflows/build.yml):
    AlborgLabFetcher.exe --input patients.xlsx --output master_labs.xlsx

First run will open an Edge window and ask you to log in to the portal
once. After that, your session is remembered (see browser.py) and future
runs won't ask again, unless the site logs you out.
"""

from __future__ import annotations

import argparse
import sys
import tempfile
import time
from datetime import date
from pathlib import Path

from browser import (
    launch_browser, ensure_logged_in, search_patient, get_result_rows,
    pick_target_row, fetch_report_pdf, dump_debug, DEFAULT_PROFILE_DIR,
)
from excel_io import read_patient_list, write_master_workbook
from pdf_parser import parse_lab_pdf


def parse_args():
    p = argparse.ArgumentParser(description="Fetch and consolidate Al Borg lab reports.")
    p.add_argument("--input", required=True, help="Excel file with Patient No / Name columns")
    p.add_argument("--output", default="master_labs.xlsx", help="Output workbook path")
    p.add_argument("--month", default=None,
                   help="Restrict to a specific month, format YYYY-MM. "
                        "Default: the current month (today's month/year).")
    p.add_argument("--allow-older", action="store_true",
                   help="If no 'All Services' report exists for the target month, "
                        "fall back to the most recent one from an earlier month "
                        "instead of reporting it as missing. Off by default.")
    p.add_argument("--headless", action="store_true",
                   help="Run the browser without a visible window (only works "
                        "after you've logged in at least once).")
    p.add_argument("--debug", action="store_true",
                   help="Save a screenshot + HTML dump of the search page on failure.")
    return p.parse_args()


def main():
    args = parse_args()

    if args.month is None:
        args.month = date.today().strftime("%Y-%m")
    print(f"Restricting to month: {args.month}"
          + (" (falling back to older months if needed)" if args.allow_older else
             " (strict - patients with nothing this month will be listed under Errors)"))

    patients = read_patient_list(args.input)
    if not patients:
        print("No patients found in the input file. Check the column headers.")
        sys.exit(1)
    print(f"Loaded {len(patients)} patients from {args.input}")

    pw, context = launch_browser(headless=args.headless)
    page = context.new_page()

    all_results: list[dict] = []
    errors: list[dict] = []
    unparsed: list[dict] = []

    try:
        ensure_logged_in(context, page)

        for i, patient in enumerate(patients, start=1):
            label = f"[{i}/{len(patients)}] {patient.patient_id} ({patient.name or 'no name'})"
            print(label)
            try:
                search_patient(page, patient.patient_id)
                rows = get_result_rows(page)

                if not rows:
                    errors.append({
                        "patient_id": patient.patient_id, "patient_name": patient.name,
                        "stage": "search", "details": "No results returned for this patient ID",
                    })
                    if args.debug:
                        dump_debug(page, f"no_results_{patient.patient_id}")
                    continue

                target = pick_target_row(rows, args.month, allow_older=args.allow_older)
                if target is None:
                    errors.append({
                        "patient_id": patient.patient_id, "patient_name": patient.name,
                        "stage": "select_row",
                        "details": f"No 'All Services' report found for {args.month} "
                                   "(no fallback to older months - pass --allow-older to permit that)",
                    })
                    if args.debug:
                        dump_debug(page, f"no_all_services_{patient.patient_id}")
                    continue

                pdf_bytes = fetch_report_pdf(context, page, target)

                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp.write(pdf_bytes)
                    tmp_path = tmp.name

                parsed = parse_lab_pdf(tmp_path)
                Path(tmp_path).unlink(missing_ok=True)

                for r in parsed.results:
                    all_results.append({
                        "patient_id": patient.patient_id, "patient_name": patient.name,
                        "accession_no": r.accession_no, "section": r.section,
                        "category": r.category, "test": r.test, "result": r.result,
                        "flag": r.flag, "unit": r.unit, "ref_range": r.ref_range,
                        "registered_on": r.registered_on, "reported_on": r.reported_on,
                        "contract": r.contract,
                    })

                for line in parsed.unparsed_lines:
                    unparsed.append({
                        "patient_id": patient.patient_id, "patient_name": patient.name,
                        "line": line,
                    })

                print(f"    -> {len(parsed.results)} test results ({target.visit_date_text})")

            except Exception as exc:  # keep going - one bad patient shouldn't kill the batch
                errors.append({
                    "patient_id": patient.patient_id, "patient_name": patient.name,
                    "stage": "fetch_or_parse", "details": str(exc),
                })
                if args.debug:
                    dump_debug(page, f"error_{patient.patient_id}")
                print(f"    !! ERROR: {exc}")

            time.sleep(1)  # be polite to the portal between patients

    finally:
        context.close()
        pw.stop()

    write_master_workbook(args.output, all_results, errors, unparsed)
    print(f"\nDone. {len(all_results)} test rows, {len(errors)} errors -> {args.output}")
    if errors:
        print("See the 'Errors' sheet in the output file for details.")


if __name__ == "__main__":
    main()
