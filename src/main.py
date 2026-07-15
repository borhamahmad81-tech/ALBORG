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
from datetime import date, datetime
from pathlib import Path


class _Tee:
    """Write all console output to a log file as well as the screen, so the
    full run can be reviewed even after the window closes."""
    def __init__(self, log_path):
        self._log = open(log_path, "w", encoding="utf-8")
        self._stdout = sys.stdout

    def write(self, text):
        self._stdout.write(text)
        try:
            self._log.write(text)
            self._log.flush()
        except Exception:
            pass

    def flush(self):
        self._stdout.flush()
        try:
            self._log.flush()
        except Exception:
            pass

from browser import (
    launch_browser, ensure_logged_in, search_patient, get_result_rows,
    pick_target_row, fetch_report_pdf, dump_debug, close_extra_tabs,
    DEFAULT_PROFILE_DIR,
)
from excel_io import (
    read_patient_list, write_master_workbook,
    read_existing_progress, get_completed_patient_ids,
)
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
    p.add_argument("--restart", action="store_true",
                   help="Ignore any existing output file and start over from patient 1. "
                        "By default, if --output already exists, already-completed "
                        "patients are skipped and the run resumes from where it left off.")
    return p.parse_args()


def _resolve_input_path(input_arg: str) -> str:
    """If the given path exists, use it as-is. Otherwise try to find it -
    Windows hides file extensions by default, so a file you renamed to
    'patients.xlsx' may actually be saved as 'patients.xlsx.xlsx' without
    looking like it. If we can't find a sensible match, print exactly what
    IS in the folder so it's obvious what to fix."""
    here = Path(".")

    if Path(input_arg).is_file():
        return input_arg

    doubled = input_arg + Path(input_arg).suffix  # patients.xlsx -> patients.xlsx.xlsx
    if Path(doubled).is_file():
        print(f"Note: '{input_arg}' wasn't found, but '{doubled}' was - using that. "
              f"(Windows likely hid the real file extension when you renamed it.)")
        return doubled

    candidates = [
        f for f in here.glob("*.xlsx")
        if f.name.lower() not in ("patients_template.xlsx",)
        and "master_labs" not in f.name.lower()
    ]
    if len(candidates) == 1:
        print(f"Note: '{input_arg}' wasn't found, but '{candidates[0]}' was the only "
              f"other Excel file here - using that.")
        return str(candidates[0])

    all_files = [f.name for f in here.iterdir() if f.is_file()]
    print(f"\nCouldn't find the input file '{input_arg}' in this folder.")
    print("Files actually in this folder:")
    for f in all_files:
        print(f"  - {f}")
    print("\nMake sure your patient list is an .xlsx file in this same folder, "
          "and that run.bat / the --input value matches its exact name.")
    sys.exit(1)


def main():
    args = parse_args()

    sys.stdout = _Tee("run_log.txt")
    print(f"=== Run started {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    if args.month is None:
        args.month = date.today().strftime("%Y-%m")
    print(f"Restricting to month: {args.month}"
          + (" (falling back to older months if needed)" if args.allow_older else
             " (strict - patients with nothing this month will be listed under Errors)"))

    input_path = _resolve_input_path(args.input)
    patients = read_patient_list(input_path)
    if not patients:
        print("No patients found in the input file. Check the column headers.")
        sys.exit(1)
    print(f"Loaded {len(patients)} patients from {args.input}")

    all_results: list[dict] = []
    errors: list[dict] = []
    unparsed: list[dict] = []

    if args.restart:
        Path(args.output).unlink(missing_ok=True)
    else:
        all_results, _old_errors, unparsed = read_existing_progress(args.output)
        done_ids = get_completed_patient_ids(all_results)
        if done_ids:
            remaining = [p for p in patients if p.patient_id not in done_ids]
            print(f"Found existing '{args.output}' with {len(done_ids)} patients already "
                  f"done - resuming with the remaining {len(remaining)}. "
                  "(Use --restart to ignore this and start over.)")
            patients = remaining

    if not patients:
        print("Nothing left to do - every patient is already in the output file.")
        sys.exit(0)

    max_attempts = 8
    attempt = 0
    remaining = list(patients)

    while remaining and attempt < max_attempts:
        attempt += 1
        if attempt > 1:
            print(f"\n>> Restarting the browser automatically (attempt {attempt}/{max_attempts}) "
                  f"- {len(remaining)} patient(s) left...")
            time.sleep(3)

        pw = context = page = None
        try:
            pw, context = launch_browser(headless=args.headless)

            # If the profile restored leftover tabs from a previous crash,
            # reuse the first one and close the rest - always exactly one tab.
            if context.pages:
                page = context.pages[0]
                for extra in context.pages[1:]:
                    try:
                        extra.close()
                    except Exception:
                        pass
            else:
                page = context.new_page()

            ensure_logged_in(context, page)

            for i, patient in enumerate(remaining, start=1):
                label = f"[{i}/{len(remaining)}] {patient.patient_id} ({patient.name or 'no name'})"
                print(label)
                close_extra_tabs(context, page)
                try:
                    search_patient(page, patient.patient_id)
                    rows = get_result_rows(page)
                    print(f"    found {len(rows)} result rows; "
                          f"departments: {sorted({r.test_department for r in rows})}")

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

                    print(f"    picked All Services row dated {target.visit_date_text}")
                    pdf_bytes = fetch_report_pdf(context, page, target)
                    print(f"    downloaded {len(pdf_bytes)} bytes")

                    if not pdf_bytes.startswith(b"%PDF"):
                        saved_note = ""
                        if args.debug:
                            Path("debug").mkdir(exist_ok=True)
                            bad_path = Path("debug") / f"not_a_pdf_{patient.patient_id}.bin"
                            bad_path.write_bytes(pdf_bytes)
                            saved_note = f" Saved a copy to {bad_path} - send this file over."
                        raise RuntimeError(
                            f"Downloaded report was not a valid PDF ({len(pdf_bytes)} bytes)."
                            + saved_note
                            + (" Re-run with --debug to save a copy next time." if not args.debug else "")
                        )

                    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                        tmp.write(pdf_bytes)
                        tmp_path = tmp.name

                    parsed = parse_lab_pdf(tmp_path)
                    Path(tmp_path).unlink(missing_ok=True)
                    print(f"    PDF patient no: '{parsed.patient_no}'; "
                          f"parsed {len(parsed.results)} test rows")

                    searched_id = str(patient.patient_id).strip()
                    pdf_id = str(parsed.patient_no).strip()
                    if pdf_id and pdf_id != searched_id:
                        errors.append({
                            "patient_id": patient.patient_id, "patient_name": patient.name,
                            "stage": "id_mismatch",
                            "details": f"Searched for {searched_id} but the downloaded "
                                       f"PDF's own 'Patient No.' field says {pdf_id}. "
                                       "Results NOT saved - re-check this patient manually.",
                        })
                        if args.debug:
                            Path("debug").mkdir(exist_ok=True)
                            (Path("debug") / f"id_mismatch_{patient.patient_id}.pdf").write_bytes(pdf_bytes)
                        print(f"    !! ID MISMATCH: searched {searched_id}, "
                              f"PDF says {pdf_id} - skipped")
                        continue

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

                except Exception as exc:  # one bad patient shouldn't kill the batch
                    msg = str(exc).lower()
                    connection_dead = any(s in msg for s in (
                        "connection closed", "target page, context or browser has been closed",
                        "target closed", "browser has been closed", "websocket",
                        "page.goto", "has been closed",
                    ))
                    if connection_dead:
                        # The browser itself died - don't mark this patient as a
                        # permanent error. Re-raise so the outer loop relaunches
                        # the browser and retries this patient.
                        print(f"    !! Browser connection lost on this patient - "
                              f"will relaunch and retry.")
                        raise
                    errors.append({
                        "patient_id": patient.patient_id, "patient_name": patient.name,
                        "stage": "fetch_or_parse", "details": str(exc),
                    })
                    if args.debug:
                        try:
                            dump_debug(page, f"error_{patient.patient_id}")
                        except Exception:
                            pass
                    print(f"    !! ERROR: {exc}")

                try:
                    write_master_workbook(args.output, all_results, errors, unparsed)
                except Exception as save_exc:
                    print(f"    (warning: could not checkpoint-save this round: {save_exc})")

                time.sleep(2)  # be polite to the portal between patients

        except Exception as exc:
            # The browser/driver itself crashed (not a single-patient error) -
            # loop back around and relaunch automatically instead of stopping.
            print(f"\n!! Browser crashed: {exc}")

        finally:
            if context is not None:
                try:
                    context.close()
                except Exception:
                    pass
            if pw is not None:
                try:
                    pw.stop()
                except Exception:
                    pass

        # Recompute what's actually left: anyone who succeeded or already got
        # a logged error (e.g. "no data this month") is done; only patients
        # that never got a chance to finish (crashed mid-way) get retried.
        done_ids = get_completed_patient_ids(all_results)
        errored_ids = {str(e["patient_id"]) for e in errors}
        remaining = [p for p in patients
                     if p.patient_id not in done_ids and p.patient_id not in errored_ids]

    write_master_workbook(args.output, all_results, errors, unparsed)
    print(f"\nDone. {len(all_results)} test rows, {len(errors)} errors -> {args.output}")
    if remaining:
        print(f"{len(remaining)} patient(s) still incomplete after {max_attempts} attempts - "
              "run again later, or check run_log.txt.")
    if errors:
        print("See the 'Errors' sheet in the output file for details.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        print("\n" + "=" * 60)
        print("The program hit an error and stopped. Details below:")
        print("=" * 60)
        traceback.print_exc()
    finally:
        # Keep the window open so the message is readable when double-clicked.
        try:
            input("\nPress Enter to close this window...")
        except EOFError:
            pass
