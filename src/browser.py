"""
browser.py

Drives the Al Borg results portal (https://results.alborgdx.com/Result)
using Playwright with a PERSISTENT Edge profile. Because the profile is
saved to disk, you only log in once - every future run reuses the saved
session cookies automatically.

NOTE ON SELECTORS: this was built from a screenshot of the search page,
not the live HTML, so the selectors below use flexible, label-based
locators (get_by_label / get_by_role / get_by_text) rather than brittle
CSS IDs, which should survive most of the site's actual markup. If a
selector doesn't match on your first real run, set DEBUG_DUMP=True (or
pass --debug) so a screenshot + full HTML of the page is saved to
debug/ - send me that and I'll fix the selector immediately.
"""

from __future__ import annotations

import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, Page, BrowserContext, TimeoutError as PWTimeout

SITE_URL = "https://results.alborgdx.com/Result"
ALL_SERVICES_LABEL = "All Services"

# Where the persistent Edge profile (and thus your login session) lives.
DEFAULT_PROFILE_DIR = Path.home() / "AppData" / "Local" / "AlborgLabFetcher" / "edge_profile"

DEBUG_DIR = Path("debug")


@dataclass
class ResultRow:
    visit_date_text: str
    visit_date: datetime | None
    accession_no: str
    test_department: str
    row_index: int  # index among matching rows, used to re-locate the row


def launch_browser(headless: bool = False, profile_dir: Path = DEFAULT_PROFILE_DIR):
    """Launch (or reuse) a persistent Edge profile. Returns (playwright, context)."""
    profile_dir.mkdir(parents=True, exist_ok=True)
    pw = sync_playwright().start()
    context = pw.chromium.launch_persistent_context(
        user_data_dir=str(profile_dir),
        channel="msedge",
        headless=headless,
        viewport={"width": 1400, "height": 900},
    )
    return pw, context


def ensure_logged_in(context: BrowserContext, page: Page, timeout_seconds: int = 300) -> None:
    """Navigate to the site and, if a login form appears, pause and wait for
    the user to log in manually. Only needed the first time per profile."""
    page.goto(SITE_URL, wait_until="domcontentloaded")

    if _on_search_page(page):
        return

    print(">> Login required. Please log in to the Al Borg portal in the "
          "browser window that just opened. Waiting up to "
          f"{timeout_seconds}s for the Search Criteria page to appear...")

    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if _on_search_page(page):
            print(">> Logged in. This session is now saved for future runs.")
            return
        time.sleep(1.5)

    raise TimeoutError("Timed out waiting for manual login. Re-run and log in faster, "
                        "or increase timeout_seconds.")


def _on_search_page(page: Page) -> bool:
    try:
        return page.get_by_text("Search Criteria").first.is_visible(timeout=2000)
    except PWTimeout:
        return False


def search_patient(page: Page, patient_id: str) -> None:
    """Fill the Patient No. field and click Search."""
    if not _on_search_page(page):
        page.goto(SITE_URL, wait_until="domcontentloaded")

    patient_field = _find_patient_no_field(page)
    patient_field.click()
    patient_field.fill("")
    patient_field.fill(patient_id)

    search_button = page.get_by_role("button", name=re.compile("Search", re.I))
    search_button.click()

    # Wait for the results table (or a "no results" state) to render.
    page.wait_for_load_state("networkidle", timeout=20000)


def _find_patient_no_field(page: Page):
    """Locate the 'Patient No.' input robustly (label text, then fallback
    to the input right after the 'Patient No.' text label)."""
    try:
        return page.get_by_label(re.compile("Patient No", re.I))
    except Exception:
        pass
    # Fallback: find the label element, then the nearby input.
    label = page.get_by_text(re.compile(r"^Patient No\.?$"), exact=False).first
    return label.locator("xpath=following::input[1]")


def get_result_rows(page: Page) -> list[ResultRow]:
    """Parse the Results table into structured rows."""
    rows_locator = page.get_by_role("row")
    count = rows_locator.count()
    parsed: list[ResultRow] = []

    for i in range(count):
        row = rows_locator.nth(i)
        text = row.inner_text()
        if not text.strip() or "Visit Date" in text:
            continue  # header row

        visit_date_text = _extract_visit_date(text)
        visit_date = _parse_visit_date(visit_date_text) if visit_date_text else None
        accession_no = _extract_first_number(text)
        department = _extract_department(text)

        if department is None:
            continue

        parsed.append(ResultRow(
            visit_date_text=visit_date_text or "",
            visit_date=visit_date,
            accession_no=accession_no or "",
            test_department=department,
            row_index=i,
        ))

    return parsed


def _extract_visit_date(row_text: str) -> str | None:
    m = re.search(r"\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*[AP]M", row_text)
    return m.group(0) if m else None


def _parse_visit_date(text: str) -> datetime | None:
    try:
        return datetime.strptime(text, "%m/%d/%Y %I:%M:%S %p")
    except ValueError:
        return None


def _extract_first_number(row_text: str) -> str | None:
    m = re.search(r"\b\d{6,}\b", row_text)
    return m.group(0) if m else None


def _extract_department(row_text: str) -> str | None:
    # Known department labels seen in the portal; extend this list if the
    # lab adds new departments.
    known = ["All Services", "Chemistry Unit", "Complete Blood Count",
             "MISC. Unit", "Hormones Unit", "Serology Unit", "Urine Unit"]
    for dep in known:
        if dep.lower() in row_text.lower():
            return dep
    return None


def pick_target_row(rows: list[ResultRow], month_filter: str | None,
                     allow_older: bool = False) -> ResultRow | None:
    """Pick the most recent 'All Services' row within the given month
    ('YYYY-MM'). By default this is STRICT: if nothing matches that month,
    returns None (caller records it as "no result this month") rather than
    silently falling back to an older visit. Pass allow_older=True to permit
    that fallback."""
    all_services = [r for r in rows if r.test_department == ALL_SERVICES_LABEL and r.visit_date]
    if not all_services:
        return None

    if month_filter:
        year, month = map(int, month_filter.split("-"))
        in_month = [r for r in all_services
                    if r.visit_date.year == year and r.visit_date.month == month]
        if in_month:
            return max(in_month, key=lambda r: r.visit_date)
        if not allow_older:
            return None

    return max(all_services, key=lambda r: r.visit_date)


def fetch_report_pdf(context: BrowserContext, page: Page, row: ResultRow) -> bytes:
    """Click the Report link for the given row and capture the raw PDF bytes
    from the network response of the new tab that opens."""
    rows_locator = page.get_by_role("row")
    target_row = rows_locator.nth(row.row_index)
    report_link = target_row.get_by_text("Report", exact=True)

    pdf_bytes_holder: dict = {}

    def handle_response(response):
        try:
            ctype = response.headers.get("content-type", "")
            if "pdf" in ctype.lower() and "pdf" not in pdf_bytes_holder:
                pdf_bytes_holder["pdf"] = response.body()
        except Exception:
            pass

    with context.expect_page(timeout=20000) as new_page_info:
        report_link.click()
    new_page = new_page_info.value
    new_page.on("response", handle_response)

    try:
        new_page.wait_for_load_state("networkidle", timeout=20000)
    except PWTimeout:
        pass

    # Give any late PDF network responses a moment to arrive.
    deadline = time.time() + 10
    while "pdf" not in pdf_bytes_holder and time.time() < deadline:
        time.sleep(0.5)

    if "pdf" not in pdf_bytes_holder:
        # Fallback: some viewers embed the PDF in an <embed>/<iframe> whose
        # src we can fetch directly via the page's own request context.
        src = None
        for sel in ("embed[type='application/pdf']", "iframe"):
            try:
                src = new_page.locator(sel).first.get_attribute("src", timeout=2000)
                if src:
                    break
            except PWTimeout:
                continue
        if src:
            resp = new_page.request.get(src)
            pdf_bytes_holder["pdf"] = resp.body()

    new_page.close()

    if "pdf" not in pdf_bytes_holder:
        raise RuntimeError(
            "Could not capture the PDF response for this report. "
            "Run with debug=True to save a screenshot/HTML for troubleshooting."
        )

    return pdf_bytes_holder["pdf"]


def dump_debug(page: Page, name: str) -> None:
    DEBUG_DIR.mkdir(exist_ok=True)
    page.screenshot(path=str(DEBUG_DIR / f"{name}.png"), full_page=True)
    (DEBUG_DIR / f"{name}.html").write_text(page.content(), encoding="utf-8")
