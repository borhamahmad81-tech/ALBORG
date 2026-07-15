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
    """Launch (or reuse) a persistent Edge profile. Returns (playwright, context).

    Uses the copy of Microsoft Edge already installed on the machine
    (channel='msedge'), so nothing extra needs downloading for the browser
    itself. Only Playwright's small internal driver is needed, which the
    build bundles; if it's somehow missing we surface a clear message."""
    profile_dir.mkdir(parents=True, exist_ok=True)
    _clear_stale_profile_locks(profile_dir)

    try:
        pw = sync_playwright().start()
    except Exception as exc:
        raise RuntimeError(
            "Playwright's internal driver could not start. This usually means "
            "the driver wasn't bundled into the .exe. If you built it yourself, "
            "make sure the build ran 'playwright install' and used "
            "--collect-all playwright. Original error: " + str(exc)
        )

    common_args = dict(
        user_data_dir=str(profile_dir),
        headless=headless,
        viewport={"width": 1400, "height": 900},
    )

    # Prefer the user's installed Microsoft Edge. If that can't launch for any
    # reason, fall back to the Chromium engine bundled inside the .exe so the
    # program still works out of the box.
    try:
        context = pw.chromium.launch_persistent_context(channel="msedge", **common_args)
    except Exception:
        try:
            context = pw.chromium.launch_persistent_context(**common_args)
        except Exception as exc:
            pw.stop()
            raise RuntimeError(
                "Could not launch a browser (tried Edge, then bundled Chromium). "
                "Original error: " + str(exc)
            )

    return pw, context


def _clear_stale_profile_locks(profile_dir: Path) -> None:
    """If a previous run crashed instead of shutting down cleanly, Chromium
    can leave lock files behind that prevent the profile from opening again.
    Since this profile folder is dedicated to this program, it's safe to
    clear these before every launch."""
    for name in ("SingletonLock", "SingletonSocket", "SingletonCookie"):
        try:
            (profile_dir / name).unlink(missing_ok=True)
        except Exception:
            pass


def close_extra_tabs(context: BrowserContext, keep_page: Page) -> None:
    """Safety net: close any tab other than the main search tab. Guards
    against leftover report tabs from an earlier error piling up."""
    for p in list(context.pages):
        if p is not keep_page and not p.is_closed():
            try:
                p.close()
            except Exception:
                pass


def ensure_logged_in(context: BrowserContext, page: Page, timeout_seconds: int = 300) -> None:
    """Navigate to the site and, if a login form appears, pause and wait for
    the user to log in manually. Only needed the first time per profile."""
    page.goto(SITE_URL, wait_until="domcontentloaded")

    if _on_search_page(page):
        return

    print(">> Login required. Please log in to the Al Borg portal in the "
          "browser window that just opened. Waiting up to "
          f"{timeout_seconds}s for the Search Criteria page to appear...")
    print(">> IMPORTANT: don't close that browser window - just log in inside it.")

    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        status = _on_search_page(page)
        if status is True:
            print(">> Logged in. This session is now saved for future runs.")
            return
        if status is None:
            raise RuntimeError(
                "The browser window was closed (or crashed) while waiting for you "
                "to log in. Please run the program again and leave the browser "
                "window open until it starts searching patients on its own."
            )
        time.sleep(1.5)

    raise TimeoutError("Timed out waiting for manual login. Re-run and log in faster, "
                        "or increase timeout_seconds.")


def _on_search_page(page: Page) -> bool | None:
    """Returns True if the search page is visible, False if not (yet), or
    None if the page/browser itself has been closed."""
    try:
        return page.get_by_text("Search Criteria").first.is_visible(timeout=2000)
    except PWTimeout:
        return False
    except Exception as exc:
        if "closed" in str(exc).lower():
            return None
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
    """Click the Report link for the given row, wait for the embedded PDF
    viewer to load (this can take a few seconds for a real report), then
    click the viewer's built-in 'PDF' button and capture the resulting
    browser download - this is the same mechanism a human clicking it would
    trigger, so it's the most reliable way to get the actual file."""
    rows_locator = page.get_by_role("row")
    target_row = rows_locator.nth(row.row_index)
    report_link = target_row.get_by_text("Report", exact=True)

    with context.expect_page(timeout=20000) as new_page_info:
        report_link.click()
    viewer_page = new_page_info.value

    try:
        try:
            viewer_page.wait_for_load_state("load", timeout=30000)
        except PWTimeout:
            pass
        # Give the embedded viewer time to actually render the report - real
        # reports can take a few seconds, per live testing.
        viewer_page.wait_for_timeout(4000)

        pdf_bytes = _download_via_pdf_button(viewer_page)

        if pdf_bytes is None:
            # Fallback: fetch the resolved report URL directly with the same
            # session cookies, in case the PDF button isn't found this time.
            report_url = viewer_page.url
            referer = page.url
            pdf_bytes = _direct_fetch(context, report_url, referer)
            if not pdf_bytes.startswith(b"%PDF"):
                html_text = pdf_bytes.decode("utf-8", errors="ignore")
                nested_url = _find_nested_report_url(html_text, report_url)
                if nested_url:
                    pdf_bytes = _direct_fetch(context, nested_url, referer)

        return pdf_bytes
    finally:
        # No matter what happened above (success, error, or a mid-step
        # crash), always close this tab - leftover tabs from failed patients
        # were piling up and could crash the whole run if one got closed
        # manually while still in use.
        try:
            if not viewer_page.is_closed():
                viewer_page.close()
        except Exception:
            pass


def _download_via_pdf_button(viewer_page: Page) -> bytes | None:
    """Click the viewer's blue 'PDF' button (confirmed present in the Al
    Borg viewer) and capture the real file it downloads. Returns None if the
    button can't be found/clicked, so the caller can fall back."""
    try:
        pdf_button = viewer_page.get_by_role("button", name=re.compile(r"^PDF$", re.I))
        if pdf_button.count() == 0:
            pdf_button = viewer_page.get_by_text("PDF", exact=True)
        pdf_button.first.wait_for(state="visible", timeout=15000)
    except PWTimeout:
        return None

    try:
        with viewer_page.expect_download(timeout=20000) as download_info:
            pdf_button.first.click()
        download = download_info.value
        return Path(download.path()).read_bytes()
    except Exception:
        return None


def _direct_fetch(context: BrowserContext, url: str, referer: str) -> bytes:
    resp = context.request.get(url, headers={"Referer": referer})
    if not resp.ok:
        raise RuntimeError(f"Report request failed: HTTP {resp.status} for {url}")
    return resp.body()


def _find_nested_report_url(html_text: str, base_url: str) -> str | None:
    from urllib.parse import urljoin
    for pattern in (
        r'<embed[^>]+src=["\']([^"\']+)["\']',
        r'<iframe[^>]+src=["\']([^"\']+)["\']',
        r'<object[^>]+data=["\']([^"\']+)["\']',
    ):
        m = re.search(pattern, html_text, re.IGNORECASE)
        if m:
            return urljoin(base_url, m.group(1))
    return None


def dump_debug(page: Page, name: str) -> None:
    DEBUG_DIR.mkdir(exist_ok=True)
    page.screenshot(path=str(DEBUG_DIR / f"{name}.png"), full_page=True)
    (DEBUG_DIR / f"{name}.html").write_text(page.content(), encoding="utf-8")
