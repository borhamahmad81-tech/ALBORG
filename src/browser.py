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
    report_url: str = ""  # direct href of the Report link, if capturable


def launch_browser(headless: bool = False, profile_dir: Path = DEFAULT_PROFILE_DIR,
                   browser_choice: str = "edge"):
    """Launch (or reuse) a persistent browser profile. Returns (playwright, context).

    browser_choice: 'edge' (default) uses installed Microsoft Edge, 'chrome'
    uses installed Google Chrome. Each browser gets its own profile folder so
    logins don't clash. Falls back to bundled Chromium if the chosen browser
    can't launch."""
    # Separate profile per browser so switching doesn't mix up sessions.
    profile_dir = profile_dir.parent / f"profile_{browser_choice}"
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
        accept_downloads=True,
        # Suppress the "file downloaded" notification bubble that pops up in
        # the corner after each report download - possible source of focus
        # interference between patients.
        args=["--disable-features=DownloadBubble,DownloadBubbleV2",
              "--disable-popup-blocking"],
    )

    channel = "chrome" if browser_choice == "chrome" else "msedge"

    # Try the chosen browser; if it can't launch, fall back to bundled Chromium
    # so the program still works out of the box.
    try:
        context = pw.chromium.launch_persistent_context(channel=channel, **common_args)
    except Exception:
        try:
            context = pw.chromium.launch_persistent_context(**common_args)
        except Exception as exc:
            pw.stop()
            raise RuntimeError(
                f"Could not launch {browser_choice} (also tried bundled Chromium). "
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
    """Fill the Patient No. field and click Search.

    Always reloads a fresh copy of the search page first, so results and
    field values from the PREVIOUS patient can never linger and get picked
    up by mistake."""
    page.goto(SITE_URL, wait_until="domcontentloaded")
    page.bring_to_front()  # make sure THIS tab has focus, not any popup/notification
    # Let the search form finish rendering.
    try:
        page.get_by_text("Search Criteria").first.wait_for(state="visible", timeout=15000)
    except PWTimeout:
        pass

    patient_field = _find_patient_no_field(page)
    page.bring_to_front()
    patient_field.click(timeout=8000)
    patient_field.fill("", timeout=8000)
    patient_field.fill(patient_id, timeout=8000)

    search_button = page.get_by_role("button", name=re.compile("Search", re.I))
    search_button.click(timeout=8000)

    # Wait for the results table (or a "no results" state) to render.
    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except PWTimeout:
        pass
    page.wait_for_timeout(1500)


def _find_patient_no_field(page: Page):
    """Locate the 'Patient No.' input, trying several real strategies in
    order and actually checking each one matches something (.count() > 0)
    before using it - a locator that LOOKS valid but matches nothing was
    previously causing the whole run to hang on click() instead of failing
    fast, so every strategy here is verified before use."""

    # 1. Proper <label for="..."> association.
    try:
        loc = page.get_by_label(re.compile("Patient No", re.I))
        if loc.count() > 0:
            return loc.first
    except Exception:
        pass

    # 2. ASP.NET-style forms often give inputs descriptive IDs/names even
    # without a real <label> link, e.g. id="...txtPatientNo...".
    try:
        loc = page.locator(
            "input[id*='patient' i], input[name*='patient' i], "
            "input[id*='PatNo' i], input[name*='PatNo' i]"
        )
        if loc.count() > 0:
            return loc.first
    except Exception:
        pass

    # 3. Placeholder text.
    try:
        loc = page.get_by_placeholder(re.compile("Patient", re.I))
        if loc.count() > 0:
            return loc.first
    except Exception:
        pass

    # 4. Plain visible text label, then the nearest input after it in the DOM.
    try:
        label = page.get_by_text(re.compile(r"Patient No\.?", re.I)).first
        if label.count() > 0:
            candidate = label.locator("xpath=following::input[1]")
            if candidate.count() > 0:
                return candidate
    except Exception:
        pass

    raise RuntimeError(
        "Could not find the 'Patient No.' field on the search page using any "
        "known strategy. Run with --debug so a screenshot/HTML of the page "
        "gets saved - send that over and I'll pinpoint the exact field."
    )


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

        # Try to grab the Report link's href directly, so we can fetch the
        # PDF without opening a viewer tab (that tab-open sequence was
        # crashing the browser connection between patients).
        report_url = ""
        try:
            link = row.locator("a").filter(has_text=re.compile("Report", re.I))
            if link.count() > 0:
                href = link.first.get_attribute("href", timeout=2000)
                if href:
                    report_url = href
        except Exception:
            pass

        parsed.append(ResultRow(
            visit_date_text=visit_date_text or "",
            visit_date=visit_date,
            accession_no=accession_no or "",
            test_department=department,
            row_index=i,
            report_url=report_url,
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
    """Get the report PDF.

    PRIMARY method (no new tab): if we captured the Report link's URL from
    the results table, fetch it directly using the logged-in session. This
    avoids opening Edge's PDF viewer tab entirely - that tab open/close
    sequence was crashing the browser connection between patients.

    FALLBACK method: if we don't have a URL or the direct fetch doesn't yield
    a PDF, fall back to the old approach of clicking Report and reading the
    viewer tab."""
    referer = page.url

    # --- Primary: direct fetch, no tab ---
    if row.report_url:
        full_url = row.report_url
        if full_url.startswith("/"):
            from urllib.parse import urljoin
            full_url = urljoin(SITE_URL, full_url)
        try:
            body = _direct_fetch(context, full_url, referer)
            if body.startswith(b"%PDF"):
                return body
            # Might be an HTML wrapper pointing at the real PDF.
            html_text = body.decode("utf-8", errors="ignore")
            nested = _find_nested_report_url(html_text, full_url)
            if nested:
                body2 = _direct_fetch(context, nested, referer)
                if body2.startswith(b"%PDF"):
                    return body2
        except Exception:
            pass  # fall through to tab-based method

    # --- Fallback: click Report, read the viewer tab ---
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
        viewer_page.wait_for_timeout(4000)

        report_url = viewer_page.url
        pdf_bytes = _direct_fetch(context, report_url, referer)
        if not pdf_bytes.startswith(b"%PDF"):
            html_text = pdf_bytes.decode("utf-8", errors="ignore")
            nested_url = _find_nested_report_url(html_text, report_url)
            if nested_url:
                pdf_bytes = _direct_fetch(context, nested_url, referer)
            else:
                maybe = _download_via_pdf_button(viewer_page)
                if maybe is not None:
                    pdf_bytes = maybe

        return pdf_bytes
    finally:
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
