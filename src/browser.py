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

    # Use the browser already installed on the machine (Edge or Chrome). The
    # engine is no longer bundled into the exe (that made it ~200MB), so the
    # chosen browser must be installed - virtually all Windows PCs have Edge.
    try:
        context = pw.chromium.launch_persistent_context(channel=channel, **common_args)
    except Exception as exc:
        # If the requested one isn't found, try the other installed browser.
        other = "msedge" if channel == "chrome" else "chrome"
        try:
            context = pw.chromium.launch_persistent_context(channel=other, **common_args)
        except Exception:
            pw.stop()
            raise RuntimeError(
                f"Could not launch {browser_choice}. Make sure Microsoft Edge "
                f"or Google Chrome is installed on this PC. Original error: {exc}"
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
    """Fill the Patient No. field (id='patientNo') and click Search.

    This site is an AngularJS single-page app. Hard-reloading it with a
    navigation between every patient was tearing down and rebuilding all its
    scripts, which crashed the browser connection. So we only load the page
    ONCE (if we're not already on it), then for each patient just clear the
    field, type the new ID, and click Search - exactly like a human does,
    staying within the same loaded app."""
    # Only navigate if we're not already on the search page.
    on_page = False
    try:
        on_page = page.locator("#patientNo").count() > 0
    except Exception:
        on_page = False

    if not on_page:
        page.goto(SITE_URL, wait_until="domcontentloaded")
        try:
            page.locator("#patientNo").wait_for(state="visible", timeout=20000)
        except PWTimeout:
            pass

    page.bring_to_front()

    field = page.locator("#patientNo")
    field.click(timeout=8000)
    field.fill("", timeout=8000)
    field.fill(patient_id, timeout=8000)

    # Click the Search button (ng-click="SearchResult(true)").
    clicked = False
    try:
        btn = page.locator("[ng-click*='SearchResult']")
        if btn.count() > 0:
            btn.first.click(timeout=8000)
            clicked = True
    except Exception:
        pass
    if not clicked:
        try:
            page.get_by_role("button", name=re.compile("Search", re.I)).first.click(timeout=8000)
            clicked = True
        except Exception:
            pass
    if not clicked:
        # Last resort: press Enter in the field.
        field.press("Enter")

    # Wait for the grid to refresh. Angular updates the grid in place, so
    # give it a moment rather than waiting on navigation.
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass
    page.wait_for_timeout(2500)


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
    """Read the results grid.

    The site uses AngularJS ui-grid, which only renders the visible rows and
    recycles DOM nodes as you scroll - so scraping <div role='row'> elements
    misses rows and gives stale indexes. Instead we ask Angular for the
    grid's underlying data array directly via its scope, which gives us every
    row reliably regardless of scroll position. Falls back to DOM scraping if
    the Angular read doesn't work."""
    data = _read_grid_data_from_angular(page)
    if data:
        parsed: list[ResultRow] = []
        for i, entity in enumerate(data):
            department = _match_department(str(entity.get("TestDepartment", "")
                                               or entity.get("Department", "")))
            if department is None:
                continue
            visit_text = str(entity.get("VisitDate", "") or "")
            visit_dt = _parse_any_date(visit_text)
            parsed.append(ResultRow(
                visit_date_text=visit_text,
                visit_date=visit_dt,
                accession_no=str(entity.get("Accession", "") or entity.get("AccessionNo", "")),
                test_department=department,
                row_index=i,
                report_url="",  # reports open via JS using row ID, handled in fetch
            ))
        if parsed:
            return parsed

    # --- Fallback: DOM scrape (older behavior) ---
    rows_locator = page.get_by_role("row")
    count = rows_locator.count()
    parsed = []
    for i in range(count):
        row = rows_locator.nth(i)
        try:
            text = row.inner_text()
        except Exception:
            continue
        if not text.strip() or "Visit Date" in text:
            continue
        department = _extract_department(text)
        if department is None:
            continue
        visit_date_text = _extract_visit_date(text)
        parsed.append(ResultRow(
            visit_date_text=visit_date_text or "",
            visit_date=_parse_visit_date(visit_date_text) if visit_date_text else None,
            accession_no=_extract_first_number(text) or "",
            test_department=department,
            row_index=i,
            report_url="",
        ))
    return parsed


def _read_grid_data_from_angular(page: Page):
    """Pull the ui-grid's full data array out of Angular's scope. Returns a
    list of row dicts, or None if it can't be read."""
    js = r"""
    () => {
      try {
        if (!window.angular) return null;
        // Find the ui-grid element and read its grid data.
        const gridEl = document.querySelector('[ui-grid]') ||
                       document.querySelector('.ui-grid');
        if (gridEl && window.angular.element) {
          const scope = window.angular.element(gridEl).scope();
          if (scope) {
            // Try common locations for the grid's data array.
            const opts = scope.gridOptions || (scope.grid && scope.grid.options) ||
                         (scope.$parent && scope.$parent.gridOptions);
            if (opts && Array.isArray(opts.data)) {
              return opts.data.map(r => JSON.parse(JSON.stringify(r)));
            }
            if (scope.grid && Array.isArray(scope.grid.rows)) {
              return scope.grid.rows.map(row => JSON.parse(JSON.stringify(row.entity)));
            }
          }
        }
        return null;
      } catch (e) { return null; }
    }
    """
    try:
        return page.evaluate(js)
    except Exception:
        return None


def _match_department(text: str) -> str | None:
    known = ["All Services", "Chemistry Unit", "Complete Blood Count",
             "MISC. Unit", "Hormones Unit", "Serology Unit", "Urine Unit",
             "Haematology Unit", "Hematology Unit"]
    for dep in known:
        if dep.lower() in (text or "").lower():
            return "All Services" if dep == "All Services" else dep
    return None


def _parse_any_date(text: str) -> datetime | None:
    text = (text or "").strip()
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S",
                "%Y-%m-%dT%H:%M:%S", "%m/%d/%Y"):
        try:
            return datetime.strptime(text[:len(text)], fmt)
        except ValueError:
            continue
    # Try ISO-ish with milliseconds/timezone chopped.
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*[AP]M)", text)
    if m:
        return _parse_visit_date(m.group(1))
    return None


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
    """Open the report for the given row and return the PDF bytes.

    The report opens via an Angular ng-click handler
    (reportLinkClicked(row.entity.ID, ...)) in a NEW tab showing the viewer.
    We find the correct 'All Services' Report link in the DOM by matching on
    the visit date, click it, grab the new tab's URL, fetch that URL's bytes
    in the background with the logged-in session, then close the tab right
    away. We deliberately do NOT interact with the PDF viewer's controls."""
    referer = page.url

    report_link = _find_report_link_for_row(page, row)
    if report_link is None:
        raise RuntimeError("Could not locate the Report link for the chosen "
                           "All Services row in the results grid.")

    viewer_page = None
    for attempt in range(2):
        try:
            with context.expect_page(timeout=25000) as new_page_info:
                report_link.click(timeout=8000)
            viewer_page = new_page_info.value
            break
        except PWTimeout:
            if attempt == 0:
                page.wait_for_timeout(1500)
                continue
            raise RuntimeError(
                "The report tab did not open in time after clicking Report. "
                "This patient will be retried on the next browser restart."
            )
    if viewer_page is None:
        raise RuntimeError("Report tab failed to open.")

    try:
        try:
            viewer_page.wait_for_load_state("domcontentloaded", timeout=20000)
        except PWTimeout:
            pass
        viewer_page.wait_for_timeout(2500)

        report_url = viewer_page.url

        # The viewer page (ReportViewer.htm?<id>) loads the real PDF into a
        # hidden frame from Report_PDF.aspx?<same id>. So we take everything
        # after "ReportViewer.htm?" and fetch Report_PDF.aspx with it.
        pdf_bytes = b""
        if "ReportViewer.htm?" in report_url:
            query = report_url.split("ReportViewer.htm?", 1)[1]
            from urllib.parse import urljoin
            base = report_url.split("Pages/Master/", 1)[0]  # site root up to Pages/Master
            pdf_endpoint = urljoin(report_url, "Report_PDF.aspx?") + query
            # urljoin can mangle the raw query; build it manually to be safe.
            viewer_root = report_url.split("ReportViewer.htm?", 1)[0]
            pdf_endpoint = viewer_root + "Report_PDF.aspx?" + query
            try:
                pdf_bytes = _direct_fetch(context, pdf_endpoint, report_url)
            except Exception:
                pdf_bytes = b""

        # If that didn't yield a PDF, fall back to fetching the viewer URL and
        # looking inside for a nested frame source.
        if not pdf_bytes.startswith(b"%PDF"):
            body = _direct_fetch(context, report_url, referer)
            if body.startswith(b"%PDF"):
                pdf_bytes = body
            else:
                html_text = body.decode("utf-8", errors="ignore")
                nested_url = _find_nested_report_url(html_text, report_url)
                if nested_url:
                    nb = _direct_fetch(context, nested_url, report_url)
                    if nb.startswith(b"%PDF"):
                        pdf_bytes = nb

        return pdf_bytes
    finally:
        try:
            if not viewer_page.is_closed():
                viewer_page.close()
        except Exception:
            pass


def _find_report_link_for_row(page: Page, row: ResultRow):
    """Find the clickable 'Report' element in the grid row matching this
    ResultRow (matched by its visit date text, then department). Returns a
    locator or None."""
    # All the Report links in the grid.
    links = page.locator("[ng-click*='reportLinkClicked']")
    try:
        n = links.count()
    except Exception:
        n = 0

    # Prefer matching by the row's visit date text appearing in the same
    # grid row as the link.
    if row.visit_date_text:
        for i in range(n):
            link = links.nth(i)
            try:
                # Walk up to the row container and check its text.
                container = link.locator("xpath=ancestor::div[contains(@class,'ui-grid-row')][1]")
                if container.count() == 0:
                    container = link.locator("xpath=ancestor::tr[1]")
                text = container.first.inner_text(timeout=2000) if container.count() else ""
                if row.accession_no and row.accession_no in text:
                    return link
                if row.visit_date_text and row.visit_date_text[:10] in text \
                        and "All Services".lower() in text.lower():
                    return link
            except Exception:
                continue

    # Fallback: the row_index-th report link.
    if 0 <= row.row_index < n:
        return links.nth(row.row_index)
    if n > 0:
        return links.first
    return None


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
