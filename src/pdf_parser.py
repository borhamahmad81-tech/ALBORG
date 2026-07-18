"""
pdf_parser.py

Parses Al Borg Diagnostics lab report PDFs (the "All Services" report,
which stacks one section per test department - CBC, Chemistry Unit, etc.
- one after another, usually one section per page).

Built and tested against a real sample report. The line-based regex
below matches the CBC and Chemistry Unit sections seen so far. Lab
systems often add other departments (urinalysis, serology, hormones)
with slightly different result formats (e.g. "Positive"/"Negative"
instead of a number). Anything that doesn't match a known pattern is
NOT silently dropped - it's collected in `unparsed_lines` so you can
see it in the output Excel and extend the regex if needed.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pdfplumber

# --- Header field patterns -------------------------------------------------

PATIENT_NO_RE = re.compile(r"Patient No\.\s*:\s*(\d+)")
ACCESSION_NO_RE = re.compile(r"Accession No\.\s*:\s*(\d+)")
REGISTERED_ON_RE = re.compile(r"Registered on\s*:\s*([\d/]+\s+[\d:]+\s*[AP]M)")
REPORTED_ON_RE = re.compile(r"Reported on\s*:\s*([\d/]+\s+[\d:]+\s*[AP]M)")
CONTRACT_RE = re.compile(r"Contract\s*:\s*(.+)")
AGE_SEX_RE = re.compile(r"Age\s*/\s*Sex\s*:\s*(\d+)\s*Year\s*/\s*(\w+)")

# Section title is the line directly above "Test Result Unit Ref. Range"
SECTION_TITLE_RE = re.compile(r"\n([^\n]+)\nTest\s+Result\s+Unit\s+Ref\.\s*Range")

TABLE_HEADER_LINE = "Test Result Unit Ref. Range"

# End-of-table markers - stop reading rows once one of these lines appears
STOP_MARKERS = ("Reviewed By:", "Verified By", "Printed By", "Page ", "All Rights Reseved")

# A normal numeric test result row, e.g.:
#   "Hemoglobin 10.6 L g/dl 12 - 15.5"
#   "MCHC 31.5 g/dl 30 - 37"                (no H/L flag)
#   "Alanine Aminotransferase (ALT) 11 U/L 5 - 31"
#   "Ferritin In Serum 2,000 H ng/mL 13 - 150"   (comma thousands separator)
#   "Troponin <0.01 ng/mL 0 - 0.04"              (< or > prefixed result)
#
# NOTE: the test-name group deliberately EXCLUDES digits. Previously it
# allowed them, so a result the number pattern couldn't match (e.g. the
# comma in "2,000") got swallowed into the test name and the parser then
# grabbed the first number of the REFERENCE RANGE as the result. Keeping
# digits out of the name means an unmatched result can no longer be hidden
# inside it - the line simply won't match and lands in unparsed_lines where
# it's visible instead of silently wrong.
NUMERIC_ROW_RE = re.compile(
    r"^(?P<test>[A-Za-z()/,.\-\s]+?)\s+"
    r"(?P<result>[<>]?\s*-?\d[\d,]*\.?\d*)\s+"
    r"(?:(?P<flag>[HL])\s+)?"
    r"(?P<unit>\S+)\s+"
    r"(?P<ref>.+)$"
)

# Same as above but for tests whose NAME legitimately contains a number,
# e.g. "Vitamin B12", "Hepatitis B", "CD4 Count". Tried only if the strict
# pattern above fails, and it requires a reference range to be present so a
# result can't be confused with part of the name.
NUMERIC_ROW_WITH_NUM_IN_NAME_RE = re.compile(
    r"^(?P<test>[A-Za-z][A-Za-z0-9()/,.\-\s]*?)\s+"
    r"(?P<result>[<>]?\s*-?\d[\d,]*\.?\d*)\s+"
    r"(?:(?P<flag>[HL])\s+)?"
    r"(?P<unit>\S+)\s+"
    r"(?P<ref>[\d<>].*)$"
)

# A qualitative result row (no ref range, e.g. "Blood Group A Positive"),
# kept as a fallback - less strict, only used if the numeric patterns fail.
QUALITATIVE_ROW_RE = re.compile(
    r"^(?P<test>[A-Za-z0-9()/,.\-\s]+?)\s+"
    r"(?P<result>[A-Za-z][A-Za-z\s]*)$"
)


def _clean_number(text: str) -> str:
    """Normalize a parsed result: strip spaces and thousands separators, so
    '2,000' becomes '2000' and '< 0.01' becomes '<0.01'."""
    return text.replace(",", "").replace(" ", "").strip()


@dataclass
class LabResult:
    patient_no: str
    accession_no: str
    section: str
    category: str
    test: str
    result: str
    flag: str
    unit: str
    ref_range: str
    registered_on: str
    reported_on: str
    contract: str


@dataclass
class ParsedReport:
    patient_no: str = ""
    accession_no: str = ""
    registered_on: str = ""
    reported_on: str = ""
    contract: str = ""
    results: list[LabResult] = field(default_factory=list)
    unparsed_lines: list[str] = field(default_factory=list)


def _extract_header_fields(page_text: str) -> dict:
    fields = {"patient_no": "", "accession_no": "", "registered_on": "",
              "reported_on": "", "contract": ""}
    if m := PATIENT_NO_RE.search(page_text):
        fields["patient_no"] = m.group(1)
    if m := ACCESSION_NO_RE.search(page_text):
        fields["accession_no"] = m.group(1)
    if m := REGISTERED_ON_RE.search(page_text):
        fields["registered_on"] = m.group(1)
    if m := REPORTED_ON_RE.search(page_text):
        fields["reported_on"] = m.group(1)
    if m := CONTRACT_RE.search(page_text):
        # stop at newline - Contract value is on its own visual line
        fields["contract"] = m.group(1).split("\n")[0].strip()
    return fields


def _extract_section_title(page_text: str) -> str:
    if m := SECTION_TITLE_RE.search(page_text):
        return m.group(1).strip()
    return "Unknown Section"


def _parse_rows(page_text: str) -> tuple[list[dict], list[str]]:
    """Return (rows, unparsed_lines) for the test table on a page."""
    if TABLE_HEADER_LINE not in page_text:
        return [], []

    after_header = page_text.split(TABLE_HEADER_LINE, 1)[1]
    lines = [ln.strip() for ln in after_header.split("\n") if ln.strip()]

    rows = []
    unparsed = []
    category = ""

    for line in lines:
        if any(line.startswith(marker) for marker in STOP_MARKERS):
            break

        m = NUMERIC_ROW_RE.match(line) or NUMERIC_ROW_WITH_NUM_IN_NAME_RE.match(line)
        if m:
            rows.append({
                "category": category,
                "test": m.group("test").strip(),
                "result": _clean_number(m.group("result")),
                "flag": m.group("flag") or "",
                "unit": m.group("unit").strip(),
                "ref_range": m.group("ref").strip(),
            })
            continue

        m = QUALITATIVE_ROW_RE.match(line)
        if m and len(m.group("result").split()) <= 3:
            rows.append({
                "category": category,
                "test": m.group("test").strip(),
                "result": m.group("result").strip(),
                "flag": "",
                "unit": "",
                "ref_range": "",
            })
            continue

        # No numeric/qualitative result on this line -> treat as a
        # sub-category heading (e.g. "Hemoglobin Level (Kidney)") and
        # remember it for subsequent rows.
        if re.search(r"[A-Za-z]", line) and not re.search(r"\d", line):
            category = line
        else:
            unparsed.append(line)

    return rows, unparsed


def parse_lab_pdf(path: str) -> ParsedReport:
    """Parse an Al Borg 'All Services' PDF into structured lab results.

    Each page is treated as one department section (CBC, Chemistry Unit,
    etc). Patient-level header fields (Patient No, Accession No, dates,
    contract) are read from the first page that contains them and are
    assumed constant across the whole report.
    """
    report = ParsedReport()

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text.strip():
                continue

            header = _extract_header_fields(text)
            for key in ("patient_no", "accession_no", "registered_on",
                        "reported_on", "contract"):
                if not getattr(report, key) and header[key]:
                    setattr(report, key, header[key])

            section_title = _extract_section_title(text)
            rows, unparsed = _parse_rows(text)

            for row in rows:
                report.results.append(LabResult(
                    patient_no=header["patient_no"] or report.patient_no,
                    accession_no=header["accession_no"] or report.accession_no,
                    section=section_title,
                    category=row["category"],
                    test=row["test"],
                    result=row["result"],
                    flag=row["flag"],
                    unit=row["unit"],
                    ref_range=row["ref_range"],
                    registered_on=header["registered_on"] or report.registered_on,
                    reported_on=header["reported_on"] or report.reported_on,
                    contract=header["contract"] or report.contract,
                ))

            for line in unparsed:
                report.unparsed_lines.append(f"[{section_title}] {line}")

    return report


if __name__ == "__main__":
    import sys
    import json

    target = sys.argv[1] if len(sys.argv) > 1 else "sample_report.pdf"
    parsed = parse_lab_pdf(target)
    print(f"Patient No: {parsed.patient_no}")
    print(f"Accession No: {parsed.accession_no}")
    print(f"Registered on: {parsed.registered_on}")
    print(f"Reported on: {parsed.reported_on}")
    print(f"Contract: {parsed.contract}")
    print(f"\n{len(parsed.results)} test results parsed:\n")
    for r in parsed.results:
        print(f"  [{r.section}] {r.test}: {r.result} {r.flag} {r.unit} (ref {r.ref_range})")
    if parsed.unparsed_lines:
        print(f"\n{len(parsed.unparsed_lines)} UNPARSED lines (review needed):")
        for line in parsed.unparsed_lines:
            print(f"  ! {line}")
