# Alborg Lab Fetcher

Pulls each patient's most recent "All Services" lab report from the Al
Borg Diagnostics results portal (`results.alborgdx.com`), parses the PDF,
and combines everyone's results into a single master Excel workbook.

## How it works

1. You give it an Excel file with just two columns — `Patient No` and
   `Name` — listing who to look up. See `patients_template.xlsx`. This
   file is only a list of patients; it has no lab data in it. The lab
   results only appear in the output file (`master_labs.xlsx`) after a run.
2. It opens Microsoft Edge (your real, installed Edge — not a separate
   download), searches each Patient No on the portal, finds the most
   recent "All Services" row, and opens its report.
3. It captures the PDF straight from the browser's network traffic (no
   manual downloading), parses out every test/result/unit/reference
   range, and writes it all into `master_labs.xlsx`.
4. Login is only needed once — your session is saved in a local Edge
   profile folder and reused on every future run.

## Getting the .exe (no Python needed on your machine)

This repo is set up to build the Windows `.exe` automatically:

1. Create a new GitHub repository and push this folder to it:
   ```
   git init
   git add .
   git commit -m "Initial version"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
2. Go to your repo's **Actions** tab on GitHub. The "Build Windows EXE"
   workflow runs automatically on push and takes a few minutes.
3. Once it finishes, either:
   - open the workflow run and download the `AlborgLabFetcher-windows`
     artifact (a zip containing the `.exe`), or
   - go to the repo's **Releases** page, where a release is created
     automatically with the `.exe` attached.
4. Copy `AlborgLabFetcher.exe` anywhere on your Windows PC and run it
   from a terminal (see Usage below) — no installation required.

Every time you push a change, a fresh `.exe` is built automatically.

## Usage

Open Command Prompt / PowerShell in the folder with the `.exe`:

```
AlborgLabFetcher.exe --input patients.xlsx --output master_labs.xlsx
```

By default it only pulls each patient's "All Services" report **from the
current month** — if a patient has nothing yet this month, they're listed
in the Errors sheet instead of pulling an older result.

Optional flags:
- `--month 2026-07` — check a specific month instead of the current one
- `--allow-older` — if nothing matches the target month, fall back to
  the most recent older report instead of reporting it as missing (off
  by default, since you said you don't want older-month data mixed in)
- `--headless` — hide the browser window (only use this after you've
  already logged in once with a visible window)
- `--debug` — if a patient fails, saves a screenshot + HTML of the page
  to a `debug/` folder so we can fix a selector quickly

**First run:** an Edge window opens on the portal. If it asks you to log
in, do so normally — the tool waits for you. After that it searches
every patient automatically.

## Output

`master_labs.xlsx` has:
- **Lab Results** — one row per test (Patient ID, Name, Accession No,
  Section/Department, Category, Test, Result, Flag, Unit, Ref Range,
  dates, Contract). High/Low results are colored for quick scanning.
- **Errors** — any patient whose search or report fetch failed, with a
  reason, so nothing silently goes missing.
- **Unparsed Lines** — any report line the parser didn't recognize
  (e.g. a test type not seen yet, like urinalysis or serology). Send me
  a sample and I'll extend the parser.

## Running from source instead (optional)

```
pip install -r requirements.txt
python src/main.py --input patients.xlsx --output master_labs.xlsx
```

## A note on the site automation

The parts of this code that read the *PDF report* (`src/pdf_parser.py`)
were built and tested against a real sample report and are solid. The
parts that drive the *website itself* (`src/browser.py`) were built from
a screenshot, not the live page, using flexible label-based selectors
that should work — but websites vary in small ways a screenshot can't
show. If the search field, Search button, or Report link isn't found on
your first real run, run with `--debug`, send me the files from the
`debug/` folder (a screenshot + the page's HTML), and I'll adjust the
selector — it's usually a one-line fix.

## Data handling note

This tool downloads and stores patient lab data locally in Excel files.
Handle `master_labs.xlsx` and the `debug/` folder the same way you'd
handle any other export of patient health information from the portal.
