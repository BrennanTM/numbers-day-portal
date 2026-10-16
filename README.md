# VETS Numbers Day

Tools that build the monthly Numbers Day compilation of veteran labor market figures after each Bureau of Labor Statistics (BLS) Employment Situation release: data retrieval, calculations, formatting, and quality checks are automated; a person reviews the result and a person sends it.

- **Portal (browser):** https://brennantm.github.io/numbers-day-portal/ — builds the email from live BLS data and copies it, formatted, for Outlook. Nothing leaves the browser except calls to api.bls.gov.
- **Command-line companion (Windows, Python 3.9+):** [`numbersday.pyz`](https://github.com/BrennanTM/numbers-day-portal/releases/latest/download/numbersday.pyz), a single file. Writes the Word file, reads the BLS disability PDF, opens an Outlook draft, registers scheduled runs.
- **Full process documentation:** [PROCESS.md](PROCESS.md) (also [PROCESS.docx](PROCESS.docx)): products, timeline, every data source with the complete series table, calculations and wording rules, validation, remaining manual steps, operations, maintenance calendar.

## What it produces each month

| Product | Contents |
|---|---|
| Email compilation | Intro paragraph, headline, unemployment and labor force participation sentences for veterans and nonveterans, the seasonal-adjustment note, and three tables (disability-status moving average; sex; race and Hispanic or Latino ethnicity), each with an unemployment-rate row and a participation-rate row where the data exist |
| Word file | The same content, used as the archive copy and as the body of the Outlook draft |
| Web page text blocks | The five text blocks of dol.gov/agencies/vets/latest-numbers in the page's own markup, ready to paste into the CMS |
| Audit file (JSON) | Every value with its series identifier, previous-month and year-ago values, BLS standard error, the disability calculation, the Table 16 reading, the quality-assurance results, warnings, and the time of generation |

Not produced: the four chart images on the web page (their design is documented; generating them is a possible later step).

## Quick start

### Portal

1. Open https://brennantm.github.io/numbers-day-portal/ in Edge or Chrome.
2. Paste your BLS API registration key once and click Save (it stays in that browser only; register at https://data.bls.gov/registrationEngine/ with a DOL address).
3. Seed the disability history once by importing the calculation workbook exported as CSV, or the tool's own `table16.csv`. Each month after that, drop the BLS file `disability_YYYY.MM.pdf` on the page and it reads Table 16 itself.
4. Click **Build this month's email**, read the email and the Checks panel, then **Copy email for Outlook** and paste into a new message. **Copy web page blocks** gives the CMS text.
5. Click **Export CSV** after adding a month and keep the file with the records.

### Command line

```powershell
$dir = Join-Path $env:OneDriveCommercial "Numbers Day"
Invoke-WebRequest https://github.com/BrennanTM/numbers-day-portal/releases/latest/download/numbersday.pyz -OutFile "$dir\numbersday.pyz"
py "$dir\numbersday.pyz" --version
```

Settings file at `C:\Users\<you>\.numbersday.env`:

```
BLS_API_KEY=your-key
NUMBERS_DAY_TO=
NUMBERS_DAY_CC=
NUMBERS_DAY_SUBJECT=Numbers Day - {month_year}
NUMBERS_DAY_ATTACH=no
```

| Command | Purpose |
|---|---|
| `run --wait --out DIR --disability CSV [--table16 PDF] [--web] [--draft]` | The monthly build: wait for the release, write the report, run the checks, optionally the web blocks and the Outlook draft |
| `table16 PDF [--csv CSV]` | Read the four disability rates from the BLS PDF and store the month's row |
| `draft FILE.docx` | Open an Outlook draft from an existing report |
| `check [--expect YYYY-MM]` | Is the expected month published? Exit 0 yes, 3 not yet |
| `fetch [--json]` | Print every series with current, previous-month and year-ago values |
| `schedule [--register ...] [--unregister]` | Upcoming release dates; create or remove one scheduled task per release |
| `config` | Show where the key and settings come from and that the template is intact |

## Release morning

| When (Eastern) | What happens |
|---|---|
| First Friday, 8:30 | BLS publishes; the API carries the new month within minutes. The tool (scheduled task, one command, or the portal's Build) waits for all 32 series, writes the report, runs the checks, and opens the first draft. The disability block reads TBD at this point. |
| about 9:15 | BLS emails the disability package PDF to VETS. Drop it on the portal or rerun with `--table16`; the month's row is stored, the 12-month average recomputed, and the draft rebuilt with the disability block filled. |
| same morning | A reviewer reads the draft and the checks, edits if needed, and sends. |
| same day | The web page text blocks are pasted into the CMS and the chart images updated. |

Scheduled runs: `schedule --register --count 6 --out "Numbers Day\Reports" --disability "Numbers Day\table16.csv" --web --draft` creates one Windows task per upcoming release date (laptop awake and signed in); `schedule --unregister` removes them.

## Data sources

- **BLS Public Data API** (api.bls.gov): 32 series, 18 years and over: unemployment rates and labor force participation rates for veterans and nonveterans (seasonally adjusted, for the headline and sentences) and for women, men, White, Black or African American, Asian, and Hispanic or Latino veterans and nonveterans (not seasonally adjusted, for the tables), with BLS's published standard errors. A registered key allows 500 requests a day; the full series table is in [PROCESS.md](PROCESS.md#41-bls-public-data-api-apiblsgov).
- **BLS disability package, Table 16 (unpublished):** emailed by BLS on release morning. The first page of Table 16 near the end of the monthly file carries the Veterans and Nonveterans rows for persons with and without a disability. The quarterly file is rejected; rates with a labor force below 75,000 (BLS's monthly publication threshold) are flagged; a newer file replaces a month's row. These series exist nowhere public.
- **VETS calculation workbook:** the monthly rates since January 2019 and the 12-month averages. Its CSV export is read as is and its averages are used as a cross-check.
- **BLS release schedule:** copied by hand into the tools once a year (currently through December 2027), because bls.gov refuses scripted requests.

## Calculations

- Sentences follow the standard wording: "up from X% the previous month", "down from", or "unchanged from", for the previous month and the prior year; gaps in the data (October 2025) fall back to the last published month and are named.
- Tables carry the published monthly values, one decimal, not seasonally adjusted.
- The disability moving average is the simple mean of the monthly rates over the twelve calendar months ending with the reference month, missing months skipped and footnoted, the count stated in the caption. This is the VETS workbook's method and reproduces its results.

## Quality assurance

Every run records these checks in the audit file; the portal shows them after each build. A fail returns exit code 1 and opens no draft.

| Check | Fail | Warn |
|---|---|---|
| Template filled | A placeholder left unfilled, or TBD although disability data were supplied | TBD because no Table 16 data this run |
| Series agree on latest month | | The API reports different latest months (mid-update) |
| Rates within plausible ranges | A missing value or a rate outside the plausible range | |
| Sentences match the numbers | A sentence misstates its rate or its direction words contradict the numbers | |
| Consistent with the previous run | | Last month's values differ from what last month's run recorded (BLS revision) |
| Changes against sampling error | | (informational) each headline change marked within or beyond sampling error |
| Disability window | | A missing month that is not a known survey gap; or no data |
| Workbook cross-check | | The workbook's own average differs from the computed one |
| Table 16 reliability | | A suppressed rate, a labor force below 75,000, or a rate inconsistent with its counts |
| Release schedule | | The month is not in the schedule, or its release date is in the future |

Verification: the September 2026 compilation was reproduced exactly from the API; the disability average matches the workbook; the Table 16 reading matches the workbook row; 59 unit tests run offline against a recorded API response.

## Remaining manual steps

1. Review and send. Nothing in the tools sends mail.
2. Drop the BLS PDF when the email arrives (about 45 minutes after the release).
3. Paste the web page text blocks into the CMS and update the four chart images.
4. Keep `table16.csv` and the audit files in the Numbers Day folder.
5. Once a year, copy the next year's release dates into the schedule.
6. Each January, expect BLS revisions and new population controls; the consistency check flags them.
7. Register one BLS key per user.

## Troubleshooting

| Symptom | Meaning | Action |
|---|---|---|
| Exit 3, "Not yet" | The expected month is not published | Use `--wait` on release morning |
| Exit 1, "QA failed" | A check failed; files written for inspection | Read the check's detail; do not send |
| "invalid key" | Wrong or missing registration key | Fix the settings file |
| HTTP 503 from the API | BLS outage; the tool retries | Try again in a few minutes |
| "Outlook drafts need Windows with classic Outlook" | Not Windows, or new Outlook | Use the portal's copy, or the Word file |
| Portal: "Could not load the PDF reader" | cdnjs.cloudflare.com blocked | Type the four rates, or use the command-line tool |

## Privacy

The key lives in the user's settings file or browser and travels only to api.bls.gov; it is never printed, logged or committed. No data leave the machine other than those requests. This repository contains no key and no data; unpublished BLS figures stay in the user's files.

## Repository contents

| File | Purpose |
|---|---|
| `index.html` | The portal (served by GitHub Pages) |
| `PROCESS.md`, `PROCESS.docx` | Full process documentation |
| Releases | `numbersday.pyz`, the command-line companion |

## Version history

| Version | Date | Change |
|---|---|---|
| 0.1.0 | 2026-10-13 | Report from the API, Word output, audit file, release schedule |
| 0.2.x | 2026-10-14 | Outlook draft; key masked in error output |
| 0.3.x | 2026-10-14 | Table 16 read from the BLS PDF; workbook CSV accepted |
| 0.4.0 | 2026-10-14 | Labor force participation rates in sentences and tables |
| 0.5.0 | 2026-10-16 | Quality-assurance checks in the audit; web page text blocks; scheduled tasks per release; process documentation |
