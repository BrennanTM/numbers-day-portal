# Numbers Day automation: process documentation

Veterans' Employment and Training Service (VETS), U.S. Department of Labor. Version 0.5.0 of the tools, documentation dated October 16, 2026.

## 1. Purpose and products

Numbers Day is the monthly compilation of veteran labor market figures that VETS circulates by email after each Bureau of Labor Statistics (BLS) Employment Situation release and posts on dol.gov/agencies/vets/latest-numbers. The automation retrieves the figures, performs the calculations, produces the products in their standard form, checks them, and leaves two deliberate manual steps: a person reviews the result and a person sends it.

Products produced each month:

- **The email compilation** (the primary product): intro paragraph, headline, unemployment and labor force participation sentences for veterans and nonveterans, the seasonal-adjustment note, and three tables (disability status moving average, sex, race and Hispanic or Latino ethnicity), each table carrying an unemployment-rate row and a participation-rate row where the data exist.
- **The Word file** of the same content, used as the archive copy and as the body of the Outlook draft.
- **The web page text blocks**: the five text blocks of the Latest Numbers page in the page's own markup, ready to paste into the content management system.
- **The audit file** (JSON): every value with its series identifier, previous-month and year-ago values, BLS standard error, the disability calculation, the Table 16 reading, the quality-assurance results, warnings, and the time of generation.

Not produced: the four chart images on the web page. Their design is captured under `reference/`; generating them is a possible later step.

## 2. The monthly timeline

| When (Eastern) | What happens | Who or what |
|---|---|---|
| First Friday, 8:30 | BLS publishes the Employment Situation; the API carries the new month within minutes. | BLS |
| 8:30 | The tool waits for all 32 series to show the new month, writes the report, runs the checks, and opens an Outlook draft or, in the portal, shows the email for copying. The disability block reads TBD at this point. | Scheduled task, or one command, or the portal's Build button |
| about 9:15 | BLS emails "CPS-Unpublished Tables" to VETS with the monthly disability package PDF (`disability_YYYY.MM.pdf`). | BLS (James Borbely, cc Karen Kosanovich) |
| on arrival | The PDF is dropped on the portal or passed to the tool; Table 16 is read, the month's row is stored, the 12-month average is recomputed, and the draft is rebuilt with the disability block filled. | One command or one drag-and-drop |
| same morning | A reviewer reads the draft and the checks, edits if needed, and sends. | Person |
| same day | The web page text blocks are pasted into the CMS and the four chart images updated. | Person with CMS access |
| after sending | The disability CSV is exported from the portal (or already updated by the tool) and kept in the Numbers Day folder on OneDrive, with the audit file. | Person, or automatic |

The release dates come from the BLS schedule bundled with the tools (currently through December 2027). The tools know which month to expect on any given day, and `numbers-day check` answers whether it is published yet.

## 3. Components and where they live

| Component | What it is | Where |
|---|---|---|
| `numbersday.pyz` | Single-file command-line tool for Windows machines with Python 3.9 or newer; no installation and no packages beyond Python. Writes the Word file, the web blocks and the audit, reads the BLS PDF, opens the Outlook draft, registers scheduled tasks. | GitHub release assets at github.com/BrennanTM/numbers-day-portal/releases; a copy in the Numbers Day folder on OneDrive |
| The portal | A single web page that does the same job in the browser and copies the finished email, formatted, for Outlook. No server; nothing leaves the browser except calls to api.bls.gov. | brennantm.github.io/numbers-day-portal (source: the same repository) |
| Settings file | The BLS key, recipients, subject. Per user, never in the repository. | `C:\Users\<user>\.numbersday.env` |
| Disability CSV | One row per month: the four Table 16 rates and, from October 2026 on, the labor-force and unemployed counts. The record for the moving average. | Numbers Day folder on OneDrive (`table16.csv`) |
| Release schedule | Dates and reference months of the Employment Situation releases. | Bundled in the tool and the portal (`schedule.json`); refreshed once a year by hand |
| Word template | The standard email layout with placeholders; built from the September 2026 example. | Bundled in the tool |
| Source code and tests | Python package, unit tests, the template builder, the DOL machine checks. | Project folder on the development machine; the portal repository holds the page, this document and the release assets |

Commands (run `py numbersday.pyz --help` for all options):

| Command | Purpose |
|---|---|
| `run --wait --out DIR --disability CSV [--table16 PDF] [--web] [--draft]` | The monthly build: wait for the release, write the report, run the checks, optionally the web blocks and the Outlook draft |
| `table16 PDF [--csv CSV]` | Read the four disability rates from the BLS PDF and store the month's row |
| `draft FILE.docx` | Open an Outlook draft from an existing report |
| `check [--expect YYYY-MM]` | Is the expected month published? Exit 0 yes, 3 not yet |
| `fetch [--json]` | Print every series with current, previous-month and year-ago values |
| `schedule [--register ...] [--unregister]` | Upcoming release dates; create or remove one scheduled task per release |
| `config` | Show where the key and settings come from and that the template is intact |

## 4. Data sources

### 4.1 BLS Public Data API (api.bls.gov)

All rates in the email except the disability block. 32 series, requested together with a registered key (500 requests a day; without a key the limit is 25 a day for the whole DOL network). The key also returns the official series titles and BLS's published standard errors (available from January 2026 data onward).

| Series | Measure | Group (18 years and over) | Adjustment |
|---|---|---|---|
| LNS14049526 | Unemployment rate | Veterans (seasonally adjusted) | Seasonally adjusted |
| LNS14049601 | Unemployment rate | Nonveterans (seasonally adjusted) | Seasonally adjusted |
| LNU04049528 | Unemployment rate | Women, veterans | Not seasonally adjusted |
| LNU04049603 | Unemployment rate | Women, nonveterans | Not seasonally adjusted |
| LNU04049527 | Unemployment rate | Men, veterans | Not seasonally adjusted |
| LNU04049602 | Unemployment rate | Men, nonveterans | Not seasonally adjusted |
| LNU04049526 | Unemployment rate | Veterans (not seasonally adjusted) | Not seasonally adjusted |
| LNU04049601 | Unemployment rate | Nonveterans (not seasonally adjusted) | Not seasonally adjusted |
| LNU04049532 | Unemployment rate | White, veterans | Not seasonally adjusted |
| LNU04049607 | Unemployment rate | White, nonveterans | Not seasonally adjusted |
| LNU04049535 | Unemployment rate | Black or African American, veterans | Not seasonally adjusted |
| LNU04049610 | Unemployment rate | Black or African American, nonveterans | Not seasonally adjusted |
| LNU04049538 | Unemployment rate | Asian, veterans | Not seasonally adjusted |
| LNU04049613 | Unemployment rate | Asian, nonveterans | Not seasonally adjusted |
| LNU04049529 | Unemployment rate | Hispanic or Latino, veterans | Not seasonally adjusted |
| LNU04049604 | Unemployment rate | Hispanic or Latino, nonveterans | Not seasonally adjusted |
| LNS11349526 | Labor force participation rate | Veterans (seasonally adjusted) | Seasonally adjusted |
| LNS11349601 | Labor force participation rate | Nonveterans (seasonally adjusted) | Seasonally adjusted |
| LNU01349528 | Labor force participation rate | Women, veterans | Not seasonally adjusted |
| LNU01349603 | Labor force participation rate | Women, nonveterans | Not seasonally adjusted |
| LNU01349527 | Labor force participation rate | Men, veterans | Not seasonally adjusted |
| LNU01349602 | Labor force participation rate | Men, nonveterans | Not seasonally adjusted |
| LNU01349526 | Labor force participation rate | Veterans (not seasonally adjusted) | Not seasonally adjusted |
| LNU01349601 | Labor force participation rate | Nonveterans (not seasonally adjusted) | Not seasonally adjusted |
| LNU01349532 | Labor force participation rate | White, veterans | Not seasonally adjusted |
| LNU01349607 | Labor force participation rate | White, nonveterans | Not seasonally adjusted |
| LNU01349535 | Labor force participation rate | Black or African American, veterans | Not seasonally adjusted |
| LNU01349610 | Labor force participation rate | Black or African American, nonveterans | Not seasonally adjusted |
| LNU01349538 | Labor force participation rate | Asian, veterans | Not seasonally adjusted |
| LNU01349613 | Labor force participation rate | Asian, nonveterans | Not seasonally adjusted |
| LNU01349529 | Labor force participation rate | Hispanic or Latino, veterans | Not seasonally adjusted |
| LNU01349604 | Labor force participation rate | Hispanic or Latino, nonveterans | Not seasonally adjusted |

The seasonally adjusted veteran and nonveteran series feed the headline and the sentences; the not-seasonally-adjusted series feed the tables, as in the standard email. Missing months (October 2025, lapse in appropriations) arrive as a dash with a footnote and are treated as absent.

### 4.2 BLS disability package, Table 16 (unpublished)

The four disability-status rates are not on the API. BLS emails the package on release morning; Table 16, "Employment status of persons 18 years and over by veteran status, period of service, sex, race, Hispanic or Latino ethnicity, and disability status", begins on the first page near the end of the monthly file whose title is not marked Continued (page 381 of 387 in September 2026). The tool reads that page's "Total" block: in the "Persons with a disability" and "Persons with no disability" sections, the Veterans and Nonveterans rows, each eight numbers (population, labor force, participation rate, employed, employment-population ratio, unemployed, unemployment rate, not in labor force; thousands; a dash means suppressed). The quarterly file is rejected. BLS's covering note states that monthly rates with a denominator below 75,000 do not meet publication standards; such rates are flagged. Months can be reissued (BLS reissued April 2025 in June 2025); a newer file replaces the month's row.

These series exist nowhere public: the full BLS labor force database (68,628 series) has no veterans-by-disability cross, and estimates from the public CPS microdata differ from BLS's because BLS's veteran reweighting is not in the public weights. The email remains the source.

### 4.3 The VETS calculation workbook

"TEMPLATE FOR CALCULATIONS - Veterans with a Disability - Monthly Rolling Average" holds the monthly rates since January 2019 and the 12-month moving averages. Its sheet exported as CSV is read as is (title rows, a Month column in Excel's various date spellings, the four rate columns, and the workbook's own averages, which are used as a cross-check). It seeded the disability history.

### 4.4 BLS release schedule

www.bls.gov/schedule/news_release/empsit.htm, copied by hand into `schedule.json` because bls.gov refuses scripted requests. Releases are at 8:30 Eastern.

### 4.5 BLS reliability documentation

Standard errors per observation come from the API (aspects). The factors that turn a monthly standard error into the standard error of a one-month or twelve-month change come from BLS's parameters-and-factors workbook, table PF-5 (veterans), and are used only to annotate changes as within or beyond sampling error at the 90 percent level.

## 5. Calculations and wording

- **Headline**: "Veteran Unemployment Rate was X.X% in Month", from the seasonally adjusted veteran rate, one decimal as published.
- **Sentences**: "In Month YYYY, the veteran unemployment rate was X.X%, up from Y.Y% the previous month and down from Z.Z% the prior year." "Up from" when the current value is higher, "down from" when lower, "unchanged from the previous month" when equal; the same for nonveterans ("Also in Month YYYY, the comparable non-veteran unemployment rate was ..."), and for the participation rates ("The veteran labor force participation rate was ..."; "The comparable non-veteran participation rate was ..."). When the previous month has no data, the comparison uses the last published month and names it; when the prior-year month has no data, the sentence carries a visible marker and a warning, and the wording is settled by the reviewer. BLS's own wording for such gaps ("little changed from September") is the model.
- **Tables**: published monthly values, one decimal, not seasonally adjusted.
- **Disability moving average**: the simple mean of the monthly published rates over the twelve calendar months ending with the reference month; months without data are skipped and named in the footnote; the number of months averaged is stated in the caption and the table title. This is the method of the VETS workbook (Excel AVERAGE over the trailing rows) and reproduces its results; it is also the method BLS used for its 11-month 2025 annual averages. The alternative of summed unemployed over summed labor force is available once enough months of counts have accumulated, and would be a team decision.
- **Subject line**: "Numbers Day - Month YYYY" unless configured otherwise.
- **Web page blocks**: the page's own wording and markup (for example "Male Veterans"), the participation rates as an added line in the sex and race blocks, the Table 16 footnote once.

## 6. Validation and quality assurance

### 6.1 Checks run on every build

Recorded under `qa` in the audit file and shown in the portal's Checks panel. A fail returns exit code 1, opens no draft, and leaves the files for inspection.

| Check | Fail | Warn |
|---|---|---|
| Template filled | A placeholder left unfilled, or TBD although disability data were supplied | TBD because no Table 16 data this run |
| Series agree on latest month | | The API reports different latest months (mid-update); rebuild in a few minutes |
| Rates within plausible ranges | A missing value, or an unemployment rate outside 0 to 40, a participation rate outside 10 to 95 | |
| Sentences match the numbers | A sentence does not state its own rate, its "up from / down from" contradicts the numbers, or a comparison has no value | |
| Consistent with the previous run | | Last month's values differ from what last month's run recorded (BLS revision), or a rerun of the same month shows changed values |
| Changes against sampling error | | (informational) each headline change marked within or beyond sampling error |
| Disability window | | A missing month that is not a known survey gap; or no data (TBD) |
| Workbook cross-check | | The workbook's own average differs from the computed one |
| Table 16 reliability | | A suppressed rate, a labor force below 75,000, or a rate inconsistent with its counts |
| Release schedule | | The month is not in the schedule, or its release date is in the future |

### 6.2 Verification history

- September 2026: the tool reproduced the published email exactly (all 16 unemployment values and both sentences) and the live web page's values; the Word output renders identically.
- The disability average computed from the workbook CSV rounds to the workbook's own figures for September 2026 (6.5 / 7.9 / 3.4 / 4.0) and September 2025.
- Table 16 read from the September PDF (8.3 / 7.1 / 2.8 / 3.7) matches the workbook row.
- Standard errors computed with BLS's parameters agree with BLS's published values (0.46 against 0.50 for the September veteran rate).
- 59 unit tests cover parsing, sentences, the moving average and its gaps, the template, the PDF and workbook readers, the Outlook step (mocked), the checks, the web blocks and the scheduled tasks; they run offline against a recorded API response.
- On the DOL laptop: Python 3.14 runs the single file, Windows PowerShell 5.1 is unrestricted, Python reaches api.bls.gov directly, classic Outlook accepts the draft, a standard user can create scheduled tasks, and headless Edge renders HTML to PNG.

## 7. Remaining manual steps

1. **Review and send.** The reviewer reads the draft and the checks, edits wording if needed, and sends. Nothing in the tools sends mail.
2. **Table 16 arrival.** Until BLS offers the figures another way, someone drops the monthly PDF on the portal or reruns with `--table16` when the BLS email arrives (about 45 minutes after the release). A watcher that picks the attachment out of the mailbox is a possible next step.
3. **Web page.** Paste the five text blocks into the CMS and update the four chart images.
4. **Records.** Keep `table16.csv` and the audit files in the Numbers Day folder; export the CSV from the portal after each month if the portal was used.
5. **Once a year.** Copy the next year's release dates from the BLS schedule page into `schedule.json` and rebuild; the tools warn when fewer than 90 days remain.
6. **Each January.** Expect revisions to seasonally adjusted data and new population controls; the "Consistent with the previous run" check will flag them. BLS also publishes annual averages then.
7. **Key registration.** One BLS key per user, registered with a DOL address and kept in the settings file or the browser.
8. **Laptop state for scheduled runs.** The machine must be awake and signed in at 8:30; the task runs in the user's session so the draft can appear.
9. **Workbook upkeep** if the workbook is still maintained alongside: from the October 2026 row the 11-row AVERAGE ranges should return to 12 rows, or the published figure would drop November 2025; the cross-check will say so.
10. **Design changes.** Changes to the standard email layout are made in the template builder and reviewed on a rendered sample before use.

## 8. Operations

### 8.1 Install or update on a DOL machine

```powershell
$dir = Join-Path $env:OneDriveCommercial "Numbers Day"
Invoke-WebRequest https://github.com/BrennanTM/numbers-day-portal/releases/latest/download/numbersday.pyz -OutFile "$dir\numbersday.pyz"
py "$dir\numbersday.pyz" --version
```

### 8.2 Settings file

`C:\Users\<user>\.numbersday.env`:

```
BLS_API_KEY=...
NUMBERS_DAY_TO=
NUMBERS_DAY_CC=
NUMBERS_DAY_SUBJECT=Numbers Day - {month_year}
NUMBERS_DAY_ATTACH=no
```

### 8.3 Exit codes and troubleshooting

| Symptom | Meaning | Action |
|---|---|---|
| Exit 3, "Not yet" | The expected month is not published | Use `--wait` on release morning |
| Exit 1, "QA failed" | A check failed; files written for inspection | Read the check's detail; do not send |
| "invalid key" | Wrong or missing registration key | Fix the settings file; the key is masked in output |
| HTTP 503 from the API | BLS outage; the tool retries three times | Try again in a few minutes |
| "Outlook drafts need Windows with classic Outlook" | Not Windows, or new Outlook | Use the portal's copy, or the Word file |
| Portal: "Could not load the PDF reader" | cdnjs.cloudflare.com blocked | Type the four rates, or use the command-line tool |
| Different latest months across series | API mid-update | Rebuild in a few minutes |

### 8.4 Security and privacy

The key lives in the user's settings file or browser and travels only to api.bls.gov; it is never printed, logged or committed, and error messages from BLS that echo it are masked. No data leave the machine other than those requests. The portal repository contains no key and no data. Unpublished BLS figures (Table 16) stay in the user's files.

## 9. Maintenance calendar

| When | Task |
|---|---|
| Monthly, release day | Run, review, send; drop the PDF; paste the web blocks; export the CSV |
| January | Check the revision warnings; consider the annual-average note; confirm the disability window has twelve months again |
| November and December | Confirm the next year's release schedule is in `schedule.json` |
| When BLS changes a series, the PDF layout or the email format | Update the series map, the Table 16 reader or the template; the tests and the checks will show what broke |

## 10. Version history

| Version | Date | Change |
|---|---|---|
| 0.1.0 | 2026-10-13 | Report from the API, Word output, audit file, release schedule |
| 0.2.x | 2026-10-14 | Outlook draft; key masked in error output |
| 0.3.x | 2026-10-14 | Table 16 read from the BLS PDF; workbook CSV accepted |
| 0.4.0 | 2026-10-14 | Labor force participation rates in sentences and tables |
| 0.5.0 | 2026-10-16 | Quality-assurance checks recorded in the audit; web page text blocks; scheduled tasks per release date; this document |
