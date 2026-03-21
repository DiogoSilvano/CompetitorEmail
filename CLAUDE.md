# CLAUDE.md — CompetitorEmail

This file provides context for AI assistants working on this codebase.

---

## Project Overview

**CompetitorEmail** is a Python-based business intelligence automation tool that:

1. Scrapes article content from competitor email newsletter links (fed in via Microsoft Power Automate)
2. Fetches and filters relevant articles from RSS feeds (UK telecom / streaming sector)
3. Scores and ranks articles by relevance
4. Deduplicates articles against an existing database using a three-stage algorithm
5. Persists results **incrementally** to an Excel workbook synced to SharePoint via Power Automate

**Domain:** UK telecommunications and streaming media competitive intelligence (Sky internal tooling).

**Known sources:** ISPreview, BroadbandTVNews, ThinkBroadband, FibreProvider, and similar UK industry outlets.

---

## Repository Structure

```
CompetitorEmail/
├── scraper.py        # Main email-link scraper and deduplication engine (904 lines)
├── rss_scraper.py    # RSS feed fetcher and scraper (489 lines)
└── Scraper           # Legacy binary duplicate of scraper.py — do not edit
```

There is no package manager manifest (`requirements.txt`, `pyproject.toml`, etc.). Dependencies are managed manually.

---

## Technology Stack

| Concern | Library |
|---|---|
| HTTP scraping | `requests` (with persistent `SESSION`) |
| HTML parsing | `BeautifulSoup4` (`lxml` parser) |
| JavaScript rendering | `selenium` + headless Chrome via `webdriver-manager` |
| PDF extraction | `pdfplumber` (optional, imported conditionally) |
| RSS parsing | `feedparser` |
| Data frames | `pandas` |
| Numerical ops | `numpy` |
| TF-IDF dedup | `scikit-learn` (`TfidfVectorizer`, `cosine_similarity`) |
| Excel I/O | `pandas` (read/write) + `openpyxl` (table metadata only) |

---

## Running the Scripts

```bash
# Process email newsletter links from the database
python scraper.py

# Fetch and process RSS feeds
python rss_scraper.py
```

Both scripts are designed for **Windows execution** and will prompt `input()` at exit (for batch/PowerShell windows that close automatically). They log detailed progress to stdout.

---

## Hardcoded Paths (Windows)

Both scripts use absolute Windows paths that must be updated if the environment changes:

**scraper.py** (lines 706–708):
```python
DATA_DIR   = r'C:\Users\dsn24\OneDrive - Sky\...\data'
OUTPUT_DIR = r'C:\Users\dsn24\OneDrive - Sky\...\output'
RSS_DIR    = r'C:\Users\dsn24\OneDrive - Sky\...\RSS Parser\output'
```

**rss_scraper.py** (lines 396–398):
```python
SCRIPT_DIR = r'C:\Users\dsn24\OneDrive - Sky\...\RSS Parser'
RSS_FILE   = os.path.join(SCRIPT_DIR, 'rss_feeds.txt')
```

> When porting to a new machine or environment, update all `DATA_DIR`, `OUTPUT_DIR`, `RSS_DIR`, and `SCRIPT_DIR` constants.

---

## Data Flow

```
Power Automate
     │  writes pending email links
     ▼
data/Competitor_Email_DB.xlsx   ──►  scraper.py  ──►  output/Competitor_Email_DB.xlsx
                                          ▲                       ▲
                                    deduplicates             merges RSS
                                          │
                              RSS Parser/output/RSS_Articles_W##_YYYY_from########.xlsx
                                          ▲
                                    rss_scraper.py
                                          ▲
                                    rss_feeds.txt  (one URL per line)
```

### Excel Files

| File | Sheet | Named Table | Style |
|---|---|---|---|
| `Competitor_Email_DB.xlsx` | `Articles` | `ArticlesTable` | `TableStyleMedium9` |
| `RSS_Articles_W{wk}_{yr}_from{date}.xlsx` | `RSS_Articles` | `RSSArticlesTable` | `TableStyleMedium9` |

---

## Key Global Constants

| Constant | Location | Purpose |
|---|---|---|
| `HEADERS_LIST` | both files | Two rotating User-Agent headers (Chrome/Safari) |
| `SESSION` | both files | Persistent `requests.Session()` for connection reuse |
| `SIMILARITY_THRESHOLD` | `scraper.py:49` | TF-IDF cosine similarity cutoff (`0.65`) |
| `SHEET_NAME` | `scraper.py` | Target Excel sheet name |
| `TABLE_NAME` | `scraper.py` | Named Excel table identifier |

---

## Core Algorithms

### Scraping Strategy (scraper.py & rss_scraper.py)

Articles are fetched with a two-method fallback:

1. **`scrape_with_requests()`** — Plain HTTP with rotating User-Agents and `requests.Session`. Parses with BeautifulSoup (`lxml`). Validates minimum text length (100–200 chars).
2. **`scrape_with_selenium()`** — Headless Chrome fallback for JS-rendered pages. Anti-detection options enabled. Random delays (3–5 s).
3. **`scrape_pdf()`** — Invoked when URL ends in `.pdf`; uses `pdfplumber` for text extraction.

Between requests, random delays of **1.5–3.5 seconds** are injected to avoid rate-limiting.

### Three-Stage Deduplication (`deduplicate_rss_against_db()` in scraper.py)

RSS articles are checked against the existing database in three stages (any match = duplicate):

| Stage | Method | Detail |
|---|---|---|
| 1 | URL normalisation | Strip query strings and fragments, compare canonical URLs |
| 2 | Title entity overlap | Extract named entities from titles; check set intersection |
| 3 | TF-IDF cosine similarity | Max 5 000 features, unigrams+bigrams, English stop words; threshold `0.65` |

### Reporting Week Logic

The system uses **Friday–Thursday reporting weeks**, labelled by the closing Thursday.

- `get_week_ending()` in `scraper.py` — returns the EmailWeekEnding date string (`DD/MM/YYYY`)
- `get_current_week_friday()` / `get_last_friday()` in `rss_scraper.py` — returns the Friday that opened the current week

Dates are stored as `DD/MM/YYYY  00:00:00` to match Power Automate's expected format.

### RSS Relevance Filtering (`is_relevant_article()`)

Articles are accepted only if their title or summary contains at least one keyword from a curated UK telecom / streaming keyword list (case-insensitive). Modify this list in `rss_scraper.py` to broaden or narrow coverage.

---

## Code Conventions

### Naming
- **Functions and variables:** `snake_case`
- **Module-level constants:** `UPPER_SNAKE_CASE`
- **Function names:** descriptive verb phrases — `get_source_name()`, `normalise_url()`, `sanitise_str()`, `extract_entities()`

### Error Handling
- Every network call is wrapped in `try/except`; failures log a warning and return `None` or an empty record — never crash the whole run.
- HTTP 403 / 429 responses trigger immediate fallback to Selenium.
- Binary / non-text responses are detected by content-type and skipped.

### Excel Safety
- `sanitise_str()` strips XML 1.0 illegal characters (control chars, surrogates, Unicode non-characters) before writing cells, preventing workbook corruption.
- Final Excel writes use a **temp-file + `shutil.move()`** pattern for atomicity.
- `openpyxl` is used **only** for adding/updating the named table metadata after `pandas` writes the data; do not mix them for cell-level access.

### Console Output
- Progress is reported as `n/N (xx%)` with emoji indicators for success/warning/error.
- Deduplication logs each decision with the matching stage and matched title for auditability.

### Section Headers
Code sections are delimited with:
```python
# =============================================================================
# SECTION NAME
# =============================================================================
```

---

## Secrets and Credentials

Credentials are stored in **Windows Credential Manager** using `keyring` / `msal`. **Never hardcode secrets, passwords, or tokens in source files.** Any new authentication requirement must use the existing Credential Manager pattern.

---

## STRICT RULES — Do Not Change Without Explicit Instruction

These constraints exist because downstream Power Automate flows and SharePoint integrations depend on exact formats and table structures.

| Rule | Reason |
|---|---|
| **Do NOT replace the Excel file wholesale** | Rows must be deleted and re-added incrementally; wholesale replacement breaks Power Automate flow triggers |
| **Do NOT change Excel table structure or column names** | Power Automate reads columns by exact name; any rename silently breaks downstream flows |
| **Do NOT change the scoring/ranking formula** | Output ranking is calibrated and validated; changes affect business decisions |
| **Do NOT modify the deduplication logic** (URL match + TF-IDF + named entity overlap) | Three-stage logic is intentionally tuned; changes alter what appears in the weekly digest |
| **DigestID format must always be `YYYY-WXX`** (e.g. `2026-W09`) | Used as a stable key by Power Automate and SharePoint lookups |

---

## No Tests

There is currently **no test suite**. When adding tests, use `pytest` and place files under a `tests/` directory. Mirror the module structure (`tests/test_scraper.py`, `tests/test_rss_scraper.py`).

---

## No CI/CD

There is no CI pipeline. Scripts are run manually on a Windows machine or scheduled via Windows Task Scheduler / Power Automate.

---

## Dependencies (Inferred — no requirements.txt)

```
requests
beautifulsoup4
lxml
selenium
webdriver-manager
pandas
numpy
scikit-learn
openpyxl
feedparser
pdfplumber  # optional
```

Install with:
```bash
pip install requests beautifulsoup4 lxml selenium webdriver-manager pandas numpy scikit-learn openpyxl feedparser pdfplumber
```

Chrome or Chromium must be installed; `webdriver-manager` handles the matching ChromeDriver automatically.

---

## Things to Watch Out For

1. **Windows-only paths** — All `DATA_DIR` / `OUTPUT_DIR` / `SCRIPT_DIR` values are Windows absolute paths. Update them before running on any other OS.
2. **No config file** — All thresholds, paths, and keyword lists are hardcoded in the scripts. Extract to a config file if the tool is to be shared across environments.
3. **`Scraper` binary** — The file named `Scraper` (no extension) is a binary duplicate of `scraper.py`. Do not edit it; it is not executed by any script.
4. **Excel table management** — Always use `rewrite_excel_table()` / `save_weekly_file()` rather than writing Excel directly; they handle table deletion, recreation, and style re-application needed to avoid corruption.
5. **Similarity threshold** — Changing `SIMILARITY_THRESHOLD` in `scraper.py` affects deduplication aggressiveness. Lower values deduplicate more aggressively; higher values allow more near-duplicates through.
6. **RSS feed list** — `rss_feeds.txt` (in `SCRIPT_DIR`) must exist and contain one RSS URL per line. Missing or empty file causes `rss_scraper.py` to exit early.
