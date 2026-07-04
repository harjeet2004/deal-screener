# MGA Prep — Finance Analyst Tools

A suite of three standalone finance tools, each mapping to a service offered
by MGA Ventures, a Mumbai-based family office investing across unlisted
equity (seed to pre-IPO), debt, listed equity, and real estate. Each tool
lives in its own subfolder, runs independently, and produces a polished,
presentable output file.

## Repository structure
```
mga-prep/
├── CLAUDE.md              # This file
├── README.md              # Short human-facing summary
├── 1-portfolio-mis/       # Tool 1: portfolio monitoring dashboard
│   ├── sample_data/       # Input company financial files
│   ├── mis_report.py      # Main script
│   └── MIS_Dashboard.xlsx # Generated output
├── 2-valuation/           # Tool 2: comparable-company valuation
│   ├── comps.csv          # Comparable companies + target
│   ├── valuation.py       # Main script
│   └── Valuation_Model.xlsx
└── 3-deal-screener/       # Tool 3: pitch-deck screening memo
    ├── deck.pdf           # Input pitch deck
    ├── screener.py        # Text-extraction helper
    └── screening_memo.md  # Generated output
```

## Tools

### 1. Portfolio MIS (`1-portfolio-mis/`)
**Maps to:** MGA's MIS Reporting service.
**Purpose:** Consolidate quarterly financials from portfolio companies in
inconsistent formats into one standardised health dashboard and interactive
Streamlit app. Designed for a long-term family office that thinks in years,
not months — all metrics and thresholds are in quarters.

- **Input:** Per-company Excel files in `sample_data/`, each with 12 quarters
  of P&L (FY23 Q1 – FY25 Q4): revenue, COGS, operating expenses, cash balance.
  Column names and layouts vary deliberately; the parser handles them robustly.
- **Processing:** Normalise to a common schema; compute metrics below per company.
- **Output (Excel):** `MIS_Dashboard.xlsx` — one summary sheet, one company per
  row, metrics in columns, descriptive RAG status cell. Header bold/navy, frozen.
- **Output (Streamlit):** `streamlit_app.py` — portfolio overview (status cards,
  RAG table, revenue + runway charts) plus a per-company detail view (metric
  cards, auto-generated analyst note, 4 trend charts).
- **Run Excel report:** `python3 mis_report.py`
- **Run dashboard:** `streamlit run streamlit_app.py`

**Metric definitions**
- *QoQ revenue growth* = (this quarter − previous quarter) / previous quarter
- *YoY revenue growth* = (this quarter − same quarter last year) / same quarter
  last year  [Q_n vs Q_{n-4}; available from FY24 Q1 onward]
- *Gross margin %* = (revenue − COGS) / revenue, latest quarter
- *Revenue CAGR (3Y)* = (Q12 revenue / Q1 revenue)^(1/3) − 1
  CAGR is the right long-horizon metric for a family office: it smooths
  quarter-to-quarter noise and expresses growth as a single annualised rate
  directly comparable across companies and asset classes.
- *Quarterly burn* = average quarterly net cash outflow over all available quarters
- *Runway (quarters)* = current cash / quarterly burn
- *Status:* Critical if runway < 2 quarters, Watch if 2–4, Healthy if > 4

**Money formatting**
All monetary figures display in Indian Rupees using the Indian numbering system:
groups of 3 for the last three digits, then groups of 2 from the right, with a
₹ symbol. Example: 2345600 → ₹23,45,600; 12500000 → ₹1,25,00,000.
Raw numbers are kept as plain floats internally; `fmt_inr()` in `mis_report.py`
is the single formatting function used everywhere money is displayed.

### 2. Valuation (`2-valuation/`)
**Maps to:** MGA's Valuation Services.
**Purpose:** Produce a defensible implied valuation range for a target
company using comparable-company analysis.

- **Input:** `comps.csv` — a set of comparable companies (name, market cap,
  total debt, cash, revenue, EBITDA) plus one target company with revenue
  and EBITDA but no market cap.
- **Processing:** For each comparable, compute Enterprise Value
  (EV = market cap + total debt − cash) and the multiples EV/Revenue and
  EV/EBITDA. Take the median and mean of each multiple across the set, then
  apply them to the target's revenue and EBITDA.
- **Output:** `Valuation_Model.xlsx` — a comps table with per-company
  multiples, a summary box of median/mean multiples, and an implied
  valuation section showing a low/median/high range.
- **Run:** `python3 valuation.py`

**Notes**
- EV/Revenue is the more meaningful multiple for early-stage or unprofitable
  companies; EV/EBITDA applies once the business is profitable.
- Present results as a range, not a single point estimate.

### 3. Deal Screener / Autonomous Scout (`3-deal-screener/`)
**Maps to:** MGA's Due Diligence service (top-of-funnel screening).
**Purpose:** Fully autonomous weekly deal-sourcing pipeline. Searches for one
promising early-stage Indian startup each week, researches it with public web
sources, synthesises findings with Gemini, produces a 9-slide PowerPoint deal
memo, and emails it — zero human input required.

- **Input:** None (autonomous). Sector rotates Consumer → Fintech → SaaS by ISO week.
- **Processing pipeline:**
  1. Google Custom Search (2 queries) → score & pick best startup candidate
  2. Google Custom Search (4 follow-up queries) + page fetching → raw research corpus
  3. Google Gemini 2.0 Flash (1 synthesis call) → structured JSON per section
  4. python-pptx → 9-slide .pptx deal memo saved to this folder
  5. Gmail API (OAuth2) → .pptx emailed as attachment
- **Output:** `<CompanyName>_<YYYY-MM-DD>.pptx` — 9 slides:
  Title · Company Overview · What It Does · Products · Management Team ·
  Market Opportunity · Funding & Traction (with matplotlib chart) ·
  Thesis Fit scorecard · Sources
- **Run:** `python deal_scout.py`
- **Schedule (Windows):** Task Scheduler → run every Monday at 08:00
- **Schedule (Mac/Linux):** `0 8 * * 1 cd /path/to/3-deal-screener && python deal_scout.py`

**APIs used (all free-tier):**
- Google Custom Search JSON API — 100 queries/day free; ~6 used per run
- Google Gemini 2.0 Flash — 1,500 requests/day free; ~1 used per run
- Gmail API — free with OAuth2

**Credential files required (in `3-deal-screener/`, never committed):**
- `.env` — copied from `.env.example`; holds GOOGLE_CSE_API_KEY, GOOGLE_CSE_CX,
  GEMINI_API_KEY, RECIPIENT_EMAIL
- `credentials.json` — Gmail OAuth2 client secrets from Google Cloud Console
- `token.json` — auto-generated on first Gmail auth; refreshes automatically

**Fact-sourcing policy:** The synthesis prompt instructs Gemini to write
"Not publicly disclosed." for any fact not found in the fetched sources.
No numbers or names are ever invented. Every claim is attributed to a URL.

## Investment thesis (used by the Deal Screener)
- **Stage:** Seed to Series B
- **Sectors:** Agnostic, with focus on Consumer, Fintech, SaaS
- **Ticket size:** Up to $2M

## Dependencies

**Tools 1 & 2** (Portfolio MIS + Valuation):
```
pip install pandas openpyxl pdfplumber matplotlib streamlit plotly
```

**Tool 3** (Deal Scout) — see `3-deal-screener/requirements.txt`:
```
pip install requests beautifulsoup4 lxml python-dotenv google-genai \
            google-api-python-client google-auth-httplib2 google-auth-oauthlib \
            python-pptx matplotlib
```

## Conventions
- Use only free, open-source Python libraries. Do not call any paid API or
  wire in a billed Anthropic API key.
- Each tool is self-contained in its subfolder; output files stay alongside
  their script.
- Keep scripts simple, readable, and well-commented — the logic must be
  explainable out loud.
- Prefer clear variable names and small functions over cleverness.

## Working notes
- Before running a script, state briefly what it will do.
- After generating any Excel or document output, open and sanity-check it
  rather than assuming it's correct.
- For any web-based dashboard or visual UI, take a screenshot of the rendered
  output, review it, and iterate on layout, spacing, and styling until it
  looks clean and professional.
- Generated output files can be regenerated at any time by re-running the
  relevant script.
