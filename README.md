# MGA Deal Scout 🏦

> Autonomous weekly deal-sourcing pipeline for early-stage Indian startups — built for **MGA Ventures**, a Mumbai-based family office.

[![Live Demo](https://img.shields.io/badge/Streamlit-Live%20Demo-FF4B4B?logo=streamlit&logoColor=white)](https://deal-screener-uykhmzbcoce9iwmbhscovd.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

---

## What It Does

Every week, with **zero human input**, this tool:

1. **Discovers** a promising early-stage Indian startup using targeted Google searches
2. **Researches** it across multiple public sources (news, LinkedIn, company site)
3. **Synthesises** findings using Groq Llama 3.3 into structured JSON
4. **Builds** a 9-slide PowerPoint deal memo — ready to present
5. **Emails** the memo as an attachment to any recipient

All findings are grounded in cited, publicly verifiable sources. The AI is instructed to write *"Not publicly disclosed."* for any fact it cannot find — no numbers or names are ever invented.

---

## Live Dashboard

**[https://deal-screener-uykhmzbcoce9iwmbhscovd.streamlit.app/](https://deal-screener-uykhmzbcoce9iwmbhscovd.streamlit.app/)**

Select a sector, hit **Run Deal Scout**, and watch the pipeline execute live. Download the deck or have it emailed directly.

![Dashboard Preview](https://img.shields.io/badge/UI-Streamlit%20Dashboard-FF4B4B?logo=streamlit)

---

## The 9-Slide Deal Memo

| Slide | Content |
|-------|---------|
| 1 | **Title** — Company name, sector, date, MGA branding |
| 2 | **Company Overview** — Stage, HQ, business model, latest news ticker |
| 3 | **What It Does** — Problem, solution, and market context |
| 4 | **Products & Services** — Product list with descriptions |
| 5 | **Management Team** — Founders, roles, backgrounds |
| 6 | **Market Opportunity** — TAM / SAM / SOM with sources |
| 7 | **Funding & Traction** — Funding table, timeline, traction metrics |
| 8 | **Thesis Fit Scorecard** — Scored against MGA's investment criteria |
| 9 | **Sources & Attribution** — All clickable numbered references |

---

## Pipeline Architecture

```
Serper.dev (search)
    ↓  2 discovery queries → score & pick best startup
    ↓  4 research queries  → fetch page content
Groq Llama 3.3 (synthesis)
    ↓  1 structured JSON call
python-pptx (deck builder)
    ↓  9-slide .pptx with charts, tables, hyperlinks
Gmail SMTP (delivery)
    ↓  .pptx emailed as attachment
```

Each run uses ~6 Serper queries and 1 Groq call — well within both free tiers.

---

## Investment Thesis (Scoring Criteria)

| Criterion | Target |
|-----------|--------|
| Stage | Seed → Series B |
| Geography | India HQ |
| Sectors | Consumer · Fintech · SaaS |
| Ticket size | Up to $2M |

The Thesis Fit slide scores each company across these dimensions and flags mismatches.

---

## Tech Stack

| Layer | Tool | Why |
|-------|------|-----|
| Search | [Serper.dev](https://serper.dev) | 2,500 free queries/month |
| AI | [Groq](https://console.groq.com) Llama 3.3-70B | 14,400 free calls/day |
| Deck | [python-pptx](https://python-pptx.readthedocs.io) | No PowerPoint license needed |
| Email | Gmail SMTP + App Password | No OAuth, no Cloud Console |
| UI | [Streamlit](https://streamlit.io) | Zero-config web deployment |
| Charts | Matplotlib | Funding timeline + traction bar |

All APIs used are **free tier** — no credit card required.

---

## Setup

### 1. Clone & install

```bash
git clone https://github.com/harjeet2004/deal-screener.git
cd deal-screener/3-deal-screener
pip install -r requirements.txt
```

### 2. Configure API keys

```bash
cp .env.example .env
# Edit .env and fill in your keys
```

| Variable | Where to get it | Required? |
|----------|----------------|-----------|
| `SERPER_API_KEY` | [serper.dev](https://serper.dev) → free account | ✅ |
| `GROQ_API_KEY` | [console.groq.com](https://console.groq.com) → free account | ✅ |
| `RECIPIENT_EMAIL` | Any email address | Optional |
| `GMAIL_USER` | Your Gmail address (for sending) | Optional |
| `GMAIL_APP_PASSWORD` | [Google App Password](https://myaccount.google.com/apppasswords) | Optional |

> Email is completely optional — the deck is always available as a download.

### 3. Run locally

```bash
# Streamlit dashboard
streamlit run deal_scout_app.py

# Headless CLI (useful for scheduling)
python deal_scout.py --sector Consumer
```

### 4. Deploy to Streamlit Cloud

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io) → New app → point to `3-deal-screener/deal_scout_app.py`
3. Add secrets (Settings → Secrets) — copy from `.streamlit/secrets.toml.example`

---

## Gmail Email Setup (Optional)

To enable email delivery, you need a **Gmail App Password** — not your regular password:

1. Enable **2-Step Verification** on your Google account
2. Go to [myaccount.google.com](https://myaccount.google.com) → Security → **App passwords**
3. Create one → copy the 16-character code
4. Add to `.env` or Streamlit Cloud secrets:

```toml
GMAIL_USER         = "your.account@gmail.com"
GMAIL_APP_PASSWORD = "abcd efgh ijkl mnop"
RECIPIENT_EMAIL    = "anyone@example.com"
```

---

## Scheduling (Automated Weekly Runs)

**Windows Task Scheduler:**
```
Action: python C:\path\to\3-deal-screener\deal_scout.py
Trigger: Every Monday at 08:00
```

**Mac / Linux cron:**
```cron
0 8 * * 1 cd /path/to/3-deal-screener && python deal_scout.py
```

The sector rotates automatically each week: Consumer → Fintech → SaaS → Consumer …

---

## Project Structure

```
deal-screener/
└── 3-deal-screener/
    ├── deal_scout_app.py        # Streamlit dashboard
    ├── deal_scout.py            # Core pipeline (search → AI → deck → email)
    ├── requirements.txt
    ├── .env.example             # Environment variable template
    └── .streamlit/
        └── secrets.toml.example # Streamlit Cloud secrets template
```

---

## Guardrails & Quality

- **Source grounding** — every fact attributed to a URL; Groq is instructed never to invent data
- **Scope filtering** — companies outside India, at Series C+, or in wrong sectors are automatically rejected
- **Retry logic** — up to 6 search attempts per run if the first candidate doesn't pass filters
- **Graceful degradation** — email failure never blocks the deck download

---

Built for [MGA Ventures](https://mgaventures.in) · Mumbai
