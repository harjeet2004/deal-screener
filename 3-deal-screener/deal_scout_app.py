#!/usr/bin/env python3
"""
deal_scout_app.py — MGA Ventures Deal Scout Dashboard
Streamlit UI for the autonomous deal-sourcing pipeline.
Select a sector, watch live progress, download the PPTX, get it emailed.
"""

import datetime
import io
import os
import queue
import re
import sys
import threading
from pathlib import Path

import streamlit as st

# ── Bootstrap ─────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

st.set_page_config(
    page_title="MGA Deal Scout",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Import pipeline first, then inject secrets directly into the module ────────
# deal_scout.py reads SERPER_API_KEY / GROQ_API_KEY as module-level globals at
# import time.  On Streamlit Cloud there is no .env, so we must patch those
# globals after import.  We also set os.environ as a belt-and-suspenders backup.
try:
    import deal_scout as _ds
    from deal_scout import (
        _fetch_logo,
        build_deck,
        deep_research,
        discover_startup,
        email_deck,
        synthesise,
    )
    IMPORT_OK = True
except ImportError as _e:
    IMPORT_OK = False
    IMPORT_ERR = str(_e)
    _ds = None

# Patch module globals with Streamlit Cloud secrets (runs every script execution,
# which is fine — Streamlit caches the module object between reruns).
try:
    if _ds is not None:
        if "SERPER_API_KEY" in st.secrets:
            _ds.SERPER_API_KEY = st.secrets["SERPER_API_KEY"]
            os.environ["SERPER_API_KEY"] = st.secrets["SERPER_API_KEY"]
        if "GROQ_API_KEY" in st.secrets:
            _ds.GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
            os.environ["GROQ_API_KEY"] = st.secrets["GROQ_API_KEY"]
        if "RECIPIENT_EMAIL" in st.secrets:
            _ds.RECIPIENT_EMAIL = st.secrets["RECIPIENT_EMAIL"]
            os.environ["RECIPIENT_EMAIL"] = st.secrets["RECIPIENT_EMAIL"]
        if "GMAIL_TOKEN_JSON" in st.secrets:
            (BASE_DIR / "token.json").write_text(st.secrets["GMAIL_TOKEN_JSON"])
except Exception:
    pass  # Running locally — .env + load_dotenv() inside deal_scout handles it

# ── Palette ────────────────────────────────────────────────────────────────────
NAVY  = "#1A2B4A"
GOLD  = "#C9A84C"
GREEN = "#277A27"
RED   = "#BB2222"
AMBER = "#CC7A00"
LTBG  = "#F4F6FA"

ALL_SECTORS = ["Consumer", "Fintech", "SaaS"]

STEP_LABELS = [
    "Discover",
    "Research",
    "AI Synthesis",
    "Build Deck",
    "Send Email",
]

STEP_ICONS = ["🔍", "📚", "🤖", "📊", "📧"]

# ── Global CSS ─────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
/* ── Base ─────────────────────────────────────────────────── */
[data-testid="stAppViewContainer"] {{ background: #F0F2F8; }}
[data-testid="block-container"] {{ padding-top: 1rem; }}

/* ── Sidebar ──────────────────────────────────────────────── */
[data-testid="stSidebar"] > div:first-child {{
  background: linear-gradient(175deg, {NAVY} 0%, #0d1b33 100%);
  border-right: 2px solid {GOLD}44;
}}
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] label {{ color: #b0bfd4 !important; }}
[data-testid="stSidebar"] hr {{ border-color: {GOLD}33 !important; }}

/* Sidebar buttons */
[data-testid="stSidebar"] .stButton > button {{
  background: linear-gradient(135deg, {GOLD}, #b8912e) !important;
  color: {NAVY} !important;
  font-weight: 800 !important;
  border: none !important;
  border-radius: 8px !important;
  width: 100% !important;
  padding: 0.7rem !important;
  font-size: 0.95rem !important;
  letter-spacing: 0.03em !important;
  box-shadow: 0 3px 12px {GOLD}44 !important;
  transition: all 0.2s ease !important;
}}
[data-testid="stSidebar"] .stButton > button:hover:not(:disabled) {{
  transform: translateY(-2px) !important;
  box-shadow: 0 6px 20px {GOLD}66 !important;
}}
[data-testid="stSidebar"] .stButton > button:disabled {{
  background: #3a4a5a !important;
  color: #7890aa !important;
  box-shadow: none !important;
}}

/* ── Download button ────────────────────────────────────────── */
.stDownloadButton > button {{
  background: linear-gradient(135deg, {GOLD}, #b8912e) !important;
  color: {NAVY} !important;
  font-weight: 700 !important;
  border: none !important;
  border-radius: 8px !important;
  padding: 0.65rem 1.4rem !important;
  font-size: 0.95rem !important;
  box-shadow: 0 3px 10px {GOLD}44 !important;
  width: 100% !important;
}}

/* ── Step tracker ────────────────────────────────────────────── */
.step-row {{
  display: flex;
  gap: 0.5rem;
  margin-bottom: 1.5rem;
}}
.step-item {{
  flex: 1;
  background: white;
  border-radius: 10px;
  padding: 0.8rem 0.5rem 0.7rem;
  text-align: center;
  border: 2px solid #dde4f0;
  position: relative;
  transition: all 0.3s ease;
}}
.step-item.done {{
  border-color: {GREEN};
  background: #f0fff5;
  box-shadow: 0 2px 10px {GREEN}22;
}}
.step-item.active {{
  border-color: {GOLD};
  background: linear-gradient(160deg, #fffdf5 0%, #fff8e8 100%);
  box-shadow: 0 0 0 2px {GOLD}55, 0 4px 24px {GOLD}44;
  animation: glow-sweep 2s ease-in-out infinite;
}}
.step-item.pending {{
  opacity: 0.38;
}}
@keyframes glow-sweep {{
  0%,100% {{ box-shadow: 0 0 0 2px {GOLD}44, 0 4px 20px {GOLD}33; }}
  50%      {{ box-shadow: 0 0 0 2px {GOLD}99, 0 6px 32px {GOLD}66; }}
}}

/* ── Finance bar-chart loader ─────────────────────────────────── */
@keyframes fin-bar {{
  0%,100% {{ height: 5px;  opacity: 0.45; }}
  50%      {{ height: 22px; opacity: 1.0;  }}
}}
.fin-loader {{
  display: flex;
  align-items: flex-end;
  gap: 3px;
  height: 26px;
  justify-content: center;
  margin: 2px auto 2px;
}}
.fin-loader b {{
  display: inline-block;
  width: 5px;
  background: linear-gradient(180deg, {GOLD} 0%, #a87820 100%);
  border-radius: 2px 2px 1px 1px;
  animation: fin-bar 1.0s ease-in-out infinite;
}}
.fin-loader b:nth-child(1) {{ animation-delay: 0s;    }}
.fin-loader b:nth-child(2) {{ animation-delay: 0.18s; }}
.fin-loader b:nth-child(3) {{ animation-delay: 0.36s; }}
.fin-loader b:nth-child(4) {{ animation-delay: 0.54s; }}
.fin-loader b:nth-child(5) {{ animation-delay: 0.72s; }}

/* ── Pulse dot for status badge ───────────────────────────────── */
@keyframes pulse-dot {{
  0%,100% {{ transform: scale(1);   opacity: 1;   }}
  50%      {{ transform: scale(0.6); opacity: 0.4; }}
}}
.pulse-dot {{
  width: 8px; height: 8px;
  background: {GOLD};
  border-radius: 50%;
  display: inline-block;
  flex-shrink: 0;
  animation: pulse-dot 1.2s ease-in-out infinite;
}}

/* ── Scan dots ("Analysing...") ───────────────────────────────── */
@keyframes scan-flicker {{
  0%,20%   {{ opacity: 0; }}
  50%       {{ opacity: 1; }}
  80%,100% {{ opacity: 0; }}
}}
.sdot {{ display: inline-block; animation: scan-flicker 1.5s infinite; }}
.sdot:nth-child(2) {{ animation-delay: 0.3s; }}
.sdot:nth-child(3) {{ animation-delay: 0.6s; }}

.step-num  {{ font-size: 0.6rem; font-weight: 700; letter-spacing: 0.1em; color: #99aabb; text-transform: uppercase; margin-bottom: 0.25rem; }}
.step-icon {{ font-size: 1.5rem; line-height: 1; margin-bottom: 0.25rem; }}
.step-lbl  {{ font-size: 0.72rem; font-weight: 700; color: #334; }}

/* ── Company hero card ───────────────────────────────────────── */
.hero {{
  background: linear-gradient(135deg, {NAVY} 0%, #223966 100%);
  border-radius: 14px;
  padding: 2rem 2.5rem 1.8rem;
  margin-bottom: 1.5rem;
  border: 1px solid {GOLD}44;
  position: relative;
  overflow: hidden;
}}
.hero::after {{
  content: '';
  position: absolute;
  top: -60px; right: -60px;
  width: 220px; height: 220px;
  background: radial-gradient({GOLD}22, transparent 65%);
  border-radius: 50%;
}}
.hero-eyebrow  {{ font-size: 0.65rem; font-weight: 700; color: {GOLD}bb; letter-spacing: 0.18em; text-transform: uppercase; margin-bottom: 0.4rem; }}
.hero-title    {{ font-size: 2.3rem; font-weight: 900; color: white; margin: 0 0 0.35rem; line-height: 1.1; }}
.hero-tagline  {{ font-size: 1rem; color: #9aafc8; font-style: italic; margin-bottom: 1.3rem; }}
.chip-row      {{ display: flex; flex-wrap: wrap; gap: 0.5rem; }}
.chip {{
  display: inline-flex;
  flex-direction: column;
  background: #1f3558;
  border: 1px solid {GOLD}55;
  border-radius: 7px;
  padding: 0.4rem 0.9rem;
  min-width: 88px;
}}
.chip-lbl  {{ font-size: 0.58rem; font-weight: 800; color: {GOLD}cc; letter-spacing: 0.12em; text-transform: uppercase; }}
.chip-val  {{ font-size: 0.88rem; font-weight: 700; color: white; margin-top: 0.1rem; }}

/* ── Info card ───────────────────────────────────────────────── */
.icard {{
  background: white;
  border-radius: 10px;
  padding: 1.1rem 1.4rem;
  margin-bottom: 0.85rem;
  border: 1px solid #e0e8f4;
  border-left: 5px solid {NAVY};
}}
.icard-title {{
  font-size: 0.65rem; font-weight: 800; color: #8899bb;
  letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 0.45rem;
}}
.icard-body {{ font-size: 0.88rem; color: #334; line-height: 1.6; margin: 0; }}

/* ── Thesis cards ────────────────────────────────────────────── */
.thesis-grid {{
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 0.6rem;
  margin-bottom: 1rem;
}}
.tcard {{
  border-radius: 10px;
  padding: 0.9rem 1rem;
  color: white;
  position: relative;
  overflow: hidden;
}}
.tcard::before {{
  content: '';
  position: absolute;
  bottom: -20px; right: -20px;
  width: 70px; height: 70px;
  background: rgba(255,255,255,0.06);
  border-radius: 50%;
}}
.tcard-label   {{ font-size: 0.62rem; font-weight: 800; opacity: 0.8; letter-spacing: 0.1em; text-transform: uppercase; }}
.tcard-verdict {{ font-size: 1.1rem; font-weight: 900; margin: 0.3rem 0 0.2rem; }}
.tcard-detail  {{ font-size: 0.72rem; opacity: 0.88; line-height: 1.45; }}

/* ── Metric grid ─────────────────────────────────────────────── */
.metric-grid {{
  display: flex;
  flex-wrap: wrap;
  gap: 0.6rem;
  margin-bottom: 1rem;
}}
.mcard {{
  flex: 1 1 130px;
  background: white;
  border: 1px solid #e0e8f4;
  border-radius: 10px;
  padding: 0.85rem 1rem;
}}
.mcard-label  {{ font-size: 0.65rem; color: #8899aa; font-weight: 700; text-transform: uppercase; letter-spacing: 0.07em; }}
.mcard-value  {{ font-size: 1.25rem; font-weight: 800; color: {GREEN}; margin-top: 0.15rem; }}
.mcard-source {{ font-size: 0.62rem; color: #aaa; margin-top: 0.1rem; }}

/* ── Bullet highlights ───────────────────────────────────────── */
.hl-item {{
  padding: 0.4rem 0 0.4rem 1rem;
  border-left: 3px solid {GREEN};
  margin-bottom: 0.4rem;
  font-size: 0.87rem;
  color: #334;
  line-height: 1.5;
}}
.rf-item {{
  padding: 0.4rem 0 0.4rem 1rem;
  border-left: 3px solid {RED};
  margin-bottom: 0.4rem;
  font-size: 0.87rem;
  color: #334;
  line-height: 1.5;
}}

/* ── Founder card ────────────────────────────────────────────── */
.founder-card {{
  background: white;
  border: 1px solid #e0e8f4;
  border-radius: 10px;
  padding: 0.85rem 1rem;
  margin-bottom: 0.5rem;
  border-top: 3px solid {GOLD};
}}
.founder-name  {{ font-weight: 800; color: {NAVY}; font-size: 0.95rem; }}
.founder-title {{ color: #778; font-size: 0.78rem; margin: 0.1rem 0; }}
.founder-bio   {{ color: #445; font-size: 0.8rem; margin-top: 0.3rem; line-height: 1.45; }}

/* ── Funding row ─────────────────────────────────────────────── */
.fr-card {{
  background: white;
  border: 1px solid #e0e8f4;
  border-radius: 10px;
  padding: 0.75rem 1rem;
  margin-bottom: 0.45rem;
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 0.5rem;
}}
.fr-round  {{ font-weight: 800; color: {NAVY}; font-size: 0.9rem; }}
.fr-amount {{ color: {GOLD}; font-weight: 800; margin-left: 0.5rem; }}
.fr-meta   {{ color: #778; font-size: 0.75rem; margin-top: 0.1rem; }}

/* ── Log ─────────────────────────────────────────────────────── */
.log-box {{
  background: #0d1117;
  border: 1px solid #30363d;
  border-radius: 10px;
  padding: 1rem 1.2rem;
  font-family: 'JetBrains Mono', 'Fira Code', 'Courier New', monospace;
  font-size: 0.77rem;
  color: #8b949e;
  max-height: 260px;
  overflow-y: auto;
  line-height: 1.65;
}}
.log-ok   {{ color: #3fb950; }}
.log-info {{ color: #79c0ff; }}
.log-ai   {{ color: #d2a8ff; }}
.log-warn {{ color: #d29922; }}
.log-err  {{ color: #f85149; }}

/* ── Section label ───────────────────────────────────────────── */
.sec-label {{
  font-size: 0.68rem;
  font-weight: 800;
  letter-spacing: 0.1em;
  text-transform: uppercase;
  margin-bottom: 0.55rem;
  padding-bottom: 0.3rem;
  border-bottom: 2px solid #e0e8f4;
}}

/* ── Status badge ────────────────────────────────────────────── */
.status-running {{
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  background: {GOLD}22;
  border: 1px solid {GOLD};
  border-radius: 20px;
  padding: 0.3rem 0.9rem;
  font-size: 0.8rem;
  font-weight: 700;
  color: {GOLD};
  margin-bottom: 1rem;
}}
.status-done {{
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  background: {GREEN}22;
  border: 1px solid {GREEN};
  border-radius: 20px;
  padding: 0.3rem 0.9rem;
  font-size: 0.8rem;
  font-weight: 700;
  color: {GREEN};
  margin-bottom: 1rem;
}}
</style>
""", unsafe_allow_html=True)


MAX_PIPELINE_ATTEMPTS = 6

# ── Session state ──────────────────────────────────────────────────────────────
def _init(k, v):
    if k not in st.session_state:
        st.session_state[k] = v

_init("running",        False)
_init("result",         None)
_init("pptx_path",      None)
_init("log_lines",      [])
_init("email_sent",     False)
_init("email_error",    None)
_init("error",          None)
_init("step",           0)
_init("sector_used",    "")
_init("candidate",      "")
_init("attempt_cur",    1)
_init("attempt_max",    MAX_PIPELINE_ATTEMPTS)


# ── Stdout tee for live log capture ───────────────────────────────────────────
class _TeeIO:
    def __init__(self, original, q: queue.Queue):
        self._orig = original
        self._q    = q

    def write(self, s: str):
        self._orig.write(s)
        line = s.rstrip()
        if line:
            self._q.put(("log", line))
        return len(s)

    def flush(self):
        self._orig.flush()

    def __getattr__(self, name):
        return getattr(self._orig, name)


# ── Pipeline thread ────────────────────────────────────────────────────────────
_INDIA_CITIES = {
    "india", "bengaluru", "bangalore", "mumbai", "delhi", "hyderabad",
    "pune", "chennai", "noida", "gurugram", "gurgaon", "kolkata",
    "ahmedabad", "jaipur", "indore", "surat", "kochi", "coimbatore",
}

def _run_pipeline(sector: str, send_email: bool, log_q: queue.Queue):
    orig = sys.stdout
    sys.stdout = _TeeIO(orig, log_q)
    try:
        today = datetime.date.today().strftime("%Y-%m-%d")
        tried: set[str] = set()   # company names already attempted this run
        structured = None
        pptx_path  = None

        for attempt in range(1, MAX_PIPELINE_ATTEMPTS + 1):
            log_q.put(("attempt", (attempt, MAX_PIPELINE_ATTEMPTS)))
            if attempt > 1:
                log_q.put(("log",
                    f"[Attempt {attempt}/{MAX_PIPELINE_ATTEMPTS}] "
                    f"Searching for a different {sector} candidate…"))

            # ── Step 1: Discover ──────────────────────────────────────────
            log_q.put(("step", 1))
            try:
                candidate = discover_startup(sector, exclude=tried)
            except SystemExit:
                if attempt < MAX_PIPELINE_ATTEMPTS:
                    log_q.put(("log",
                        f"[Attempt {attempt}] No new candidate found in results. Retrying…"))
                    continue
                log_q.put(("error",
                    f"Could not find a suitable {sector} startup after "
                    f"{MAX_PIPELINE_ATTEMPTS} attempts. Please try again later."))
                return

            company = candidate["name"]
            tried.add(company)
            log_q.put(("candidate", company))

            # ── Step 2: Research ─────────────────────────────────────────
            log_q.put(("step", 2))
            research = deep_research(company, sector, candidate["url"])

            # ── Step 3: AI Synthesis ─────────────────────────────────────
            log_q.put(("step", 3))
            structured = synthesise(company, sector, research)

            # Guard 1: India HQ check
            hq = structured.get("hq", "")
            if hq and not any(c in hq.lower() for c in _INDIA_CITIES) \
                    and "not publicly" not in hq.lower():
                log_q.put(("log",
                    f"[Attempt {attempt}] '{company}' HQ is '{hq}' — not India. "
                    "Skipping and searching again…"))
                continue

            # Guard 2: Thesis-fit check — skip if 2+ dimensions are out of scope
            fit_stage  = structured.get("thesis_fit_stage",  "").lower()
            fit_sector = structured.get("thesis_fit_sector", "").lower()
            fit_ticket = structured.get("thesis_fit_ticket", "").lower()
            out_of_scope = [f for f in [fit_stage, fit_sector, fit_ticket]
                            if "out of scope" in f]
            if len(out_of_scope) >= 2:
                log_q.put(("log",
                    f"[Attempt {attempt}] '{company}' is outside MGA's scope "
                    "(stage/sector/ticket). Searching for a better match…"))
                continue

            # Good candidate — exit the retry loop
            break

        else:
            # Loop finished without a break → all attempts exhausted
            log_q.put(("error",
                f"Could not find a thesis-aligned {sector} startup after "
                f"{MAX_PIPELINE_ATTEMPTS} attempts. Please try again later."))
            return

        # ── Step 4: Build deck ────────────────────────────────────────────
        logo_bytes = _fetch_logo(structured.get("website", ""))
        log_q.put(("step", 4))
        pptx_path = build_deck(structured, today, logo_bytes=logo_bytes)

        # Mark done right after deck is built — download is available from here on
        # even if the email step below fails or is skipped.
        log_q.put(("done", (structured, str(pptx_path))))

        # ── Step 5: Email (non-fatal) ─────────────────────────────────────
        log_q.put(("step", 5))
        if send_email:
            summary = structured.get(
                "email_summary", f"{company} — see attached deal memo.")
            try:
                email_deck(pptx_path,
                           structured.get("company_name", company),
                           summary, today)
                log_q.put(("email_ok", None))
            except BaseException as em:
                # Catch SystemExit too (e.g. missing credentials.json)
                log_q.put(("email_err", str(em)))
        else:
            log_q.put(("log", "[Email] Skipped (checkbox unchecked)."))

        log_q.put(("finish", None))

    except SystemExit:
        log_q.put(("error",
            "No candidate found after retries. Try a different sector or re-run."))
    except Exception as exc:
        import traceback
        log_q.put(("error", f"{exc}\n{traceback.format_exc()}"))
    finally:
        sys.stdout = orig


# ── Queue drain ────────────────────────────────────────────────────────────────
def _drain():
    if not (st.session_state.running and "_log_q" in st.session_state):
        return
    q = st.session_state["_log_q"]
    while True:
        try:
            kind, payload = q.get_nowait()
            if kind == "log":
                st.session_state.log_lines.append(payload)
            elif kind == "step":
                st.session_state.step = payload
            elif kind == "attempt":
                st.session_state.attempt_cur, st.session_state.attempt_max = payload
            elif kind == "candidate":
                st.session_state.candidate = payload
            elif kind == "done":
                # Deck is ready — store result but keep running=True so email
                # messages (email_ok / email_err) can still be drained.
                structured, pptx = payload
                st.session_state.result    = structured
                st.session_state.pptx_path = pptx
            elif kind == "finish":
                # Pipeline fully complete (after email step)
                st.session_state.running = False
                st.session_state.step    = 5
            elif kind == "email_ok":
                st.session_state.email_sent = True
            elif kind == "email_err":
                st.session_state.email_error = payload
            elif kind == "error":
                st.session_state.error   = payload
                st.session_state.running = False
        except queue.Empty:
            break


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center;padding:1.8rem 0.5rem 1.4rem;">
      <div style="font-size:3rem;margin-bottom:0.5rem;filter:drop-shadow(0 0 10px {GOLD}66);">🏦</div>
      <div style="font-size:1.05rem;font-weight:900;color:{GOLD};letter-spacing:0.08em;">MGA VENTURES</div>
      <div style="font-size:0.72rem;color:#607080;margin-top:0.2rem;letter-spacing:0.05em;">DEAL SCOUT DASHBOARD</div>
    </div>
    <hr>
    """, unsafe_allow_html=True)

    st.markdown(f"<div style='font-size:0.68rem;font-weight:800;color:{GOLD}cc;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:0.4rem;'>Sector / Domain</div>", unsafe_allow_html=True)

    sector_final = st.selectbox(
        "Sector",
        ALL_SECTORS,
        label_visibility="collapsed",
        disabled=st.session_state.running,
    )

    st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)

    send_email_opt = st.checkbox(
        "Send email when done",
        value=True,
        disabled=st.session_state.running,
    )

    st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)

    btn_disabled = (
        st.session_state.running
        or not IMPORT_OK
    )
    btn_label = "⏳  Running pipeline…" if st.session_state.running else "▶  Run Deal Scout"
    run_btn = st.button(btn_label, disabled=btn_disabled, use_container_width=True)

    # Download once ready
    if st.session_state.pptx_path and Path(st.session_state.pptx_path).exists():
        st.markdown(f"<hr>", unsafe_allow_html=True)
        pptx_bytes = Path(st.session_state.pptx_path).read_bytes()
        st.download_button(
            "⬇  Download Deal Memo",
            data=pptx_bytes,
            file_name=Path(st.session_state.pptx_path).name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
        if st.session_state.email_sent:
            st.markdown(f"<div style='color:{GREEN};font-size:0.8rem;text-align:center;margin-top:0.4rem;'>✅ Email sent</div>", unsafe_allow_html=True)
        if st.session_state.email_error:
            st.markdown(f"<div style='color:{AMBER};font-size:0.75rem;text-align:center;margin-top:0.4rem;'>⚠ Email failed</div>", unsafe_allow_html=True)

    st.markdown(f"""
    <hr>
    <div style="font-size:0.7rem;color:#50637a;line-height:2.0;">
      <div style="color:{GOLD}cc;font-weight:800;font-size:0.7rem;letter-spacing:0.08em;margin-bottom:0.3rem;">HOW IT WORKS</div>
      🔍 Serper.dev → find startup<br>
      📚 4 searches + page fetch<br>
      🤖 Groq Llama-3.3 → JSON<br>
      📊 python-pptx → 9-slide deck<br>
      📧 Gmail API → email attach
    </div>
    """, unsafe_allow_html=True)


# ── Trigger ────────────────────────────────────────────────────────────────────
if run_btn and not st.session_state.running:
    st.session_state.running     = True
    st.session_state.result      = None
    st.session_state.pptx_path   = None
    st.session_state.log_lines   = []
    st.session_state.email_sent  = False
    st.session_state.email_error = None
    st.session_state.error       = None
    st.session_state.step        = 1
    st.session_state.sector_used = sector_final
    st.session_state.candidate   = ""
    st.session_state.attempt_cur = 1
    st.session_state.attempt_max = MAX_PIPELINE_ATTEMPTS

    lq = queue.Queue()
    st.session_state["_log_q"] = lq
    threading.Thread(
        target=_run_pipeline,
        args=(sector_final, send_email_opt, lq),
        daemon=True,
    ).start()

_FIN_LOADER = (
    '<div class="fin-loader">'
    '<b></b><b></b><b></b><b></b><b></b>'
    '</div>'
)


def _step_html(current_step, is_running, is_done):
    html = '<div class="step-row">'
    for i, (label, icon) in enumerate(zip(STEP_LABELS, STEP_ICONS)):
        n = i + 1
        if is_done or n < current_step:
            cls = "done"
            disp_icon = '<div class="step-icon">✅</div>'
        elif n == current_step and is_running:
            cls = "active"
            disp_icon = _FIN_LOADER
        elif n == current_step and not is_running:
            cls = "done"
            disp_icon = '<div class="step-icon">✅</div>'
        else:
            cls = "pending"
            disp_icon = f'<div class="step-icon">{icon}</div>'
        html += f"""
        <div class="step-item {cls}">
          <div class="step-num">Step {n}</div>
          {disp_icon}
          <div class="step-lbl">{label}</div>
        </div>"""
    html += "</div>"
    return html


# ── Live status panel ──────────────────────────────────────────────────────────
# Uses st.fragment so only this section auto-refreshes every 1.5 s while the
# pipeline runs — the result cards below never blink or flicker.
_poll_interval: float | None = 1.5 if st.session_state.running else None

@st.fragment(run_every=_poll_interval)
def _status_panel():
    _was = st.session_state.running
    _drain()

    active      = st.session_state.running
    has_result  = bool(st.session_state.result)
    done        = has_result and not active        # finished + deck ready
    emailing    = has_result and active            # deck ready, email in flight
    error       = bool(st.session_state.error)
    cur         = st.session_state.step

    if active or done or error or cur > 0:
        _sdots = ('<span class="sdot">.</span>'
                  '<span class="sdot">.</span>'
                  '<span class="sdot">.</span>')
        if emailing:
            company_name = st.session_state.result.get("company_name", "")
            st.markdown(
                f'<div class="status-running">'
                f'<span class="pulse-dot"></span>&nbsp; '
                f'Sending email — <strong>{company_name}</strong>{_sdots}'
                f'</div>',
                unsafe_allow_html=True,
            )
        elif active:
            cname  = st.session_state.candidate
            action = (f"Analysing <strong>{cname}</strong>"
                      if cname else f"Scanning {st.session_state.sector_used}")
            cur_a  = st.session_state.attempt_cur
            max_a  = st.session_state.attempt_max
            retry_tag = (
                f"&nbsp;<span style='font-size:0.7rem;opacity:0.7;font-weight:600;'>"
                f"[Attempt {cur_a}/{max_a}]</span>"
                if cur_a > 1 else ""
            )
            st.markdown(
                f'<div class="status-running">'
                f'<span class="pulse-dot"></span>&nbsp; {action}{_sdots}{retry_tag}'
                f'</div>',
                unsafe_allow_html=True,
            )
        elif done:
            company_name = st.session_state.result.get("company_name", "")
            st.markdown(
                f'<div class="status-done">✅ &nbsp;Deal memo ready — '
                f'<strong>{company_name}</strong></div>',
                unsafe_allow_html=True,
            )
        elif error:
            st.markdown(
                f'<div class="status-running" style="border-color:{RED};'
                f'color:{RED};background:{RED}22;">'
                f'❌ &nbsp;Pipeline stopped — see details below</div>',
                unsafe_allow_html=True,
            )

        st.markdown(_step_html(cur, active, done), unsafe_allow_html=True)

    if st.session_state.log_lines:
        def _log_class(line: str) -> str:
            ll = line.lower()
            if any(k in ll for k in ["[discovery]", "[disco", "candidate:"]):
                return "log-ok"
            if any(k in ll for k in ["[research]", "fetching", "searching"]):
                return "log-info"
            if any(k in ll for k in ["[groq]", "[gemini]", "synthesising"]):
                return "log-ai"
            if "[warn]" in ll:
                return "log-warn"
            if "[error]" in ll:
                return "log-err"
            return ""

        log_html = "\n".join(
            f'<div class="log-line {_log_class(l)}">{l}</div>'
            for l in st.session_state.log_lines
        )
        with st.expander("📋 Live pipeline logs", expanded=active):
            st.markdown(f'<div class="log-box">{log_html}</div>',
                        unsafe_allow_html=True)

    # Pipeline just finished → full app rerun so sidebar button updates
    if _was and not st.session_state.running:
        st.rerun()


_status_panel()


# ── Header ─────────────────────────────────────────────────────────────────────
if not IMPORT_OK:
    st.error(f"Could not import deal_scout.py: {IMPORT_ERR}")
    st.stop()

st.markdown(f"""
<div style="background:linear-gradient(135deg,{NAVY} 0%,#22396a 100%);
            border-radius:14px;padding:1.4rem 2rem;margin-bottom:1.5rem;
            border:1px solid {GOLD}44;display:flex;align-items:center;gap:1.2rem;">
  <div style="font-size:2.5rem;filter:drop-shadow(0 0 8px {GOLD}88);">🏦</div>
  <div>
    <div style="font-size:0.65rem;font-weight:800;color:{GOLD}aa;letter-spacing:0.18em;
                text-transform:uppercase;">MGA Ventures · Internal Tool</div>
    <div style="font-size:1.7rem;font-weight:900;color:white;line-height:1.15;">
      Deal Scout Dashboard
    </div>
    <div style="font-size:0.82rem;color:#8099b8;margin-top:0.1rem;">
      Autonomous startup discovery · AI research · PowerPoint · Email delivery
    </div>
  </div>
</div>
""", unsafe_allow_html=True)






# Convenience aliases used by the result/empty sections below
active = st.session_state.running
done   = bool(st.session_state.result)   # deck ready (running may still be True for email)
error  = bool(st.session_state.error)

# ── Error ──────────────────────────────────────────────────────────────────────
if st.session_state.error:
    st.error(f"**Pipeline failed:** {st.session_state.error}")


# ── Result ─────────────────────────────────────────────────────────────────────
if st.session_state.result:
    d        = st.session_state.result
    company  = d.get("company_name", "Unknown")
    tagline  = d.get("tagline", "")
    sector_v = d.get("sector", "")
    stage    = d.get("stage", "N/A")
    hq       = d.get("hq", "India")
    founded  = d.get("founded_year", "N/A")
    website  = d.get("website", "")

    # Hero
    total_funding = ""
    for fr in (d.get("funding_rounds") or []):
        if fr.get("amount"):
            total_funding = fr["amount"]
            break

    st.markdown(f"""
    <div class="hero">
      <div class="hero-eyebrow">{st.session_state.sector_used} · India · {datetime.date.today().strftime("%d %b %Y")}</div>
      <div class="hero-title">{company}</div>
      <div class="hero-tagline">{tagline or "—"}</div>
      <div class="chip-row">
        <div class="chip">
          <span class="chip-lbl">Sector</span>
          <span class="chip-val">{sector_v or "—"}</span>
        </div>
        <div class="chip">
          <span class="chip-lbl">Stage</span>
          <span class="chip-val">{stage}</span>
        </div>
        <div class="chip">
          <span class="chip-lbl">HQ</span>
          <span class="chip-val">{hq}</span>
        </div>
        <div class="chip">
          <span class="chip-lbl">Founded</span>
          <span class="chip-val">{founded}</span>
        </div>
        {f'<div class="chip"><span class="chip-lbl">Latest Funding</span><span class="chip-val">{total_funding}</span></div>' if total_funding else ""}
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Main columns ────────────────────────────────────────────────────────────
    lcol, rcol = st.columns([3, 2], gap="large")

    with lcol:
        # Problem / Solution / Business model
        for key, label, border in [
            ("problem",        "Problem",        RED),
            ("solution",       "Solution",       GREEN),
            ("business_model", "Business Model", NAVY),
        ]:
            text = d.get(key) or "Not publicly disclosed."
            st.markdown(f"""
            <div class="icard" style="border-left-color:{border}">
              <div class="icard-title">{label}</div>
              <p class="icard-body">{text}</p>
            </div>
            """, unsafe_allow_html=True)

        # Traction metrics
        metrics = d.get("traction_metrics") or []
        if metrics:
            st.markdown(f'<div class="sec-label" style="color:{GREEN}">Traction Metrics</div>', unsafe_allow_html=True)
            mhtml = '<div class="metric-grid">'
            for m in metrics[:6]:
                src = m.get("source_url", "")
                outlet = re.sub(r"https?://(?:www\.)?([^/]+).*", r"\1", src) if src else ""
                mhtml += f"""<div class="mcard">
                  <div class="mcard-label">{m.get('metric','')}</div>
                  <div class="mcard-value">{m.get('value','')}</div>
                  {"<div class='mcard-source'>" + outlet + "</div>" if outlet else ""}
                </div>"""
            mhtml += "</div>"
            st.markdown(mhtml, unsafe_allow_html=True)

        # Key differentiators
        diffs = d.get("key_differentiators") or []
        if diffs:
            st.markdown(f'<div class="sec-label" style="color:{NAVY}">Key Differentiators</div>', unsafe_allow_html=True)
            for diff in diffs[:4]:
                st.markdown(f'<div class="hl-item" style="border-color:{GOLD}">⭐ {diff}</div>', unsafe_allow_html=True)

        # Investment highlights
        highlights = d.get("investment_highlights") or []
        if highlights:
            st.markdown(f'<div class="sec-label" style="color:{GREEN};margin-top:0.6rem;">Investment Highlights</div>', unsafe_allow_html=True)
            for h in highlights[:3]:
                st.markdown(f'<div class="hl-item">✦ {h}</div>', unsafe_allow_html=True)

        # Red flags
        red_flags = d.get("red_flags") or []
        if red_flags:
            st.markdown(f'<div class="sec-label" style="color:{RED};margin-top:0.6rem;">Red Flags</div>', unsafe_allow_html=True)
            for f in red_flags[:3]:
                st.markdown(f'<div class="rf-item">⚠ {f}</div>', unsafe_allow_html=True)

    with rcol:
        # Thesis fit
        criteria = [
            ("Investment Stage", d.get("thesis_fit_stage", "")),
            ("Sector Fit",       d.get("thesis_fit_sector", "")),
            ("Ticket Size",      d.get("thesis_fit_ticket", "")),
        ]

        def _fit_color(text: str) -> str:
            t = text.lower()
            if "out of scope" in t or "not publicly" in t:
                return RED
            if "partial" in t:
                return AMBER
            return GREEN

        def _fit_verdict(text: str) -> str:
            t = text.lower()
            if "out of scope" in t:
                return "OUT OF SCOPE"
            if "partial" in t:
                return "PARTIAL FIT"
            if "not publicly" in t:
                return "UNKNOWN"
            return "FIT ✓"

        st.markdown(f'<div class="sec-label" style="color:{NAVY}">MGA Thesis Fit</div>', unsafe_allow_html=True)
        thtml = '<div class="thesis-grid">'
        for lbl, text in criteria:
            col_hex = _fit_color(text)
            verdict = _fit_verdict(text)
            short   = (text[:75] + "…") if len(text) > 75 else text
            thtml += f"""<div class="tcard" style="background:{col_hex};">
              <div class="tcard-label">{lbl}</div>
              <div class="tcard-verdict">{verdict}</div>
              <div class="tcard-detail">{short}</div>
            </div>"""
        thtml += "</div>"
        st.markdown(thtml, unsafe_allow_html=True)

        # Founders
        founders = d.get("founders") or []
        if founders:
            st.markdown(f'<div class="sec-label" style="color:{NAVY};margin-top:0.3rem;">Founders</div>', unsafe_allow_html=True)
            for f in founders[:3]:
                li = f.get("linkedin_url", "")
                li_html = f'<a href="{li}" target="_blank" style="font-size:0.65rem;color:#0066cc;text-decoration:none;">in LinkedIn ↗</a>' if li else ""
                st.markdown(f"""<div class="founder-card">
                  <div class="founder-name">{f.get('name','')}</div>
                  <div class="founder-title">{f.get('title','')}</div>
                  <div class="founder-bio">{f.get('background','')}</div>
                  {li_html}
                </div>""", unsafe_allow_html=True)

        # Funding rounds
        rounds = d.get("funding_rounds") or []
        if rounds:
            st.markdown(f'<div class="sec-label" style="color:{NAVY};margin-top:0.5rem;">Funding History</div>', unsafe_allow_html=True)
            for r in rounds[:4]:
                src = r.get("source_url", "")
                outlet = re.sub(r"https?://(?:www\.)?([^/]+).*", r"\1", src) if src else ""
                src_link = f'<a href="{src}" target="_blank" style="font-size:0.65rem;color:#0066cc;text-decoration:none;">{outlet} ↗</a>' if src and outlet else ""
                st.markdown(f"""<div class="fr-card">
                  <div>
                    <div><span class="fr-round">{r.get('round','')}</span>
                         <span class="fr-amount">{r.get('amount','')}</span></div>
                    <div class="fr-meta">{r.get('investors','')}  ·  {r.get('date','')}</div>
                    {src_link}
                  </div>
                </div>""", unsafe_allow_html=True)

        # Use of funds
        uof = d.get("use_of_funds", "")
        if uof and "not publicly" not in uof.lower():
            st.markdown(f"""
            <div class="icard" style="border-left-color:{GOLD};margin-top:0.4rem;">
              <div class="icard-title" style="color:{GOLD};">Use of Funds</div>
              <p class="icard-body">{uof}</p>
            </div>
            """, unsafe_allow_html=True)

    # ── Download + email row ────────────────────────────────────────────────────
    st.markdown("---")
    dc, ec, _ = st.columns([2, 2, 1])
    with dc:
        if st.session_state.pptx_path and Path(st.session_state.pptx_path).exists():
            st.download_button(
                "⬇  Download 9-Slide Deal Memo (PPTX)",
                data=Path(st.session_state.pptx_path).read_bytes(),
                file_name=Path(st.session_state.pptx_path).name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )
    with ec:
        if st.session_state.email_sent:
            st.success("✅ Email sent successfully")
        if st.session_state.email_error:
            st.warning(f"⚠ Email failed: {st.session_state.email_error[:120]}")
        if st.session_state.pptx_path and not st.session_state.email_sent and not st.session_state.email_error:
            st.info("Email sending was skipped.")

# ── Empty state — layman workflow explanation ──────────────────────────────────
elif not active and not error:
    st.markdown(f"""
    <div style="text-align:center;padding:3rem 1rem 1.5rem;">
      <div style="font-size:4rem;filter:drop-shadow(0 4px 12px {GOLD}44);">🔭</div>
      <h2 style="color:#223;font-weight:900;margin:0.6rem 0 0.3rem;font-size:1.8rem;">How Deal Scout Works</h2>
      <p style="font-size:0.9rem;color:#778;margin-bottom:2rem;">
        Pick a sector on the left and click <b style="color:{NAVY};">▶ Run Deal Scout</b>.
        The whole process takes about 60–90 seconds and needs no input from you.
      </p>
    </div>
    """, unsafe_allow_html=True)

    steps_exp = [
        ("🔍", "Step 1 — Find a startup",
         "Consumer", "Finance", "Fintech",
         f"We type search queries like <i>\"Fintech startup India raises crore 2025\"</i> into Google "
         f"(via Serper.dev) and look at results from trusted Indian business news sites — "
         f"<b>Inc42, YourStory, Entrackr, Economic Times</b>. Each result gets a score based on "
         f"signals like city names (Bengaluru, Mumbai), funding keywords (seed, Series A), "
         f"and recency. The highest-scoring result that names a real company becomes our candidate."),
        ("📚", "Step 2 — Deep-dive research",
         None, None, None,
         f"Once we have a company name, we fire off <b>4 more targeted searches</b> — one each for "
         f"products, founders, funding history, and traction metrics. We then open the most relevant "
         f"article from each search and read its full text, building up a research corpus of "
         f"raw facts sourced only from credible outlets."),
        ("🤖", "Step 3 — AI reads everything",
         None, None, None,
         f"All those articles are sent to <b>Groq's Llama 3.3 AI</b> with a strict prompt: "
         f"<i>\"Extract only facts that are explicitly stated. Write 'Not publicly disclosed.' for "
         f"anything you can't find. Never invent numbers.\"</i> The AI returns a structured JSON "
         f"covering 25+ fields — problem, solution, funding rounds, founders, thesis fit, and more."),
        ("📊", "Step 4 — Build the PowerPoint",
         None, None, None,
         f"The structured data is turned into a <b>9-slide professional deal memo</b> using python-pptx. "
         f"Every number shown in the deck is traced back to a source URL. Slides include: "
         f"Title · Company Overview · What It Does · Products · Team · "
         f"Market Opportunity · Funding & Traction · Thesis Fit · Sources."),
        ("📧", "Step 5 — Email it",
         None, None, None,
         f"The finished PPTX is attached to an email and sent via <b>Gmail API</b> to the address "
         f"in your .env file. You can also download it directly from this page using the button "
         f"that appears in the sidebar once the run completes."),
    ]

    STEP_COLORS = [NAVY, "#2E5FA3", GREEN, AMBER, "#8B2252"]

    for idx, (icon, title, *_, body) in enumerate(steps_exp):
        col_accent = STEP_COLORS[idx]
        st.markdown(f"""
        <div style="display:flex;gap:1.2rem;background:white;border-radius:12px;
                    padding:1.2rem 1.5rem;margin-bottom:0.85rem;
                    border:1px solid #e0e8f4;border-left:5px solid {col_accent};
                    box-shadow:0 2px 8px #0001;">
          <div style="font-size:2rem;flex-shrink:0;margin-top:0.15rem;">{icon}</div>
          <div>
            <div style="font-weight:800;color:{col_accent};font-size:1rem;margin-bottom:0.35rem;">{title}</div>
            <div style="font-size:0.88rem;color:#445;line-height:1.65;">{body}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{NAVY},#22396a);border-radius:12px;
                padding:1.2rem 1.8rem;margin-top:0.5rem;border:1px solid {GOLD}44;">
      <div style="font-size:0.7rem;font-weight:800;color:{GOLD};letter-spacing:0.12em;
                  text-transform:uppercase;margin-bottom:0.6rem;">APIs used — all free tier</div>
      <div style="display:flex;gap:2rem;flex-wrap:wrap;">
        <div style="color:#b8c8e0;font-size:0.83rem;">
          <b style="color:white;">Serper.dev</b><br>Google search wrapper<br>
          <span style="color:{GOLD}88;font-size:0.75rem;">~6 queries per run · 2,500/mo free</span>
        </div>
        <div style="color:#b8c8e0;font-size:0.83rem;">
          <b style="color:white;">Groq (Llama 3.3)</b><br>AI synthesis engine<br>
          <span style="color:{GOLD}88;font-size:0.75rem;">1 call per run · 14,400/day free</span>
        </div>
        <div style="color:#b8c8e0;font-size:0.83rem;">
          <b style="color:white;">Gmail API</b><br>Email delivery<br>
          <span style="color:{GOLD}88;font-size:0.75rem;">OAuth2 · free</span>
        </div>
        <div style="color:#b8c8e0;font-size:0.83rem;">
          <b style="color:white;">Clearbit Logo</b><br>Company logo fetch<br>
          <span style="color:{GOLD}88;font-size:0.75rem;">Free, no auth needed</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)


