import streamlit as st
import io
import os
from datetime import date, timedelta

# ── Page config ────────────────────────────────────────────────
st.set_page_config(
    page_title="Carlsquare Timeline Generator",
    layout="centered",
)

# ── Password gate ───────────────────────────────────────────────
PASSWORD = "carlsquare2026"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("## Carlsquare Timeline Generator")
    st.markdown("Enter the password to continue.")
    pw = st.text_input("Password", type="password", key="pw_input")
    if st.button("Login"):
        if pw == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password. Please try again.")
    st.stop()

# ── Import generator ──────────────────────────────────────────
from timeline_slide_generator import render_timeline_slide, TEMPLATES, compute_schedule

# ── Styles ─────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #f7f8fc; }
    .card {
        background: white; border-radius: 12px; padding: 2rem 2.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08); margin-bottom: 1.5rem;
    }
    .section-label {
        font-size: 0.75rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.08em; color: #8892a4; margin-bottom: 0.5rem;
    }
    div.stButton > button, div.stFormSubmitButton > button {
        background-color: #12213D !important; color: white !important;
        border: none !important; border-radius: 8px !important;
        padding: 0.65rem 2.5rem !important; font-size: 1rem !important;
        font-weight: 600 !important; width: 100% !important; transition: background 0.2s;
    }
    div.stButton > button:hover, div.stFormSubmitButton > button:hover {
        background-color: #1a3060 !important; color: white !important;
    }
    div.stButton > button p, div.stButton > button span,
    div.stFormSubmitButton > button p, div.stFormSubmitButton > button span {
        color: white !important;
    }
    div.stDownloadButton > button {
        background-color: #009A8F !important; color: white !important;
        border: none !important; border-radius: 8px !important;
        padding: 0.65rem 2.5rem !important; font-size: 1rem !important;
        font-weight: 600 !important; width: 100% !important;
    }
    div.stDownloadButton > button:hover { background-color: #007a72 !important; color: white !important; }
    div.stDownloadButton > button p, div.stDownloadButton > button span { color: white !important; }
    hr { border-color: #e8ecf2; margin: 1.5rem 0; }
    h1, h2, h3, h4, p, label, .stMarkdown, .stText { color: #12213D !important; }
    [data-testid="stAppViewContainer"] { background-color: #f7f8fc !important; }
    [data-testid="stHeader"] { background-color: #f7f8fc !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}

    /* Hidden placeholder submit button */
    div.stFormSubmitButton > button:disabled { display: none !important; }

    /* Delete (✕) buttons — force white text in all child elements */
    button[data-testid="baseButton-secondary"] { color: white !important; }
    button[data-testid="baseButton-secondary"] p,
    button[data-testid="baseButton-secondary"] span,
    button[data-testid="baseButton-secondary"] div { color: white !important; }


    /* Force ALL button text white — covers p, span, div inside buttons */
    div.stButton > button *, div.stButton > button {
        color: white !important;
    }
    [data-testid="stExpander"] details summary {
        background-color: #12213D !important; border-radius: 8px; color: white !important;
    }
    [data-testid="stExpander"] details summary p,
    [data-testid="stExpander"] details summary svg { color: white !important; fill: white !important; }
    [data-testid="stDateInput"] input,
    [data-testid="stSelectbox"] > div > div {
        background-color: #12213D !important; color: white !important;
        border-radius: 8px !important; border: none !important;
    }
    [data-testid="stSelectbox"] svg { fill: white !important; }
    .ws-phase-header {
        background: #12213D; color: white !important; font-size: 0.8rem;
        font-weight: 700; padding: 0.4rem 0.8rem; border-radius: 6px;
        margin: 0.8rem 0 0.4rem 0;
    }
    .ws-help { font-size: 0.78rem; color: #8892a4; margin-bottom: 0.6rem; font-style: italic; }

</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────
_, col_title = st.columns([1, 4])
with col_title:
    st.markdown("<h2 style='color:#12213D;margin-bottom:0'>Timeline Generator</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#8892a4;margin-top:-0.5rem'>Generate a Carlsquare M&A process timeline slide</p>",
                unsafe_allow_html=True)

st.markdown("""
<div style="background:white; border-radius:12px; padding:1.2rem 1.6rem; margin-bottom:1rem; border-left:4px solid #009A8F;">
    <p style="margin:0 0 0.6rem 0; color:#12213D; font-size:0.95rem;">
        Fill in the options below and click <strong>Generate Timeline</strong> to download a ready-made PowerPoint slide.
        The output is a fully editable <strong>.pptx file</strong>.
    </p>
    <ul style="margin:0; padding-left:1.2rem; color:#12213D; font-size:0.9rem; line-height:1.8;">
        <li><strong>Close Date</strong> — the expected signing/closing date. Everything else is calculated from this.</li>
        <li><strong>Standard</strong> — full ~8 month process: buyer list, CIM, outreach, IOI, management presentations, LOI, DD, and close.</li>
        <li><strong>Accelerated</strong> — compressed (~6 months) when an offer is already on the table.</li>
        <li><strong>Dark theme</strong> — navy background. <strong>Light theme</strong> — white background.</li>
        <li><strong>Customise Text</strong> — override the auto-generated subtitle and the section label in the top-left corner.</li>
        <li><strong>Edit Workstreams</strong> — toggle individual rows on/off, rename labels, adjust start/end dates, or add brand new rows to any phase. Use 'Reset to defaults' to undo all changes.</li>
    </ul>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# Workstream editor helpers
# ══════════════════════════════════════════════════════════════════

def build_ws_state_from_template(process: str, close_date: date) -> list:
    """Flatten all template rows into an editable list of dicts."""
    tmpl = TEMPLATES[process]
    rows = []
    for phase in tmpl["phases"]:
        for r in phase["rows"]:
            is_milestone = r.get("type") == "milestone_row"
            rows.append({
                "row_id":             r["row_id"],
                "phase_label":        phase["phase_label"],
                "phase_id":           phase["phase_id"],
                "label":              r["label"],
                "include":            r.get("include", True),
                "is_milestone":       is_milestone,
                "milestone_key":      r.get("milestone_key"),
                "start_date":         None if is_milestone else (
                    close_date + timedelta(days=int(r["start_offset_weeks"]) * 7)),
                "end_date":           None if is_milestone else (
                    close_date + timedelta(days=int(r["end_offset_weeks"]) * 7)),
                "start_offset_weeks": None if is_milestone else r.get("start_offset_weeks"),
                "end_offset_weeks":   None if is_milestone else r.get("end_offset_weeks"),
                "is_custom":          False,
            })
    return rows


def ws_state_key(process: str) -> str:
    return f"ws_rows_{process}"


def strip_subnumber(label: str) -> str:
    """Remove leading 'a. ', 'b. ' etc from a label so we can re-add it automatically."""
    import re
    return re.sub(r'^[a-zA-Z]\.\s+', '', label.strip())


def apply_subnumbers(rows: list):
    """Re-apply a. b. c. prefixes to non-milestone rows within each phase, in order."""
    import string
    phase_counters = {}
    for r in rows:
        if r['is_milestone']:
            continue
        pid = r['phase_id']
        phase_counters[pid] = phase_counters.get(pid, 0) + 1
        letter = string.ascii_lowercase[phase_counters[pid] - 1]
        base = strip_subnumber(r['label'])
        r['label'] = f"{letter}. {base}"


def ensure_ws_state(process: str, close_date: date):
    key = ws_state_key(process)
    if key not in st.session_state:
        rows = build_ws_state_from_template(process, close_date)
        apply_subnumbers(rows)
        st.session_state[key] = rows


def recalc_from_offsets(rows: list, close_date: date):
    """Recalculate dates from week offsets for non-custom rows when close date changes."""
    for r in rows:
        if not r["is_milestone"] and not r["is_custom"] and r["start_offset_weeks"] is not None:
            r["start_date"] = close_date + timedelta(days=int(r["start_offset_weeks"]) * 7)
            r["end_date"]   = close_date + timedelta(days=int(r["end_offset_weeks"]) * 7)


def rows_to_custom_template(rows: list, process: str) -> dict:
    """Convert the edited flat list back into a template dict for the generator."""
    import copy
    base = copy.deepcopy(TEMPLATES[process])

    # Group rows by phase preserving order
    seen, phase_rows_map = [], {}
    for r in rows:
        pid = r["phase_id"]
        if pid not in seen:
            seen.append(pid)
        phase_rows_map.setdefault(pid, []).append(r)

    new_phases = []
    for pid in seen:
        phase_rows = phase_rows_map[pid]
        task_rows = []
        for r in phase_rows:
            if not r["include"]:
                continue
            if r["is_milestone"]:
                task_rows.append({
                    "row_id":        r["row_id"],
                    "label":         r["label"],
                    "include":       True,
                    "type":          "milestone_row",
                    "milestone_key": r["milestone_key"],
                })
            else:
                task_rows.append({
                    "row_id":              r["row_id"],
                    "label":               r["label"],
                    "include":             True,
                    "start_date_override": r["start_date"],
                    "end_date_override":   r["end_date"],
                    "start_offset_weeks":  0,
                    "end_offset_weeks":    0,
                })
        if task_rows:
            new_phases.append({
                "phase_id":    pid,
                "phase_label": phase_rows[0]["phase_label"],
                "rows":        task_rows,
            })

    return {"anchors": base["anchors"], "phases": new_phases}


# ── Workstream editor UI ────────────────────────────────────────
def render_ws_editor(process: str, close_date: date):
    key  = ws_state_key(process)
    rows = st.session_state[key]

    # Re-sync dates from offsets if close date has changed
    recalc_from_offsets(rows, close_date)

    st.markdown(
        "<p class='ws-help'>Toggle rows on/off, rename labels, or adjust start/end dates. "
        "Milestone rows (IOI, LOI, Close) are date-locked to the anchor offsets. "
        "Add completely new rows at the bottom.</p>",
        unsafe_allow_html=True,
    )

    # Column headers
    h = st.columns([0.45, 3.3, 1.9, 1.9, 0.55])
    for col, lbl in zip(h, ["On", "Workstream label", "Start date", "End date", "Del"]):
        col.markdown(f"<small><b>{lbl}</b></small>", unsafe_allow_html=True)
    st.markdown("<hr style='margin:0.25rem 0 0.4rem 0'>", unsafe_allow_html=True)

    current_phase = None
    to_delete = []

    for idx, r in enumerate(rows):
        if r["phase_label"] != current_phase:
            current_phase = r["phase_label"]
            st.markdown(
                f"<div class='ws-phase-header'>{current_phase}</div>",
                unsafe_allow_html=True,
            )

        cols = st.columns([0.45, 3.3, 1.9, 1.9, 0.55])

        r["include"] = cols[0].checkbox(
            "on", value=r["include"],
            key=f"wsinc_{process}_{idx}",
            label_visibility="collapsed",
        )
        r["label"] = cols[1].text_input(
            "lbl", value=r["label"],
            key=f"wslbl_{process}_{idx}",
            label_visibility="collapsed",
            disabled=r["is_milestone"],
        )

        if r["is_milestone"]:
            cols[2].caption("— milestone —")
            cols[3].caption("— milestone —")
        else:
            prev_sd, prev_ed = r["start_date"], r["end_date"]
            new_sd = cols[2].date_input(
                "sd", value=r["start_date"],
                key=f"wssd_{process}_{idx}",
                label_visibility="collapsed",
            )
            new_ed = cols[3].date_input(
                "ed", value=r["end_date"],
                key=f"wsed_{process}_{idx}",
                label_visibility="collapsed",
            )
            if new_sd != prev_sd or new_ed != prev_ed:
                r["is_custom"] = True
            r["start_date"] = new_sd
            r["end_date"]   = new_ed

        if not r["is_milestone"]:
            if cols[4].button("✕", key=f"wsdel_{process}_{idx}", help="Remove row"):
                to_delete.append(idx)
        else:
            cols[4].write("")

    for idx in sorted(to_delete, reverse=True):
        st.session_state[key].pop(idx)
    if to_delete:
        apply_subnumbers(st.session_state[key])
        st.rerun()

    # ── Add new row ─────────────────────────────────────────────
    st.markdown("---")
    st.markdown("**➕ Add a new workstream row**")
    ac = st.columns([2, 2.5, 1.8, 1.8])

    phase_options = list(dict.fromkeys(r["phase_label"] for r in rows))
    new_phase = ac[0].selectbox("Add to phase", phase_options, key=f"wsnewphase_{process}")
    new_label = ac[1].text_input("Label", placeholder="e.g. Legal review", key=f"wsnewlbl_{process}")
    new_sd    = ac[2].date_input("Start", value=close_date - timedelta(weeks=12), key=f"wsnewsd_{process}")
    new_ed    = ac[3].date_input("End",   value=close_date - timedelta(weeks=8),  key=f"wsnewed_{process}")

    if st.button("➕ Add row", key=f"wsaddrow_{process}"):
        if new_label.strip():
            phase_id = next((r["phase_id"] for r in rows if r["phase_label"] == new_phase), "custom")
            existing_ids = {r["row_id"] for r in rows}
            new_id = f"custom_{len(rows)}"
            while new_id in existing_ids:
                new_id += "_x"

            insert_at = max(
                (i for i, r in enumerate(rows) if r["phase_label"] == new_phase),
                default=len(rows) - 1,
            ) + 1

            st.session_state[key].insert(insert_at, {
                "row_id":             new_id,
                "phase_label":        new_phase,
                "phase_id":           phase_id,
                "label":              new_label.strip(),
                "include":            True,
                "is_milestone":       False,
                "milestone_key":      None,
                "start_date":         new_sd,
                "end_date":           new_ed,
                "start_offset_weeks": None,
                "end_offset_weeks":   None,
                "is_custom":          True,
            })
            apply_subnumbers(st.session_state[key])
            st.rerun()
        else:
            st.warning("Please enter a label before adding.")

    # ── Reset ────────────────────────────────────────────────────
    st.markdown("")
    if st.button("↺ Reset workstreams to defaults", key=f"wsreset_{process}"):
        rows = build_ws_state_from_template(process, close_date)
        apply_subnumbers(rows)
        st.session_state[key] = rows
        st.rerun()


# ══════════════════════════════════════════════════════════════════
# Main form
# ══════════════════════════════════════════════════════════════════
with st.form("timeline_form"):

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("<p class='section-label'>Close Date</p>", unsafe_allow_html=True)
        close_date = st.date_input(
            "Close Date",
            value=date.today() + timedelta(weeks=32),
            min_value=date.today(),
            label_visibility="collapsed",
        )
    with col2:
        st.markdown("<p class='section-label'>Process Type</p>", unsafe_allow_html=True)
        process_type = st.selectbox(
            "Process Type",
            options=["Standard", "Accelerated"],
            help="Standard: ~8 month process.  Accelerated: ~6 month process.",
            label_visibility="collapsed",
        )

    st.markdown("<p class='section-label'>Slide Theme</p>", unsafe_allow_html=True)
    theme_choice = st.radio(
        "Theme",
        options=["🌙 Dark", "☀️ Light"],
        horizontal=True,
        label_visibility="collapsed",
    )

    with st.expander("Customise Text (optional)"):
        st.markdown("<p class='section-label'>Subtitle Line</p>", unsafe_allow_html=True)
        subtitle = st.text_input(
            "Subtitle",
            placeholder='e.g. "Begin process April 1st, launch to market early June, expected close in November"',
            label_visibility="collapsed",
        )
        st.markdown("<p class='section-label' style='margin-top:1rem'>Section Label (top-left)</p>",
                    unsafe_allow_html=True)
        top_label = st.text_input(
            "Top label",
            value="3 | Process design and investor discussion",
            label_visibility="collapsed",
        )

    # hidden placeholder — form requires a submit button but we use the one outside
    st.form_submit_button("⚡  Generate Timeline", disabled=True, use_container_width=False)


# ── Workstream editor (outside the form so widgets are interactive)
process_key = process_type.lower()
ensure_ws_state(process_key, close_date)

with st.expander("Edit Workstreams (optional)"):
    render_ws_editor(process_key, close_date)

st.markdown(" ")
submitted = st.button("⚡  Generate Timeline", key="generate_btn", use_container_width=True)

# ── Generate ────────────────────────────────────────────────────
if submitted:
    theme    = "dark"  if "Dark"  in theme_choice else "light"
    process  = process_type.lower()
    out_name = f"Carlsquare_Timeline_{process.title()}_{theme.title()}_{close_date.strftime('%Y-%m-%d')}.pptx"

    ws_key      = ws_state_key(process)
    apply_subnumbers(st.session_state[ws_key])   # ensure letters are current before render
    custom_tmpl = rows_to_custom_template(st.session_state[ws_key], process)

    with st.spinner("Building your timeline slide…"):
        buf = io.BytesIO()
        try:
            render_timeline_slide(
                close_date      = close_date,
                process         = process,
                theme_name      = theme,
                out_path        = buf,
                subtitle        = subtitle.strip(),
                top_label       = top_label.strip() or "[Insert section label here]",
                custom_template = custom_tmpl,
            )
            buf.seek(0)

            st.success("✅ Your timeline is ready!")
            st.download_button(
                label     = "⬇️ Download .pptx",
                data      = buf,
                file_name = out_name,
                mime      = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

            _, milestone_dates = compute_schedule(close_date, custom_tmpl)
            launch = milestone_dates.get("launch_to_market")
            ioi    = milestone_dates.get("ioi_due")
            loi    = milestone_dates.get("loi_due")

            st.markdown("---")
            st.markdown("#### Key dates in your timeline")
            m1, m2, m3, m4 = st.columns(4)
            if launch: m1.metric("Launch to market", launch.strftime("%b %d, %Y"))
            if ioi:    m2.metric("IOI due",          ioi.strftime("%b %d, %Y"))
            if loi:    m3.metric("LOI due",           loi.strftime("%b %d, %Y"))
            m4.metric("Close", close_date.strftime("%b %d, %Y"))

        except Exception as e:
            st.error(f"Something went wrong: {e}")
            import traceback
            st.code(traceback.format_exc())
