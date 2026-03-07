import streamlit as st
import io
import os
from datetime import date, timedelta

# ── Page config ────────────────────────────────────────────────
st.set_page_config(
    page_title="Carlsquare Timeline Generator",
    page_icon="📊",
    layout="centered",
)

# ── Import generator (must be in same folder) ──────────────────
from timeline_slide_generator import render_timeline_slide

# ── Styles ─────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #f7f8fc; }

    /* Card container */
    .card {
        background: white;
        border-radius: 12px;
        padding: 2rem 2.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        margin-bottom: 1.5rem;
    }

    /* Section headings inside cards */
    .section-label {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #8892a4;
        margin-bottom: 0.5rem;
    }

    /* All buttons — generate + download */
    div.stButton > button,
    div.stFormSubmitButton > button {
        background-color: #12213D !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.65rem 2.5rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        transition: background 0.2s;
    }
    div.stButton > button:hover,
    div.stFormSubmitButton > button:hover {
        background-color: #1a3060 !important;
        color: white !important;
    }
    div.stButton > button p,
    div.stButton > button span,
    div.stFormSubmitButton > button p,
    div.stFormSubmitButton > button span {
        color: white !important;
    }

    /* Download button */
    div.stDownloadButton > button {
        background-color: #009A8F !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.65rem 2.5rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        width: 100% !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #007a72 !important;
        color: white !important;
    }
    div.stDownloadButton > button p,
    div.stDownloadButton > button span {
        color: white !important;
    }
    /* Divider */
    hr { border-color: #e8ecf2; margin: 1.5rem 0; }

    /* Force all headings and body text to dark */
    h1, h2, h3, h4, p, label, .stMarkdown, .stText { color: #12213D !important; }
    /* Override any dark mode that Streamlit might inject */
    [data-testid="stAppViewContainer"] { background-color: #f7f8fc !important; }
    [data-testid="stHeader"] { background-color: #f7f8fc !important; }

    /* Hide streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Expander header — navy background, white text */
    [data-testid="stExpander"] details summary {
        background-color: #12213D !important;
        border-radius: 8px;
        color: white !important;
    }
    [data-testid="stExpander"] details summary p,
    [data-testid="stExpander"] details summary svg {
        color: white !important;
        fill: white !important;
    }

    /* Input fields, dropdowns — navy background, white text */
    [data-testid="stDateInput"] input,
    [data-testid="stSelectbox"] > div > div {
        background-color: #12213D !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
    }
    /* Dropdown arrow icon */
    [data-testid="stSelectbox"] svg {
        fill: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 4])
with col_logo:
    logo_path = os.path.join(os.path.dirname(__file__), "logo_light.png")
    if os.path.exists(logo_path):
        st.image(logo_path, width=90)
with col_title:
    st.markdown("<h2 style='color:#12213D;margin-bottom:0'>Timeline Generator</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#8892a4;margin-top:-0.5rem'>Generate a Carlsquare M&A process timeline slide</p>",
                unsafe_allow_html=True)

st.markdown("""
<div style="background:white; border-radius:12px; padding:1.2rem 1.6rem; margin-bottom:1rem; border-left:4px solid #009A8F;">
    <p style="margin:0 0 0.6rem 0; color:#12213D; font-size:0.95rem;">
        Fill in the options below and click <strong>Generate Timeline</strong> to download a ready-made PowerPoint slide
        you can drop straight into any existing presentation. The output is a fully editable <strong>.pptx file</strong>.
    </p>
    <ul style="margin:0; padding-left:1.2rem; color:#12213D; font-size:0.9rem; line-height:1.8;">
        <li><strong>Close Date</strong> — the expected signing/closing date of the deal. Everything else is calculated from this.</li>
        <li><strong>Standard</strong> — a full ~8 month process built from scratch: buyer list, CIM, outreach, IOI, management presentations, LOI, due diligence, and close.</li>
        <li><strong>Accelerated</strong> — a compressed process (~6 months) used when an offer is already on the table. Skips some early prep steps and moves faster to due diligence.</li>
        <li><strong>Dark theme</strong> — navy background.</li>
        <li><strong>Light theme</strong> — white background.</li>
    </ul>
</div>
""", unsafe_allow_html=True)

# ── Form ────────────────────────────────────────────────────────
with st.form("timeline_form"):

    # ── Row 1: Close date + Process type ──
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
            help="Standard: ~8 month process.  Accelerated: ~6 month process (existing offer on table).",
            label_visibility="collapsed",
        )

    # ── Row 2: Theme ──
    st.markdown("<p class='section-label'>Slide Theme</p>", unsafe_allow_html=True)
    theme_choice = st.radio(
        "Theme",
        options=["Dark 🌙", "Light ☀️"],
        horizontal=True,
        label_visibility="collapsed",
    )

    # ── Optional fields ──
    with st.expander("Optional: Customise Text"):
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

    st.markdown(" ")

    submitted = st.form_submit_button("⚡  Generate Timeline")

# ── Generate ────────────────────────────────────────────────────
if submitted:
    theme     = "dark"  if "Dark"  in theme_choice else "light"
    process   = process_type.lower()
    out_name  = f"Carlsquare_Timeline_{process.title()}_{theme.title()}_{close_date.strftime('%Y-%m-%d')}.pptx"

    with st.spinner("Building your timeline slide…"):
        buf = io.BytesIO()
        try:
            render_timeline_slide(
                close_date  = close_date,
                process     = process,
                theme_name  = theme,
                out_path    = buf,
                subtitle    = subtitle.strip(),
                top_label   = top_label.strip() or "[Insert section label here]",
            )
            buf.seek(0)

            st.success("Your timeline is ready! ✅")
            st.download_button(
                label     = "Download .pptx ⬇️",
                data      = buf,
                file_name = out_name,
                mime      = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

            # Preview info card
            from timeline_slide_generator import TEMPLATES, compute_schedule
            tmpl = TEMPLATES[process]
            _, milestone_dates = compute_schedule(close_date, tmpl)
            launch = milestone_dates["launch_to_market"]
            ioi    = milestone_dates["ioi_due"]
            loi    = milestone_dates["loi_due"]

            st.markdown("---")
            st.markdown("#### Key dates in your timeline")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Launch to market", launch.strftime("%b %d, %Y"))
            m2.metric("IOI due",          ioi.strftime("%b %d, %Y"))
            m3.metric("LOI due",          loi.strftime("%b %d, %Y"))
            m4.metric("Close",            close_date.strftime("%b %d, %Y"))

        except Exception as e:
            st.error(f"Something went wrong: {e}")
