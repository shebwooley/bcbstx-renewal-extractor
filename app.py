from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from bcbstx_extractor import run_extraction, build_validation_notes, build_plan_summary_df

# ── Feature flags ──────────────────────────────────────────────────────────────
# Flip to True when the quoting tool is ready to consume JSON output.
SHOW_JSON_DOWNLOAD = False

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BCBSTX Renewal Extractor",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Give enough top room so the header's rounded corners aren't clipped */
    .block-container { padding-top: 3.5rem !important; }

    .app-header {
        background: linear-gradient(135deg, #003087 0%, #1a5fb4 100%);
        padding: 1.6rem 2rem 1.4rem;
        border-radius: 0.85rem;
        margin-bottom: 1.5rem;
        color: white;
    }
    .app-header h1 {
        margin: 0 0 0.3rem;
        font-size: 1.75rem;
        font-weight: 700;
        letter-spacing: -0.01em;
    }
    .app-header p {
        margin: 0;
        opacity: 0.88;
        font-size: 0.92rem;
        line-height: 1.5;
    }

    [data-testid="metric-container"] {
        background: #f0f4ff;
        border: 1px solid #c7d7f5;
        border-radius: 0.75rem;
        padding: 0.85rem 1rem;
    }

    .stDownloadButton > button {
        width: 100%;
        border-radius: 0.5rem;
        font-weight: 600;
    }

    .footer-cap {
        text-align: center;
        font-size: 0.78rem;
        color: #9ca3af;
        margin-top: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📋 BCBSTX Renewal Extractor</h1>
    <p>Upload a Blue Cross Blue Shield of Texas fully-insured small group renewal PDF.<br>
    The tool extracts plan data and census details into a formatted Excel review file — no AI, no data storage.</p>
</div>
""", unsafe_allow_html=True)

# ── File upload ────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload BCBSTX renewal PDF",
    type=["pdf"],
    help="The renewal offer PDF provided by BCBSTX.",
)

if uploaded_file is None:
    st.info("👆 Upload a renewal PDF above to get started. Your file is never saved to our servers — it's processed and immediately discarded.")
    st.stop()

st.write(f"📄 **{uploaded_file.name}**  —  {uploaded_file.size / 1_024:.0f} KB")
st.markdown("---")

# ── Extract buttons ────────────────────────────────────────────────────────────
if "play_music" not in st.session_state:
    st.session_state.play_music = False

btn_col1, btn_col2, _ = st.columns([2, 2, 5])
with btn_col1:
    extract_btn = st.button(
        "⚡ Extract Renewal",
        type="primary",
        use_container_width=True,
    )
with btn_col2:
    music_btn = st.button(
        "🎵 Extract with Music",
        type="secondary",
        use_container_width=True,
    )

do_extraction = extract_btn or music_btn
if music_btn:
    st.session_state.play_music = True
elif extract_btn:
    st.session_state.play_music = False

if not do_extraction:
    st.stop()

# ── Audio — starts immediately, slot is cleared after extraction to stop it ────
audio_slot = st.empty()
if st.session_state.play_music:
    mp3_path = Path(__file__).parent / "jeopardy-themelq.mp3"
    if mp3_path.exists():
        with audio_slot:
            st.audio(mp3_path.read_bytes(), format="audio/mpeg", autoplay=True)

# ── Extraction ─────────────────────────────────────────────────────────────────
with st.spinner("Extracting renewal data… this usually takes about 30 seconds. Hang tight."):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        pdf_path = tmpdir_path / uploaded_file.name
        pdf_path.write_bytes(uploaded_file.getbuffer())

        try:
            result      = run_extraction(pdf_path, output_dir=tmpdir_path)
            data        = result["data"]
            excel_bytes = Path(result["excel_path"]).read_bytes()
            json_bytes  = Path(result["json_path"]).read_bytes()
        except Exception as exc:
            audio_slot.empty()
            st.error("The extractor could not process this PDF. See the error details below.")
            st.exception(exc)
            st.stop()

# Stop the music now that extraction is complete
audio_slot.empty()

st.success("✅ Extraction complete!")

group       = data.get("group_info",  {}) or {}
plans       = data.get("plans",       []) or []
census_rows = data.get("census_rows", []) or []

# ── Summary metrics ────────────────────────────────────────────────────────────
st.markdown("### Group Summary")
group_display  = group.get("group_name") or group.get("mailing_name") or "Not found"
rating_area    = group.get("rating_area")
area_str       = f"  ·  Rating Area {rating_area}" if rating_area else ""
st.markdown(f"<div style='font-size:1.15rem; font-weight:600; margin-bottom:0.75rem;'>{group_display}{area_str}</div>", unsafe_allow_html=True)
mc1, mc2, mc3 = st.columns(3)
mc1.metric("Rating Type", str(data.get("rating_type", "unknown")).upper())
mc2.metric("Plans Found", len(plans))
mc3.metric("Census Rows", len(census_rows))

st.markdown("---")

# ── Downloads ──────────────────────────────────────────────────────────────────
st.markdown("### Downloads")
dl1, dl2, _ = st.columns([2, 2, 5])

with dl1:
    st.download_button(
        "📊 Download Excel Review",
        data=excel_bytes,
        file_name=f"{Path(uploaded_file.name).stem}_broker_review.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# JSON download — code preserved; surfaced once SHOW_JSON_DOWNLOAD = True
if SHOW_JSON_DOWNLOAD:
    with dl2:
        st.download_button(
            "📦 Download JSON Data",
            data=json_bytes,
            file_name=f"{Path(uploaded_file.name).stem}_extract.json",
            mime="application/json",
            use_container_width=True,
        )

st.markdown("---")

# ── Validation notes ───────────────────────────────────────────────────────────
st.markdown("### Validation Notes")
validation_df = build_validation_notes(data)

def _style_validation(df: pd.DataFrame):
    def row_color(row):
        color = "#fff3cd" if row["Severity"] == "Warning" else "#d1fae5"
        return [f"background-color: {color}"] * len(row)
    return df.style.apply(row_color, axis=1)

st.dataframe(_style_validation(validation_df), use_container_width=True, hide_index=True)

# ── Plan summary ───────────────────────────────────────────────────────────────
st.markdown("### Plan Summary")
plan_df = build_plan_summary_df(data)
if plan_df.empty:
    st.warning("No plans were extracted from this PDF.")
else:
    # Format dollar amounts and percentages for display
    DOLLAR_COLS = {"Current Total", "Renewal Total", "Current EO", "Renewal EO"}
    PCT_COLS    = {"Total % Change", "EO % Change"}
    fmt = {}
    for col in plan_df.columns:
        if col in PCT_COLS:
            fmt[col] = "{:.2%}"       # 0.2496 → "24.96%"
        elif col in DOLLAR_COLS:
            fmt[col] = "${:,.2f}"     # 1086.49 → "$1,086.49"
    styled_plan = plan_df.style.format(fmt, na_rep="—")
    st.dataframe(styled_plan, use_container_width=True, hide_index=True)

# ── Census preview ─────────────────────────────────────────────────────────────
with st.expander("📋 Census Preview"):
    census_df = pd.DataFrame(census_rows)
    if census_df.empty:
        st.write("No census rows found.")
    else:
        st.dataframe(census_df, use_container_width=True, hide_index=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(
    '<p class="footer-cap">BCBSTX Renewal Extractor - No AI used - '
    "Your file is never saved to our servers - it's processed and immediately discarded</p>",
    unsafe_allow_html=True,
)
