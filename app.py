from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

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
    /* Hide default Streamlit header padding */
    .block-container { padding-top: 1.5rem; }

    /* App header banner */
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

    /* Metric cards */
    [data-testid="metric-container"] {
        background: #f0f4ff;
        border: 1px solid #c7d7f5;
        border-radius: 0.75rem;
        padding: 0.85rem 1rem;
    }

    /* Download button: full width */
    .stDownloadButton > button {
        width: 100%;
        border-radius: 0.5rem;
        font-weight: 600;
    }

    /* Dataframe container */
    [data-testid="stDataFrame"] {
        border-radius: 0.5rem;
        overflow: hidden;
    }

    /* Footer caption */
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
    label_visibility="visible",
)

if uploaded_file is None:
    st.info("👆 Upload a renewal PDF above to get started. Files are processed in memory and immediately discarded.")
    st.stop()

st.write(f"📄 **{uploaded_file.name}**  —  {uploaded_file.size / 1_024:.0f} KB")
st.markdown("---")


# ── Extract buttons ────────────────────────────────────────────────────────────
if "play_music" not in st.session_state:
    st.session_state.play_music = False

btn_col1, btn_col2, _ = st.columns([2, 2.4, 5])
with btn_col1:
    extract_btn = st.button(
        "⚡ Extract Renewal",
        type="primary",
        use_container_width=True,
    )
with btn_col2:
    music_btn = st.button(
        "🎵 Extract with Jeopardy Music",
        type="secondary",
        use_container_width=True,
        help="Plays the Jeopardy Think! theme while the extractor runs. Hold tight.",
    )

do_extraction = extract_btn or music_btn
if music_btn:
    st.session_state.play_music = True
elif extract_btn:
    st.session_state.play_music = False

if not do_extraction:
    st.stop()


# ── Jeopardy Think! music (Web Audio API — no external files or copyright issues) ──
# Notes are synthesized in the browser via the Web Audio API.
# The iframe is removed by Streamlit on the next page rerun, which stops playback.
if st.session_state.play_music:
    components.html("""
    <script>
    (function () {
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) return;
        const ctx = new AC();

        // BPM = 118; beat = one quarter note in seconds
        const beat = 60 / 118;

        // Jeopardy "Think!" theme — approximate transcription
        // Format: [frequency_Hz, duration_in_quarter_note_beats]
        const phrase = [
            // Bar 1
            [392, 1], [523, 1], [392, 1], [262, 1],
            // Bar 2
            [392, 1], [523, 1], [659, 1], [784, 1],
            // Bar 3 (descending run)
            [698, 0.5], [659, 0.5], [587, 0.5], [523, 0.5],
            [494, 0.5], [523, 0.5], [587, 0.5], [659, 0.5],
            // Bar 4 (repeat of bar 1)
            [392, 1], [523, 1], [392, 1], [262, 1],
            // Bar 5
            [392, 1], [523, 1], [659, 2],
            // Bar 6
            [587, 1], [523, 1], [494, 1], [440, 1],
            // Bar 7 – resolve
            [392, 2], [0, 1],
            // Second half mirrors first with slight variation
            [392, 1], [523, 1], [392, 1], [262, 1],
            [392, 1], [523, 1], [659, 1], [784, 1],
            [784, 0.5], [698, 0.5], [659, 0.5], [587, 0.5],
            [523, 0.5], [494, 0.5], [440, 0.5], [392, 0.5],
            [392, 1], [523, 1], [392, 1], [262, 1],
            [392, 1], [523, 1], [698, 2],
            [659, 1], [587, 1], [523, 1], [494, 1],
            [392, 2], [0, 1],
        ];

        // Phase total duration in seconds
        const totalBeats = phrase.reduce((s, [, d]) => s + d, 0);

        function schedulePhrase(startT) {
            let t = startT;
            phrase.forEach(([freq, dur]) => {
                if (freq > 0) {
                    const osc  = ctx.createOscillator();
                    const gain = ctx.createGain();
                    osc.connect(gain);
                    gain.connect(ctx.destination);
                    osc.type = 'sine';
                    osc.frequency.setValueAtTime(freq, t);
                    gain.gain.setValueAtTime(0.22, t);
                    gain.gain.exponentialRampToValueAtTime(0.001, t + dur * beat * 0.88);
                    osc.start(t);
                    osc.stop(t + dur * beat);
                }
                t += dur * beat;
            });
            // Schedule next loop 200 ms before this one ends so there is no gap
            const msUntilEnd = (startT + totalBeats * beat - ctx.currentTime) * 1000 - 200;
            setTimeout(() => schedulePhrase(startT + totalBeats * beat), msUntilEnd);
        }

        schedulePhrase(ctx.currentTime + 0.05);
    })();
    </script>
    """, height=0)


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
            st.error("The extractor could not process this PDF. See the error details below.")
            st.exception(exc)
            st.stop()

st.success("✅ Extraction complete!")

group       = data.get("group_info",  {}) or {}
plans       = data.get("plans",       []) or []
census_rows = data.get("census_rows", []) or []


# ── Summary metrics ────────────────────────────────────────────────────────────
st.markdown("### Group Summary")
mc1, mc2, mc3, mc4 = st.columns(4)
mc1.metric("Group",        group.get("group_name") or group.get("mailing_name") or "Not found")
mc2.metric("Rating Type",  str(data.get("rating_type", "unknown")).upper())
mc3.metric("Plans Found",  len(plans))
mc4.metric("Census Rows",  len(census_rows))

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

# ── JSON download — code preserved; surfaced once SHOW_JSON_DOWNLOAD = True ──
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

def _style_validation(df: pd.DataFrame) -> pd.io.formats.style.Styler:
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
    st.dataframe(plan_df, use_container_width=True, hide_index=True)


# ── Census preview ─────────────────────────────────────────────────────────────
with st.expander("📋 Census Preview"):
    census_df = pd.DataFrame(census_rows)
    if census_df.empty:
        st.write("No census rows found.")
    else:
        st.dataframe(census_df, use_container_width=True, hide_index=True)


# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown(
    '<p class="footer-cap">BCBSTX Renewal Extractor · No AI used · '
    'Files are processed in memory and never stored</p>',
    unsafe_allow_html=True,
)
