import os
from dotenv import load_dotenv
load_dotenv()

api_key = os.environ.get("ANTHROPIC_API_KEY")
if api_key:
    print(f"✅ ANTHROPIC_API_KEY loaded — ends with ...{api_key[-6:]}")
else:
    print("⚠️ ANTHROPIC_API_KEY not found — running rules-only mode")
import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import time
import shutil

from config.settings import (
    LOB_OPTIONS, ALLOWED_USERS,
    INPUT_COL_SR, INPUT_COL_DATE, INPUT_COL_BEST_PART, INPUT_COL_RATING,
    INPUT_COL_IMPROVEMENT, INPUT_COL_PL, INPUT_COL_COURSE, INPUT_COL_TOPIC, INPUT_COL_LOB,
    EXCEL_OUTPUT_DIR, PDF_OUTPUT_DIR,
)
from utils.report_generator import generate_report, clean_text, fmt_date
from utils.zip_handler import create_zip_from_folder, create_lob_zip
from utils.airtable import upload_to_airtable
try:
    from utils.pdf_generator import generate_pdf
    print('✅ pdf_generator imported successfully')
except Exception as e:
    import traceback
    print(f'❌ pdf_generator import failed: {e}')
    print(traceback.format_exc())
    generate_pdf = None
from utils.sentiment import analyse_from_excel_rows

os.makedirs(EXCEL_OUTPUT_DIR, exist_ok=True)
os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)

# ✅ Airtable token
try:
    AIRTABLE_TOKEN = st.secrets["12345"]
except Exception:
    AIRTABLE_TOKEN = None

if "generated_files" not in st.session_state:
    st.session_state["generated_files"] = []

upload_logs = []
log_entries = []

# ---------- Simple Email Authentication ----------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""


def auth_screen():
    st.markdown(
        "<h2 style='text-align:center; margin-bottom:0;'>Feedback Automation Tool</h2>",
        unsafe_allow_html=True,
    )
    st.markdown("<p style='text-align:center;'>🔒 Please sign in with your official email ID</p>", unsafe_allow_html=True)

    col_left, col_center, col_right = st.columns([1, 2, 1])
    with col_center:
        with st.form("email_login_form", clear_on_submit=False):
            email = st.text_input("Official email ID", placeholder="your.name@company.com")
            submitted = st.form_submit_button("Enter Dashboard")

            if submitted:
                email_clean = email.strip().lower()
                if email_clean in ALLOWED_USERS:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = ALLOWED_USERS[email_clean]
                    st.success(f"Welcome, {st.session_state['username']} 👋")
                    st.rerun()
                else:
                    st.error("Access denied. This email is not registered for this tool.")


if not st.session_state.get("authenticated", False):
    auth_screen()
    st.stop()

# ------------------ Welcome Label ------------------
if st.session_state.get("username"):
    st.markdown(
        f"""
        <div style="
            width:100%;
            text-align:left;
            font-size:24px;
            font-weight:600;
            color:#00B050;
            margin-top:0px;
            margin-bottom:75px;
        ">
            👋 Welcome, <span style="color:white;">{st.session_state['username']}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

LOGO_PATH = Path(__file__).parent / "Emeritus.jpg"

# 🌗 Adaptive Styling
st.markdown("""
<style>
/* 🌞 Light Mode */
@media (prefers-color-scheme: light) {
    [data-testid="stAppViewContainer"] {
        background-color: #FFFFFF !important;
        color: black !important;
    }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4 {
        color: black !important;
    }
    .stDataFrame, .stTable {
        background-color: #FFFFFF !important;
        color: black !important;
        border: 1px solid #ddd !important;
    }
    div[data-testid="stProgress"] div[role="progressbar"] {
        background-color: #00B050 !important;
    }
    .stAlert {
        background-color: #F0F2F6 !important;
        color: black !important;
        border: 1px solid #ddd !important;
    }
}

/* 🌙 Dark Mode */
@media (prefers-color-scheme: dark) {
    [data-testid="stAppViewContainer"] {
        background-color: #0E1117 !important;
        color: white !important;
    }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4 {
        color: white !important;
    }
    .stDataFrame, .stTable {
        background-color: #1E1E1E !important;
        color: white !important;
        border: 1px solid #333 !important;
    }
    div[data-testid="stProgress"] div[role="progressbar"] {
        background-color: #00B050 !important;
    }
    .stAlert {
        background-color: #1E1E1E !important;
        color: white !important;
        border: 1px solid #333 !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ------------------ Header ------------------
st.markdown(
    """
    <div style='width:100%; text-align:center; padding-top:10px; padding-bottom:6px;'>
        <h1 style='margin:0; font-size:34px; color:white;'>Feedback Automation Tool</h1>
    </div>
    <hr style='margin-top:8px; border-color:#555;'>
    """,
    unsafe_allow_html=True
)

# ------------------ Sidebar ------------------
with st.sidebar:
    st.image(str(LOGO_PATH), width=220)
    st.markdown("<h2 style='color:white;'>⚙️ Work-Flow</h2>", unsafe_allow_html=True)

    # ---------------- TEMPLATE DOWNLOAD ----------------
    template_path = Path(__file__).parent / "Automation Template.xlsx"
    st.markdown("### 📥 Download Automation Template")
    if template_path.exists():
        with open(template_path, "rb") as tfile:
            st.download_button(
                label="⬇️ Download Excel Template",
                data=tfile,
                file_name="Feedback_Automation Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="template_download_btn"
            )
    else:
        st.warning("Template file not found — please add 'Automation Template.xlsx' to the app folder.")

    # ---------------- FILE UPLOAD ----------------
    uploaded = st.file_uploader(
        "📁 Upload Automation Excel Sheet",
        type=["xlsx"]
    )

    # ---------------- LOB SELECTION ----------------
    lob_choice = st.selectbox(
        "🏷️ Select LOB",
        LOB_OPTIONS,
        index=0
    )

    # ---------------- OUTPUT PATH ----------------
    downloads_path = Path.home() / "Downloads"
    out_folder_input = st.text_input(
        "💾 Output folder path",
        str(downloads_path)
    )

    # ---------------- GENERATE BUTTON ----------------
    generate_btn = st.button("🚀 Generate Reports")

    # ✅ Upload Button
    upload_btn = st.button(
        "📤 Upload to Airtable",
        disabled=not st.session_state.get("generated_files")
    )

# ------------------ Summary Dashboard ------------------
st.markdown("<h3 style='color:white; text-align:center;'>📊 Summary Dashboard</h3>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
with col1:
    total_reports_ph = st.empty()
    total_reports_ph.metric("Total Reports Generated", "0", delta="↗ Ready")
with col2:
    avg_rating_ph = st.empty()
    avg_rating_ph.metric("Average Rating", "0.00 ⭐", delta="No Data Yet")
with col3:
    total_lob_ph = st.empty()
    total_lob_ph.metric("LOBs Processed", "0", delta="Stable")

st.markdown("""
<div style='text-align:center; color:#00B050; font-style:italic;'>
Dashboard auto-refreshes after each report batch generation.
</div>
<hr style='margin-top:10px; border-color:#555;'>
""", unsafe_allow_html=True)

# ------------------ Main Generation Logic ------------------
if uploaded and generate_btn:
    base_outroot = Path(out_folder_input).expanduser().resolve()
    if base_outroot.exists():
        for item in base_outroot.iterdir():
            if item.is_dir() and item.name.startswith("Feedback Reports for"):
                shutil.rmtree(item)
    else:
        base_outroot.mkdir(parents=True, exist_ok=True)

    with st.spinner("🧹 Resetting workspace..."):
        time.sleep(1)
        st.success("✅ Workspace cleaned and ready!")

    if lob_choice == "Select":
        st.warning("⚠️ Please select a valid LOB before generating reports.")
        st.stop()

    tmp_path = Path("temp_automation.xlsx")
    with open(tmp_path, "wb") as f:
        f.write(uploaded.getbuffer())

    raw = pd.read_excel(tmp_path, sheet_name="Automation", engine="openpyxl", dtype=object)
    raw.columns = [str(c).strip() for c in raw.columns]

    rows = []
    for _, r in raw.iterrows():
        rows.append({
            "Sr": clean_text(r.get(INPUT_COL_SR, "")),
            "Date": fmt_date(r.get(INPUT_COL_DATE, "")),
            "BestPart": clean_text(r.get(INPUT_COL_BEST_PART, "")),
            "Rating": clean_text(r.get(INPUT_COL_RATING, "")),
            "Improvement": clean_text(r.get(INPUT_COL_IMPROVEMENT, "")),
            "PL": clean_text(r.get(INPUT_COL_PL, "")),
            "Course": clean_text(r.get(INPUT_COL_COURSE, "")),
            "Topic": clean_text(r.get(INPUT_COL_TOPIC, "")),
            "LOB": clean_text(r.get(INPUT_COL_LOB, "")).upper(),
        })
    df = pd.DataFrame(rows)

    if lob_choice == "All":
        lob_list = sorted([l for l in df["LOB"].unique() if l])
    else:
        lob_list = [lob_choice.upper()]

    base_outroot.mkdir(parents=True, exist_ok=True)

    log_entries = []
    all_ratings = []
    lob_results = {}
    pdf_count = 0

    groups = []
    for lob in lob_list:
        subset = df[df["LOB"] == lob]
        if not subset.empty:
            grouped = subset.groupby(["Course", "PL", "Date", "LOB"], dropna=False)
            for key, grp in grouped:
                groups.append((key, grp))

    total_groups = len(groups)
    if total_groups == 0:
        st.warning("No records found for the selected LOB.")
        st.stop()

    with st.status("⚙️ Generating reports...", expanded=True) as status:
        progress = st.progress(0)

        for i, ((course, pl, date_str, lob), grp) in enumerate(groups, start=1):
            st.write(f"⚙️ Processing **{lob}** — {course} ({date_str})...")
            progress.progress(int(i / total_groups * 100))

            session_folder_name = f"Feedback Reports for {date_str}" if date_str else "Feedback Reports"
            xlsx_out_folder = base_outroot / session_folder_name / "excel" / lob
            pdf_out_folder  = base_outroot / session_folder_name / "pdf"   / lob

            file_path, group_ratings = generate_report(grp, course, pl, date_str, xlsx_out_folder)
            all_ratings.extend(group_ratings)
            xlsx_stem = Path(file_path).stem

            # Build rows for PDF: (sr_no, best_part, rating_int, improvement)
            rows_for_pdf = []
            for idx, row in enumerate(grp.itertuples(), start=1):
                try:
                    rv = int(float(row.Rating)) if row.Rating not in ("", "0", "None") else 3
                except Exception:
                    rv = 3
                rows_for_pdf.append((idx, row.BestPart, rv, str(row.Improvement).strip()))

            st.write(f"📄 Generating PDF for **{lob}**...")
            pdf_path = None
            if generate_pdf is None:
                print('⚠️ PDF skipped — pdf_generator failed to import')
            else:
                try:
                    pdf_sentiments = analyse_from_excel_rows(rows_for_pdf)
                    topic_val = grp["Topic"].dropna()
                    session_title = topic_val.iloc[0] if len(topic_val) > 0 else course
                    avg_score = round(sum(group_ratings) / len(group_ratings), 2) if group_ratings else 0.0
                    pdf_path = generate_pdf({
                        "run_code": course,
                        "title": session_title,
                        "date": date_str,
                        "pl_name": pl,
                        "lob": lob,
                        "rows": rows_for_pdf,
                        "avg_score": avg_score,
                        "sentiments": pdf_sentiments,
                        "output_dir": str(pdf_out_folder),
                        "filename": xlsx_stem,
                    })
                    pdf_count += 1
                    print(f"📄 PDF generated at: {pdf_path}")
                    print(f"📄 PDF exists on disk: {os.path.exists(pdf_path)}")
                    print(f"📄 PDF size: {os.path.getsize(pdf_path) if os.path.exists(pdf_path) else 'FILE NOT FOUND'}")
                except Exception as e:
                    import traceback
                    print(f'❌ PDF generation failed: {e}')
                    print(traceback.format_exc())
                    st.write(f"⚠️ PDF skipped for {lob}: {e}")
                    pdf_path = None

            # Track per-LOB results
            if lob not in lob_results:
                lob_results[lob] = {"files": [], "ratings": []}
            lob_results[lob]["files"].append({
                "course": course, "pl": pl, "date": date_str,
                "xlsx_path": str(file_path),
                "pdf_path": pdf_path,
            })
            lob_results[lob]["ratings"].extend(group_ratings)

            st.write(f"✅ **{lob}** — {course} complete")

            st.session_state["generated_files"].append({
                "file_path": file_path,
                "course": course,
                "pl": pl,
                "date": date_str,
            })

            log_entries.append({
                "S.No": i,
                "Course": course,
                "PL": pl,
                "LOB": lob,
                "Session Date": date_str,
                "File": str(file_path),
            })

        status.update(label="✅ All reports ready!", state="complete")

    # Update Summary Dashboard metrics
    total_reports_ph.metric("Total Reports Generated", str(len(log_entries)), delta="Complete")
    avg_rating_ph.metric(
        "Average Rating",
        f"{sum(all_ratings)/len(all_ratings):.2f} ⭐" if all_ratings else "N/A",
        delta="Final"
    )
    total_lob_ph.metric("LOBs Processed", str(len(lob_list)), delta="Done")

    # Post-generation 4-column KPIs
    st.markdown("---")
    with st.container():
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📋 Reports Generated", len(log_entries))
        k2.metric("⭐ Average Rating", f"{sum(all_ratings)/len(all_ratings):.2f}" if all_ratings else "N/A")
        k3.metric("🏷 LOBs Processed", len(set(e["LOB"] for e in log_entries)))
        k4.metric("📄 PDFs Generated", pdf_count)

    # Per-LOB expanders
    for lob_name, lob_data in lob_results.items():
        with st.expander(f"📁 {lob_name} — {len(lob_data['files'])} report(s)"):
            for entry in lob_data["files"]:
                st.write(f"📁 Excel: `{os.path.basename(entry['xlsx_path'])}`")
                if entry.get("pdf_path"):
                    st.write(f"📄 PDF: `{os.path.basename(entry['pdf_path'])}`")
            if lob_data["ratings"]:
                st.caption(f"LOB Average Rating: {sum(lob_data['ratings'])/len(lob_data['ratings']):.2f} ⭐")

    # Show generation log
    df_log = pd.DataFrame(log_entries)
    df_log.index = range(1, len(df_log) + 1)
    df_log.index.name = "S.No"
    st.dataframe(df_log, height=300)

    # Download Log
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_fname = f"Feedback_Report_Log_{ts}.xlsx"
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_log.to_excel(writer, index=True)
    bio.seek(0)
    st.download_button("⬇️ Download Generation Log (Excel)", bio, log_fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.success(f"✅ Feedback Reports Generated Successfully for {lob_choice}")

# ✅ Airtable Upload Trigger
if upload_btn:
    if not st.session_state.get("generated_files"):
        st.warning("⚠️ Generate reports first")
    else:
        st.info("🚀 Uploading to Airtable...")
        for item in st.session_state["generated_files"]:
            upload_to_airtable(
                item["file_path"],
                item["course"],
                item["pl"],
                item["date"],
                AIRTABLE_TOKEN,
                upload_logs,
            )
        st.success("✅ Upload completed")

# ------------------ ZIP Download Section ------------------
with st.expander("📦 Download Feedback Reports"):
    if log_entries:
        try:
            session_folders = [
                f for f in base_outroot.iterdir()
                if f.is_dir() and f.name.startswith("Feedback Reports for")
            ]

            if not session_folders:
                st.warning("No generated session folders found.")
            else:
                for session_folder in session_folders:

                    # Case 1 — All LOBs Selected
                    if lob_choice.upper() == "ALL":
                        zip_name = f"Feedback_Reports_All_LOBs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                        zip_path = base_outroot / zip_name
                        zip_bytes = create_zip_from_folder(session_folder, zip_path)

                        st.download_button(
                            label="⬇️ Download All LOB Reports (ZIP)",
                            data=zip_bytes,
                            file_name=zip_name,
                            mime="application/zip",
                            key=f"all_lobs_zip_{datetime.now().strftime('%H%M%S%f')}"
                        )
                        st.caption(f"🕒 Generated at {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}")
                        st.success("✅ Master ZIP for All LOBs ready for download")
                        time.sleep(2)

                    # Case 2 — Specific LOB Selected
                    else:
                        selected_lob = lob_choice.upper()
                        excel_base = session_folder / "excel"
                        found = False
                        if excel_base.exists():
                            for lob_dir in excel_base.iterdir():
                                if lob_dir.is_dir() and lob_dir.name.upper() == selected_lob:
                                    found = True
                                    zip_name = f"Feedback_Reports_{lob_dir.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                                    zip_path = base_outroot / zip_name
                                    zip_bytes = create_lob_zip(session_folder, lob_dir.name, zip_path)

                                    st.download_button(
                                        label=f"⬇️ Download {lob_dir.name} Reports (ZIP)",
                                        data=zip_bytes,
                                        file_name=zip_name,
                                        mime="application/zip",
                                        key=f"{lob_dir.name.lower()}_{datetime.now().strftime('%H%M%S%f')}_zip"
                                    )
                                    st.caption(f"🕒 Generated at {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}")
                                    st.success(f"✅ ZIP ready for {lob_dir.name}")
                                    time.sleep(2)
                                    break

                        if not found:
                            st.warning(f"No reports found for selected LOB: {lob_choice}")

        except Exception as e:
            st.error(f"Error while creating ZIP files: {e}")
    else:
        st.info("ℹ️ Generate reports first to enable ZIP downloads.")

# ✅ Airtable Upload Log Download
if upload_logs:
    df_upload = pd.DataFrame(upload_logs)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_upload.to_excel(writer, index=False)
    bio.seek(0)
    st.download_button(
        "⬇️ Download Airtable Upload Log",
        bio,
        "Airtable_Upload_Log.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------ Footer ------------------
st.markdown("""
<hr>
<div style='text-align:center; background-color:#2B2B2B; color:white; padding:12px; border-radius:5px;'>
<b>Developed by EMERITUS - Tech Certs (INDIA APAC)</b><br>
Version 1.1.4 | © 2025 All Rights Reserved
</div>
""", unsafe_allow_html=True)
