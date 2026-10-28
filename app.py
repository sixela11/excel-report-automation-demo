"""Streamlit demo: Excel consolidation and validation on synthetic sample data.

Run locally:  streamlit run app.py
"""
import streamlit as st

from demo_service import describe_inputs, run_sample

REPO_URL = "https://github.com/sixela11/excel-report-automation-demo"
PROFILE_URL = "https://github.com/sixela11"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Excel Report Automation Demo", page_icon="📊", layout="centered")

st.title("Monthly Excel reports, automated")
st.write("See how a folder of messy monthly spreadsheets becomes one clean, checked report in a single click.")
st.info(
    "**Sample demonstration only.** All data is synthetic (every customer name starts with "
    "\"Sample\"). This is not a completed client project, and no real business data is used.",
    icon="ℹ️",
)

st.header("The problem")
st.write(
    "Many teams rebuild the same report every month by hand. They open several spreadsheets, "
    "copy the rows into a master file, remove duplicates, and fix IDs that lost their leading zeros. "
    "It's slow and easy to get wrong, and it's hard to tell which rows were changed or left out."
)

st.header("The solution")
st.write(
    "A small Python automation reads every file, checks each row against clear rules, and builds "
    "one formatted Excel report. Clean rows go into the report. Anything wrong is listed separately "
    "with the exact file and row, so a person can review it. The original files are never changed."
)

before, after = st.columns(2)
with before:
    st.subheader("Before")
    st.markdown(
        "- Open each monthly file\n"
        "- Copy and paste into a master sheet\n"
        "- Hunt for duplicates by eye\n"
        "- Re-type IDs like `000104` that became `104`\n"
        "- Hope nothing was missed"
    )
with after:
    st.subheader("After")
    st.markdown(
        "- One click processes every file\n"
        "- Every row checked automatically\n"
        "- Duplicates found and listed\n"
        "- IDs and postcodes kept intact\n"
        "- Every problem reported with file and row"
    )

st.header("1. The sample files")
st.write("Seven months of synthetic sales exports, with realistic problems added on purpose:")
st.dataframe(describe_inputs(), hide_index=True, width="stretch")

st.header("2. Run it")
if st.button("▶ Run Automation", type="primary", width="stretch"):
    with st.spinner("Processing files..."):
        result, report = run_sample()
    # Held in this visitor's session memory only; nothing is written to shared storage.
    st.session_state["demo"] = {"result": result, "report": report}

demo = st.session_state.get("demo")
if demo:
    result, stats = demo["result"], demo["result"].stats
    st.success("Done. These figures are calculated live from the sample files.")

    c1, c2 = st.columns(2)
    c1.metric("Files processed", f"{stats['Files processed']} of {stats['Files found']}")
    c2.metric("Rows in final report", stats["Rows in final report"])
    c3, c4 = st.columns(2)
    c3.metric("Duplicates removed", stats["Duplicate rows removed"])
    c4.metric("Rows excluded (errors)", stats["Rows with errors (excluded)"])
    c5, c6 = st.columns(2)
    c5.metric("Warnings for review", stats["Warnings"])
    c6.metric("Files rejected", stats["Files rejected"])
    if result.rejected_files:
        st.warning(f"Rejected (missing required columns): {', '.join(result.rejected_files)}")

    st.download_button("⬇ Download the Excel report", data=demo["report"],
                       file_name="consolidated_report.xlsx", mime=XLSX_MIME, width="stretch")

    st.header("3. What's in the report")
    summary, clean, dups, exc = st.tabs(["Run statistics", "Clean data", "Duplicates", "Exceptions"])
    with summary:
        st.table({"Metric": list(stats), "Value": [f"{v:,}" if isinstance(v, (int, float)) else v
                                                   for v in stats.values()]})
        st.caption("Processing time is measured on this server and will vary.")
    with clean:
        st.caption("Each row keeps its source file and Excel row number, so it can be traced back.")
        st.dataframe(result.consolidated, hide_index=True, width="stretch")
    with dups:
        st.caption("For a conflicting duplicate (same invoice, different values), the first file's "
                   "row is kept and the other is flagged for a person to decide.")
        st.dataframe(result.duplicates, hide_index=True, width="stretch")
    with exc:
        st.caption("Errors exclude a row. Warnings keep it but flag it for review.")
        st.dataframe(result.exceptions, hide_index=True, width="stretch")

st.header("Limitations of this demo")
st.markdown(
    "- Runs only on the bundled synthetic files. You can't upload files here, by design.\n"
    "- The expected columns are fixed. A real project would be set up for your own layout.\n"
    "- Reads `.xlsx` files and only the first sheet of each.\n"
    "- Conflicting duplicates keep the row from the earlier file (by filename) until someone reviews them.\n"
    "- Restoring leading zeros assumes fixed-width IDs. Each repair is logged.\n"
    "- Accepts Excel dates, `yyyy-mm-dd` and Australian `dd/mm/yyyy`, but not US `mm/dd/yyyy`.\n"
    "- No time savings are claimed. Real benefit depends on your files and process."
)

st.header("Have a repetitive spreadsheet task?")
st.write(
    "I build fixed-price Python and Excel automations for Australian businesses. "
    "Describe the task in general terms first. Please don't send confidential files."
)
st.link_button("Contact me via GitHub", PROFILE_URL, width="stretch")
st.markdown(f"Source code and documentation: [{REPO_URL}]({REPO_URL})")
