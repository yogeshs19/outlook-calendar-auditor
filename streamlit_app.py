import streamlit as st
import pandas as pd

st.set_page_config(page_title="Outlook Calendar Auditor", layout="wide")
st.title("📅 Outlook Calendar Auditor")

st.sidebar.header("Settings")
timezone = st.sidebar.selectbox("Select Timezone", ["UTC", "US/Eastern", "Europe/London", "Asia/Kolkata"])
work_start = st.sidebar.time_input("Workday start", value=pd.to_datetime("09:00").time())
work_end = st.sidebar.time_input("Workday end", value=pd.to_datetime("17:00").time())

uploaded_file = st.file_uploader("Upload Outlook Calendar CSV", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.subheader("Raw Calendar Data")
    st.dataframe(df.head())

    issues = []
    if "Start Date" in df.columns and "Start Time" in df.columns:
        df["Start"] = pd.to_datetime(df["Start Date"] + " " + df["Start Time"], errors="coerce")
        df["End"] = pd.to_datetime(df["End Date"] + " " + df["End Time"], errors="coerce")

        for i, row in df.iterrows():
            if row["Start"].time() < work_start or row["End"].time() > work_end:
                issues.append({"Subject": row.get("Subject", ""), "Issue": "Outside working hours"})
            if not pd.notnull(row.get("Location")):
                issues.append({"Subject": row.get("Subject", ""), "Issue": "Missing location"})
    
    if issues:
        st.subheader("🚨 Issues Found")
        issues_df = pd.DataFrame(issues)
        st.dataframe(issues_df)
        st.download_button("Download Issues CSV", issues_df.to_csv(index=False), "issues.csv")
    else:
        st.success("No issues found! 🎉")
