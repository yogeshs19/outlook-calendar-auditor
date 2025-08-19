import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Outlook Calendar Auditor", layout="wide")
st.title("📅 Outlook Calendar Auditor — SAFE build")

# -------- Sidebar --------
st.sidebar.header("Settings")
timezone = st.sidebar.selectbox("Timezone (display only)", ["UTC","US/Eastern","Europe/London","Asia/Kolkata"], index=3)
work_start = st.sidebar.time_input("Workday start", value=pd.to_datetime("09:00").time())
work_end = st.sidebar.time_input("Workday end", value=pd.to_datetime("17:00").time())
st.sidebar.markdown("---")
exclude_all_day = st.sidebar.checkbox("Exclude all‑day events", value=True)
exclude_holidays = st.sidebar.checkbox("Exclude holidays", value=True)
exclude_birthdays = st.sidebar.checkbox("Exclude birthdays", value=True)

up = st.file_uploader("Upload Outlook Calendar CSV", type=["csv"])

def read_csv_safely(file):
    for enc in ("utf-8","utf-8-sig","cp1252","latin-1"):
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=enc, engine="python")
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file, dtype=str, engine="python")

def parse_start_end(df):
    # Try combined first
    start = pd.to_datetime(df.get("Start"), errors="coerce", utc=False)
    end = pd.to_datetime(df.get("End"), errors="coerce", utc=False)
    # If still NaT, try split
    if start.isna().all() and {"Start Date","Start Time"}.issubset(df.columns):
        start = pd.to_datetime(df["Start Date"].astype(str) + " " + df["Start Time"].astype(str), errors="coerce", utc=False)
    if end.isna().all() and {"End Date","End Time"}.issubset(df.columns):
        end = pd.to_datetime(df["End Date"].astype(str) + " " + df["End Time"].astype(str), errors="coerce", utc=False)
    return start, end

if not up:
    st.info("Export your *Calendar* (not Holidays/Birthdays) to CSV and upload it here.")
    st.stop()

# Read & normalize
df = read_csv_safely(up)

# Subject column
subject_col = None
for c in df.columns:
    if c.strip().lower() == "subject":
        subject_col = c
        break

# Optional filters before parsing
if exclude_all_day and "All Day Event" in df.columns:
    df = df[df["All Day Event"].astype(str).str.lower().isin(["false","0","no"]) | df["All Day Event"].isna()]
if exclude_holidays and subject_col:
    df = df[~df[subject_col].astype(str).str.contains("holiday", case=False, na=False)]
if exclude_birthdays and subject_col:
    df = df[~df[subject_col].astype(str).str.contains("birthday", case=False, na=False)]

# Parse datetimes (vectorized)
start_dt, end_dt = parse_start_end(df)
df["StartDT"] = start_dt
df["EndDT"] = end_dt

# Keep only valid rows (prevents NaT errors)
valid = df["StartDT"].notna() & df["EndDT"].notna()
events = df.loc[valid].copy()

if events.empty:
    st.warning("No events with valid Start/End after filtering. Check CSV columns and date range.")
    st.stop()

# Compute numeric hours for safe comparisons (no .time() on NaT)
events["StartHour"] = events["StartDT"].dt.hour + events["StartDT"].dt.minute/60.0
events["EndHour"] = events["EndDT"].dt.hour + events["EndDT"].dt.minute/60.0
ws = work_start.hour + work_start.minute/60.0
we = work_end.hour + work_end.minute/60.0

events = events.sort_values("StartDT").reset_index(drop=True)

# KPIs / flags
events["AfterHours"] = (events["StartHour"] < ws) | (events["EndHour"] > we)
prev_end = events["EndDT"].shift(1)
events["Conflict"] = events["StartDT"] < prev_end
gap = (events["StartDT"] - prev_end).dt.total_seconds() / 60.0
events["GapPrevMin"] = gap
events["BackToBack"] = gap.between(0, 5, inclusive="left")

# Link / agenda
desc_col = next((c for c in events.columns if c.strip().lower() in ("body","description","notes")), None)
link_col = next((c for c in events.columns if c.strip().lower() in ("online meeting join url","meeting link","teams link","join link","link")), None)

has_link = pd.Series(False, index=events.index)
if link_col:
    has_link = has_link | events[link_col].astype(str).str.contains(r"https?://", case=False, regex=True)
if desc_col:
    has_link = has_link | events[desc_col].astype(str).str.contains(r"https?://", case=False, regex=True)
events["HasJoinLink"] = has_link

has_agenda = pd.Series(False, index=events.index)
if desc_col:
    has_agenda = events[desc_col].astype(str).str.strip().str.len() >= 15
events["HasAgenda"] = has_agenda

subj = events[subject_col] if subject_col else pd.Series("", index=events.index)
events["ShortTitle"] = subj.astype(str).str.strip().str.len() < 5
events["ALLCAPS"] = subj.astype(str).apply(lambda x: x.isupper() and len(x.strip()) >= 5)

# KPI summary
kpi = {
    "Meetings": len(events),
    "Conflicts": int(events["Conflict"].sum()),
    "After-hours": int(events["AfterHours"].sum()),
    "Missing link": int((~events["HasJoinLink"]).sum()),
    "Missing agenda": int((~events["HasAgenda"]).sum()),
    "Back-to-back (<5m)": int(events["BackToBack"].sum()),
    "Short titles": int(events["ShortTitle"].sum()),
    "ALL CAPS titles": int(events["ALLCAPS"].sum()),
    "Median duration (min)": float((events["EndDT"] - events["StartDT"]).dt.total_seconds().median() / 60.0),
}

c1, c2, c3, c4 = st.columns(4)
c1.metric("Meetings", kpi["Meetings"])
c2.metric("Conflicts", kpi["Conflicts"])
c3.metric("After-hours", kpi["After-hours"])
c4.metric("Back-to-back (<5m)", kpi["Back-to-back (<5m)"])
c5, c6, c7, c8 = st.columns(4)
c5.metric("Missing link", kpi["Missing link"])
c6.metric("Missing agenda", kpi["Missing agenda"])
c7.metric("Short titles", kpi["Short titles"])
c8.metric("ALL CAPS titles", kpi["ALL CAPS titles"])

show_cols = [c for c in ["StartDT","EndDT","Subject","AfterHours","Conflict","BackToBack","GapPrevMin","HasJoinLink","HasAgenda"] if c in events.columns]

st.subheader("Parsed Events (first 1000)")
st.dataframe(events[show_cols].head(1000))

issues = events[(events["Conflict"]) | (events["AfterHours"]) | (~events["HasJoinLink"]) | (~events["HasAgenda"]) | (events["BackToBack"]) | (events["ShortTitle"]) | (events["ALLCAPS"])].copy()
st.subheader("Issues")
st.dataframe(issues[show_cols].head(1000))

st.download_button("Download Issues CSV", issues.to_csv(index=False).encode("utf-8"), "issues.csv", "text/csv")
