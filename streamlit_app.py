
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time

st.set_page_config(page_title="Team Free Slot Finder (Outlook CSV)", layout="wide")
st.title("🗓️ Team Free Slot Finder — Outlook CSV")

with st.sidebar:
    st.header("⚙️ Settings")
    work_start = st.time_input("Workday start", value=pd.to_datetime("09:00").time())
    work_end = st.time_input("Workday end", value=pd.to_datetime("18:00").time())
    slot_minutes = st.selectbox("Slot size (minutes)", [15, 30, 45, 60], index=1)
    exclude_all_day = st.checkbox("Exclude all‑day events", value=True)
    exclude_holidays = st.checkbox("Exclude holidays", value=True)
    exclude_birthdays = st.checkbox("Exclude birthdays", value=True)
    horizon_days = st.slider("Look‑ahead days", min_value=1, max_value=30, value=14)

uploads = st.file_uploader(
    "Upload one or more Outlook calendar CSVs (one per person)",
    type=["csv"], accept_multiple_files=True
)

st.caption("Tip: Export **Calendar** → CSV in Outlook (classic). Avoid Holidays/Birthdays calendars.")

def read_csv_safely(file):
    for enc in ("utf-8","utf-8-sig","cp1252","latin-1"):
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=enc, engine="python")
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file, dtype=str, engine="python")

def parse_start_end(df: pd.DataFrame):
    # Always return Series aligned to df.index
    if "Start" in df.columns:
        start = pd.to_datetime(df["Start"].astype(str), errors="coerce")
    else:
        start = pd.Series(pd.NaT, index=df.index)
    if "End" in df.columns:
        end = pd.to_datetime(df["End"].astype(str), errors="coerce")
    else:
        end = pd.Series(pd.NaT, index=df.index)

    # fallback to split columns
    if start.isna().all() and {"Start Date","Start Time"}.issubset(df.columns):
        start = pd.to_datetime((df["Start Date"].astype(str) + " " + df["Start Time"].astype(str)), errors="coerce")
    if end.isna().all() and {"End Date","End Time"}.issubset(df.columns):
        end = pd.to_datetime((df["End Date"].astype(str) + " " + df["End Time"].astype(str)), errors="coerce")

    # other possible variants
    if start.isna().all():
        for c in ["Start DateTime","Begin","Start time","Starts","StartTime"]:
            if c in df.columns:
                start = pd.to_datetime(df[c].astype(str), errors="coerce")
                break
    if end.isna().all():
        for c in ["End DateTime","Finish","End time","Ends","EndTime"]:
            if c in df.columns:
                end = pd.to_datetime(df[c].astype(str), errors="coerce")
                break

    if not isinstance(start, pd.Series):
        start = pd.Series(start, index=df.index)
    if not isinstance(end, pd.Series):
        end = pd.Series(end, index=df.index)
    return start, end

def clean_events(df, exclude_all_day=True, exclude_holidays=True, exclude_birthdays=True):
    # Standardize column names we care about
    subj_col = next((c for c in df.columns if c.strip().lower()=="subject"), None)
    if exclude_all_day and "All Day Event" in df.columns:
        mask_all_day = df["All Day Event"].astype(str).str.lower().isin(["true","1","yes"])
        df = df[~mask_all_day | df["All Day Event"].isna()]
    if subj_col:
        if exclude_holidays:
            df = df[~df[subj_col].astype(str).str.contains("holiday", case=False, na=False)]
        if exclude_birthdays:
            df = df[~df[subj_col].astype(str).str.contains("birthday", case=False, na=False)]
    start, end = parse_start_end(df)
    df = df.assign(StartDT=start, EndDT=end)
    df = df[df["StartDT"].notna() & df["EndDT"].notna()].copy()
    df = df.sort_values("StartDT").reset_index(drop=True)
    return df

def build_busy_mask(events: pd.DataFrame, day_start: datetime, day_end: datetime, step: timedelta):
    # Return boolean mask of length N slots; True = busy
    slots = []
    t = day_start
    while t < day_end:
        slots.append(t)
        t += step
    if not len(slots):
        return np.array([], dtype=bool), []
    start_arr = events["StartDT"].values
    end_arr = events["EndDT"].values
    busy = np.zeros(len(slots), dtype=bool)
    for i, s in enumerate(slots):
        e = s + step
        # overlapping if event start < slot_end and event end > slot_start
        overlap = ((start_arr < e) & (end_arr > s)).any()
        busy[i] = bool(overlap)
    return busy, slots

def find_common_free(calendars, work_start: time, work_end: time, slot_minutes: int, horizon_days: int):
    today = datetime.now().date()
    step = timedelta(minutes=slot_minutes)
    suggestions = []
    for d in range(horizon_days):
        day = today + timedelta(days=d)
        day_start = datetime.combine(day, work_start)
        day_end = datetime.combine(day, work_end)
        # skip weekends? (Optional) — comment out next two lines to include weekends
        # if day.weekday() >= 5:
        #     continue
        # Build each person's busy mask
        masks = []
        slots_ref = None
        for ev in calendars:
            busy, slots = build_busy_mask(ev, day_start, day_end, step)
            if slots_ref is None:
                slots_ref = slots
            # Align lengths
            if len(busy) != len(slots_ref):
                # if mismatch (shouldn't happen), pad/trim
                L = len(slots_ref)
                b = np.zeros(L, dtype=bool)
                b[:min(L, len(busy))] = busy[:min(L, len(busy))]
                busy = b
            masks.append(busy)
        if not masks or slots_ref is None:
            continue
        combined_busy = np.zeros_like(masks[0])
        for m in masks:
            combined_busy |= m
        free_mask = ~combined_busy
        # group consecutive free slots
        i = 0
        while i < len(free_mask):
            if free_mask[i]:
                j = i
                while j < len(free_mask) and free_mask[j]:
                    j += 1
                start_ts = slots_ref[i]
                end_ts = slots_ref[min(j, len(slots_ref)-1)] + step
                suggestions.append({"Date": day.isoformat(), "Start": start_ts, "End": end_ts, "DurationMin": int((end_ts - start_ts).total_seconds()/60)})
                i = j
            else:
                i += 1
    sug = pd.DataFrame(suggestions)
    if not sug.empty:
        sug = sug.sort_values(["Date","Start"]).reset_index(drop=True)
    return sug

if not uploads:
    st.info("Upload **multiple** CSVs (one per teammate) to find common free time.")
    st.stop()

# Load and clean per person
calendars = []
names = []
for up in uploads:
    df = read_csv_safely(up)
    ev = clean_events(df, exclude_all_day=exclude_all_day, exclude_holidays=exclude_holidays, exclude_birthdays=exclude_birthdays)
    calendars.append(ev)
    names.append(up.name.split('.')[0])

# Suggest free slots
suggestions = find_common_free(calendars, work_start, work_end, slot_minutes, horizon_days)

if suggestions.empty:
    st.warning("No common free slots found within the selected horizon and working hours.")
else:
    st.success(f"Found {len(suggestions)} free blocks across the next {horizon_days} days.")
    # Show top suggestions (longest first), then by soonest
    top = suggestions.sort_values(["DurationMin","Date","Start"], ascending=[False, True, True]).head(50)
    st.subheader("⭐ Top suggestions (longest first)")
    st.dataframe(top)

    st.subheader("📋 All suggestions (chronological) — first 500 rows")
    st.dataframe(suggestions.head(500))

    st.download_button("Download all suggestions (CSV)", suggestions.to_csv(index=False).encode("utf-8"), "common_free_slots.csv", "text/csv")
