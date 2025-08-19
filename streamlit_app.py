
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import StringIO
from datetime import datetime, date, time, timedelta

st.set_page_config(page_title="Scheduling Assistant — CSV/ICS (Busy-aware)", layout="wide")
st.title("📆 Scheduling Assistant — CSV/ICS (Busy-aware)")

# ================= Sidebar Controls =================
with st.sidebar:
    st.header("⚙️ Settings")
    work_start = st.time_input("Workday start", value=pd.to_datetime("09:00").time())
    work_end = st.time_input("Workday end", value=pd.to_datetime("18:00").time())
    slot_minutes = st.selectbox("Slot size (minutes)", [15, 30, 60], index=1)
    show_date = st.date_input("Grid date", value=date.today())
    st.markdown("---")
    exclude_all_day = st.checkbox("Exclude all‑day events", value=True)
    exclude_holidays = st.checkbox("Exclude holidays", value=True)
    exclude_birthdays = st.checkbox("Exclude birthdays", value=True)
    st.markdown("---")
    st.caption("Which 'Show As' count as BUSY?")
    busy_count_free = st.checkbox("Treat 'Free' as busy", value=False)
    busy_count_tentative = st.checkbox("Treat 'Tentative' as busy", value=True)
    busy_count_working_elsewhere = st.checkbox("Treat 'Working Elsewhere' as busy", value=True)
    busy_count_ooo = st.checkbox("Treat 'Out of Office' as busy", value=True)
    horizon_days = st.slider("Find common free slots for next N days", 1, 45, 14)

uploads = st.file_uploader(
    "Upload **one CSV or ICS per person** (Outlook export or calendar ICS). Rename participants below if needed.",
    type=["csv","ics"],
    accept_multiple_files=True
)

st.caption("Outlook CSV: File → Open & Export → Import/Export → Export to a file → CSV → Calendar. Outlook on the web: export ICS.")

# ================= Utilities =================

def read_csv_safely(file):
    for enc in ("utf-8","utf-8-sig","cp1252","latin-1"):
        try:
            file.seek(0)
            return pd.read_csv(file, dtype=str, encoding=enc, engine="python")
        except Exception:
            continue
    file.seek(0)
    return pd.read_csv(file, dtype=str, engine="python")

def unfold_lines(text):
    out = []
    for line in text.splitlines():
        if not out:
            out.append(line.rstrip("\r\n"))
        else:
            if line.startswith((" ", "\t")):
                out[-1] += line[1:].rstrip("\r\n")
            else:
                out.append(line.rstrip("\r\n"))
    return out

def parse_dt_ics(val):
    v = val.strip()
    # date only
    if re.fullmatch(r"\d{8}$", v):
        try:
            dt = datetime.strptime(v, "%Y%m%d")
            return dt, True
        except Exception:
            return None, True
    # datetime
    try:
        if v.endswith("Z"):
            dt = datetime.strptime(v, "%Y%m%dT%H%M%SZ")
        else:
            dt = datetime.strptime(v, "%Y%m%dT%H%M%S")
        return dt, False
    except Exception:
        return None, False

def parse_prop(line):
    if ":" not in line:
        return None, None, None
    head, value = line.split(":", 1)
    parts = head.split(";")
    name = parts[0].upper()
    params = {}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            params[k.upper()] = v
        else:
            params[p.upper()] = True
    return name, params, value

def ics_to_df(text):
    lines = unfold_lines(text)
    events = []
    in_ev = False
    cur = {}
    for ln in lines:
        if ln.startswith("BEGIN:VEVENT"):
            in_ev = True
            cur = {}
            continue
        if ln.startswith("END:VEVENT"):
            if cur:
                events.append(cur)
            in_ev = False
            cur = {}
            continue
        if not in_ev:
            continue
        name, params, value = parse_prop(ln)
        if not name:
            continue
        if name == "DTSTART":
            dt, allday = parse_dt_ics(value)
            cur["StartDT"] = dt
            cur["_ALLDAY"] = cur.get("_ALLDAY", False) or allday
        elif name == "DTEND":
            dt, allday = parse_dt_ics(value)
            cur["EndDT"] = dt
            cur["_ALLDAY"] = cur.get("_ALLDAY", False) or allday
        elif name == "SUMMARY":
            cur["Subject"] = value
        elif name == "LOCATION":
            cur["Location"] = value
        elif name == "TRANSP":
            # TRANSPARENT => free, OPAQUE => busy
            cur["ShowAs"] = "Free" if value.strip().upper()=="TRANSPARENT" else "Busy"
    df = pd.DataFrame(events)
    if not df.empty:
        df = df[df["StartDT"].notna() & df["EndDT"].notna()].copy()
        df["StartDT"] = pd.to_datetime(df["StartDT"], errors="coerce")
        df["EndDT"] = pd.to_datetime(df["EndDT"], errors="coerce")
        df = df[df["StartDT"].notna() & df["EndDT"].notna()].copy()
    return df

def parse_start_end_csv(df: pd.DataFrame):
    # Prefer combined columns
    if "Start" in df.columns:
        start = pd.to_datetime(df["Start"].astype(str), errors="coerce")
    else:
        start = pd.Series(pd.NaT, index=df.index)
    if "End" in df.columns:
        end = pd.to_datetime(df["End"].astype(str), errors="coerce")
    else:
        end = pd.Series(pd.NaT, index=df.index)

    # fallback to split
    if start.isna().all() and {"Start Date","Start Time"}.issubset(df.columns):
        start = pd.to_datetime(df["Start Date"].astype(str) + " " + df["Start Time"].astype(str), errors="coerce", dayfirst=False)
    if end.isna().all() and {"End Date","End Time"}.issubset(df.columns):
        end = pd.to_datetime(df["End Date"].astype(str) + " " + df["End Time"].astype(str), errors="coerce", dayfirst=False)

    # common variants
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

def csv_show_as_column(df):
    # Try common Outlook export names
    for c in df.columns:
        cl = c.strip().lower()
        if cl in ("show time as","show as","busystatus","busy status","showtimeas"):
            return c
    return None

def clean_events_csv(df, exclude_all_day=True, exclude_holidays=True, exclude_birthdays=True):
    subj_col = next((c for c in df.columns if c.strip().lower()=="subject"), None)
    if exclude_all_day and "All Day Event" in df.columns:
        mask_all_day = df["All Day Event"].astype(str).str.lower().isin(["true","1","yes"])
        df = df[~mask_all_day | df["All Day Event"].isna()]
    if subj_col:
        if exclude_holidays:
            df = df[~df[subj_col].astype(str).str.contains("holiday", case=False, na=False)]
        if exclude_birthdays:
            df = df[~df[subj_col].astype(str).str.contains("birthday", case=False, na=False)]
    start, end = parse_start_end_csv(df)
    df = df.assign(StartDT=start, EndDT=end)
    df = df[df["StartDT"].notna() & df["EndDT"].notna()].copy()
    df["StartDT"] = pd.to_datetime(df["StartDT"], errors="coerce")
    df["EndDT"] = pd.to_datetime(df["EndDT"], errors="coerce")
    df = df[df["StartDT"].notna() & df["EndDT"].notna()].copy()
    # Standardize 'ShowAs'
    show_col = csv_show_as_column(df)
    if show_col:
        df["ShowAs"] = df[show_col].astype(str).str.title()
    return df.sort_values("StartDT").reset_index(drop=True)

def is_event_busy(show_as: str) -> bool:
    if not isinstance(show_as, str) or not show_as:
        return True  # assume busy if unknown
    val = show_as.strip().lower()
    if val == "free":
        return busy_count_free
    if val in ("tentative", "maybe"):
        return busy_count_tentative
    if val in ("working elsewhere", "workingelsewhere"):
        return busy_count_working_elsewhere
    if val in ("out of office", "ooo"):
        return busy_count_ooo
    # "Busy" or anything else -> busy
    return True

def build_slots(day: date, start: time, end: time, step_min: int):
    slots = []
    t = datetime.combine(day, start)
    end_dt = datetime.combine(day, end)
    step = timedelta(minutes=step_min)
    while t < end_dt:
        slots.append(t)
        t += step
    return slots, step

def person_busy_mask(events: pd.DataFrame, slots, step):
    if events.empty:
        return np.zeros(len(slots), dtype=bool)
    # Precompute which events count as BUSY
    if "ShowAs" in events.columns:
        busy_flags = events["ShowAs"].apply(is_event_busy).to_numpy(dtype=bool)
    else:
        busy_flags = np.ones(len(events), dtype=bool)
    starts = events["StartDT"].to_numpy(dtype="datetime64[ns]")
    ends = events["EndDT"].to_numpy(dtype="datetime64[ns]")
    busy = np.zeros(len(slots), dtype=bool)
    for i, s in enumerate(slots):
        e = s + step
        s64 = np.datetime64(s)
        e64 = np.datetime64(e)
        mask = (starts < e64) & (ends > s64) & busy_flags
        busy[i] = bool(mask.any())
    return busy

def common_free(calendars, start: time, end: time, slot_min: int, days: int):
    today = date.today()
    all_blocks = []
    for d in range(days):
        day = today + timedelta(days=d)
        slots, step = build_slots(day, start, end, slot_min)
        if not slots:
            continue
        combined_busy = np.zeros(len(slots), dtype=bool)
        for ev in calendars:
            combined_busy |= person_busy_mask(ev, slots, step)
        free_mask = ~combined_busy
        i = 0
        while i < len(free_mask):
            if free_mask[i]:
                j = i
                while j < len(free_mask) and free_mask[j]:
                    j += 1
                start_ts = slots[i]
                end_ts = slots[min(j, len(slots)-1)] + step
                all_blocks.append({
                    "Date": day.isoformat(),
                    "Start": start_ts,
                    "End": end_ts,
                    "DurationMin": int((end_ts - start_ts).total_seconds()/60)
                })
                i = j
            else:
                i += 1
    df = pd.DataFrame(all_blocks)
    if not df.empty:
        df = df.sort_values(["Date","Start"]).reset_index(drop=True)
    return df

# ================= Main Flow =================
if not uploads:
    st.info("Upload calendars to begin. One file per person (CSV or ICS).")
    st.stop()

# Infer participant names from filenames; allow editing
default_names = [u.name.rsplit(".",1)[0] for u in uploads]
with st.form("names_form"):
    st.subheader("Participants")
    cols = st.columns(min(3, len(uploads)) or 1)
    names = []
    for i, u in enumerate(uploads):
        names.append(cols[i % len(cols)].text_input(f"Name for: {u.name}", value=default_names[i]))
    submitted = st.form_submit_button("Confirm names")

# Parse each file
calendars = []
summary_rows = []
for i, up in enumerate(uploads):
    name = names[i] if submitted else default_names[i]
    if up.name.lower().endswith(".csv"):
        df = read_csv_safely(up)
        ev = clean_events_csv(df, exclude_all_day, exclude_holidays, exclude_birthdays)
    else:  # ICS
        txt = up.read().decode("utf-8", errors="replace")
        ev = ics_to_df(txt)
        if exclude_all_day and "_ALLDAY" in ev.columns:
            ev = ev[~ev["_ALLDAY"]]
        if exclude_holidays and "Subject" in ev.columns:
            ev = ev[~ev["Subject"].astype(str).str.contains("holiday", case=False, na=False)]
        if exclude_birthdays and "Subject" in ev.columns:
            ev = ev[~ev["Subject"].astype(str).str.contains("birthday", case=False, na=False)]
        ev = ev.sort_values("StartDT").reset_index(drop=True)
    calendars.append({"name": name, "events": ev})
    cnt = len(ev)
    dr = f"{ev['StartDT'].min()} → {ev['EndDT'].max()}" if cnt else "—"
    summary_rows.append({"Name": name, "Events": cnt, "Range": dr})

st.subheader("📊 Files parsed")
st.dataframe(pd.DataFrame(summary_rows))

# ======== Day Grid ========
st.subheader(f"🗓️ Availability grid — {show_date.isoformat()}")
slots, step = build_slots(show_date, work_start, work_end, slot_minutes)
if not slots:
    st.warning("No slots in the selected window.")
else:
    grid = []
    for person in calendars:
        busy = person_busy_mask(person["events"], slots, step)
        row = {"Person": person["name"]}
        for i, s in enumerate(slots):
            label = s.strftime("%H:%M")
            row[label] = "Busy" if busy[i] else "Free"
        grid.append(row)
    grid_df = pd.DataFrame(grid)
    st.dataframe(grid_df)

# ======== Per-day diagnostics (selected date) ========
st.subheader("🔎 Events on selected day (per person)")
for person in calendars:
    ev = person["events"]
    day_start = datetime.combine(show_date, time.min)
    day_end = datetime.combine(show_date, time.max)
    day_ev = ev[(ev["StartDT"] <= day_end) & (ev["EndDT"] >= day_start)].copy()
    if not day_ev.empty:
        day_ev["Start"] = day_ev["StartDT"].dt.strftime("%Y-%m-%d %H:%M")
        day_ev["End"] = day_ev["EndDT"].dt.strftime("%Y-%m-%d %H:%M")
        cols = [c for c in ["Start","End","Subject","ShowAs","Location"] if c in day_ev.columns]
        with st.expander(f"📄 {person['name']} — {len(day_ev)} event(s) on {show_date.isoformat()}"):
            st.dataframe(day_ev[cols])
    else:
        st.info(f"{person['name']}: no events on {show_date.isoformat()}")

# ======== Common Free Suggestions ========
st.subheader(f"✅ Common free slots — next {horizon_days} days")
events_only = [p["events"] for p in calendars]
sug = common_free(events_only, work_start, work_end, slot_minutes, horizon_days)
if sug.empty:
    st.warning("No common free slots found. Try widening work hours or horizon, or toggling which 'Show As' count as busy.")
else:
    st.dataframe(sug.head(200))
    st.download_button("Download suggestions (CSV)", sug.to_csv(index=False).encode("utf-8"), "common_free_slots.csv", "text/csv")
