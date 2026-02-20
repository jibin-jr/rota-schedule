import streamlit as st
import pandas as pd
from datetime import timedelta, time, datetime
from io import BytesIO

base="dark"
# =========================
# Configuration & Helpers
# =========================
st.set_page_config(page_title="Rota Assignment + Shifts + Conflicts", layout="wide")

PEAK_BINS = [-float("inf"), 40, 80, 140, 199.75, float("inf")]
PEAK_LABELS = ["Platinum", "Gold", "Silver", "Bronze", "Ungraded"]
# =========================
# Caching heavy only
# =========================
@st.cache_data
def load_excel(upload_file):
    # Adapt column names here to your sheet (kept generic)
    df_fixtures = pd.read_excel(upload_file, sheet_name="Fixtures")
    st.session_state.df_fixtures = df_fixtures[df_fixtures["Match ID"].notna() & (df_fixtures["Match ID"].astype(str).str.strip() != "")]
    df_score = pd.read_excel(upload_file, sheet_name="Historical Score")
    df_availability = pd.read_excel(upload_file, sheet_name="Analyst Availability")
    df_qindex = pd.read_excel(upload_file, sheet_name="QIndex")
    return df_fixtures, df_score, df_availability, df_qindex

# =========================
# Precompute analyst summary once
# =========================
def precompute_best_analyst(df_score: pd.DataFrame) -> pd.DataFrame:
    # Expect df_score has columns: Analyst, Team, Score
    df = df_score.copy()
    df["Grade"] = pd.cut(df["Score"], bins=PEAK_BINS, labels=PEAK_LABELS)

    summary = df.groupby(["Analyst", "Team"], as_index=False).agg(
        match_count=("Score", "count"),
        average_score=("Score", "mean"),
    )

    grade_counts = df.pivot_table(
        index=["Analyst", "Team"],
        columns="Grade",
        values="Score",
        aggfunc="count",
        fill_value=0
    ).reset_index()

    final = summary.merge(grade_counts, on=["Analyst", "Team"], how="left")
    final["Grade"] = pd.cut(final["average_score"], bins=PEAK_BINS, labels=PEAK_LABELS)
    final["average_score"] = final["average_score"].round(2)
    final["merge"] = (
        final["Analyst"]
        + " | "
        + final["match_count"].astype(str)
        + " | "
        + final["average_score"].astype(str)
    )
    return final

def get_best_analyst(analyst_summary: pd.DataFrame, team: str, top_n: int = 10) -> list[str]:
    df = analyst_summary.loc[analyst_summary["Team"] == team]
    if df.empty:
        return pd.DataFrame(columns=["Analyst","match_count", "average_score"])
    # df = df.sort_values(by=["match_count", "average_score"], ascending=[False, True])
    return df[["merge","Analyst","average_score","match_count"]]

def calculate_shift_times(first_ko, shift_length_minutes):
    """
    first_ko: pd.Timestamp
    shift_length_minutes: int
    """
    ko_time = first_ko.time()

    # Night window: 23:00 → 05:00
    if ko_time >= time(23, 0) or ko_time <= time(5, 0):
        shift_start = (
            first_ko.normalize() - timedelta(days=1)
            if ko_time <= time(5, 0)
            else first_ko.normalize()
        ) + timedelta(hours=23)
    else:
        shift_start = first_ko - timedelta(minutes=90)

    shift_end = shift_start + timedelta(hours=shift_length_minutes)

    return shift_start, shift_end

def update_analyst_availability(
    analysts_df: pd.DataFrame,
    analyst_name: str,
    match_start,
    match_end,
    shift_length_hours
):
    """
    Update start/end availability for a given analyst after assignment.
    
    start time  = match end time
    end time    = match start time + shift length
    """

    # Ensure datetimes
    match_start = pd.to_datetime(match_start)
    match_end   = pd.to_datetime(match_end)



    mask = analysts_df["Analyst"] == analyst_name
    # st.write("mask",mask)

    assignmentCount = analysts_df[analysts_df["Analyst"]==analyst_name]["Assignment Count"].to_list()[0]
    # st.write("assignmentCount",assignmentCount)

   
    new_start_time = match_end
    if assignmentCount == 0: 
        firstMatchStartTime = match_start - timedelta(minutes=90)    
        shift_start, shift_end = calculate_shift_times(first_ko=match_start,
                                                       shift_length_minutes=shift_length_hours)
        # st.write("shift_start", shift_start,"shift_end", shift_end)
        analysts_df.loc[mask, "shift_start"]   = shift_start
        analysts_df.loc[mask, "shift_end"]   = shift_end

        new_end_time   = firstMatchStartTime + timedelta(hours=shift_length_hours)
        analysts_df.loc[mask, "End time available"]   = new_end_time
    else:
        new_end_time   = match_start + timedelta(hours=shift_length_hours)

    analysts_df.loc[mask, "start time available"] = new_start_time
    # st.write("new_start_time",new_start_time)
    # st.write("new_end_time",new_end_time)

    # Optional: increment assignment count
    if "Assignment Count" in analysts_df.columns:
        analysts_df.loc[mask, "Assignment Count"] += 1
    

    return analysts_df

def adjust_start_time(current_df, previous_df, shiftInterval):
    """
    current_df: today's analyst availability dataframe
    previous_df: yesterday's used analysts with Shift End
    shiftInterval: hours between shifts
    """

    # Build quick lookup: Analyst -> yesterday shift end
    prev_end_map = (
        previous_df
        .set_index("Analyst")["Shift End"]
        .to_dict()
    )
    # st.write("prev_end_map",prev_end_map)

    def compute_start(row):
        analyst = row["Analyst"]
        # st.write("anlayst",analyst)
        current_start = row["start time available"]
        # st.write("current_start",current_start)

        if analyst in prev_end_map:
            min_start = prev_end_map[analyst] + pd.Timedelta(hours=shiftInterval)
            return max(current_start, min_start)

        return current_start

    current_df["start time available"] = current_df.apply(compute_start, axis=1)

    return current_df

def run_assignment():
    st.session_state.run_assignment_clicked = True
    st.session_state.assignment_completed = False
    disabled=st.session_state.assignment_completed

# =========================
# App UI
# =========================

# Initialize session state
if 'df_fixtures' not in st.session_state:
    st.session_state.df_fixtures = None

if "run_assignment_clicked" not in st.session_state:
    st.session_state.run_assignment_clicked = False

if "assignment_completed" not in st.session_state:
    st.session_state.assignment_completed = False


logo_link = "https://omsstats.wpenginepowered.com/wp-content/themes/orbit-media-bootstrap4/resources/images/logo.png"
st.logo(logo_link, link="https://www.statsperform.com/")
st.set_page_config(page_title="T13 Rota Assignment", layout='wide')

st.markdown(""" 
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&display=swap');

.magic-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 90px;
    color: white;
    letter-spacing: 2px;
    line-height: 85px;
    text-transform: uppercase;
    text-align: center;
    margin-bottom: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 15px; /* space between logo and text */
}

.magic-title img {
    height: 80px;  /* adjust logo size */
}

.gradient-text {
    background: linear-gradient(90deg, #FF0000, #FF7A00, #FFD700); /* red → orange → yellow */
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    color: transparent;
}

.subtitle {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 120px;
    color: white;
    letter-spacing: 8px;
    text-transform: uppercase;
    text-align: center;
    margin-top: -20px;
    margin-bottom: 0;
}
            
</style>

<div class="magic-title">
    <img src="https://www.pngall.com/wp-content/uploads/13/Soccer-PNG-Images.png">
    Tier 13 <span class="gradient-text">ROTA</span> Assignment
</div>
""", unsafe_allow_html=True)

st.markdown(
    """
    <style>
    /* Main title */
    h1 {
        color: #FF4B4B !important;
    }

    /* Headers */
    h2 {
        color: #FF4B4B !important;
    }

    /* Subheaders */
    h3 {
        color: #FF4B4B !important;
    }

    /* Markdown headers (fallback) */
    h4, h5, h6 {
        color: #FF4B4B !important;
    }

    /* Optional: make headers slightly bolder */
    h1, h2, h3 {
        font-weight: 700 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.sidebar.header("Upload Files")
with st.sidebar.expander("Upload Input File",expanded=True):
    uploaded = st.file_uploader("Upload Input Excel file", type=["xlsx"])



if not uploaded:
    col1, col2 = st.columns([3, 1])
    with col1:
        # Welcome/instructions
        st.markdown("""
        <div style="background: white; padding: 30px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <h2 style="color: #4f46e5;">🚀 How to Use</h2>
                <ol style="line-height: 2; color: #374151; font-size: 15px;">
                <li>
                    Refresh the <b>Historical Score / Performance</b> query to ensure the latest data is used in the input file.
                </li>
                <li>
                    Update the <b>Analyst Availability</b> sheet with correct availability.
                </li>
                <li>
                    Update the <b>Fixtures</b> sheet with correct format.
                </li>
                <li>
                    Upload the <b>Input Excel file</b> using the sidebar upload option.
                </li>
                <li>
                    Review fixture dates and confirm the rota range is within <b>7 days</b>.
                </li>
                <li>
                    Click <b>“Run Assignment”</b> to generate analyst assignments.
                </li>
                <li>
                    Validate assignments and workloads in the <b>Preview Tables</b>.
                </li>
                <li>
                    Export the final schedule using the <b>Download Excel</b> option.
                </li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("### 📊 Status Panel")
        st.info("📂 Please upload a fixture file to begin.")
        st.stop()

st.session_state.df_fixtures, df_score, df_availability, df_qindex = load_excel(uploaded)
# Ensure Kick Off is datetime
st.session_state.df_fixtures["Kick Off"] = pd.to_datetime(st.session_state.df_fixtures["Kick Off"], errors="coerce")
# st.session_state.df_fixtures["StartTime"] = st.session_state.df_fixtures["Kick Off"] - timedelta(minutes=30)
# st.session_state.df_fixtures["EndTime"] = st.session_state.df_fixtures["Kick Off"] + timedelta(minutes=120)

st.session_state.df_fixtures["Day"] = pd.to_datetime(st.session_state.df_fixtures["Kick Off"]).dt.strftime("%a")

# Apply 6-hour shift rule + format nicely
st.session_state.df_fixtures["DateKey"] = (
    st.session_state.df_fixtures["Kick Off"]).dt.strftime("%Y-%m-%d %a")

# st.write("df_fixtures",st.session_state.df_fixtures)

analyst_summary = precompute_best_analyst(df_score)
# st.write("analyst_summary",analyst_summary)
today = pd.Timestamp.today().normalize()

df_availability["DOJ in Department"] = pd.to_datetime(
    df_availability["DOJ in Department"],
    errors="coerce"
)

df_availability["Experience (Days)"] = (
    today - df_availability["DOJ in Department"]
).dt.days

# Optional cleanup
df_availability.loc[df_availability["Experience (Days)"] < 0, "Experience (Days)"] = None

# st.write("df_availability",df_availability)


st.header("Scheduling controls ")

with st.expander(label="Peak-Days  :material/trending_down: Non-peak Days Settings",expanded=True,icon=":material/moving:", width='stretch'):
    # st.subheader("Peak/Non-Peak Settings")
    dates_list = st.session_state.df_fixtures["DateKey"].unique().tolist()
    default_peak = [d for d in dates_list if d.endswith("Sat") or d.endswith("Sun")]

    if "peakDays" not in st.session_state:
        st.session_state.peakDays = default_peak
    if "nonPeakDays" not in st.session_state:
        st.session_state.nonPeakDays = [d for d in dates_list if d not in st.session_state.peakDays]

    def _sync_peak():
        st.session_state.nonPeakDays = [d for d in dates_list if d not in st.session_state.peakDays]
    def _sync_nonpeak():
        st.session_state.peakDays = [d for d in dates_list if d not in st.session_state.nonPeakDays]

    colPeakDay1, colPeakDay2 = st.columns([1,1])
    with colPeakDay1:
        st.info("Non-peakDays will have minimum of 09 hours shift duration and minimum of 15 hours shift interval :material/trending_down: ",icon=":material/info:")
        colShiftLength1, colShiftInterval1 = st.columns(2)
        with colShiftLength1:
            nonPeekDayShiftLength = st.number_input("Shift duration (Non-Peak, hrs)", 8, 15, value=9, key= 'nonPeekDayShiftLength')
        with colShiftInterval1:
            nonPeekDayShiftInterval = st.number_input("Shift Interval (Non-Peak, hrs)",min_value=15,max_value=24, key= 'nonPeekDayShiftInterval')
        st.multiselect("Non-Peak Days", dates_list, key="nonPeakDays", default=st.session_state.nonPeakDays, on_change=_sync_nonpeak)
    with colPeakDay2:
        st.info("PeakDays will have minimum of 12 hours shift duration and minimum of only 12 hours shift interval :material/moving:",icon=":material/info:")
        colShiftLength2, colShiftInterval2 = st.columns(2)
        with colShiftLength2:
            peekDayShiftLength = st.number_input("Shift duration (Peak Day, hrs)", 8, 15, value=12, key= 'peekDayShiftLength')
        with colShiftInterval2:
            peekDayShiftInterval = st.number_input("Shift Interval (Peak Day, hrs)",min_value=12,max_value=24 , key= 'peekDayShiftInterval')
        st.multiselect("Peak Days", dates_list, key="peakDays", default=st.session_state.peakDays, on_change=_sync_peak)

    peak_day_set = set(st.session_state.peakDays)

st.markdown("---")
rotaStartDate = dates_list[0]
rotaEndDate = dates_list[-1]
rotaRange = (pd.to_datetime(rotaEndDate) - pd.to_datetime(rotaStartDate)).days
# st.write(rotaStartDate)
# st.write(rotaEndDate)
# st.write(rotaRange)
# st.write(st.session_state.nonPeakDays)
# st.write(st.session_state.peakDays)



# Main content area
col1, col2 = st.columns([3, 1])

with col1:
    # Display extracted table
    if st.session_state.df_fixtures.empty==False:
        # Display DataFrame below
        st.markdown("### 📈 Uploaded Fixture Preview")
        st.dataframe(st.session_state.df_fixtures, use_container_width=True)
        
# st.write("rotaRange",rotaRange)
with col2:
    # Status panel
    st.markdown("### 📊 Status Panel")

    status_card = st.container()
    with status_card:

        if st.session_state.df_fixtures.empty is False:
            st.success("✅ File Loaded")

            # ---------------------------
            # 1) Rota Range Validation
            # ---------------------------
            if rotaRange > 7:
                st.warning(
                    f"⚠️ **Rota Range Too Long!**\n\n"
                    f"You’ve uploaded fixtures spanning **{rotaRange + 1} days** 📅.\n"
                    "This tool supports **only a single-week** rota planning.\n\n"
                    "👉 Please upload fixtures within a 7-day range and try again."
                )
                st.stop()
            else:
                st.success(
                    f"📆 **Fixture Date Range OK**\n\n"
                    f"Your fixtures span **{rotaRange + 1} days**.\n"
                    "Within supported 1-week range ✅"
                )

            # ---------------------------
            # 2) Q-Index Competition Validation
            # ---------------------------
            fixture_df = st.session_state.df_fixtures

            fixture_comps = set(fixture_df["Competition"].dropna().unique())
            qindex_comps = set(df_qindex["Competition"].dropna().unique())

            missing_comps = sorted(fixture_comps - qindex_comps)

            if missing_comps:
                missing_list = "\n".join([f"• {c}" for c in missing_comps])

                st.error(
                    "❌ **Missing Q-Index Data!**\n\n"
                    "The following competitions are present in your fixtures "
                    "but **not found** in the Q-Index file:\n\n"
                    f"{missing_list}\n\n"
                    "👉 Please add these competitions to the Q-Index sheet before running assignments."
                )

                # Optional: show count + expander
                st.metric("Missing Competitions", len(missing_comps))
                with st.expander("🔍 View Missing Competitions"):
                    for c in missing_comps:
                        st.write(f"• {c}")

                st.stop()
            else:
                st.success(
                    "🎯 **Q-Index Validation Passed!**\n\n"
                    "Every competition in the fixture file has a matching Q-Index entry.\n"
                    "You're good to proceed with assignments! 🚀"
                )





        else:
            st.info("📂 Please upload a fixture file to begin.")

df_overall_shifts = pd.DataFrame({
    'Date': pd.Series(dtype='int'),
    'Analyst': pd.Series(dtype='str'),
    'Shift Start': pd.Series(dtype='datetime64[ns]'),
    'Shift End': pd.Series(dtype='datetime64[ns]'),
    'Shift Type': pd.Series(dtype='str')
})


processingDateUpdated_df = []
# st.date_input("start Date", value="today")
for i in range(0, (int(rotaRange)+1)):
    # currentDayAssignedAnalyst = []
    # st.write(i)
    # st.write(st.session_state)
    currentDate = pd.to_datetime(rotaStartDate)+timedelta(days=i)
    colDateHeader =  pd.to_datetime(currentDate).strftime("%A, %B %#d, %Y")
    st.header(colDateHeader) 

    col1, col2, col3, col4, col5, col6=  st.columns([0.5, 0.5, 0.5,0.5,0.5, 0.5])
    with col1:
        matchLen = st.number_input(f" Match Length in Minutes", min_value=120, max_value=240, value=120, key=f"{colDateHeader}_MatchLength")
    with col2:
        if i > 0:
            previousDay = (currentDate - timedelta(days=1)).strftime("%A, %B %#d, %Y")
            previousDayEnd = st.session_state.get(f"{previousDay}_matchDayEnd")
            # st.write("previousDay start",previousDayEnd)
            dt = pd.to_datetime(previousDayEnd)
            # hour = dt.hour
            # st.write("hour", hour)
            startTimeHourValue = dt.hour
            startTimeMinuteValue = dt.minute
        else:
            startTimeHourValue = 12
            startTimeMinuteValue = 0
        # st.write("startTimeValue",startTimeHourValue,startTimeMinuteValue)
        MatchDayStart = st.time_input("Match Day Start (hrs)",value=time(startTimeHourValue,startTimeMinuteValue),key=f"{colDateHeader}_MatchDayStart")

        # MatchDayStart = st.number_input(f" Match Day Start (hrs)", min_value=1, max_value=24, value=startTimeValue, key=f"{colDateHeader}_MatchDayStart")
    with col4:
        MatchDayEnd = st.time_input("Match Day End (hrs)", value=time(6,0), key=f"{colDateHeader}_MatchDayEnd")

        # MatchDayEnd = st.number_input(f" Match Day End (hrs)", min_value=1, max_value=40, value=30, key=f"{colDateHeader}_MatchDayEnd")
    with col3:


        keyStart = f"{colDateHeader}_matchDayStart" 

        if keyStart not in st.session_state:
            st.session_state[keyStart] = (
                pd.to_datetime(currentDate)+timedelta(minutes=((MatchDayStart.hour)*60)+MatchDayStart.minute)
            )
        st.session_state[keyStart] = (
                pd.to_datetime(currentDate)+timedelta(minutes=((MatchDayStart.hour)*60)+MatchDayStart.minute)
            )
        # st.write("MatchDayStart",MatchDayStart)

        st.info(f"### Matchday start from  \n{(st.session_state[keyStart])}")

    with col5:
        keyEnd = f"{colDateHeader}_matchDayEnd" 

        if keyEnd not in st.session_state:
            st.session_state[keyEnd] = (
                pd.to_datetime(currentDate)+timedelta(minutes=(1440+(MatchDayEnd.hour*60)+MatchDayEnd.minute))
            )
        st.session_state[keyEnd] = (
                pd.to_datetime(currentDate)+timedelta(minutes=(1440+(MatchDayEnd.hour*60)+MatchDayEnd.minute))
            )
        # st.info(f"### Matchday start from  \n{(st.session_state[keyEnd])}")
        # matchDayEndTimeByDay = pd.to_datetime(currentDate)+timedelta(hours=MatchDayEnd)
        st.info(f"### Matchday Ends at  \n{(st.session_state[keyEnd])}")
    with col6:
        st.success(f"### Processing Date \n{(st.session_state[keyStart]).strftime('%A, %B %#d, %Y')}")

    dateStr = pd.to_datetime(currentDate).strftime("%A, %d %B %Y")
    # st.write(dateStr)
    if dateStr not in st.session_state.peakDays:
        isPeakDay = False
    else:
        isPeakDay = True
    if isPeakDay:
        maximumAssignmentCount = 3
        shiftLength = peekDayShiftLength
        shiftInterval = peekDayShiftInterval
    else:
        maximumAssignmentCount = 2
        shiftLength = nonPeekDayShiftLength
        shiftInterval = nonPeekDayShiftInterval
    # st.write(isPeakDay, shiftLength, shiftInterval, maximumAssignmentCount)
    colDate = pd.to_datetime(currentDate).strftime("%A")
    # st.write(colDate)
    matchDayStartTime = currentDate + timedelta(hours=4)
    matchDayEndTime = currentDate + timedelta(hours=30)
    # st.write(matchDayStartTime,matchDayEndTime)
    currenDateAnalyst = df_availability[df_availability[colDate]=='Y'][['Oracle ID','Batch','Analyst',colDate,'Experience (Days)']]
    currenDateAnalyst['start time available'] = matchDayStartTime
    currenDateAnalyst['End time available'] = matchDayEndTime + timedelta(minutes=180)
    currenDateAnalyst['Assignment Count'] = 0    
    currenDateAnalyst[['shift_start','shift_end']] = None
    # st.write("currenDateAnalyst",currenDateAnalyst)
    st.session_state.df_fixtures["Match Processing Date"] = (st.session_state[keyStart]).strftime('%A, %B %#d, %Y')

    currentDayFixtures = st.session_state.df_fixtures[(st.session_state.df_fixtures['Kick Off']>=(st.session_state[keyStart])) &
                                                        (st.session_state.df_fixtures['Kick Off']<(st.session_state[keyEnd]))]

    st.write("### Fixtures")
    st.dataframe(currentDayFixtures)
    if currentDayFixtures.empty == False:
        processingDateUpdated_df.append(currentDayFixtures)
    st.markdown("---")

final_fixtures = pd.concat(processingDateUpdated_df,
                           ignore_index=True
                           )

# st.dataframe(final_fixtures)
st.button(
        "🚀 Run Assignment",
        type="primary",
        on_click=run_assignment
                )

dates_list2 = final_fixtures["Match Processing Date"].unique().tolist()
# st.write(dates_list2)
rotaStartDate2 = dates_list2[0]
rotaEndDate2 = dates_list2[-1]
rotaRange2 = (pd.to_datetime(rotaEndDate2) - pd.to_datetime(rotaStartDate2)).days
# st.write(rotaStartDate2)
# st.write(rotaEndDate2)
# st.write(rotaRange2)
# st.write(st.session_state)
if st.session_state.run_assignment_clicked :
    with st.spinner("Running assignment... ⏳"):
        shifts = []
        ovarallAssignmentsList = []
        nonUsed = []
        
        for i in range(0, (int(rotaRange2)+1)):
            # currentDayAssignedAnalyst = []
            # st.write(i)
            # st.write(st.session_state)
            currentDate = pd.to_datetime(rotaStartDate2)+timedelta(days=i)
            # st.write(currentDate)
            colDateHeader =  pd.to_datetime(currentDate).strftime("%A, %B %#d, %Y")
            # st.header(colDateHeader) 


            dateStr = pd.to_datetime(currentDate).strftime("%Y-%m-%d %a")
            # st.write(dateStr)
            # st.write("st.session_state.peakDays",st.session_state.peakDays)
            if dateStr not in st.session_state.peakDays:
                isPeakDay = False
            else:
                isPeakDay = True
            if isPeakDay:
                maximumAssignmentCount = 3
                shiftLength = peekDayShiftLength
                shiftInterval = peekDayShiftInterval 
            else:
                maximumAssignmentCount = 2
                shiftLength = nonPeekDayShiftLength
                shiftInterval = nonPeekDayShiftInterval
            # st.write(isPeakDay, shiftLength, shiftInterval, maximumAssignmentCount)
            colDate = pd.to_datetime(currentDate).strftime("%A")
            # st.write(colDate)
            keyStart = f"{colDateHeader}_matchDayStart"
            # st.write("Keystart", st.session_state[keyStart])
            keyEnd = f"{colDateHeader}_matchDayEnd"
            # st.write("KeyEnd", st.session_state[keyEnd])              
            matchDayStartTime = st.session_state[keyStart]
            matchDayEndTime =  st.session_state[keyEnd]
            # st.write(matchDayStartTime,matchDayEndTime)
            currenDateAnalyst = df_availability[df_availability[colDate]=='Y'][['Oracle ID','Batch','Analyst',colDate,'Experience (Days)']]
          
            currenDateAnalyst['start time available'] = matchDayStartTime
            currenDateAnalyst['End time available'] = matchDayEndTime + timedelta(minutes=180)
            currenDateAnalyst['Assignment Count'] = 0    
            currenDateAnalyst[['shift_start','shift_end']] = None
            # st.write("currenDateAnalyst",currenDateAnalyst)
            # st.write("final_fixtures",final_fixtures)
            currentProcessingDate = pd.to_datetime(dateStr).strftime("%A, %B %#d, %Y")
            # st.write(currentProcessingDate)

            currentDayFixtures = final_fixtures[final_fixtures['Match Processing Date']==currentProcessingDate]

            # st.write("currentDayFixtures",currentDayFixtures)
            currentDayFixtures = pd.merge(currentDayFixtures, df_qindex, on='Competition', how='inner').sort_values(by=['Tier','QIndex Target', 'Kick Off','Is_PMT'], ascending=[False,False,True,False])
            # st.write("currentDayFixtures",currentDayFixtures)
            
            #Handling First Day assignment
            if i == 0:
                currenDateAnalyst = currenDateAnalyst
                # st.write("No previous Day")
            else:
                previousDay = (currentDate - timedelta(days=1)).strftime("%A, %B %#d, %Y")
                # st.write(f"Previous Day {previousDay} used Analyst")
                previousDayUsedAnalyst = df_shifts[df_shifts['Date']==previousDay]
                # st.write("previousDayUsedAnalyst",previousDayUsedAnalyst)
                currenDateAnalyst = adjust_start_time(current_df= currenDateAnalyst,
                                                        previous_df= previousDayUsedAnalyst,
                                                        shiftInterval= shiftInterval)

                # st.write("currenDateAnalystNew",currenDateAnalyst)
                # currenDateAnalyst = currenDateAnalyst

            # st.write("currenDateAnalyst",currenDateAnalyst)
            assignmentsList = []
            for index, row in currentDayFixtures.iterrows():
                # st.write(row)
                KickOff = row["Kick Off"]
                matchStartTime = KickOff 
                matchEndTime = KickOff + timedelta(minutes=matchLen)
                # st.write("matchEndTime",matchEndTime)
                homeTeam = row["Home Team"]
                awayTeam =  row["Away Team"]
                Is_PMT = row["Is_PMT"]
                # st.write("homeTeam",homeTeam)

                h_opts = get_best_analyst(analyst_summary, homeTeam) 
                # st.write("h_opts",h_opts)

                if h_opts.empty == False:
                    HomeAvailableAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                (currenDateAnalyst['End time available']>=matchEndTime) &
                                                (currenDateAnalyst['Analyst'].isin(h_opts['Analyst'])) &
                                                 (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                ].sort_values(by=['Assignment Count'],ascending=False)
                else:
                    HomeAvailableAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                (currenDateAnalyst['End time available']>=matchEndTime) &
                                (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                ].sort_values(by=['Assignment Count'],ascending=False)

                # st.write("HomeAvailableAnalyst",HomeAvailableAnalyst)
                mergeHome = pd.merge(HomeAvailableAnalyst, 
                        h_opts, 
                        on='Analyst', 
                        how='inner').sort_values(by=["match_count", "average_score"], 
                                                ascending=[False, True])

                # st.write("mergeHome",mergeHome)
                if mergeHome.empty == False:
                    HomeAnalyst =  mergeHome[mergeHome['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                else:
                    if Is_PMT == "Yes":
                        # st.write("PMT Match")
                        # st.write("currenDateAnalyst",currenDateAnalyst)
                        ExeperienceAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                            (currenDateAnalyst['End time available']>=matchEndTime) &
                                                            (currenDateAnalyst["Experience (Days)"]>= 365) &
                                                            (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                               ].sort_values(by=["Assignment Count", "Experience (Days)"], 
                                                                ascending=[True, False])
                        # st.write("ExeperienceAnalyst",ExeperienceAnalyst)
                        if ExeperienceAnalyst.empty == False:
                            HomeAnalyst =  ExeperienceAnalyst[ExeperienceAnalyst['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                        else:
                            NonExeperienceAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                            (currenDateAnalyst['End time available']>=matchEndTime) &
                                                            (currenDateAnalyst["Experience (Days)"]<= 365) &
                                                            (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                                ].sort_values(by=["Assignment Count", "Experience (Days)"], 
                                                                ascending=[True, False])
                            # st.write("NonExeperienceAnalyst",NonExeperienceAnalyst)
                            if NonExeperienceAnalyst.empty == False:
                                HomeAnalyst =  NonExeperienceAnalyst[NonExeperienceAnalyst['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                            else:
                                HomeAnalyst = None
                    if Is_PMT == "No":
                        # st.write("Non PMT Match")
                        NonExeperienceAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                            (currenDateAnalyst['End time available']>=matchEndTime)&
                                                            (currenDateAnalyst["Experience (Days)"]<= 365) &
                                                                  (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                                  ].sort_values(by=["Assignment Count", "Experience (Days)"], 
                                                                ascending=[True, False])
                        # st.write("NonExeperienceAnalyst",NonExeperienceAnalyst)
                        if NonExeperienceAnalyst.empty == False:
                            HomeAnalyst =  NonExeperienceAnalyst[NonExeperienceAnalyst['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                        else:
                            HomeAnalyst = None

                # HomeAnalyst =  mergeHome["Analyst"].to_list()[0]
                # st.write("HomeAnalyst",HomeAnalyst)

                if HomeAnalyst != None:
                    # st.write("Home analyst not none")
                    currenDateAnalyst = update_analyst_availability(currenDateAnalyst, HomeAnalyst, matchStartTime,
                                                                matchEndTime, shiftLength).sort_values(by=['Assignment Count'],ascending=False)
                # if i > 0:
                #     st.write("homeTeam",homeTeam)
                #     st.write("HomeAnalyst",HomeAnalyst)
                #     st.write("currenDateAnalyst",currenDateAnalyst)

                # st.write("awayTeam",awayTeam)
                a_opts = get_best_analyst(analyst_summary, awayTeam)
                # st.write("a_opts",a_opts)

                if a_opts.empty == False:
                    AwayAvailableAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                (currenDateAnalyst['End time available']>=matchEndTime) &
                                                (currenDateAnalyst['Analyst'].isin(a_opts['Analyst'])) &
                                                (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                ].sort_values(by=['Assignment Count'],ascending=False)
                else:
                    AwayAvailableAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                (currenDateAnalyst['End time available']>=matchEndTime) &
                                (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                ].sort_values(by=['Assignment Count'],ascending=False)
                # st.write("AwayAvailableAnalyst",AwayAvailableAnalyst)

                mergeAway = pd.merge(AwayAvailableAnalyst, 
                                        a_opts, 
                                        on='Analyst', 
                                        how='inner').sort_values(by=["match_count", "average_score"], 
                                                    ascending=[False, True])
                # if i >0:
                #     st.write("mergeAway",mergeAway)
                if mergeAway.empty == False:
                    AwayAnalyst =  mergeAway[mergeAway['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                else:
                    if Is_PMT == "Yes":
                        # st.write("PMT Match")
                        # st.write("currenDateAnalyst",currenDateAnalyst)
                        ExeperienceAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                            (currenDateAnalyst['End time available']>=matchEndTime) &
                                                            (currenDateAnalyst["Experience (Days)"]>= 365) &
                                                               (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                               ].sort_values(by=["Assignment Count", "Experience (Days)"], 
                                                                ascending=[True, False])
                        # if i > 0:
                        #     st.write("ExeperienceAnalyst",ExeperienceAnalyst)
                        if ExeperienceAnalyst.empty == False:
                            AwayAnalyst =  ExeperienceAnalyst[ExeperienceAnalyst['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                        else:
                            NonExeperienceAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                            (currenDateAnalyst['End time available']>=matchEndTime) &
                                                            (currenDateAnalyst["Experience (Days)"]<= 365) &
                                                                      (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                                      ].sort_values(by=["Assignment Count", "Experience (Days)"], 
                                                                ascending=[True, False])
                            # if i > 0:
                            #     st.write("NonExeperienceAnalyst",NonExeperienceAnalyst)
                            if NonExeperienceAnalyst.empty == False:
                                AwayAnalyst =  NonExeperienceAnalyst[NonExeperienceAnalyst['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                            else:
                                AwayAnalyst = None
                    if Is_PMT == "No":
                        # st.write("Non PMT Match")
                        NonExeperienceAnalyst = currenDateAnalyst[(currenDateAnalyst['start time available']<=matchStartTime) &
                                                            (currenDateAnalyst['End time available']>=matchEndTime) &
                                                            (currenDateAnalyst["Experience (Days)"]<= 365) &
                                                                  (currenDateAnalyst['Assignment Count']<maximumAssignmentCount)
                                                                  ].sort_values(by=["Assignment Count", "Experience (Days)"], 
                                                                ascending=[True, False])
                        # st.write("NonExeperienceAnalyst",NonExeperienceAnalyst)
                        if NonExeperienceAnalyst.empty == False:
                            AwayAnalyst =  NonExeperienceAnalyst[NonExeperienceAnalyst['Assignment Count']<maximumAssignmentCount]["Analyst"].to_list()[0]
                        else:
                            AwayAnalyst = None

                # st.write("AwayAnalyst",AwayAnalyst)

                if AwayAnalyst != None:
                    currenDateAnalyst = update_analyst_availability(currenDateAnalyst, AwayAnalyst, matchStartTime,
                                                                matchEndTime, shiftLength).sort_values(by=['Assignment Count'],ascending=False)
                # if i > 0:
                #     st.write("AwayeTeam",awayTeam)
                #     st.write("AwayAnalyst",AwayAnalyst)
                #     st.write("currenDateAnalyst",currenDateAnalyst)
                # st.write("currenDateAnalyst",currenDateAnalyst)            

                data = {
                        'Processing Date': row['Match Processing Date'],
                        'Tier': row['Tier'],
                        'Kick Off': row['Kick Off'],
                        'Match ID': row['Match ID'],
                        'Competition': row['Competition'],
                        'Home Team': row['Home Team'],
                        'Away Team': row['Away Team'],
                        'Home Analyst': HomeAnalyst,
                        'Away Analyst': AwayAnalyst,
                        'StartTime': matchStartTime,
                        'EndTime': matchEndTime
                        }
                # st.write(data)
                assignmentsList.append(data)
                ovarallAssignmentsList.append(data)
                # if i >0:
                #     st.write(data)

            
            currentDateMatchAssignment = pd.DataFrame(assignmentsList)
            # st.subheader("Assignment")
            # st.dataframe(currentDateMatchAssignment)
            # st.markdown("---")


            # st.write("currenDateAnalyst",currenDateAnalyst)

            currentDayUsedAnalystList = (pd.concat([currentDateMatchAssignment["Home Analyst"],
                                                currentDateMatchAssignment["Away Analyst"]])).dropna().unique().tolist()
            # st.write("currentDayUsedAnalystList",currentDayUsedAnalystList)

            currentDayNonUsedAnalys_df = currenDateAnalyst[~currenDateAnalyst["Analyst"].isin(currentDayUsedAnalystList)]
            # st.write("currentDayNonUsedAnalys_df",currentDayNonUsedAnalys_df)

            currentDayUsedAnalysShift = currenDateAnalyst[currenDateAnalyst["Analyst"].isin(currentDayUsedAnalystList)]
            # st.write("currentDayUsedAnalysShift",currentDayUsedAnalysShift)    

            
            for index, row in currentDayUsedAnalysShift.iterrows():
                shifts.append({
                    "Date": colDateHeader,
                    "Analyst": row['Analyst'],
                    "Shift Start": row['shift_start'],
                    "Shift End": row['shift_end'],
                    "Assignment Count": row['Assignment Count']
                })

            df_shifts = pd.DataFrame(shifts)
            df_currentDay_shifts = df_shifts[df_shifts['Date']==colDateHeader]
            # st.subheader("Shifts")
            # st.dataframe(df_currentDay_shifts)

            for index, row in currentDayNonUsedAnalys_df.iterrows():
                nonUsed.append({
                    "Date": colDateHeader,
                    "Analyst": row['Analyst'],
                    "Assignment Count": row['Assignment Count']
                })
        
            df_NonUsedAnalyst = pd.DataFrame(nonUsed)
            # st.write("df_NonUsedAnalyst",df_NonUsedAnalyst)
            if df_NonUsedAnalyst.empty == False:
                df_currentDay_NonUsedAnalyst = df_NonUsedAnalyst[df_NonUsedAnalyst['Date']==colDateHeader]
            # st.write("df_currentDay_NonUsedAnalyst",df_currentDay_NonUsedAnalyst)

        overallMatchAssignment_df = pd.DataFrame(ovarallAssignmentsList).sort_values(by='Kick Off',ascending=True)

        st.subheader("📅 Assignment Overview")
        st.dataframe(overallMatchAssignment_df, hide_index=True)
        # st.dataframe("df_shifts Overall",df_shifts)
        df = df_shifts.copy()

        df["Shift Start"] = pd.to_datetime(df["Shift Start"])
        df["Shift End"]   = pd.to_datetime(df["Shift End"])

        df["Shift Date"] = df["Shift Start"].dt.strftime("%d-%b %a")
        df["Shift Start Display"] = df["Shift Start"].dt.strftime("%I:%M %p").str.lstrip("0")
        rota_df = (
                    df.pivot_table(
                        index="Analyst",
                        columns="Shift Date",
                        values="Shift Start Display",
                        aggfunc="first"   # safe: one shift per analyst per day
                    )
                    .reset_index()
                )
        date_cols = sorted(
            rota_df.columns[1:], 
            key=lambda x: pd.to_datetime(x, format="%d-%b %a")
        )

        rota_df = rota_df[["Analyst"] + date_cols]
        st.subheader("📅 Analyst Shift Overview")
        st.dataframe(rota_df, use_container_width=True, hide_index=True)





        # st.write("df Overall Non Used analyst",df_NonUsedAnalyst)

        cms_df = overallMatchAssignment_df[['Match ID','Home Team','Away Team','Home Analyst','Away Analyst']]
        # st.write("cms_df",cms_df)
        # Build Analyst Info column
        cms_df["Analyst Info"] = cms_df["Home Analyst"] + ", " + cms_df["Away Analyst"]

        cmsfinal_df = cms_df[["Match ID", "Analyst Info"]]
        # output = BytesIO()

        import io
        # import pandas as pd

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook  = writer.book

            # ---------------- Sheets ----------------
            overallMatchAssignment_df.to_excel(writer, index=False, sheet_name="Assignments")
            df_shifts.to_excel(writer, index=False, sheet_name="Shifts")
            cmsfinal_df.to_excel(writer, index=False, sheet_name="cms upload")
            df_NonUsedAnalyst.to_excel(writer, index=False, sheet_name="Non Used Analyst")
            analyst_summary.to_excel(writer, index=False, sheet_name="Analyst performance summary")
            rota_df.to_excel(writer, index=False, sheet_name="Rota")


            # ---------------- Formats ----------------
            high_load_fmt = workbook.add_format({
                "bg_color": "#FFC7CE",  # red
                "font_color": "#9C0006"
            })

            medium_load_fmt = workbook.add_format({
                "bg_color": "#FFE699",  # orange
                "font_color": "#9C6500"
            })

            low_load_fmt = workbook.add_format({
                "bg_color": "#C6EFCE",  # green
                "font_color": "#006100"
            })

            # ---------------- Apply Analyst Colouring (df_shifts only) ----------------
            shifts_ws = writer.sheets["Shifts"]

            analyst_col = df_shifts.columns.get_loc("Analyst")
            assignment_col = df_shifts.columns.get_loc("Assignment Count")

            for row_idx, assignment_count in enumerate(df_shifts["Assignment Count"], start=1):
                cell_row = row_idx  # Excel row (skip header)

                if assignment_count == 3:
                    shifts_ws.write(cell_row, analyst_col, df_shifts.iloc[row_idx - 1, analyst_col], high_load_fmt)
                elif assignment_count == 2:
                    shifts_ws.write(cell_row, analyst_col, df_shifts.iloc[row_idx - 1, analyst_col], medium_load_fmt)
                elif assignment_count == 1:
                    shifts_ws.write(cell_row, analyst_col, df_shifts.iloc[row_idx - 1, analyst_col], low_load_fmt)

            # ---------------- Auto-fit Columns for ALL Sheets ----------------
            for sheet_name, df in {
                "Assignments": overallMatchAssignment_df,
                "Shifts": df_shifts,
                "cms upload": cmsfinal_df,
                "Non Used Analyst": df_NonUsedAnalyst,
                "Rota": rota_df,
                "Analyst performance summary": analyst_summary
            }.items():
                worksheet = writer.sheets[sheet_name]

                for idx, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2
                    worksheet.set_column(idx, idx, max_len)

        # Convert to bytes
        excel_data = output.getvalue()

        # st.session_state.assignment_completed = True
        # st.session_state.run_assignment_clicked = False

        # Streamlit download button
        if st.download_button(
            label="Download as Excel File",
            data=excel_data,
            file_name="assignment_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ):
            st.success("File has been exported successfully !!!")



    st.success("✅ Assignment completed successfully!")            


