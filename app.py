import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time

st.set_page_config(page_title="T13 Rota Assignment", layout='wide',page_icon='⚽')

st.title("T13 Rota Assignment")
st.markdown('###')

upload_file = st.file_uploader("Upload your excel file", type=["xlsx", "xls"])
st.markdown("---")

if upload_file:

    col1, col2 = st.columns(2)
    df_fixtures = pd.read_excel(upload_file, sheet_name="Fixtures", usecols='A:H').sort_values(by=['Kick Off'])
    df_score = pd.read_excel(upload_file, sheet_name="Historical Score")
    df_availability = pd.read_excel(upload_file, sheet_name="Analyst Availability")

    df_fixtures['Kick Off'] = pd.to_datetime(df_fixtures['Kick Off'])
    cutoff = time(6,0,0)
    df_fixtures['Dates'] = df_fixtures['Kick Off'].apply(lambda ko: ko.date() if ko.time() >= cutoff else (ko - pd.Timedelta(days=1)).date() ).astype(str)
    df_fixtures['DatesFormatted'] = pd.to_datetime(df_fixtures["Dates"]).dt.strftime('%Y-%m-%d %a')
    dates_list = df_fixtures['DatesFormatted'].unique().tolist()
    df_fixtures['WeekdayName'] = pd.to_datetime(df_fixtures["Dates"]).dt.day_name()

    with col1:
        st.caption("Peek Days will have 10 hrs shift")
        colShiftLength1, colShiftInterval1 = st.columns(2)
        with colShiftLength1:
            peekDayShiftLength = st.number_input("Shift Length (Peek Day)",min_value=10,max_value=15)
        with colShiftInterval1:
            peekDayShiftInterval = st.number_input("Shift Interval (Peek Day)",min_value=12,max_value=24)
        defaultPeakDayList = df_fixtures[df_fixtures["WeekdayName"].isin(['Saturday','Sunday'])]['DatesFormatted'].unique().tolist()
        
        peekday = st.multiselect("Select the Peek Day",dates_list,default=defaultPeakDayList)
        # st.write(peekday)
    with col2:
        NonPeakDayList = df_fixtures[~df_fixtures["DatesFormatted"].isin(peekday)]['DatesFormatted'].unique().tolist()
        st.caption("Non-Peek Days will have 9 hrs shift")
        # st.write(nonPeakDay)
        colShiftLength2, colShiftInterval2 = st.columns(2)
        with colShiftLength2:
            nonPeekDayShiftLength = st.number_input("Shift Length (Non-Peek Day)",min_value=9,max_value=15)
        with colShiftInterval2:
            nonPeekDayShiftInterval = st.number_input("Shift Interval (Non-Peek Day)",min_value=15,max_value=24)
        nonPeakDay = st.multiselect("Select Non-Peak Days", NonPeakDayList,default=NonPeakDayList, width="stretch")

    st.markdown("---")
    # st.write(df_fixtures)
