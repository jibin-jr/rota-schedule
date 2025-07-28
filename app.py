import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time



st.title("T13 Rota Assignment")

upload_file = st.file_uploader("Upload your excel file", type=["xlsx", "xls"])

if upload_file:
    df_fixtures = pd.read_excel(upload_file, sheet_name="Fixtures").sort_values(by=['Kick Off'])
    df_score = pd.read_excel(upload_file, sheet_name="Historical Score")
    df_availability = pd.read_excel(upload_file, sheet_name="Analyst Availability")
    df_fixtures['Kick Off'] = pd.to_datetime(df_fixtures['Kick Off'])
    cutoff = time(6,0,0)
    df_fixtures['Dates'] = df_fixtures['Kick Off'].apply(lambda ko: ko.date() if ko.time() >= cutoff else (ko - pd.Timedelta(days=1)).date() ).astype(str)
    dates_list = df_fixtures['Dates'].unique().tolist()
    st.multiselect("Select the Peek Day",dates_list)
    # RotaStartDate = 
        