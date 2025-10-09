import streamlit as st
import pandas as pd
import requests
import datetime
import matplotlib.pyplot as plt
from io import BytesIO

# Initialize session state for logs
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []

if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []

if "last_fifteen_log_time" not in st.session_state:
    st.session_state.last_fifteen_log_time = None

# Function to fetch NIFTY option chain data
def fetch_option_chain():
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.nseindia.com/"
    }
    session = requests.Session()
    session.get("https://www.nseindia.com", headers=headers)
    response = session.get(url, headers=headers)
    data = response.json()
    return data

# Function to process option chain data
def process_data(data):
    records = data["records"]["data"]
    ce_oi_change = 0
    pe_oi_change = 0
    for item in records:
        ce = item.get("CE")
        pe = item.get("PE")
        if ce:
            ce_oi_change += ce.get("changeinOpenInterest", 0)
        if pe:
            pe_oi_change += pe.get("changeinOpenInterest", 0)

    signal = "Neutral"
    if ce_oi_change < pe_oi_change:
        signal = "Buy"
    elif ce_oi_change > pe_oi_change:
        signal = "Sell"

    pcr = round(pe_oi_change / ce_oi_change, 2) if ce_oi_change != 0 else None

    return {
        "Time": datetime.datetime.now().strftime("%H:%M:%S"),
        "CE OI Change": ce_oi_change,
        "PE OI Change": pe_oi_change,
        "Signal": signal,
        "PCR": pcr
    }

# Sidebar button to fetch data
st.sidebar.title("OI Dashboard")
if st.sidebar.button("Fetch Latest Data"):
    try:
        data = fetch_option_chain()
        entry = process_data(data)
        st.session_state.five_min_log.append(entry)

        # Log to 15-min if time difference is >= 15 minutes
        now = datetime.datetime.now()
        if st.session_state.last_fifteen_log_time is None or (now - st.session_state.last_fifteen_log_time).seconds >= 900:
            st.session_state.fifteen_min_log.append(entry)
            st.session_state.last_fifteen_log_time = now

    except Exception as e:
        st.error(f"Error fetching data: {e}")

# Display 5-min log
st.subheader("5-Minute Log")
df_5min = pd.DataFrame(st.session_state.five_min_log)
st.dataframe(df_5min)

# PCR Trend Chart
if not df_5min.empty and "PCR" in df_5min.columns:
    st.subheader("PCR Trend Chart")
    fig, ax = plt.subplots()
    ax.plot(df_5min["Time"], df_5min["PCR"], marker='o', linestyle='-')
    ax.set_xlabel("Time")
    ax.set_ylabel("PCR")
    ax.set_title("Put-Call Ratio Over Time")
    plt.xticks(rotation=45)
    st.pyplot(fig)

# Display 15-min log
st.subheader("15-Minute Log")
df_15min = pd.DataFrame(st.session_state.fifteen_min_log)
st.dataframe(df_15min)

# Download Excel
def generate_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_5min.to_excel(writer, index=False, sheet_name="FiveMinLog")
        df_15min.to_excel(writer, index=False, sheet_name="FifteenMinLog")
    output.seek(0)
    return output

st.sidebar.download_button(
    label="Download Excel",
    data=generate_excel(),
    file_name="OIAnalysisDashboard.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
