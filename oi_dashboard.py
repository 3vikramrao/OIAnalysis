import streamlit as st
import pandas as pd
import requests
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

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
    return data["records"]["data"]

# Function to process option chain data
def process_option_chain(data):
    rows = []
    for item in data:
        strike = item.get("strikePrice")
        ce_oi = item.get("CE", {}).get("openInterest", 0)
        pe_oi = item.get("PE", {}).get("openInterest", 0)
        total_oi = ce_oi + pe_oi
        rows.append({
            "Strike Price": strike,
            "CE OI": ce_oi,
            "PE OI": pe_oi,
            "Total OI": total_oi
        })
    df = pd.DataFrame(rows)
    df_sorted = df.sort_values(by="Total OI", ascending=False).reset_index(drop=True)
    return df_sorted

# Initialize session state for logs
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []

if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []

if "last_fifteen_log_time" not in st.session_state:
    st.session_state.last_fifteen_log_time = datetime.now() - timedelta(minutes=15)

st.title("NIFTY OI Analysis Dashboard")

# Sidebar button to fetch data
if st.sidebar.button("Fetch Latest Data"):
    try:
        option_chain_data = fetch_option_chain()
        now = datetime.now().strftime("%H:%M:%S")

        # Aggregate CE and PE OI changes
        ce_oi_total = sum(item.get("CE", {}).get("openInterest", 0) for item in option_chain_data)
        pe_oi_total = sum(item.get("PE", {}).get("openInterest", 0) for item in option_chain_data)

        # Calculate PCR
        pcr = round(pe_oi_total / ce_oi_total, 2) if ce_oi_total else 0

        # Generate signal
        signal = "Buy" if pcr > 1 else "Sell"

        # Log 5-min data
        st.session_state.five_min_log.append({
            "Time": now,
            "CE OI": ce_oi_total,
            "PE OI": pe_oi_total,
            "PCR": pcr,
            "Signal": signal
        })

        # Log 15-min data if time elapsed
        current_time = datetime.now()
        if current_time - st.session_state.last_fifteen_log_time >= timedelta(minutes=15):
            st.session_state.fifteen_min_log.append({
                "Time": now,
                "CE OI": ce_oi_total,
                "PE OI": pe_oi_total,
                "PCR": pcr,
                "Signal": signal
            })
            st.session_state.last_fifteen_log_time = current_time

    except Exception as e:
        st.error(f"Error fetching data: {e}")

# Display 5-min log
st.subheader("5-Minute Log")
df_5min = pd.DataFrame(st.session_state.five_min_log)
st.dataframe(df_5min)

# Display PCR trend chart
if not df_5min.empty:
    st.subheader("PCR Trend Chart")
    fig, ax = plt.subplots()
    ax.plot(df_5min["Time"], df_5min["PCR"], marker='o')
    ax.set_xlabel("Time")
    ax.set_ylabel("PCR")
    ax.set_title("PCR Trend Over Time")
    plt.xticks(rotation=45)
    st.pyplot(fig)

# Display 15-min log
st.subheader("15-Minute Log")
df_15min = pd.DataFrame(st.session_state.fifteen_min_log)
st.dataframe(df_15min)

# Display Option Chain OI Analysis
st.subheader("Nifty Option Chain OI Analysis")
if "option_chain_data" in locals():
    df_option_chain = process_option_chain(option_chain_data)
    st.dataframe(df_option_chain)
