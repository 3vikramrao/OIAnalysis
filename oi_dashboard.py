import streamlit as st
import pandas as pd
from datetime import datetime
import requests
import matplotlib.pyplot as plt
from openpyxl import Workbook

# Initialize session state for logs
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []
if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []
if "last_ltp" not in st.session_state:
    st.session_state.last_ltp = 0

# Setup session with cookie handling
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0",
    "Accept": "*/*",
    "Referer": "https://www.nseindia.com"
})
session.get("https://www.nseindia.com")  # Establish session

# Function to fetch NIFTY LTP using NSE market status API
def fetch_nifty_ltp():
    url = "https://www.nseindia.com/api/marketStatus"
    response = session.get(url)
    data = response.json()
    for index in data["marketState"]:
        if index["index"] == "NIFTY 50":
            return float(index["last"])
    return None

# Function to fetch option chain data
def fetch_option_chain():
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
    response = session.get(url)
    data = response.json()
    return data["records"]["data"]

# Function to interpret sentiment
def interpret_sentiment(ltp, ce_oi_change, pe_oi_change):
    last_ltp = st.session_state.last_ltp
    if ce_oi_change > pe_oi_change and ltp < last_ltp:
        return "Short Buildup", "Sell"
    elif ce_oi_change < pe_oi_change and ltp > last_ltp:
        return "Long Buildup", "Buy"
    elif ce_oi_change < pe_oi_change and ltp < last_ltp:
        return "Long Unwinding", "Sell"
    elif ce_oi_change > pe_oi_change and ltp > last_ltp:
        return "Short Covering", "Buy"
    else:
        return "Neutral", "Hold"

# Function to plot OI vs Strike Price
def plot_oi_vs_ltp(option_data):
    ce_data = [(item["strikePrice"], item["CE"]["openInterest"])
               for item in option_data if "CE" in item and "openInterest" in item["CE"]]
    pe_data = [(item["strikePrice"], item["PE"]["openInterest"])
               for item in option_data if "PE" in item and "openInterest" in item["PE"]]

    ce_df = pd.DataFrame(ce_data, columns=["Strike", "CE OI"])
    pe_df = pd.DataFrame(pe_data, columns=["Strike", "PE OI"])

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(ce_df["Strike"], ce_df["CE OI"], label="Call OI", color="blue")
    ax.plot(pe_df["Strike"], pe_df["PE OI"], label="Put OI", color="red")
    ax.set_xlabel("Strike Price")
    ax.set_ylabel("Open Interest")
    ax.set_title("OI vs Strike Price")
    ax.legend()
    st.pyplot(fig)

# Streamlit UI
st.title("NIFTY OI Analysis Dashboard")
st.sidebar.header("Controls")

if st.sidebar.button("Fetch Latest Data"):
    try:
        ltp = fetch_nifty_ltp()
        option_data = fetch_option_chain()

        ce_oi_total = sum(item.get("CE", {}).get("openInterest", 0) for item in option_data if "CE" in item)
        pe_oi_total = sum(item.get("PE", {}).get("openInterest", 0) for item in option_data if "PE" in item)

        ce_oi_change = sum(item.get("CE", {}).get("changeinOpenInterest", 0) for item in option_data if "CE" in item)
        pe_oi_change = sum(item.get("PE", {}).get("changeinOpenInterest", 0) for item in option_data if "PE" in item)

        sentiment, signal = interpret_sentiment(ltp, ce_oi_change, pe_oi_change)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = {
            "Timestamp": timestamp,
            "LTP": ltp,
            "CE OI": ce_oi_total,
            "PE OI": pe_oi_total,
            "CE OI Change": ce_oi_change,
            "PE OI Change": pe_oi_change,
            "Sentiment": sentiment,
            "Signal": signal
        }

        st.session_state.five_min_log.append(log_entry)
        if len(st.session_state.five_min_log) % 3 == 0:
            st.session_state.fifteen_min_log.append(log_entry)

        st.session_state.last_ltp = ltp
        st.success("Data fetched and logged successfully.")

    except Exception as e:
        st.error(f"Error fetching data: {e}")

# Display logs
st.subheader("5-Minute Log")
df_5min = pd.DataFrame(st.session_state.five_min_log)
st.dataframe(df_5min)

st.subheader("15-Minute Log")
df_15min = pd.DataFrame(st.session_state.fifteen_min_log)
st.dataframe(df_15min)

# Display Option Chain OI Analysis
st.subheader("Nifty Option Chain OI vs Strike Price")
try:
    option_data = fetch_option_chain()
    plot_oi_vs_ltp(option_data)
except:
    st.warning("Unable to fetch option chain data for plotting.")

# Excel export
def export_to_excel():
    with pd.ExcelWriter("OIAnalysisDashboard.xlsx", engine="openpyxl") as writer:
        df_5min[["LTP", "CE OI", "PE OI", "CE OI Change", "PE OI Change", "Sentiment", "Signal"]].to_excel(writer, sheet_name="FiveMin", index=False)
        df_15min[["LTP", "CE OI", "PE OI", "CE OI Change", "PE OI Change", "Sentiment", "Signal"]].to_excel(writer, sheet_name="FifteenMin", index=False)

if st.sidebar.button("Download Excel"):
    export_to_excel()
    st.success("Excel file 'OIAnalysisDashboard.xlsx' has been created.")

