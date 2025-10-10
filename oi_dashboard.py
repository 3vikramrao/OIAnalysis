import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# Sidebar input for custom auto-refresh interval (in minutes)
refresh_interval_min = st.sidebar.number_input("Auto-refresh interval (minutes)", min_value=1, value=5)
refresh_interval_sec = refresh_interval_min * 60

# Initialize session state
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []
if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []
if "last_ltp" not in st.session_state:
    st.session_state.last_ltp = 0

# Function to fetch NIFTY LTP
def fetch_nifty_ltp(session):
    url = "https://www.nseindia.com/api/marketStatus"
    response = session.get(url)
    data = response.json()
    for index in data["marketState"]:
        if index["index"] == "NIFTY 50":
            return float(index["last"])
    return None

# Function to fetch option chain data
def fetch_option_chain(session):
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

# Streamlit UI
st.title("NIFTY OI Analysis Dashboard")
st.sidebar.header("Controls")

# Setup session with headers
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0",
    "Accept": "*/*",
    "Referer": "https://www.nseindia.com"
})
session.get("https://www.nseindia.com")

# Button to fetch latest data
if st.sidebar.button("Fetch Latest Data"):
    try:
        ltp = fetch_nifty_ltp(session)
        option_data = fetch_option_chain(session)

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

        # Display Option Chain OI Analysis Table
        st.subheader("Nifty Option Chain OI Analysis")
        table_data = []
        for item in option_data:
            strike = item.get("strikePrice")
            ce = item.get("CE", {})
            pe = item.get("PE", {})
            table_data.append({
                "Strike Price": strike,
                "CE OI": ce.get("openInterest"),
                "CE Change OI": ce.get("changeinOpenInterest"),
                "CE LTP": ce.get("lastPrice"),
                "PE OI": pe.get("openInterest"),
                "PE Change OI": pe.get("changeinOpenInterest"),
                "PE LTP": pe.get("lastPrice")
		"Sentiment": sentiment,
            	"Signal": signal
            })
        df_option_chain = pd.DataFrame(table_data)
        st.dataframe(df_option_chain)

    except Exception as e:
        st.error(f"Error fetching data: {e}")

# Display logs
st.subheader("5-Minute Log")
df_5min = pd.DataFrame(st.session_state.five_min_log)
st.dataframe(df_5min)

st.subheader("15-Minute Log")
df_15min = pd.DataFrame(st.session_state.fifteen_min_log)
st.dataframe(df_15min)

# Excel export
def export_to_excel():
    with pd.ExcelWriter("OIAnalysisDashboard.xlsx", engine="openpyxl") as writer:
        df_5min[["LTP", "CE OI", "PE OI", "CE OI Change", "PE OI Change", "Sentiment", "Signal"]].to_excel(writer, sheet_name="FiveMin", index=False)
        df_15min[["LTP", "CE OI", "PE OI", "CE OI Change", "PE OI Change", "Sentiment", "Signal"]].to_excel(writer, sheet_name="FifteenMin", index=False)
        if 'df_option_chain' in locals():
            df_option_chain.to_excel(writer, sheet_name="OptionChain", index=False)

if st.sidebar.button("Download Excel"):
    export_to_excel()
    st.success("Excel file 'OIAnalysisDashboard.xlsx' has been created.")
