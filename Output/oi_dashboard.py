import streamlit as st
import pandas as pd
import requests
import datetime
from io import BytesIO

# Set page config
st.set_page_config(page_title="OI Analysis Dashboard", layout="wide")

# Title
st.title("📈 OI Analysis Dashboard for NIFTY")

# Initialize session state for logs
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []
if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []

# Function to fetch NIFTY futures data
def fetch_futures_data():
    url = "https://www.nseindia.com/api/quote-derivative?symbol=NIFTY"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br"
    }
    try:
        response = requests.get(url, headers=headers)
        data = response.json()
        fut_data = data['stocks'][0]['metadata']
        return {
            "LTP": float(fut_data['lastPrice']),
            "OI": int(fut_data['openInterest']),
            "Volume": int(fut_data['numberOfContractsTraded'])
        }
    except:
        return None

# Function to fetch option chain data
def fetch_option_chain():
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br"
    }
    try:
        response = requests.get(url, headers=headers)
        data = response.json()
        records = data['records']['data']
        option_data = []
        for record in records:
            strike = record.get('strikePrice')
            ce_oi = record.get('CE', {}).get('openInterest', 0)
            pe_oi = record.get('PE', {}).get('openInterest', 0)
            option_data.append({
                "Strike Price": strike,
                "Call OI": ce_oi,
                "Put OI": pe_oi
            })
        return pd.DataFrame(option_data)
    except:
        return pd.DataFrame()

# Function to interpret OI changes
def interpret_oi(current, previous):
    if not previous:
        return "N/A", "N/A"
    ltp_change = current["LTP"] - previous["LTP"]
    oi_change = current["OI"] - previous["OI"]
    if ltp_change > 0 and oi_change > 0:
        return "Long Buildup", "Buy"
    elif ltp_change < 0 and oi_change > 0:
        return "Short Buildup", "Sell"
    elif ltp_change < 0 and oi_change < 0:
        return "Long Unwinding", "Sell"
    elif ltp_change > 0 and oi_change < 0:
        return "Short Covering", "Buy"
    else:
        return "Neutral", "Hold"

# Fetch data
st.sidebar.header("🔄 Refresh Data")
if st.sidebar.button("Fetch Latest Data"):
    current_time = datetime.datetime.now().strftime("%H:%M:%S")
    futures_data = fetch_futures_data()
    option_chain_df = fetch_option_chain()

    if futures_data:
        previous_data = st.session_state.five_min_log[-1] if st.session_state.five_min_log else None
        interpretation, signal = interpret_oi(futures_data, previous_data)

        log_entry = {
            "Time": current_time,
            "LTP": futures_data["LTP"],
            "OI": futures_data["OI"],
            "Volume": futures_data["Volume"],
            "Interpretation": interpretation,
            "Signal": signal
        }

        st.session_state.five_min_log.append(log_entry)
        if len(st.session_state.five_min_log) % 3 == 0:
            st.session_state.fifteen_min_log.append(log_entry)

# Display logs
st.subheader("📄 5-Minute Log")
df_5min = pd.DataFrame(st.session_state.five_min_log)
st.dataframe(df_5min)

st.subheader("📄 15-Minute Log")
df_15min = pd.DataFrame(st.session_state.fifteen_min_log)
st.dataframe(df_15min)

# Display option chain
st.subheader("📊 Option Chain Data")
if 'option_chain_df' in locals():
    st.dataframe(option_chain_df)

# Export to Excel
def export_to_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_5min.to_excel(writer, sheet_name="FiveMin", index=False)
        df_15min.to_excel(writer, sheet_name="FifteenMin", index=False)
        if 'option_chain_df' in locals():
            option_chain_df.to_excel(writer, sheet_name="OptionChain", index=False)
    output.seek(0)
    return output

st.sidebar.header("📁 Export Data")
if st.sidebar.button("Download Excel"):
    excel_data = export_to_excel()
    st.sidebar.download_button(
        label="Download OIAnalysisDashboard.xlsx",
        data=excel_data,
        file_name="OIAnalysisDashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
