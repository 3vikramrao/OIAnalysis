import streamlit as st
import pandas as pd
from nsepython import *
from datetime import datetime
import io

# Initialize session state for logs
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []
if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []

# Title
st.title("📊 NIFTY Open Interest Analysis Dashboard")

# Sidebar controls
st.sidebar.header("Controls")
fetch_data = st.sidebar.button("Fetch Latest Data")
export_excel = st.sidebar.button("Download Excel")

# Function to interpret OI changes
def interpret_oi_change(prev, curr):
    if not prev:
        return "N/A"
    ltp_change = curr["LTP"] - prev["LTP"]
    oi_change = curr["OI"] - prev["OI"]
    if ltp_change > 0 and oi_change > 0:
        return "Long Buildup"
    elif ltp_change < 0 and oi_change > 0:
        return "Short Buildup"
    elif ltp_change < 0 and oi_change < 0:
        return "Long Unwinding"
    elif ltp_change > 0 and oi_change < 0:
        return "Short Covering"
    else:
        return "Neutral"

# Function to fetch and process data
def fetch_nifty_data():
    try:
        eq_data = nse_eq("NIFTY")
        fut_data = nse_fno("NIFTY")
        oc_data = nse_optionchain_scrapper("NIFTY")

        ltp = float(eq_data["priceInfo"]["lastPrice"])
        oi = int(fut_data["marketDeptOrderBook"]["totalBuyQuantity"])
        vol = int(eq_data["priceInfo"]["quantityTraded"])

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        current_data = {"Timestamp": timestamp, "LTP": ltp, "OI": oi, "Volume": vol}

        # Interpret OI change
        prev_data = st.session_state.five_min_log[-1] if st.session_state.five_min_log else None
        signal = interpret_oi_change(prev_data, current_data)
        current_data["Signal"] = signal

        # Append to logs
        st.session_state.five_min_log.append(current_data)
        if len(st.session_state.five_min_log) % 3 == 0:
            st.session_state.fifteen_min_log.append(current_data)

        # Option Chain Table
        ce_data = []
        pe_data = []
        for item in oc_data["records"]["data"]:
            strike = item.get("strikePrice")
            ce_oi = item.get("CE", {}).get("openInterest", 0)
            pe_oi = item.get("PE", {}).get("openInterest", 0)
            ce_data.append({"Strike": strike, "CE OI": ce_oi})
            pe_data.append({"Strike": strike, "PE OI": pe_oi})
        option_chain_df = pd.DataFrame(ce_data).merge(pd.DataFrame(pe_data), on="Strike")

        return current_data, option_chain_df
    except Exception as e:
        st.error(f"Error fetching data: {e}")
        return None, None

# Fetch data on button click
if fetch_data:
    latest_data, option_chain = fetch_nifty_data()
    if latest_data:
        st.subheader("📈 Latest NIFTY Futures Data")
        st.write(pd.DataFrame([latest_data]))

        st.subheader("📊 Option Chain Data")
        st.dataframe(option_chain)

# Display logs
if st.session_state.five_min_log:
    st.subheader("🕔 5-Minute Log")
    st.dataframe(pd.DataFrame(st.session_state.five_min_log))

if st.session_state.fifteen_min_log:
    st.subheader("🕒 15-Minute Log")
    st.dataframe(pd.DataFrame(st.session_state.fifteen_min_log))

# Export to Excel
if export_excel:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(st.session_state.five_min_log).to_excel(writer, sheet_name="FiveMin", index=False)
        pd.DataFrame(st.session_state.fifteen_min_log).to_excel(writer, sheet_name="FifteenMin", index=False)
        if "option_chain" in locals():
            option_chain.to_excel(writer, sheet_name="OptionChain", index=False)
    st.download_button("📥 Download Excel File", data=output.getvalue(), file_name="OIAnalysisDashboard.xlsx")


