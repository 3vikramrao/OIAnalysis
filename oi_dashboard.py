import streamlit as st
import pandas as pd
from nsepython import nse_index, nse_optionchain_scrapper
from datetime import datetime, timedelta
import io

# Initialize session state for logs
if "five_min_log" not in st.session_state:
    st.session_state.five_min_log = []
if "fifteen_min_log" not in st.session_state:
    st.session_state.fifteen_min_log = []
if "last_fifteen_log_time" not in st.session_state:
    st.session_state.last_fifteen_log_time = None

st.title("📊 NIFTY Open Interest Dashboard")

# Sidebar controls
if st.sidebar.button("Fetch Latest Data"):
    try:
        # Fetch NIFTY index data
        index_data = nse_index("NIFTY 50")
        ltp = index_data.get("last", None)
        volume = index_data.get("volume", None)

        # Fetch option chain data
        oc_data = nse_optionchain_scrapper("NIFTY")
        records = oc_data.get("records", {}).get("data", [])

        # Process option chain data
        option_rows = []
        total_ce_oi = 0
        total_pe_oi = 0
        for item in records:
            strike = item.get("strikePrice")
            ce = item.get("CE", {})
            pe = item.get("PE", {})
            ce_oi = ce.get("openInterest", 0)
            pe_oi = pe.get("openInterest", 0)
            total_ce_oi += ce_oi
            total_pe_oi += pe_oi
            option_rows.append({
                "Strike": strike,
                "CE OI": ce_oi,
                "PE OI": pe_oi
            })

        # Create DataFrame
        option_df = pd.DataFrame(option_rows).sort_values("Strike")

        # Display index data
        st.subheader("📈 NIFTY Index Data")
        st.write(f"**LTP:** {ltp}")
        st.write(f"**Volume:** {volume}")
        st.write(f"**Total CE OI:** {total_ce_oi}")
        st.write(f"**Total PE OI:** {total_pe_oi}")

        # Display option chain
        st.subheader("🔗 Option Chain")
        st.dataframe(option_df)

        # Determine market sentiment
        sentiment = "Neutral"
        signal = "Hold"
        if total_ce_oi > total_pe_oi and ltp is not None:
            sentiment = "Bearish"
            signal = "Sell"
        elif total_pe_oi > total_ce_oi and ltp is not None:
            sentiment = "Bullish"
            signal = "Buy"

        st.subheader("📌 Market Sentiment")
        st.write(f"**Sentiment:** {sentiment}")
        st.write(f"**Signal:** {signal}")

        # Log data
        now = datetime.now()
        log_entry = {
            "Time": now.strftime("%H:%M:%S"),
            "LTP": ltp,
            "Volume": volume,
            "Total CE OI": total_ce_oi,
            "Total PE OI": total_pe_oi,
            "Sentiment": sentiment,
            "Signal": signal
        }
        st.session_state.five_min_log.append(log_entry)

        # Log every 15 minutes
        if (st.session_state.last_fifteen_log_time is None or
            now - st.session_state.last_fifteen_log_time >= timedelta(minutes=15)):
            st.session_state.fifteen_min_log.append(log_entry)
            st.session_state.last_fifteen_log_time = now

    except Exception as e:
        st.error(f"Error fetching data: {e}")

# Display logs
st.subheader("🕔 5-Minute Log")
st.dataframe(pd.DataFrame(st.session_state.five_min_log))

st.subheader("🕒 15-Minute Log")
st.dataframe(pd.DataFrame(st.session_state.fifteen_min_log))

# Export to Excel
if st.sidebar.button("Download Excel"):
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            pd.DataFrame(st.session_state.five_min_log).to_excel(writer, sheet_name="FiveMin", index=False)
            pd.DataFrame(st.session_state.fifteen_min_log).to_excel(writer, sheet_name="FifteenMin", index=False)
            if "option_df" in locals():
                option_df.to_excel(writer, sheet_name="OptionChain", index=False)
        st.download_button("Download OIAnalysisDashboard.xlsx", data=output.getvalue(), file_name="OIAnalysisDashboard.xlsx")
    except Exception as e:
        st.error(f"Error exporting Excel: {e}")
