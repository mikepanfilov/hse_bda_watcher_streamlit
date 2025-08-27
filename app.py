
import io
import time
import requests
from openpyxl import load_workbook
import streamlit as st

st.set_page_config(page_title="HSE XLS Monitor", layout="centered")

DEFAULT_URL = "https://priem44.hse.ru/ABITREPORTS/MAGREPORTS/EnrollmentList/28367398628_Commercial.xlsx"

st.title("HSE XLS Monitor")
st.caption("Parses A20 for update time, counts 'Да' in H22:H500 (contracts) and I22:I500 (paid).")

url = st.text_input("XLS(X) URL", value=DEFAULT_URL)

# Auto-refresh every hour
st_autorefresh = st.experimental_rerun  # backward compat shim if needed
st_autorefresh = st.autorefresh if hasattr(st, "autorefresh") else None
if st_autorefresh:
    st_autorefresh(interval=60 * 60 * 1000, key="hourly_refresh")

@st.cache_data(ttl=60*60, show_spinner=True)
def fetch_and_parse(u: str):
    # Fetch
    r = requests.get(u, timeout=60)
    r.raise_for_status()
    bio = io.BytesIO(r.content)
    # Parse with openpyxl for exact cells
    wb = load_workbook(bio, data_only=True)
    ws = wb.active  # first sheet

    def is_yes(val):
        if val is None:
            return False
        return str(val).strip().lower() == "да"

    # A20 date/time string
    a20 = ws["A20"].value

    # Counts
    contracts = 0
    paid = 0
    for row in range(22, 501):
        if is_yes(ws[f"H{row}"].value):
            contracts += 1
        if is_yes(ws[f"I{row}"].value):
            paid += 1

    return {
        "a20": a20,
        "contracts": contracts,
        "paid": paid,
        "ts": int(time.time())
    }

try:
    data = fetch_and_parse(url)
    c1, c2, c3 = st.columns(3)
    c1.metric("Contracts (H22:H500 == 'Да')", data["contracts"])
    c2.metric("Paid (I22:I500 == 'Да')", data["paid"])
    c3.metric("Last check (UTC)", time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(data["ts"])))

    st.write("**A20**:", data["a20"] or "—")
    st.caption("Next auto-refresh in ~1 hour. Cached results refresh hourly or when URL changes.")
except Exception as e:
    st.error(f"Error: {e}")
    st.stop()

with st.expander("Raw debug"):
    st.code(f"URL: {url}\nA20: {data['a20']}\nContracts: {data['contracts']}\nPaid: {data['paid']}\nFetched at: {data['ts']}", language="text")
