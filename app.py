
import io
import time
import requests
from openpyxl import load_workbook
import streamlit as st
from streamlit_autorefresh import st_autorefresh

st.set_page_config(page_title="HSE XLS Monitor", layout="centered")

DEFAULT_URL = "https://priem44.hse.ru/ABITREPORTS/MAGREPORTS/EnrollmentList/28367398628_Commercial.xlsx"

st.title("Монитор ВШЭ: АБД")
st.caption("Kоличество контрактов и оплаченных договоров. Считает 'Да' в соответствующих полях файла.")

# url = st.text_input("XLS(X) URL", value=DEFAULT_URL)
url = DEFAULT_URL
reg_input = st.text_input("Регистрационный номер для поиска ранга", value="", placeholder="Например: 12345678")

# Auto-refresh every hour
st_autorefresh(interval=60 * 60 * 1000, key="hourly_refresh")

@st.cache_data(ttl=60*60, show_spinner=True)
def fetch_and_parse(u: str):
    r = requests.get(u, timeout=60)
    r.raise_for_status()
    bio = io.BytesIO(r.content)
    wb = load_workbook(bio, data_only=True)
    ws = wb.active  # first sheet

    def is_yes(val):
        if val is None:
            return False
        return str(val).strip().lower() == "да"

    a20 = ws["A20"].value

    rows = []  # collect minimal info from rows 22..500
    contracts = 0
    paid = 0
    for row in range(22, 501):
        reg = ws[f"B{row}"].value
        has_contract = is_yes(ws[f"H{row}"].value)
        has_paid = is_yes(ws[f"I{row}"].value)
        if has_contract:
            contracts += 1
        if has_paid:
            paid += 1
        rows.append({
            "row": row,
            "reg": str(reg).strip() if reg is not None else None,
            "contract": has_contract,
            "paid": has_paid
        })

    # Precompute the order of rows that have 'Да' for contract/paid by sheet order
    contract_rows = [r for r in rows if r["contract"]]
    paid_rows = [r for r in rows if r["paid"]]

    return {
        "a20": a20,
        "contracts": contracts,
        "paid": paid,
        "ts": int(time.time()),
        "rows": rows,
        "contract_rows": contract_rows,
        "paid_rows": paid_rows
    }

def compute_rank_by_row(rows_yes, target_row):
    """Return 1-based rank among rows_yes according to sheet order.
    If target_row is below all rows_yes, returns len(rows_yes)+1.
    """
    # rows_yes is already ordered by sheet row ascending
    rank = 1
    for r in rows_yes:
        if r["row"] < target_row:
            rank += 1
        else:
            break
    return rank

try:
    data = fetch_and_parse(url)
    c1, c2, c3 = st.columns(3)
    c1.metric("Заключенных договоров", data["contracts"])
    c2.metric("Оплаченных договоров", data["paid"])
    # c3.metric("Последнее обновление", time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(data["ts"])))

    st.write("**A20**:", data["a20"] or "—")

    st.caption("Auto-refresh hourly. Cached results refresh hourly or when URL changes.")

    if reg_input.strip():
        reg = reg_input.strip()
        all_rows = data["rows"]
        found = next((r for r in all_rows if r["reg"] == reg), None)
        if not found:
            st.error("Регистрационный номер не найден в файле.")
        else:
            # factual ranks among 'Да' lists by sheet order
            fact_contract_rank = None
            fact_paid_rank = None
            if found["contract"]:
                # rank is index in contract_rows by row order
                fact_contract_rank = next((i+1 for i, r in enumerate(data["contract_rows"]) if r["row"] == found["row"]), None)
            if found["paid"]:
                fact_paid_rank = next((i+1 for i, r in enumerate(data["paid_rows"]) if r["row"] == found["row"]), None)

            # positional ranks even if candidate is not 'Да' (where they'd stand)
            pos_contract_rank = compute_rank_by_row(data["contract_rows"], found["row"])
            pos_paid_rank = compute_rank_by_row(data["paid_rows"], found["row"])

            c1, c2 = st.columns(2)
            c1.metric("Место среди договоров", fact_contract_rank if fact_contract_rank is not None else pos_contract_rank)
            c2.metric("Место среди оплат", fact_paid_rank if fact_paid_rank is not None else pos_paid_rank)

            note = []
            if fact_contract_rank is None:
                note.append("абитуриент пока **не в списке заключивших договор**; показана позиция по месту в общем рейтинге")
            if fact_paid_rank is None:
                note.append("абитуриент пока **не в списке оплативших**; показана позиция по месту в общем рейтинге")
            if note:
                st.caption(" ; ".join(note))


    st.caption("Кэшированные результаты обновляются каждый час.")

except Exception as e:
    st.error(f"Error: {e}")
    st.stop()

# with st.expander("Raw debug"):
#     st.code(f"URL: {url}\nA20: {data['a20']}\nContracts: {data['contracts']}\nPaid: {data['paid']}\nFetched at: {data['ts']}", language="text")

