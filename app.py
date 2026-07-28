"""
קופה — ניהול הכנסות וחייבים (Streamlit + Google Sheets)
=========================================================
אפליקציית קופה לחנות יין שמתחברת ישירות לקובץ wine_inventory ב-Google Sheets.
כל הנתונים נשמרים בגיליון עצמו — לא תלוי ב-Claude, נגיש מכל מכשיר/דפדפן.
"""

import uuid
from datetime import datetime

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------------------
# הגדרות בסיס
# ----------------------------------------------------------------------------
SHEET_ID = "1c0LCm5BGPyEfU5q48NgiRqYHIw7iDqIi8l0rXejbaTE"  # wine_inventory
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

INVENTORY_SHEET = "גיליון1"
SALES_LOG_SHEET = "sales_log"
CUSTOMERS_SHEET = "customers"
PAYMENTS_SHEET = "payments_log"

INV_COLS = ["יקב", "שם היין", "שנה", "מחיר קניה", "מחיר מכירה",
            "מלאי", "נמכר", "נמכר החודש", "חודש מעקב", "סוג", "מחיר פתוח"]
CUSTOMERS_COLS = ["מזהה", "שם", "טלפון", "יתרת חוב"]
PAYMENTS_COLS = ["תאריך", "לקוח", "סכום"]
SALES_LOG_EXTRA_COLS = ["אופן תשלום", "לקוח"]

st.set_page_config(page_title="קופה — ניהול הכנסות וחייבים", page_icon="🍷", layout="centered")

st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
div[data-testid="stMetricValue"] { direction: ltr; text-align: right; }
.stTabs [data-baseweb="tab-list"] { direction: rtl; }
thead tr th { text-align: right !important; }
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# חיבור ל-Google Sheets
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_client():
    info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    return get_client().open_by_key(SHEET_ID)


def get_or_create_ws(sh, name, headers):
    try:
        ws = sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=200, cols=max(10, len(headers)))
        ws.append_row(headers)
    existing_headers = ws.row_values(1)
    if not existing_headers:
        ws.append_row(headers)
    return ws


def ensure_extra_columns(ws, extra_cols):
    """מוסיף עמודות חדשות בסוף אם עדיין לא קיימות (בלי לפגוע בעמודות הקיימות)."""
    headers = ws.row_values(1)
    for col in extra_cols:
        if col not in headers:
            ws.update_cell(1, len(headers) + 1, col)
            headers = ws.row_values(1)
    return headers


def this_month_label():
    return datetime.now().strftime("%m-%Y")


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ----------------------------------------------------------------------------
# מלאי (גיליון1)
# ----------------------------------------------------------------------------
def load_inventory(sh):
    ws = sh.worksheet(INVENTORY_SHEET)
    ensure_extra_columns(ws, ["מחיר פתוח"])
    values = ws.get_all_values()
    headers = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    df["_row"] = range(2, 2 + len(df))

    name_col = "שם היין"
    df = df[df[name_col].astype(str).str.strip() != ""].copy()

    for col in ["מחיר מכירה", "מלאי", "נמכר", "נמכר החודש"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    if "מחיר פתוח" not in df.columns:
        df["מחיר פתוח"] = ""
    df["is_open_price"] = df["מחיר פתוח"].astype(str).str.strip().isin(
        ["TRUE", "True", "true", "1", "כן", "✓"]
    )
    df["display_name"] = (df["יקב"].astype(str).str.strip() + " — " +
                           df["שם היין"].astype(str).str.strip() + " " +
                           df["שנה"].astype(str).str.strip()).str.strip()
    return df.reset_index(drop=True), ws


def add_inventory_row(ws, producer, name, year, buy_price, sell_price, stock, wtype, open_price):
    ws.append_row([
        producer, name, year,
        buy_price if buy_price else "",
        0 if open_price else sell_price,
        stock, 0, 0, "", wtype,
        "TRUE" if open_price else "",
    ])


def record_sale_in_inventory(ws, row_num, new_stock, new_sold, new_sold_month):
    ws.update(f"F{row_num}:I{row_num}", [[new_stock, new_sold, new_sold_month, this_month_label()]])


def append_sales_log(sh, wine_name, wine_type, unit_price, qty, method, customer_name):
    ws = get_or_create_ws(sh, SALES_LOG_SHEET,
                           ["תאריך", "שם היין", "סוג", "מחיר", "כמות"] + SALES_LOG_EXTRA_COLS)
    ensure_extra_columns(ws, SALES_LOG_EXTRA_COLS)
    ws.append_row([now_str(), wine_name, wine_type, unit_price, qty,
                   "בהקפה" if method == "credit" else "מזומן/אשראי",
                   customer_name or ""])


# ----------------------------------------------------------------------------
# חייבים / לקוחות
# ----------------------------------------------------------------------------
def load_customers(sh):
    ws = get_or_create_ws(sh, CUSTOMERS_SHEET, CUSTOMERS_COLS)
    values = ws.get_all_values()
    if len(values) <= 1:
        return pd.DataFrame(columns=CUSTOMERS_COLS + ["_row"]), ws
    df = pd.DataFrame(values[1:], columns=values[0])
    df["_row"] = range(2, 2 + len(df))
    df["יתרת חוב"] = pd.to_numeric(df["יתרת חוב"], errors="coerce").fillna(0)
    return df.reset_index(drop=True), ws


def add_customer(ws, name, phone):
    cid = str(uuid.uuid4())[:8]
    ws.append_row([cid, name, phone, 0])
    return cid


def adjust_customer_balance(ws, row_num, delta):
    cell = ws.cell(row_num, 4)
    current = float(cell.value or 0)
    new_balance = max(0, current + delta)
    ws.update_cell(row_num, 4, new_balance)
    return new_balance


def append_payment(sh, customer_name, amount):
    ws = get_or_create_ws(sh, PAYMENTS_SHEET, PAYMENTS_COLS)
    ws.append_row([now_str(), customer_name, amount])


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.title("🍷 קופה — ניהול הכנסות וחייבים")
st.caption("מחובר ישירות לקובץ wine_inventory ב-Google Drive — נגיש מכל מכשיר, בלי תלות ב-Claude.")

try:
    sh = get_spreadsheet()
except Exception as e:
    st.error("שגיאה בהתחברות ל-Google Sheets. ודא שהגדרת נכון את ה-secrets ושיתפת את הגיליון עם חשבון השירות.")
    st.exception(e)
    st.stop()

if st.button("🔄 רענן נתונים מהגיליון"):
    st.cache_resource.clear()
    st.rerun()

inv_df, inv_ws = load_inventory(sh)
cust_df, cust_ws = load_customers(sh)

tab_register, tab_inventory, tab_debtors, tab_reports = st.tabs(
    ["🧾 קופה", "📦 מלאי", "📒 חייבים", "📊 דוחות"]
)

# ---- קופה ----
with tab_register:
    st.subheader("מכירה חדשה")

    if inv_df.empty:
        st.info("אין עדיין פריטים במלאי — עבור ללשונית 'מלאי' והוסף פריט ראשון.")
    else:
        options = ["בחר פריט…"] + inv_df["display_name"].tolist()
        choice = st.selectbox("פריט", options, key="sale_item")

        if choice != "בחר פריט…":
            row = inv_df[inv_df["display_name"] == choice].iloc[0]
            qty = st.number_input("כמות", min_value=1, value=1, step=1, key="sale_qty")

            if row["is_open_price"]:
                unit_price = st.number_input(
                    "מחיר מכירה לבקבוק הזה (₪) — נקבע כעת", min_value=0.0, step=1.0, key="open_price_input"
                )
            else:
                unit_price = float(row["מחיר מכירה"])
                st.write(f"מחיר ליחידה: **₪{unit_price:,.0f}**  |  מלאי זמין: **{int(row['מלאי'])}**")

            method = st.radio("אופן תשלום", ["מזומן / אשראי", "בהקפה (חוב)"], horizontal=True, key="sale_method")
            method_key = "credit" if method == "בהקפה (חוב)" else "cash"

            customer_name = None
            customer_row_num = None
            if method_key == "credit":
                if cust_df.empty:
                    st.warning("אין עדיין לקוחות רשומים — הוסף לקוח חדש למטה.")
                else:
                    cust_choice = st.selectbox("לקוח", ["בחר לקוח…"] + cust_df["שם"].tolist(), key="sale_customer")
                    if cust_choice != "בחר לקוח…":
                        crow = cust_df[cust_df["שם"] == cust_choice].iloc[0]
                        customer_name = crow["שם"]
                        customer_row_num = int(crow["_row"])

                with st.expander("+ הוסף לקוח חדש"):
                    new_cust_name = st.text_input("שם הלקוח", key="new_cust_name_reg")
                    new_cust_phone = st.text_input("טלפון (לא חובה)", key="new_cust_phone_reg")
                    if st.button("הוסף לקוח", key="add_cust_reg"):
                        if new_cust_name.strip():
                            add_customer(cust_ws, new_cust_name.strip(), new_cust_phone.strip())
                            st.success(f"נוסף לקוח: {new_cust_name}")
                            st.rerun()
                        else:
                            st.warning("הזן שם לקוח")

            if st.button("💾 רשום מכירה", type="primary", use_container_width=True):
                if row["is_open_price"] and (not unit_price or unit_price <= 0):
                    st.error("הזן מחיר מכירה לפריט הזה")
                elif method_key == "credit" and not customer_name:
                    st.error("בחר לקוח להקפה, או הוסף לקוח חדש")
                else:
                    line_total = unit_price * qty
                    new_stock = max(0, int(row["מלאי"]) - qty)
                    same_month = row["חודש מעקב"] == this_month_label()
                    new_sold_month = (int(row["נמכר החודש"]) + qty) if same_month else qty
                    new_sold = int(row["נמכר"]) + qty

                    record_sale_in_inventory(inv_ws, int(row["_row"]), new_stock, new_sold, new_sold_month)
                    append_sales_log(sh, row["שם היין"], row["סוג"], unit_price, qty, method_key, customer_name)

                    if method_key == "credit" and customer_row_num:
                        adjust_customer_balance(cust_ws, customer_row_num, line_total)

                    st.success(f"נרשם: {row['display_name']} — ₪{line_total:,.0f}")
                    st.cache_resource.clear()
                    st.rerun()

# ---- מלאי ----
with tab_inventory:
    st.subheader("מלאי ומחירים")
    if not inv_df.empty:
        show_cols = ["display_name", "מחיר מכירה", "מלאי", "is_open_price", "סוג"]
        pretty = inv_df[show_cols].rename(columns={
            "display_name": "יין", "מחיר מכירה": "מחיר (₪)",
            "מלאי": "מלאי", "is_open_price": "מחיר פתוח", "סוג": "סוג"
        })
        st.dataframe(pretty, use_container_width=True, hide_index=True)
    else:
        st.info("המלאי ריק כרגע.")

    with st.expander("+ הוסף פריט חדש"):
        c1, c2 = st.columns(2)
        producer = c1.text_input("יקב / יצרן")
        wine_name = c2.text_input("שם היין")
        c3, c4 = st.columns(2)
        year = c3.text_input("שנה")
        wtype = c4.text_input("סוג (למשל: אדום יבש)")
        open_price = st.checkbox("מחיר משתנה — ייקבע בזמן המכירה")
        c5, c6, c7 = st.columns(3)
        buy_price = c5.number_input("מחיר קניה (₪)", min_value=0.0, step=1.0)
        sell_price = c6.number_input("מחיר מכירה (₪)", min_value=0.0, step=1.0, disabled=open_price)
        stock = c7.number_input("מלאי התחלתי", min_value=0, step=1)
        if st.button("הוסף למלאי"):
            if not wine_name.strip():
                st.warning("הזן שם יין")
            elif not open_price and sell_price <= 0:
                st.warning("הזן מחיר מכירה, או סמן 'מחיר משתנה'")
            else:
                add_inventory_row(inv_ws, producer.strip(), wine_name.strip(), year.strip(),
                                   buy_price, sell_price, stock, wtype.strip(), open_price)
                st.success(f"נוסף: {wine_name}")
                st.cache_resource.clear()
                st.rerun()

# ---- חייבים ----
with tab_debtors:
    st.subheader("חייבים / הקפה")
    total_debt = cust_df["יתרת חוב"].sum() if not cust_df.empty else 0
    st.metric("סה\"כ חובות פתוחים", f"₪{total_debt:,.0f}")

    if not cust_df.empty:
        debtors = cust_df[cust_df["יתרת חוב"] > 0].sort_values("יתרת חוב", ascending=False)
        if not debtors.empty:
            for _, c in debtors.iterrows():
                col1, col2, col3 = st.columns([3, 2, 2])
                col1.write(f"**{c['שם']}**  \n{c['טלפון'] or 'ללא טלפון'}")
                col2.write(f"חייב ₪{c['יתרת חוב']:,.0f}")
                with col3:
                    with st.popover("קבל תשלום"):
                        amt = st.number_input("סכום שהתקבל (₪)", min_value=0.0,
                                               value=float(c["יתרת חוב"]), key=f"pay_{c['_row']}")
                        if st.button("רשום תשלום", key=f"paybtn_{c['_row']}"):
                            new_bal = adjust_customer_balance(cust_ws, int(c["_row"]), -amt)
                            append_payment(sh, c["שם"], amt)
                            st.success(f"נרשם תשלום ₪{amt:,.0f}")
                            st.cache_resource.clear()
                            st.rerun()
        else:
            st.write("כל הלקוחות מסודרים 🎉")
    else:
        st.info("אין עדיין לקוחות רשומים.")

    with st.expander("+ הוסף לקוח חדש"):
        cname = st.text_input("שם הלקוח", key="new_cust_name_debtors")
        cphone = st.text_input("טלפון (לא חובה)", key="new_cust_phone_debtors")
        if st.button("הוסף לקוח", key="add_cust_debtors"):
            if cname.strip():
                add_customer(cust_ws, cname.strip(), cphone.strip())
                st.success(f"נוסף לקוח: {cname}")
                st.cache_resource.clear()
                st.rerun()
            else:
                st.warning("הזן שם לקוח")

# ---- דוחות ----
with tab_reports:
    st.subheader("סיכום")
    try:
        log_ws = sh.worksheet(SALES_LOG_SHEET)
        log_values = log_ws.get_all_values()
        if len(log_values) > 1:
            log_df = pd.DataFrame(log_values[1:], columns=log_values[0])
            log_df["תאריך_dt"] = pd.to_datetime(log_df["תאריך"], errors="coerce")
            log_df["מחיר"] = pd.to_numeric(log_df["מחיר"], errors="coerce").fillna(0)
            log_df["כמות"] = pd.to_numeric(log_df["כמות"], errors="coerce").fillna(0)
            log_df["סה\"כ"] = log_df["מחיר"] * log_df["כמות"]

            today = pd.Timestamp.now().normalize()
            week_ago = today - pd.Timedelta(days=7)
            month_start = today.replace(day=1)

            today_rev = log_df[log_df["תאריך_dt"] >= today]["סה\"כ"].sum()
            week_rev = log_df[log_df["תאריך_dt"] >= week_ago]["סה\"כ"].sum()
            month_rev = log_df[log_df["תאריך_dt"] >= month_start]["סה\"כ"].sum()
        else:
            today_rev = week_rev = month_rev = 0
    except gspread.exceptions.WorksheetNotFound:
        today_rev = week_rev = month_rev = 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("הכנסות היום", f"₪{today_rev:,.0f}")
    c2.metric("7 ימים אחרונים", f"₪{week_rev:,.0f}")
    c3.metric("החודש", f"₪{month_rev:,.0f}")
    c4.metric("חובות פתוחים", f"₪{total_debt:,.0f}")

    st.subheader("תנועות אחרונות")
    try:
        log_ws = sh.worksheet(SALES_LOG_SHEET)
        log_values = log_ws.get_all_values()
        if len(log_values) > 1:
            recent = pd.DataFrame(log_values[1:], columns=log_values[0]).tail(25).iloc[::-1]
            st.dataframe(recent, use_container_width=True, hide_index=True)
        else:
            st.info("אין עדיין תנועות.")
    except gspread.exceptions.WorksheetNotFound:
        st.info("אין עדיין תנועות.")
