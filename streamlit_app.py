import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, date
import json

st.set_page_config(
    page_title="ระบบสั่งอาหารองค์กร",
    page_icon="🍱",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Sarabun', sans-serif !important; }

.main .block-container { max-width: 1100px; padding: 2rem 2rem 4rem; }

/* Header */
.app-header {
    display: flex; align-items: center; gap: 14px;
    padding: 1.25rem 1.5rem; border-radius: 14px;
    background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
    margin-bottom: 1.5rem; color: white;
}
.app-header h1 { font-size: 22px; font-weight: 700; margin: 0; }
.app-header p { font-size: 13px; margin: 0; opacity: 0.85; }
.header-icon { font-size: 36px; }

/* Metric cards */
.metric-row { display: flex; gap: 12px; margin-bottom: 1.25rem; flex-wrap: wrap; }
.metric-card {
    flex: 1; min-width: 130px;
    background: white; border-radius: 12px;
    padding: 14px 18px; border: 1px solid #e8eaed;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.metric-label { font-size: 12px; color: #5f6368; margin-bottom: 4px; }
.metric-value { font-size: 24px; font-weight: 700; color: #1a73e8; }

/* Menu card grid */
.menu-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(155px, 1fr)); gap: 12px; margin-bottom: 1.5rem; }
.menu-card {
    background: white; border-radius: 12px; padding: 14px 12px;
    border: 1.5px solid #e8eaed; cursor: pointer;
    transition: all 0.18s; text-align: center;
}
.menu-card:hover { border-color: #1a73e8; box-shadow: 0 2px 8px rgba(26,115,232,0.15); }
.menu-card.selected { border-color: #1a73e8; background: #e8f0fe; }
.menu-card-name { font-size: 14px; font-weight: 600; color: #202124; margin-bottom: 4px; }
.menu-card-price { font-size: 15px; font-weight: 700; color: #1a73e8; }
.menu-card-cat { font-size: 11px; color: #80868b; margin-top: 4px; }

/* Cart item */
.cart-item {
    display: flex; align-items: center; gap: 10px;
    padding: 10px 0; border-bottom: 1px solid #f1f3f4;
}
.cart-name { flex: 1; font-size: 14px; font-weight: 500; }
.cart-price { font-size: 14px; color: #5f6368; min-width: 70px; text-align: right; }

/* Status badges */
.badge { display: inline-block; font-size: 12px; padding: 3px 10px; border-radius: 20px; font-weight: 500; }
.badge-success { background: #e6f4ea; color: #137333; }
.badge-info { background: #e8f0fe; color: #1967d2; }
.badge-warn { background: #fef7e0; color: #b06000; }

/* Summary table */
.summary-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.summary-table th { padding: 9px 12px; background: #f8f9fa; text-align: left; font-weight: 600; color: #3c4043; border-bottom: 2px solid #e8eaed; }
.summary-table td { padding: 9px 12px; border-bottom: 1px solid #f1f3f4; color: #202124; }
.summary-table tr:hover td { background: #f8f9fa; }

/* Total bar */
.total-bar {
    display: flex; justify-content: space-between; align-items: center;
    padding: 14px 18px; background: #1a73e8; color: white;
    border-radius: 12px; margin-top: 1rem; font-size: 16px; font-weight: 600;
}

/* Section title */
.section-title { font-size: 16px; font-weight: 700; color: #202124; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }

/* Config box */
.config-box { background: #f8f9fa; border-radius: 12px; padding: 1.25rem; margin-bottom: 1rem; border: 1px solid #e8eaed; }
.config-note { font-size: 12px; color: #5f6368; line-height: 1.7; }
.config-note code { background: #e8eaed; padding: 1px 5px; border-radius: 4px; font-size: 11px; }

/* Tabs override */
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #f1f3f4; padding: 5px; border-radius: 10px; }
.stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 8px 20px; font-family: 'Sarabun', sans-serif !important; font-size: 14px; font-weight: 500; }
.stTabs [aria-selected="true"] { background: white !important; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }

/* Button */
.stButton > button { font-family: 'Sarabun', sans-serif !important; border-radius: 8px !important; font-size: 14px !important; font-weight: 500 !important; }
.stButton > button[kind="primary"] { background: #1a73e8 !important; border: none !important; }

/* Input */
.stTextInput > div > div > input, .stSelectbox > div > div { font-family: 'Sarabun', sans-serif !important; border-radius: 8px !important; }

/* Hide streamlit branding */
#MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ─── Google Sheets helper ────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

@st.cache_resource(show_spinner=False)
def get_gspread_client():
    """Build gspread client from st.secrets."""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        return None

def get_spreadsheet():
    client = get_gspread_client()
    if client is None:
        return None
    try:
        return client.open_by_key(st.secrets["spreadsheet_id"])
    except Exception:
        return None


@st.cache_data(ttl=120, show_spinner=False)
def load_menu(spreadsheet_id: str):
    """Load menu items from Sheet 'Menu'. Cache 2 min."""
    ss = get_spreadsheet()
    if ss is None:
        return []
    try:
        ws = ss.worksheet("Menu")
        rows = ws.get_all_values()[1:]  # skip header
        return [
            {"name": r[0], "price": float(r[1]) if len(r) > 1 and r[1] else 0,
             "cat": r[2] if len(r) > 2 else "", "note": r[3] if len(r) > 3 else ""}
            for r in rows if r and r[0].strip()
        ]
    except Exception as e:
        st.error(f"โหลดเมนูไม่สำเร็จ: {e}")
        return []


def append_orders(orders: list[dict]):
    """Write order rows to Sheet 'Orders'."""
    ss = get_spreadsheet()
    if ss is None:
        raise RuntimeError("ไม่สามารถเชื่อมต่อ Google Sheets ได้")
    ws = ss.worksheet("Orders")
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    rows = [[
        now,
        o["orderer"],
        o["name"],
        o["qty"],
        o["price"] * o["qty"],
        o.get("note", ""),
    ] for o in orders]
    ws.append_rows(rows, value_input_option="USER_ENTERED")


@st.cache_data(ttl=30, show_spinner=False)
def load_orders(spreadsheet_id: str, date_from: str, date_to: str):
    ss = get_spreadsheet()
    if ss is None:
        return pd.DataFrame()
    try:
        ws = ss.worksheet("Orders")
        data = ws.get_all_values()
        if len(data) < 2:
            return pd.DataFrame()
        df = pd.DataFrame(data[1:], columns=["วันที่", "ชื่อ/แผนก", "รายการ", "จำนวน", "ราคารวม", "หมายเหตุ"])
        df["ราคารวม"] = pd.to_numeric(df["ราคารวม"], errors="coerce").fillna(0)
        df["จำนวน"] = pd.to_numeric(df["จำนวน"], errors="coerce").fillna(0)
        return df[df["ชื่อ/แผนก"].str.strip() != ""]
    except Exception as e:
        st.error(f"โหลดข้อมูลออเดอร์ไม่สำเร็จ: {e}")
        return pd.DataFrame()


# ─── Session state defaults ──────────────────────────────────────────────────

if "cart" not in st.session_state:
    st.session_state.cart = {}   # {item_name: {name, price, qty, note, cat}}
if "menu_loaded" not in st.session_state:
    st.session_state.menu_loaded = False
if "menu_items" not in st.session_state:
    st.session_state.menu_items = []


# ─── Header ─────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
  <div class="header-icon">🍱</div>
  <div>
    <h1>ระบบสั่งอาหารองค์กร</h1>
    <p>เชื่อมต่อ Google Sheets &bull; บันทึกคำสั่งซื้ออัตโนมัติ</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── Tabs ────────────────────────────────────────────────────────────────────

tab_order, tab_summary, tab_setup = st.tabs(["🛒 สั่งอาหาร", "📊 สรุปยอด", "⚙️ ตั้งค่า"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — ORDER
# ══════════════════════════════════════════════════════════════════════════════
with tab_order:
    col_left, col_right = st.columns([3, 2], gap="large")

    with col_left:
        st.markdown('<div class="section-title">👤 ข้อมูลผู้สั่ง</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            orderer = st.text_input("ชื่อ / แผนก", placeholder="เช่น สมชาย / ฝ่ายบัญชี", label_visibility="visible")
        with c2:
            order_note = st.text_input("หมายเหตุ (ไม่เผ็ด, ไม่ผัก...)", label_visibility="visible")

        # Load / Reload menu
        load_col, refresh_col = st.columns([2, 1])
        with load_col:
            if st.button("📋 โหลดเมนูจาก Google Sheets", use_container_width=True, type="primary"):
                with st.spinner("กำลังโหลดเมนู..."):
                    sheet_id = st.secrets.get("spreadsheet_id", "")
                    items = load_menu(sheet_id)
                    st.session_state.menu_items = items
                    st.session_state.menu_loaded = True
                    if items:
                        st.success(f"โหลดสำเร็จ — {len(items)} รายการ")
                    else:
                        st.warning("ไม่พบเมนู หรือยังไม่ได้ตั้งค่า Google Sheets")
        with refresh_col:
            if st.button("🔄 รีเฟรช", use_container_width=True):
                load_menu.clear()
                st.rerun()

        # Menu grid
        if st.session_state.menu_items:
            st.markdown('<div class="section-title" style="margin-top:1.25rem;">🍽️ เลือกรายการอาหาร</div>', unsafe_allow_html=True)

            # Filter by category
            cats = sorted(set(m["cat"] for m in st.session_state.menu_items if m["cat"]))
            if cats:
                selected_cat = st.selectbox("หมวดหมู่", ["ทั้งหมด"] + cats, label_visibility="visible")
                filtered = [m for m in st.session_state.menu_items if selected_cat == "ทั้งหมด" or m["cat"] == selected_cat]
            else:
                filtered = st.session_state.menu_items

            # Render clickable cards using columns
            cols_per_row = 3
            rows = [filtered[i:i+cols_per_row] for i in range(0, len(filtered), cols_per_row)]
            for row in rows:
                cols = st.columns(cols_per_row)
                for ci, item in enumerate(row):
                    with cols[ci]:
                        in_cart = item["name"] in st.session_state.cart
                        btn_label = f"✅ {item['name']}" if in_cart else item["name"]
                        badge = f"  \n**฿{item['price']:.0f}**"
                        if st.button(
                            f"{btn_label}\n\n{'฿'+str(int(item['price']))}{'  •  '+item['cat'] if item['cat'] else ''}",
                            key=f"menu_{item['name']}",
                            use_container_width=True,
                            type="primary" if in_cart else "secondary",
                        ):
                            if in_cart:
                                del st.session_state.cart[item["name"]]
                            else:
                                st.session_state.cart[item["name"]] = {
                                    **item, "qty": 1, "orderer": orderer, "note": order_note
                                }
                            st.rerun()
        elif not st.session_state.menu_loaded:
            st.info("กด **โหลดเมนูจาก Google Sheets** เพื่อเริ่มต้น")

    with col_right:
        st.markdown('<div class="section-title">🛒 รายการที่เลือก</div>', unsafe_allow_html=True)

        if not st.session_state.cart:
            st.markdown('<p style="color:#80868b;font-size:13px;">ยังไม่ได้เลือกรายการ</p>', unsafe_allow_html=True)
        else:
            total = 0
            for name, item in list(st.session_state.cart.items()):
                c1, c2, c3 = st.columns([3, 2, 1])
                with c1:
                    st.markdown(f"**{name}**  \n<span style='font-size:12px;color:#5f6368'>฿{item['price']:.0f} × {item['qty']}</span>", unsafe_allow_html=True)
                with c2:
                    new_qty = st.number_input("", min_value=1, max_value=20, value=item["qty"],
                                              key=f"qty_{name}", label_visibility="collapsed")
                    st.session_state.cart[name]["qty"] = new_qty
                with c3:
                    if st.button("✕", key=f"del_{name}"):
                        del st.session_state.cart[name]
                        st.rerun()
                total += item["price"] * item["qty"]
                st.divider()

            st.markdown(f"""
            <div class="total-bar">
              <span>รวมทั้งหมด</span>
              <span>฿{total:.0f}</span>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("")
            if st.button("✅ บันทึกลง Google Sheets", type="primary", use_container_width=True):
                if not orderer.strip():
                    st.warning("กรุณาใส่ชื่อหรือแผนกก่อนบันทึก")
                else:
                    orders = [
                        {**item, "orderer": orderer, "note": order_note}
                        for item in st.session_state.cart.values()
                    ]
                    try:
                        with st.spinner("กำลังบันทึก..."):
                            append_orders(orders)
                        st.success("🎉 บันทึกคำสั่งซื้อเรียบร้อยแล้ว!")
                        st.session_state.cart = {}
                        load_orders.clear()
                        st.rerun()
                    except Exception as e:
                        st.error(f"บันทึกไม่สำเร็จ: {e}")

            if st.button("🗑️ ล้างรายการ", use_container_width=True):
                st.session_state.cart = {}
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
with tab_summary:
    st.markdown('<div class="section-title">📊 สรุปคำสั่งซื้อ</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        date_from = st.date_input("จากวันที่", value=date.today())
    with c2:
        date_to = st.date_input("ถึงวันที่", value=date.today())
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        load_btn = st.button("📥 โหลดข้อมูล", type="primary", use_container_width=True)

    if load_btn:
        sheet_id = st.secrets.get("spreadsheet_id", "")
        with st.spinner("กำลังโหลด..."):
            df = load_orders(sheet_id, str(date_from), str(date_to))

        if df.empty:
            st.info("ไม่พบข้อมูลในช่วงวันที่ที่เลือก")
        else:
            total_rows = len(df)
            total_amt = df["ราคารวม"].sum()
            persons = df["ชื่อ/แผนก"].nunique()

            m1, m2, m3 = st.columns(3)
            m1.metric("รายการทั้งหมด", total_rows)
            m2.metric("ผู้สั่ง / แผนก", persons)
            m3.metric("ยอดรวม (฿)", f"{total_amt:,.0f}")

            st.markdown("---")
            st.markdown("**สรุปแยกตามชื่อ / แผนก**")

            grouped = df.groupby("ชื่อ/แผนก").agg(
                รายการ=("รายการ", lambda x: ", ".join(x)),
                จำนวนรายการ=("รายการ", "count"),
                ยอดรวม=("ราคารวม", "sum")
            ).reset_index().sort_values("ยอดรวม", ascending=False)
            grouped["ยอดรวม"] = grouped["ยอดรวม"].apply(lambda x: f"฿{x:,.0f}")

            st.dataframe(grouped, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("**รายการทั้งหมด**")
            st.dataframe(df, use_container_width=True, hide_index=True)

            # Download CSV
            csv = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ ดาวน์โหลด CSV", csv, file_name=f"orders_{date_from}_{date_to}.csv", mime="text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — SETUP GUIDE
# ══════════════════════════════════════════════════════════════════════════════
with tab_setup:
    st.markdown('<div class="section-title">⚙️ วิธีตั้งค่าระบบ</div>', unsafe_allow_html=True)

    with st.expander("📋 ขั้นตอนที่ 1 — เตรียม Google Sheets", expanded=True):
        st.markdown("""
**สร้าง Google Spreadsheet ใหม่** แล้วตั้งชื่อ sheet ดังนี้:

**Sheet ชื่อ `Menu`** (คอลัมน์ตามลำดับ):

| A: ชื่ออาหาร | B: ราคา | C: หมวดหมู่ | D: หมายเหตุ |
|---|---|---|---|
| ข้าวผัดกุ้ง | 60 | อาหารจานเดียว | |
| ก๋วยเตี๋ยวเนื้อ | 55 | ก๋วยเตี๋ยว | เผ็ดได้ |
| น้ำส้มคั้น | 25 | เครื่องดื่ม | |

**Sheet ชื่อ `Orders`** — ระบบจะเขียนให้อัตโนมัติ (แถวแรกเป็น header):

| วันที่ | ชื่อ/แผนก | รายการ | จำนวน | ราคารวม | หมายเหตุ |
|---|---|---|---|---|---|
""")

    with st.expander("🔑 ขั้นตอนที่ 2 — สร้าง Google Service Account"):
        st.markdown("""
1. ไปที่ [Google Cloud Console](https://console.cloud.google.com/)
2. สร้าง Project ใหม่ (หรือเลือก Project เดิม)
3. ไปที่ **APIs & Services → Enable APIs** → เปิดใช้ **Google Sheets API** และ **Google Drive API**
4. ไปที่ **IAM & Admin → Service Accounts → Create Service Account**
5. ตั้งชื่อ เช่น `food-order-bot` แล้วกด Create
6. ใน Service Account ที่สร้าง ไปที่แท็บ **Keys → Add Key → JSON** — ดาวน์โหลดไฟล์ JSON
7. กลับไปที่ Google Sheets → **Share** → ใส่ email ของ Service Account (ดูจากไฟล์ JSON ช่อง `client_email`) → ให้สิทธิ์ **Editor**
""")

    with st.expander("🚀 ขั้นตอนที่ 3 — Deploy บน Streamlit Community Cloud"):
        st.markdown("""
1. **Fork หรือ Push โค้ดนี้ขึ้น GitHub** (repository ของคุณ)
2. ไปที่ [share.streamlit.io](https://share.streamlit.io) → **New app** → เลือก repo และไฟล์ `app.py`
3. ก่อน Deploy ไปที่ **Advanced settings → Secrets** แล้วใส่:

```toml
spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "abc123..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\\n...\\n-----END RSA PRIVATE KEY-----\\n"
client_email = "food-order-bot@your-project.iam.gserviceaccount.com"
client_id = "123456789"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
```

> ค่าทั้งหมดนำมาจากไฟล์ JSON ที่ดาวน์โหลดในขั้นตอนที่ 2

4. กด **Deploy** — เสร็จแล้วได้ URL สาธารณะของแอป
""")

    with st.expander("🏠 ขั้นตอนที่ 4 — ทดสอบ Local ก่อน Deploy"):
        st.markdown("""
สร้างไฟล์ `.streamlit/secrets.toml` ใน project folder แล้วใส่ค่าเดียวกับที่ตั้งบน Cloud:

```bash
pip install -r requirements.txt
streamlit run app.py
```
""")

    # Show current connection status
    st.markdown("---")
    st.markdown("**สถานะการเชื่อมต่อปัจจุบัน**")
    try:
        sid = st.secrets.get("spreadsheet_id", "")
        if sid:
            ss = get_spreadsheet()
            if ss:
                st.success(f"✅ เชื่อมต่อ Google Sheets สำเร็จ — Spreadsheet ID: `{sid[:20]}...`")
            else:
                st.error("❌ ไม่สามารถเชื่อมต่อ Spreadsheet ได้ — ตรวจสอบ Service Account และ permissions")
        else:
            st.warning("⚠️ ยังไม่ได้ตั้งค่า `spreadsheet_id` ใน Secrets")
    except Exception as e:
        st.warning(f"⚠️ ยังไม่ได้ตั้งค่า Secrets: {e}")