import streamlit as st
import sqlite3
import pandas as pd
import firebase_admin
from firebase_admin import credentials, storage
import os
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go

# --- 1. APP UI & GLOBAL STYLING ---
st.set_page_config(page_title="Ring Fence Funds Portal", layout="wide", page_icon="🛡️")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    :root {
        --brand-navy: #0b1f3a;
        --brand-navy-2: #13294b;
        --brand-blue: #1d4ed8;
        --brand-blue-light: #3b82f6;
        --surface: #ffffff;
        --surface-alt: #f4f6fb;
        --border-soft: #e4e8f1;
        --text-primary: #0f172a;
        --text-muted: #64748b;
        --radius-lg: 18px;
        --radius-md: 12px;
        --radius-sm: 8px;
        --shadow-soft: 0 4px 16px rgba(15, 23, 42, 0.06);
        --shadow-lift: 0 12px 28px rgba(15, 23, 42, 0.10);
    }

    html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; }

    /* ---------- Global App Background ---------- */
    .stApp {
        background:
            radial-gradient(circle at 0% 0%, rgba(29,78,216,0.05) 0%, transparent 45%),
            radial-gradient(circle at 100% 0%, rgba(11,31,58,0.04) 0%, transparent 45%),
            var(--surface-alt);
    }

    .block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1400px; }

    h1, h2, h3 { color: var(--text-primary); letter-spacing: -0.01em; }
    h1 { font-weight: 800 !important; }

    hr { border-color: var(--border-soft) !important; margin: 1.4rem 0 !important; }

    /* ---------- Buttons ---------- */
    div.stButton > button, div.stDownloadButton > button {
        background: linear-gradient(135deg, var(--brand-blue) 0%, var(--brand-navy) 100%) !important;
        color: white !important;
        border: none !important;
        padding: 0.7rem 1.5rem !important;
        font-weight: 600 !important;
        border-radius: var(--radius-sm) !important;
        font-size: 14.5px !important;
        letter-spacing: 0.1px;
        box-shadow: 0 6px 16px rgba(29, 78, 216, 0.22) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 22px rgba(29, 78, 216, 0.30) !important;
        filter: brightness(1.04);
    }
    div.stButton > button:active, div.stDownloadButton > button:active { transform: translateY(0); }

    /* ---------- Login Form Card ---------- */
    div[data-testid="stForm"] {
        background: var(--surface) !important;
        border-radius: var(--radius-lg) !important;
        padding: 36px 34px !important;
        box-shadow: var(--shadow-lift) !important;
        border: 1px solid var(--border-soft) !important;
    }
    div[data-testid="stForm"] label p {
        color: #334155 !important;
        font-weight: 600 !important;
        font-size: 13px !important;
    }
    div[data-testid="stForm"] input {
        background-color: var(--surface-alt) !important;
        color: var(--text-primary) !important;
        border: 1.5px solid var(--border-soft) !important;
        border-radius: var(--radius-sm) !important;
        padding: 10px 12px !important;
    }
    div[data-testid="stForm"] input:focus {
        border-color: var(--brand-blue-light) !important;
        box-shadow: 0 0 0 3px rgba(59,130,246,0.15) !important;
    }

    .login-badge {
        display: inline-block;
        background: linear-gradient(135deg, var(--brand-blue), var(--brand-navy));
        color: white;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.6px;
        text-transform: uppercase;
        padding: 5px 12px;
        border-radius: 20px;
        margin-bottom: 14px;
    }
    .login-title {
        font-size: 24px;
        font-weight: 800;
        color: var(--text-primary);
        margin-top: 6px;
        margin-bottom: 4px;
        letter-spacing: -0.02em;
    }
    .login-subtitle {
        font-size: 13.5px;
        color: var(--text-muted);
        margin-bottom: 22px;
        line-height: 1.5;
    }

    /* ---------- Header ---------- */
    .header-logo-container {
        display: flex;
        align-items: center;
        gap: 15px;
        margin-bottom: 10px;
    }
    .portal-eyebrow {
        font-size: 11.5px;
        font-weight: 700;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        color: var(--brand-blue);
        margin-bottom: 2px;
    }

    /* ---------- KPI Cards ---------- */
    .kpi-card {
        position: relative;
        overflow: hidden;
        border-radius: var(--radius-md);
        padding: 22px 22px;
        color: white;
        box-shadow: var(--shadow-soft);
        transition: transform 0.25s ease, box-shadow 0.25s ease;
        border: 1px solid rgba(255,255,255,0.08);
    }
    .kpi-card:hover { transform: translateY(-4px); box-shadow: var(--shadow-lift); }
    .kpi-title {
        font-size: 11.5px; font-weight: 700; text-transform: uppercase;
        letter-spacing: 0.9px; opacity: 0.85;
    }
    .kpi-value { font-size: 26px; font-weight: 800; margin-top: 8px; letter-spacing: -0.02em; }
    .kpi-blue   { background: linear-gradient(135deg, #0b1f3a 0%, #1d4ed8 100%); }
    .kpi-indigo { background: linear-gradient(135deg, #3b0764 0%, #7c3aed 100%); }
    .kpi-teal   { background: linear-gradient(135deg, #064e3b 0%, #10b981 100%); }
    .kpi-orange { background: linear-gradient(135deg, #7c2d12 0%, #ea580c 100%); }

    /* ---------- Tabs ---------- */
    .stTabs [data-baseweb="tab-list"] { gap: 8px; border-bottom: none; }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        background-color: var(--surface);
        border: 1px solid var(--border-soft);
        border-radius: var(--radius-sm);
        padding: 0 18px;
        font-weight: 600;
        color: var(--text-muted);
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        transition: all 0.15s ease;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, var(--brand-blue), var(--brand-navy)) !important;
        color: white !important;
        border-color: transparent !important;
    }

    /* ---------- Sidebar ---------- */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, var(--brand-navy) 0%, var(--brand-navy-2) 100%);
        border-right: none;
    }
    section[data-testid="stSidebar"] * { color: #e2e8f0; }
    section[data-testid="stSidebar"] .stMarkdown p { color: #cbd5e1; }
    section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12) !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h3 { color: #ffffff !important; }

    section[data-testid="stSidebar"] div[role="radiogroup"] label {
        background: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: var(--radius-sm);
        padding: 10px 12px;
        margin-bottom: 6px;
        transition: all 0.15s ease;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
        background: rgba(255,255,255,0.10);
    }

    section[data-testid="stSidebar"] div.stButton > button {
        background: rgba(255,255,255,0.08) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        box-shadow: none !important;
        color: #f1f5f9 !important;
    }
    section[data-testid="stSidebar"] div.stButton > button:hover {
        background: rgba(255,255,255,0.16) !important;
        transform: none;
    }

    /* Sidebar role/user badge chips */
    .role-chip {
        display: inline-block;
        font-size: 12px;
        font-weight: 700;
        padding: 5px 12px;
        border-radius: 20px;
        margin-top: 4px;
    }
    .role-chip-admin { background: rgba(16,185,129,0.18); color: #6ee7b7; border: 1px solid rgba(16,185,129,0.35); }
    .role-chip-viewer { background: rgba(250,204,21,0.15); color: #fde047; border: 1px solid rgba(250,204,21,0.35); }

    /* ---------- Filter Cards (sidebar) ---------- */
    .filter-card-box {
        background: rgba(255,255,255,0.05);
        border: 1px solid rgba(255,255,255,0.10);
        border-left: 3px solid var(--brand-blue-light);
        border-radius: var(--radius-sm);
        padding: 16px;
        margin-bottom: 16px;
    }
    .filter-header {
        font-size: 12px;
        font-weight: 700;
        color: #f1f5f9;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 12px;
    }

    /* ---------- Multiselect / Select Styling ---------- */
    div[data-baseweb="select"] > div {
        background-color: rgba(255,255,255,0.06) !important;
        border: 1.5px solid rgba(255,255,255,0.15) !important;
        border-radius: 10px !important;
        padding: 3px 6px !important;
    }
    div[data-baseweb="tag"] {
        background: linear-gradient(135deg, var(--brand-blue) 0%, var(--brand-navy) 100%) !important;
        border-radius: 20px !important;
        padding: 4px 10px !important;
        border: none !important;
    }
    div[data-baseweb="tag"] span { color: #ffffff !important; font-weight: 600 !important; font-size: 12px !important; }

    /* ---------- Dataframe / Table ---------- */
    div[data-testid="stDataFrame"] {
        border-radius: var(--radius-md);
        overflow: hidden;
        border: 1px solid var(--border-soft);
        box-shadow: var(--shadow-soft);
    }

    /* ---------- Alerts ---------- */
    div[data-testid="stAlert"] {
        border-radius: var(--radius-sm);
        box-shadow: var(--shadow-soft);
    }
</style>
""", unsafe_allow_html=True)

# --- 2. USER DATABASE & ROLE CONTROL MODULE ---
USER_DB = {
    "admin": {"password": "police123", "role": "admin", "display_name": "Administrator"},
    "user": {"password": "ringfence2026", "role": "viewer", "display_name": "Standard Viewer"}
}

def check_password():
    """Returns True if user is authenticated, otherwise displays the Login Screen."""
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
        st.session_state["user_role"] = None
        st.session_state["user_id"] = None

    if st.session_state["authenticated"]:
        return True

    # Render Login Page UI
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, login_col, _ = st.columns([1, 1.2, 1])

    with login_col:
        with st.form("login_form", clear_on_submit=False):
            st.markdown('<div style="text-align: center;">', unsafe_allow_html=True)
            if os.path.exists("punjab_police_logo.png"):
                st.image("punjab_police_logo.png", width=84)
            elif os.path.exists("punjab_police.png"):
                st.image("punjab_police.png", width=84)
            else:
                st.markdown("<h1 style='margin-bottom:0;'>🛡️</h1>", unsafe_allow_html=True)

            st.markdown("""
                <div class="login-badge">Secure Portal Access</div>
                <div class="login-title">Ring Fence Funds Portal</div>
                <div class="login-subtitle">Traffic Headquarters Punjab<br>Please sign in to continue</div>
                </div>
            """, unsafe_allow_html=True)

            username_input = st.text_input("Username", placeholder="Enter your username").strip().lower()
            password_input = st.text_input("Password", type="password", placeholder="Enter your password")
            login_btn = st.form_submit_button("🔐 Login to Portal", use_container_width=True)

            if login_btn:
                if username_input in USER_DB and USER_DB[username_input]["password"] == password_input:
                    st.session_state["authenticated"] = True
                    st.session_state["user_id"] = username_input
                    st.session_state["user_role"] = USER_DB[username_input]["role"]
                    st.session_state["display_name"] = USER_DB[username_input]["display_name"]
                    st.rerun()
                    return True
                else:
                    st.error("❌ Invalid Username or Password")
                    return False

    return False

# Security Gate
if not check_password():
    st.stop()

# --- 3. FIREBASE INITIALIZATION ---
if not firebase_admin._apps:
    import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore

# Initialize Firebase using Streamlit Secrets
if not firebase_admin._apps:
    try:
        # Convert Streamlit TOML secrets into a Python dictionary
        key_dict = dict(st.secrets["firebase"])
        
        # Handle newlines in private key if formatted as a single string
        if "private_key" in key_dict:
            key_dict["private_key"] = key_dict["private_key"].replace("\\n", "\n")
            
        cred = credentials.Certificate(key_dict)
        firebase_admin.initialize_app(cred)
    except Exception as e:
        st.error(f"Failed to connect to Firebase: {e}")

# Connect to Firestore database
db = firestore.client()
    firebase_admin.initialize_app(cred, {
        "storageBucket": "ring-fence-funds.firebasestorage.app"
    })

bucket = storage.bucket()
DB_FILE = "professional_funds_distribution.db"

# --- 4. DATA UTILITIES & DATABASE INITIALIZATION ---
def init_db():
    """Initializes local SQLite schema if starting fresh."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS dim_districts (
            district_id INTEGER PRIMARY KEY AUTOINCREMENT,
            district_name TEXT UNIQUE NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS dim_heads_of_account (
            head_id INTEGER PRIMARY KEY AUTOINCREMENT,
            head_code TEXT UNIQUE NOT NULL,
            head_description TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fact_allocations (
            allocation_id INTEGER PRIMARY KEY AUTOINCREMENT,
            district_id INTEGER,
            head_id INTEGER,
            ddo_code TEXT,
            allocation_date TEXT,
            amount REAL,
            FOREIGN KEY (district_id) REFERENCES dim_districts(district_id),
            FOREIGN KEY (head_id) REFERENCES dim_heads_of_account(head_id)
        )
    ''')
    conn.commit()
    conn.close()

def sync_db_from_cloud():
    blob = bucket.blob(DB_FILE)
    if blob.exists():
        blob.download_to_filename(DB_FILE)

def sync_db_to_cloud():
    blob = bucket.blob(DB_FILE)
    blob.upload_from_filename(DB_FILE)

# Sync and verify database structure on startup
sync_db_from_cloud()
init_db()

# --- 5. EXCEL A4 PRINT GENERATOR ---
def generate_a4_print_excel(df, start_date, end_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fund Allocations"

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    total_fill = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    
    font_header = Font(name="Arial", size=9, bold=True)
    font_bold = Font(name="Arial", size=9, bold=True)
    font_normal = Font(name="Arial", size=9, bold=False)
    
    thin_border = Border(
        left=Side(style="thin", color="888888"),
        right=Side(style="thin", color="888888"),
        top=Side(style="thin", color="888888"),
        bottom=Side(style="thin", color="888888")
    )

    num_cols = len(df.columns)

    logo_path = "punjab_police_logo.png" if os.path.exists("punjab_police_logo.png") else "punjab_police.png"
    if os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image
            img = Image(logo_path)
            img.width = 45
            img.height = 45
            ws.add_image(img, "A1")
        except Exception:
            pass

    start_str = start_date.strftime("%d-%m-%Y")
    end_str = end_date.strftime("%d-%m-%Y")
    title_text = f"RING FENCE FUNDS DETAIL FROM {start_str} TO {end_str}"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    title_cell = ws.cell(row=1, column=1, value=title_text)
    title_cell.font = Font(name="Arial", size=11, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.row_dimensions[3].height = 35
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=3, column=col_idx, value=str(col_name))
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    start_row = 4
    num_data_rows = len(df) - 1

    for r_idx in range(num_data_rows):
        current_row = start_row + r_idx
        ws.row_dimensions[current_row].height = 20
        
        for c_idx, col_name in enumerate(df.columns, 1):
            val = df.iloc[r_idx][col_name]
            cell = ws.cell(row=current_row, column=c_idx, value=val)
            cell.border = thin_border

            if col_name == "District":
                cell.font = font_bold
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_name == "DDO Code":
                cell.font = font_normal
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == "Total":
                cell.font = font_bold
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.fill = total_fill
            else:
                cell.font = font_normal
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")

    tot_row_idx = start_row + num_data_rows
    ws.row_dimensions[tot_row_idx].height = 22
    for c_idx, col_name in enumerate(df.columns, 1):
        val = df.iloc[-1][col_name]
        cell = ws.cell(row=tot_row_idx, column=c_idx, value=val)
        cell.font = font_bold
        cell.fill = total_fill
        cell.border = thin_border

        if col_name in ["District", "DDO Code"]:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 12
    for c_idx in range(3, num_cols + 1):
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = 16

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# --- 6. SIDEBAR NAVIGATION, DIG PHOTO & ROLE STATUS ---
with st.sidebar:
    # DIG Officer Photo / Header Image
    if os.path.exists("dig_waqas_nazir.jpg"):
        st.image("dig_waqas_nazir.jpg", caption="DIG Muhammad Waqas Nazir", use_container_width=True)
    elif os.path.exists("punjab_police_logo.png"):
        st.image("punjab_police_logo.png", use_container_width=True)
    elif os.path.exists("punjab_police.png"):
        st.image("punjab_police.png", use_container_width=True)
    else:
        st.warning("Logo/Image files not found.")
        
    user_role = st.session_state.get("user_role", "viewer")
    display_name = st.session_state.get("display_name", "User")

    st.markdown(f"""
        <div style="margin-top:6px; margin-bottom:10px;">
            <div style="font-size:11px; letter-spacing:0.8px; text-transform:uppercase; color:#94a3b8; font-weight:700;">Signed in as</div>
            <div style="font-size:15px; font-weight:700; color:#f8fafc; margin-top:2px;">{display_name}</div>
            <span class="role-chip {'role-chip-admin' if user_role == 'admin' else 'role-chip-viewer'}">
                {'🟢 Administrator' if user_role == 'admin' else '🟡 Standard Viewer'}
            </span>
        </div>
    """, unsafe_allow_html=True)

    if st.button("🚪 Log Out", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["user_role"] = None
        st.rerun()

    st.markdown("---")
    st.markdown('<h3 style="margin-bottom:4px;">📌 Navigation</h3>', unsafe_allow_html=True)

    # Role-Based Menu Filtering
    if user_role == "admin":
        nav_options = ["📊 Fund Allocations Matrix", "📥 Data Entry Form"]
    else:
        nav_options = ["📊 Fund Allocations Matrix"]

    app_mode = st.radio("Select Portal View:", nav_options, index=0)
    st.markdown("---")

# --- MAIN DASHBOARD HEADER WITH LOGOS ---
l_col1, l_col2, l_col3, title_col = st.columns([0.08, 0.08, 0.08, 0.76])

with l_col1:
    if os.path.exists("punjab_police_logo.png"):
        st.image("punjab_police_logo.png", width=62)
    elif os.path.exists("punjab_police.png"):
        st.image("punjab_police.png", width=62)
    else:
        st.write("🛡️")

with l_col2:
    if os.path.exists("govt_punjab_logo.png"):
        st.image("govt_punjab_logo.png", width=62)
    elif os.path.exists("govt_punjab.png"):
        st.image("govt_punjab.png", width=62)

with l_col3:
    if os.path.exists("traffic_police_logo.png"):
        st.image("traffic_police_logo.png", width=62)
    elif os.path.exists("traffic_police.png"):
        st.image("traffic_police.png", width=62)

with title_col:
    st.markdown('<div class="portal-eyebrow">Traffic Headquarters Punjab</div>', unsafe_allow_html=True)
    st.title("Ring Fence Funds Allocation Portal")
    st.caption("Official Allocation & Matrix Ledger")

st.markdown("---")

# --- 7. DATA ENTRY PAGE (ADMIN ONLY) ---
if app_mode == "📥 Data Entry Form":
    if st.session_state.get("user_role") != "admin":
        st.error("🚫 Access Denied: Only Administrator accounts are permitted to enter or modify data.")
        st.stop()

    st.subheader("📥 Input New Allocation Entry")
    
    form_col1, _ = st.columns([2, 1])
    
    with form_col1:
        with st.form("sql_entry_form", clear_on_submit=True):
            st.markdown("##### Fill in Allocation Details:")
            selected_district = st.text_input("District Name", placeholder="e.g. CPO, Faisalabad")
            ddo_input = st.text_input("DDO Code", placeholder="e.g. FD4218")
            selected_head = st.text_input("Head of Account Code", placeholder="e.g. A03955")
            head_description = st.text_input("Head Description", placeholder="e.g. Computer Stationery")
            date_input = st.date_input("Allocation Date", value=pd.to_datetime("2026-01-01"))
            amount_input = st.number_input("Amount (PKR)", min_value=0.0, step=1000.0)
            
            submit_btn = st.form_submit_button("🚀 Save and Sync Record", use_container_width=True)

    if submit_btn:
        if not selected_district or not selected_head:
            st.error("Please enter both District Name and Head of Account Code.")
        else:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("INSERT OR IGNORE INTO dim_districts (district_name) VALUES (?)", (selected_district,))
            cursor.execute("SELECT district_id FROM dim_districts WHERE district_name = ?", (selected_district,))
            d_id = cursor.fetchone()[0]
            
            cursor.execute("INSERT OR IGNORE INTO dim_heads_of_account (head_code, head_description) VALUES (?, ?)", (selected_head, head_description))
            cursor.execute("SELECT head_id FROM dim_heads_of_account WHERE head_code = ?", (selected_head,))
            h_id = cursor.fetchone()[0]
            
            cursor.execute("INSERT INTO fact_allocations (district_id, head_id, ddo_code, allocation_date, amount) VALUES (?, ?, ?, ?, ?)", 
                           (d_id, h_id, ddo_input, date_input.strftime("%Y-%m-%d"), amount_input))
            conn.commit()
            conn.close()
            sync_db_to_cloud()
            st.success("✅ Data saved and synced to cloud successfully!")

# --- 8. MATRIX TABLE & EXECUTIVE GRAPHICAL DASHBOARD PAGE ---
elif app_mode == "📊 Fund Allocations Matrix":
    with st.sidebar:
        st.markdown("### 🔍 Filter Dashboard")
        
        conn = sqlite3.connect(DB_FILE)
        districts = pd.read_sql("SELECT district_name FROM dim_districts", conn)["district_name"].tolist()
        heads = pd.read_sql("SELECT head_code FROM dim_heads_of_account", conn)["head_code"].tolist()
        conn.close()

        # Filter Card 1: Geographic & Head Filters
        st.markdown("""
        <div class="filter-card-box">
            <div class="filter-header">🎯 Category & Location</div>
        """, unsafe_allow_html=True)
        
        f_dist = st.multiselect("Select District(s)", districts, placeholder="Choose districts...")
        f_head = st.multiselect("Select Budget Head(s)", heads, placeholder="Choose budget heads...")
        
        st.markdown("</div>", unsafe_allow_html=True)

        # Filter Card 2: Date Range Filter
        st.markdown("""
        <div class="filter-card-box">
            <div class="filter-header">📅 Date Range Selector</div>
        """, unsafe_allow_html=True)
        
        start_date = st.date_input("Start Date", value=pd.to_datetime("2026-01-01"))
        end_date = st.date_input("End Date", value=pd.to_datetime("2026-12-31"))
        
        st.markdown("</div>", unsafe_allow_html=True)

        active_count = len(f_dist) + len(f_head)
        if active_count > 0:
            st.info(f"⚡ **{active_count} Filter(s) Applied**")

    # SQL Data Retrieval
    conn = sqlite3.connect(DB_FILE)
    query = '''
    SELECT f.allocation_date, d.district_name, f.ddo_code, h.head_code, h.head_description, f.amount 
    FROM fact_allocations f
    JOIN dim_districts d ON f.district_id = d.district_id
    JOIN dim_heads_of_account h ON f.head_id = h.head_id
    WHERE f.allocation_date BETWEEN ? AND ?
    '''
    params = [start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")]
    
    if f_dist:
        query += " AND d.district_name IN ({})".format(','.join(['?']*len(f_dist)))
        params.extend(f_dist)
    if f_head:
        query += " AND h.head_code IN ({})".format(','.join(['?']*len(f_head)))
        params.extend(f_head)
        
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    if not df.empty:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
        df_active = df[df["amount"] > 0].copy()

        # EXECUTIVE GRAPHICAL DASHBOARD
        st.markdown("### 📈 Executive Financial Insights")
        
        total_funds = df["amount"].sum()
        active_districts_count = df_active["district_name"].nunique() if not df_active.empty else 0
        active_heads_count = df_active["head_code"].nunique() if not df_active.empty else 0
        avg_allocation = df_active["amount"].mean() if not df_active.empty else 0

        # KPI Cards Display
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(f'<div class="kpi-card kpi-blue"><div class="kpi-title">💰 Total Allocation</div><div class="kpi-value">PKR {total_funds:,.0f}</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card kpi-indigo"><div class="kpi-title">🏛️ Active Districts</div><div class="kpi-value">{active_districts_count}</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card kpi-teal"><div class="kpi-title">📋 Active Budget Heads</div><div class="kpi-value">{active_heads_count}</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="kpi-card kpi-orange"><div class="kpi-title">📊 Avg. Allocation Size</div><div class="kpi-value">PKR {avg_allocation:,.0f}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if not df_active.empty:
            tab1, tab2, tab3 = st.tabs(["🏛️ Executive Overview", "🌳 Hierarchical Treemap", "📊 Allocation Spread"])

            with tab1:
                col_chart1, col_chart2 = st.columns([1, 1])

                with col_chart1:
                    head_pie_df = df_active.groupby(["head_code", "head_description"])["amount"].sum().reset_index()
                    head_pie_df["Head Title"] = head_pie_df["head_code"] + " (" + head_pie_df["head_description"] + ")"
                    
                    fig_donut = go.Figure(data=[go.Pie(
                        labels=head_pie_df["Head Title"],
                        values=head_pie_df["amount"],
                        hole=0.62,
                        marker=dict(colors=px.colors.qualitative.Prism),
                        textinfo="percent+label",
                        insidetextorientation="radial"
                    )])
                    
                    fig_donut.update_layout(
                        title=dict(text="<b>Budget Head Share Breakdown</b>", font=dict(size=16)),
                        annotations=[dict(
                            text=f'<b>Total</b><br>PKR {total_funds/1e6:.1f}M', 
                            x=0.5, y=0.5, 
                            font=dict(size=15, color='#0d6efd'), 
                            showarrow=False
                        )],
                        template="plotly_white",
                        margin=dict(t=50, b=20, l=10, r=10),
                        height=380,
                        showlegend=False
                    )
                    fig_donut.update_traces(hovertemplate="<b>%{label}</b><br>Amount: PKR %{value:,.0f}<br>Share: %{percent}")
                    st.plotly_chart(fig_donut, use_container_width=True)

                with col_chart2:
                    dist_bar_df = df_active.groupby("district_name")["amount"].sum().reset_index()
                    dist_bar_df = dist_bar_df.sort_values(by="amount", ascending=True).tail(10)
                    
                    fig_bar = px.bar(
                        dist_bar_df,
                        x="amount",
                        y="district_name",
                        orientation="h",
                        color="amount",
                        color_continuous_scale="Blues",
                        title="<b>Top District Allocations</b>"
                    )
                    fig_bar.update_layout(
                        template="plotly_white",
                        margin=dict(t=50, b=20, l=10, r=10),
                        height=380,
                        coloraxis_showscale=False,
                        xaxis_title="Allocated Amount (PKR)",
                        yaxis_title=""
                    )
                    fig_bar.update_traces(
                        hovertemplate="<b>%{y}</b><br>Allocation: PKR %{x:,.0f}",
                        marker_line_color='rgb(8,48,107)',
                        marker_line_width=1
                    )
                    st.plotly_chart(fig_bar, use_container_width=True)

            with tab2:
                st.markdown("##### 🌳 Nested Budget Distribution (Head of Account ➔ District)")
                fig_tree = px.treemap(
                    df_active,
                    path=['head_code', 'district_name'],
                    values='amount',
                    color='amount',
                    color_continuous_scale='Tealgrn'
                )
                fig_tree.update_layout(
                    template="plotly_white",
                    margin=dict(t=20, b=20, l=10, r=10),
                    height=420
                )
                fig_tree.update_traces(hovertemplate="<b>%{label}</b><br>Allocated: PKR %{value:,.0f}")
                st.plotly_chart(fig_tree, use_container_width=True)

            with tab3:
                st.markdown("##### 📊 Financial Distribution Density")
                fig_hist = px.histogram(
                    df_active,
                    x="amount",
                    nbins=20,
                    color_discrete_sequence=["#0d6efd"],
                    labels={"amount": "Allocation Value Ranges (PKR)"},
                    title="<b>Allocation Size Frequency Histogram</b>"
                )
                fig_hist.update_layout(
                    template="plotly_white",
                    margin=dict(t=50, b=20, l=10, r=10),
                    height=380,
                    xaxis_title="Allocation Amount Range (PKR)",
                    yaxis_title="Count of Entry Records"
                )
                st.plotly_chart(fig_hist, use_container_width=True)

        else:
            st.warning("⚠️ Selected filters yield zero active allocations.")

        st.markdown("---")
        st.subheader("📊 Fund Allocations Matrix Table")

        df["Head"] = df["head_code"] + "\n" + df["head_description"]

        pivot_df = df.pivot_table(
            index=["district_name", "ddo_code"],
            columns="Head",
            values="amount",
            aggfunc="sum",
            fill_value=0
        )

        pivot_df["Total"] = pivot_df.sum(axis=1)
        pivot_df = pivot_df.reset_index()
        pivot_df.rename(columns={"district_name": "District", "ddo_code": "DDO Code"}, inplace=True)

        numeric_cols = [c for c in pivot_df.columns if c not in ["District", "DDO Code"]]
        total_row = {"District": "Total", "DDO Code": ""}
        for col in numeric_cols:
            total_row[col] = pivot_df[col].sum()
            
        total_df = pd.DataFrame([total_row])
        pivot_df = pd.concat([pivot_df, total_df], ignore_index=True)

        column_configuration = {
            "District": st.column_config.TextColumn("District"),
            "DDO Code": st.column_config.TextColumn("DDO Code"),
        }
        for col in numeric_cols:
            column_configuration[col] = st.column_config.NumberColumn(
                col,
                format="%,.0f"
            )

        st.dataframe(
            pivot_df,
            column_config=column_configuration,
            use_container_width=True,
            hide_index=True
        )

        excel_data = generate_a4_print_excel(pivot_df, start_date, end_date)
        st.download_button(
            label="🖨️ Download Print-Ready Excel (A4 Landscape)",
            data=excel_data,
            file_name=f"Ring_Fence_Funds_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No records found for the selected filter criteria.")
