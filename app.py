"""
ELECTRIC WIRE & CABLE — SALES + PRODUCTION + INVENTORY PORTAL
A single-file Streamlit application for Coppermils covering:
• Public storefront with product features, gauge details, and price cards
• Sidebar quick-product selector & custom CSS contrast fixes
• Customer records, Production entry, Raw material stock, Scrap tracking, & Analytics
"""

import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
import io
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ------------------------------------------------------------------------------------
# 0. APP CONFIG & LIGHT-THEME STYLING FIX
# ------------------------------------------------------------------------------------
BUSINESS_NAME = "Coppermils"
CURRENCY = "PKR"
DB_FILE = "wire_cable_erp.db"

st.set_page_config(page_title=f"{BUSINESS_NAME} — Portal", layout="wide", page_icon="⚡")

# Custom CSS to force clean contrast for dark/light themes and style product cards
st.markdown("""
    <style>
    /* Force high contrast light inputs so text is crisp and clear */
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div {
        background-color: #ffffff !important;
        color: #111111 !important;
        border: 1px solid #cccccc !important;
        border-radius: 6px !important;
    }
    input {
        color: #111111 !important;
    }
    
    /* Product Card Styling */
    .product-card {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 18px;
        margin-bottom: 15px;
        box-shadow: 0px 3px 6px rgba(0,0,0,0.05);
    }
    .product-title {
        font-size: 18px;
        font-weight: bold;
        color: #1F4E78;
    }
    .product-badge {
        background-color: #e6f0ff;
        color: #0056b3;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
    }
    .product-price {
        font-size: 18px;
        font-weight: bold;
        color: #27ae60;
    }
    </style>
""", unsafe_allow_html=True)

# Running Marquee Banner with Contact Numbers
MARQUEE_TEXT = "Quality Matters for contact 0333-4534668, 03214562502"
st.markdown(
    f"""
    <div style="background-color: #fffae6; padding: 10px; border-radius: 5px; border: 1px solid #ffe066; margin-bottom: 20px;">
        <marquee behavior="scroll" direction="left" scrollamount="6" style="color: #b35900; font-weight: bold; font-size: 16px;">
            ⚡ {BUSINESS_NAME} — {MARQUEE_TEXT} ⚡
        </marquee>
    </div>
    """,
    unsafe_allow_html=True
)

# ------------------------------------------------------------------------------------
# 1. DATABASE SETUP & INITIALIZATION
# ------------------------------------------------------------------------------------
def init_db():
    """Ensure all required tables and default data exist with Gauge & Features."""
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    # Raw Materials
    c.execute('''CREATE TABLE IF NOT EXISTS raw_materials (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT,
        item_name TEXT UNIQUE,
        grade TEXT,
        unit TEXT,
        stock_qty REAL DEFAULT 0,
        min_threshold REAL DEFAULT 0
    )''')

    # Finished Product Catalog (Includes Gauge and Features)
    c.execute('''CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT UNIQUE,
        name TEXT,
        category TEXT,
        gauge TEXT,
        features TEXT,
        unit TEXT DEFAULT 'Coil (90m)',
        unit_price REAL DEFAULT 0,
        stock_qty REAL DEFAULT 0
    )''')

    # Migration check: Ensure gauge and features columns exist if database was pre-created
    c.execute("PRAGMA table_info(products)")
    columns = [col[1] for col in c.fetchall()]
    if 'gauge' not in columns:
        c.execute("ALTER TABLE products ADD COLUMN gauge TEXT DEFAULT ''")
    if 'features' not in columns:
        c.execute("ALTER TABLE products ADD COLUMN features TEXT DEFAULT ''")

    # Customers
    c.execute('''CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        phone TEXT UNIQUE,
        city TEXT,
        address TEXT
    )''')

    # Orders
    c.execute('''CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_date TEXT,
        customer_name TEXT,
        customer_phone TEXT,
        city TEXT,
        address TEXT,
        product_code TEXT,
        product_name TEXT,
        quantity REAL,
        unit_price REAL,
        total_amount REAL,
        status TEXT DEFAULT 'Pending'
    )''')

    # Production Log
    c.execute('''CREATE TABLE IF NOT EXISTS production_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        prod_date TEXT,
        manager_name TEXT,
        supervisor_name TEXT,
        labour_name TEXT,
        product_code TEXT,
        output_qty REAL,
        copper_used_kg REAL,
        copper_type TEXT,
        pvc_used_kg REAL,
        pvc_grade TEXT,
        rejected_broken_qty REAL DEFAULT 0,
        scrap_copper_kg REAL DEFAULT 0,
        scrap_pvc_kg REAL DEFAULT 0
    )''')

    # Staff Records
    c.execute('''CREATE TABLE IF NOT EXISTS staff (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        role TEXT,
        phone TEXT
    )''')

    # Scrap Tracker
    c.execute('''CREATE TABLE IF NOT EXISTS scrap_inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        material_type TEXT UNIQUE,
        stock_kg REAL DEFAULT 0
    )''')

    conn.commit()

    # Seed Default Raw Materials
    c.execute("SELECT COUNT(*) FROM raw_materials")
    if c.fetchone()[0] == 0:
        raw_items = [
            ("Copper", "Copper Rod - Pure (99.9%)", "99.9%", "Kg", 500.0, 50.0),
            ("Copper", "Copper Rod - Loose Gauge", "Loose", "Kg", 300.0, 30.0),
            ("Copper", "Copper Wire - China Grade", "China", "Kg", 200.0, 20.0),
            ("PVC", "PVC Compound - Insulation Grade", "Insulation", "Kg", 1000.0, 100.0),
            ("PVC", "PVC Compound - Sheathing Grade", "Sheathing", "Kg", 800.0, 80.0),
            ("XLPE", "XLPE Compound - High Voltage", "HV", "Kg", 400.0, 40.0),
            ("Aluminium", "Aluminium Wire Rod", "Standard", "Kg", 600.0, 60.0)
        ]
        c.executemany("INSERT OR IGNORE INTO raw_materials (category, item_name, grade, unit, stock_qty, min_threshold) VALUES (?,?,?,?,?,?)", raw_items)

    # Seed Default Products with Gauge & Features
    c.execute("SELECT COUNT(*) FROM products")
    if c.fetchone()[0] == 0:
        prods = [
            ("CBL-3070", "3/070 Single Core Copper Wire", "Building Wire", "3/0.029 (0.7mm)", "99.9% Pure Copper, Flame Retardant PVC", "Coil (90m)", 4500.0, 120.0),
            ("CBL-7029", "7/029 Single Core Copper Wire", "Building Wire", "7/0.029 (0.7mm)", "High Conductivity, Heavy Duty Insulation", "Coil (90m)", 8200.0, 85.0),
            ("CBL-7036", "7/036 Heavy Duty Copper Cable", "Building Wire", "7/0.036 (0.9mm)", "Pure Electrolytic Copper, High Heat Resistance", "Coil (90m)", 11500.0, 60.0),
            ("CBL-7044", "7/044 Commercial Power Cable", "Power Cable", "7/0.044 (1.1mm)", "Commercial Grade, Weather Resistant Outer Sheath", "Coil (90m)", 16800.0, 40.0),
            ("CBL-2C-7029", "2-Core 7/029 Flexible Twin Cable", "Twin Core", "7/0.029 Dual", "Double Layer PVC Insulated Flexible Wire", "Coil (90m)", 15500.0, 30.0),
            ("CBL-3C-4MM", "3-Core 4mm XLPE Armored Cable", "Industrial", "4.0 sq mm", "XLPE Insulation, High Voltage Heavy Industrial", "Meter", 680.0, 500.0)
        ]
        c.executemany("INSERT OR IGNORE INTO products (code, name, category, gauge, features, unit, unit_price, stock_qty) VALUES (?,?,?,?,?,?,?,?)", prods)

    # Seed Scrap
    c.execute("SELECT COUNT(*) FROM scrap_inventory")
    if c.fetchone()[0] == 0:
        c.executemany("INSERT OR IGNORE INTO scrap_inventory (material_type, stock_kg) VALUES (?,?)", [
            ("Copper Scrap", 0.0),
            ("PVC Scrap", 0.0)
        ])

    conn.commit()
    conn.close()

def get_db_connection():
    init_db()
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# ------------------------------------------------------------------------------------
# 2. HELPER FUNCTIONS
# ------------------------------------------------------------------------------------
def load_data(query, params=()):
    conn = get_db_connection()
    try:
        df = pd.read_sql_query(query, conn, params=params)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()
    return df

def execute_query(query, params=()):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute(query, params)
    conn.commit()
    conn.close()

def generate_excel_report():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center")
    border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

    ws_prod = wb.active
    ws_prod.title = "Finished Products Stock"
    df_p = load_data("SELECT code, name, category, gauge, features, unit, unit_price, stock_qty FROM products")
    ws_prod.append(["Product Code", "Product Name", "Category", "Gauge", "Features", "Unit", f"Price ({CURRENCY})", "Stock Qty"])
    if not df_p.empty:
        for row in df_p.itertuples(index=False):
            ws_prod.append(list(row))

    ws_raw = wb.create_sheet("Raw Material Inventory")
    df_r = load_data("SELECT category, item_name, grade, unit, stock_qty, min_threshold FROM raw_materials")
    ws_raw.append(["Category", "Item Name", "Grade", "Unit", "Current Stock", "Min Threshold"])
    if not df_r.empty:
        for row in df_r.itertuples(index=False):
            ws_raw.append(list(row))

    ws_sales = wb.create_sheet("Sales Orders")
    df_s = load_data("SELECT id, order_date, customer_name, customer_phone, city, product_name, quantity, unit_price, total_amount, status FROM orders")
    ws_sales.append(["Order ID", "Date", "Customer Name", "Phone", "City", "Product", "Qty", f"Price ({CURRENCY})", f"Total ({CURRENCY})", "Status"])
    if not df_s.empty:
        for row in df_s.itertuples(index=False):
            ws_sales.append(list(row))

    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    output.seek(0)
    return output

# ------------------------------------------------------------------------------------
# 3. SIDEBAR NAVIGATION & QUICK PRODUCT SELECTOR
# ------------------------------------------------------------------------------------
st.sidebar.title(f"⚡ {BUSINESS_NAME}")

nav_choice = st.sidebar.radio(
    "Select Module",
    [
        "🛍️ Storefront (Place Order)",
        "📦 Finished Product Catalog",
        "🏭 Production Entry",
        "🧱 Raw Material Inventory",
        "♻️ Scrap Tracking",
        "📋 Sales Orders",
        "👥 Customer Directory",
        "👷 Staff Management",
        "📊 Reports & Analytics"
    ]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🎯 Quick Product Select")
all_p = load_data("SELECT code, name FROM products")
if not all_p.empty:
    prod_map = {f"{row['code']} - {row['name']}": row['code'] for _, row in all_p.iterrows()}
    sidebar_selected = st.sidebar.selectbox("Jump to Product Order:", ["-- Choose Product --"] + list(prod_map.keys()))
    if sidebar_selected != "-- Choose Product --":
        st.session_state["selected_product"] = prod_map[sidebar_selected]

# ------------------------------------------------------------------------------------
# MODULE 1: STOREFRONT (PUBLIC / NO LOGIN)
# ------------------------------------------------------------------------------------
if nav_choice == "🛍️ Storefront (Place Order)":
    st.title("🛒 Public Storefront — Product Catalog & Ordering")
    st.caption("Browse available wire & cable gauges, features, and place instant orders.")

    df_p = load_data("SELECT code, name, category, gauge, features, unit, unit_price, stock_qty FROM products WHERE stock_qty > 0")

    col_cat, col_search = st.columns([1, 2])
    with col_cat:
        categories = ["All Categories"] + list(df_p["category"].unique()) if not df_p.empty else ["All Categories"]
        sel_cat = st.selectbox("Filter Category", categories)
    with col_search:
        search_term = st.text_input("Search Product Name or Code", "")

    filtered_df = df_p.copy()
    if not filtered_df.empty:
        if sel_cat != "All Categories":
            filtered_df = filtered_df[filtered_df["category"] == sel_cat]
        if search_term:
            filtered_df = filtered_df[
                filtered_df["name"].str.contains(search_term, case=False, na=False) |
                filtered_df["code"].str.contains(search_term, case=False, na=False)
            ]

    st.subheader("Available Cable & Wire Stock")

    if filtered_df.empty:
        st.info("No matching products found or stock empty.")
    else:
        # Render clean cards showing Gauge, Features, Price, and Stock
        for idx, row in filtered_df.iterrows():
            st.markdown(f"""
            <div class="product-card">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <span class="product-title">{row['name']} ({row['code']})</span>
                    <span class="product-badge">{row['category']}</span>
                </div>
                <div style="margin-top: 8px; color: #444444; font-size: 14px;">
                    📐 <b>Gauge / Thickness:</b> {row['gauge'] if row['gauge'] else 'Standard'} | 
                    🏷️ <b>Unit:</b> {row['unit']}
                </div>
                <div style="margin-top: 4px; color: #555555; font-size: 14px;">
                    ✨ <b>Key Features:</b> {row['features'] if row['features'] else 'High Performance Wire'}
                </div>
                <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 10px;">
                    <span class="product-price">{CURRENCY} {row['unit_price']:,.2f} / {row['unit']}</span>
                    <span style="color: #666666; font-size: 13px;">📦 Stock: <b>{row['stock_qty']}</b> {row['unit']}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if st.button(f"🛒 Select {row['code']} for Order", key=f"btn_{row['code']}"):
                st.session_state["selected_product"] = row['code']
                st.toast(f"Selected {row['code']} in order form below!")

            st.markdown("<br>", unsafe_allow_html=True)

    # Order Placement Form
    st.subheader("📝 Place Your Order")
    with st.form("public_order_form"):
        prod_codes = list(df_p["code"]) if not df_p.empty else []
        default_index = 0
        if "selected_product" in st.session_state and st.session_state["selected_product"] in prod_codes:
            default_index = prod_codes.index(st.session_state["selected_product"])

        f_code = st.selectbox("Select Product Code", prod_codes if prod_codes else ["No Products Available"], index=default_index if prod_codes else 0)
        f_qty = st.number_input("Quantity Required", min_value=1.0, value=1.0, step=1.0)

        st.markdown("---")
        st.markdown("**Customer Details**")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            c_name = st.text_input("Full Name *")
            c_phone = st.text_input("Phone Number *")
        with col_c2:
            c_city = st.text_input("City *")
            c_address = st.text_area("Delivery Address")

        submit_order = st.form_submit_button("🚀 Submit Order")

        if submit_order:
            if not c_name or not c_phone or not c_city:
                st.error("Please fill in Name, Phone, and City fields.")
            elif not prod_codes or f_code == "No Products Available":
                st.error("No product selected.")
            else:
                p_info = load_data("SELECT name, unit_price, stock_qty FROM products WHERE code = ?", (f_code,))
                if p_info.empty:
                    st.error("Invalid product selected.")
                else:
                    p_name = p_info.iloc[0]["name"]
                    u_price = p_info.iloc[0]["unit_price"]
                    curr_stock = p_info.iloc[0]["stock_qty"]

                    if f_qty > curr_stock:
                        st.error(f"Cannot place order. Requested quantity ({f_qty}) exceeds available stock ({curr_stock}).")
                    else:
                        tot_amt = f_qty * u_price
                        today_str = date.today().strftime("%Y-%m-%d")

                        execute_query(
                            "INSERT OR REPLACE INTO customers (name, phone, city, address) VALUES (?, ?, ?, ?)",
                            (c_name, c_phone, c_city, c_address)
                        )

                        execute_query(
                            """INSERT INTO orders 
                               (order_date, customer_name, customer_phone, city, address, product_code, product_name, quantity, unit_price, total_amount, status)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending')""",
                            (today_str, c_name, c_phone, c_city, c_address, f_code, p_name, f_qty, u_price, tot_amt)
                        )

                        execute_query("UPDATE products SET stock_qty = stock_qty - ? WHERE code = ?", (f_qty, f_code))

                        st.success(f"🎉 Order placed successfully! Total: {CURRENCY} {tot_amt:,.2f}. We will contact you shortly.")
                        st.rerun()

# ------------------------------------------------------------------------------------
# MODULE 2: FINISHED PRODUCT CATALOG
# ------------------------------------------------------------------------------------
elif nav_choice == "📦 Finished Product Catalog":
    st.title("📦 Product Catalog & Spec Manager")

    tab1, tab2 = st.tabs(["View / Edit Products", "➕ Add New Product"])

    with tab1:
        df_p = load_data("SELECT * FROM products")
        if df_p.empty:
            st.info("No products found.")
        else:
            st.dataframe(df_p, use_container_width=True)

            st.subheader("Edit Product Specifications")
            col_sel, col_p, col_s = st.columns(3)
            with col_sel:
                edit_code = st.selectbox("Select Product", df_p["code"].tolist())
            
            p_row = df_p[df_p["code"] == edit_code].iloc[0]
            
            col_g, col_f = st.columns(2)
            with col_g:
                edit_gauge = st.text_input("Gauge Specification", value=str(p_row.get("gauge", "")))
            with col_f:
                edit_features = st.text_input("Features / Description", value=str(p_row.get("features", "")))

            with col_p:
                new_price = st.number_input("Unit Price", value=float(p_row["unit_price"]), min_value=0.0, step=10.0)
            with col_s:
                new_stock = st.number_input("Stock Qty", value=float(p_row["stock_qty"]), min_value=0.0, step=1.0)

            if st.button("Update Product Record"):
                execute_query("UPDATE products SET gauge = ?, features = ?, unit_price = ?, stock_qty = ? WHERE code = ?", 
                              (edit_gauge, edit_features, new_price, new_stock, edit_code))
                st.success(f"Updated product {edit_code} successfully!")
                st.rerun()

    with tab2:
        with st.form("add_product_form"):
            c_code = st.text_input("Product Code (e.g. CBL-3070)")
            c_name = st.text_input("Product Description / Name")
            c_cat = st.selectbox("Category", ["Building Wire", "Power Cable", "Twin Core", "Industrial", "Control Cable", "Other"])
            c_gauge = st.text_input("Gauge Specification (e.g. 7/0.029, 3/0.070, 4mm)")
            c_features = st.text_area("Features (e.g. 99.9% Pure Copper, Heat Resistant PVC)")
            c_unit = st.selectbox("Unit of Measure", ["Coil (90m)", "Meter", "Roll", "Km"])
            c_price = st.number_input("Unit Price", min_value=0.0, step=10.0)
            c_stock = st.number_input("Initial Stock Qty", min_value=0.0, step=1.0)

            if st.form_submit_button("Add Product"):
                if not c_code or not c_name:
                    st.error("Please enter Product Code and Name.")
                else:
                    try:
                        execute_query(
                            "INSERT INTO products (code, name, category, gauge, features, unit, unit_price, stock_qty) VALUES (?,?,?,?,?,?,?,?)",
                            (c_code, c_name, c_cat, c_gauge, c_features, c_unit, c_price, c_stock)
                        )
                        st.success(f"Product {c_code} added successfully!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error adding product: {e}")

# ------------------------------------------------------------------------------------
# MODULE 3: PRODUCTION ENTRY
# ------------------------------------------------------------------------------------
elif nav_choice == "🏭 Production Entry":
    st.title("🏭 Daily Production & Manufacturing Log")

    df_staff = load_data("SELECT name, role FROM staff")
    managers = df_staff[df_staff["role"] == "Production Manager"]["name"].tolist() if not df_staff.empty else ["Default Manager"]
    supervisors = df_staff[df_staff["role"] == "Supervisor"]["name"].tolist() if not df_staff.empty else ["Default Supervisor"]
    labourers = df_staff[df_staff["role"] == "Labour"]["name"].tolist() if not df_staff.empty else ["Default Labour"]

    df_raw = load_data("SELECT item_name, grade, category FROM raw_materials")
    copper_items = df_raw[df_raw["category"] == "Copper"]["item_name"].tolist() if not df_raw.empty else ["Copper Rod - Pure (99.9%)"]
    pvc_items = df_raw[df_raw["category"] == "PVC"]["item_name"].tolist() if not df_raw.empty else ["PVC Compound - Insulation Grade"]

    df_prods = load_data("SELECT code, name FROM products")
    prod_options = [f"{row['code']} - {row['name']}" for _, row in df_prods.iterrows()] if not df_prods.empty else []

    with st.form("production_entry_form"):
        st.subheader("1. General & Staff Details")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            p_date = st.date_input("Production Date", date.today())
        with c2:
            p_manager = st.selectbox("Production Manager", managers)
        with c3:
            p_supervisor = st.selectbox("Supervisor", supervisors)
        with c4:
            p_labour = st.selectbox("Labour / Operator", labourers)

        st.subheader("2. Finished Product Output")
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            sel_prod_str = st.selectbox("Produced Product", prod_options if prod_options else ["No products"])
        with col_p2:
            out_qty = st.number_input("Finished Good Output Qty", min_value=0.0, step=1.0)
        with col_p3:
            rej_qty = st.number_input("Rejected / Broken Cable Qty", min_value=0.0, step=1.0)

        st.subheader("3. Raw Material Consumption")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            cop_type = st.selectbox("Copper Type Used", copper_items)
        with col_m2:
            cop_used = st.number_input("Copper Consumed (Kg)", min_value=0.0, step=0.1)
        with col_m3:
            pvc_type = st.selectbox("PVC Grade Used", pvc_items)
        with col_m4:
            pvc_used = st.number_input("PVC Consumed (Kg)", min_value=0.0, step=0.1)

        st.subheader("4. Scrap Generation")
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            scrap_cop = st.number_input("Copper Scrap Generated (Kg)", min_value=0.0, step=0.1)
        with col_s2:
            scrap_pvc = st.number_input("PVC Scrap Generated (Kg)", min_value=0.0, step=0.1)

        submit_prod = st.form_submit_button("🏭 Log Production Batch")

        if submit_prod:
            if not prod_options or sel_prod_str == "No products":
                st.error("Please register products first.")
            elif out_qty <= 0:
                st.error("Output quantity must be greater than zero.")
            else:
                p_code = sel_prod_str.split(" - ")[0]

                execute_query(
                    """INSERT INTO production_log 
                       (prod_date, manager_name, supervisor_name, labour_name, product_code, output_qty,
                        copper_used_kg, copper_type, pvc_used_kg, pvc_grade, rejected_broken_qty,
                        scrap_copper_kg, scrap_pvc_kg)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (str(p_date), p_manager, p_supervisor, p_labour, p_code, out_qty,
                     cop_used, cop_type, pvc_used, pvc_type, rej_qty, scrap_cop, scrap_pvc)
                )

                execute_query("UPDATE products SET stock_qty = stock_qty + ? WHERE code = ?", (out_qty, p_code))

                if cop_used > 0:
                    execute_query("UPDATE raw_materials SET stock_qty = stock_qty - ? WHERE item_name = ?", (cop_used, cop_type))
                if pvc_used > 0:
                    execute_query("UPDATE raw_materials SET stock_qty = stock_qty - ? WHERE item_name = ?", (pvc_used, pvc_type))

                if scrap_cop > 0:
                    execute_query("UPDATE scrap_inventory SET stock_kg = stock_kg + ? WHERE material_type = 'Copper Scrap'", (scrap_cop,))
                if scrap_pvc > 0:
                    execute_query("UPDATE scrap_inventory SET stock_kg = stock_kg + ? WHERE material_type = 'PVC Scrap'", (scrap_pvc,))

                st.success("✅ Production batch successfully logged!")
                st.rerun()

    st.subheader("📜 Recent Production Batches")
    df_logs = load_data("SELECT * FROM production_log ORDER BY id DESC LIMIT 10")
    if not df_logs.empty:
        st.dataframe(df_logs, use_container_width=True)

# ------------------------------------------------------------------------------------
# MODULE 4: RAW MATERIAL INVENTORY
# ------------------------------------------------------------------------------------
elif nav_choice == "🧱 Raw Material Inventory":
    st.title("🧱 Raw Material Stock Management")

    df_raw = load_data("SELECT * FROM raw_materials")

    st.subheader("Current Stock Levels")
    if not df_raw.empty:
        for idx, row in df_raw.iterrows():
            if row["stock_qty"] <= row["min_threshold"]:
                st.warning(f"⚠️ Low Stock Alert: {row['item_name']} ({row['category']}) is at {row['stock_qty']} {row['unit']} (Threshold: {row['min_threshold']})")

        st.dataframe(df_raw, use_container_width=True)

    tab1, tab2 = st.tabs(["📥 Adjust / Add Raw Stock", "➕ Register New Material"])

    with tab1:
        if not df_raw.empty:
            col_item, col_add = st.columns(2)
            with col_item:
                sel_item = st.selectbox("Select Material to Update", df_raw["item_name"].tolist())
            with col_add:
                add_qty = st.number_input("Quantity Received / Added (Kg)", step=10.0)

            if st.button("Add to Stock"):
                execute_query("UPDATE raw_materials SET stock_qty = stock_qty + ? WHERE item_name = ?", (add_qty, sel_item))
                st.success(f"Added {add_qty} Kg to {sel_item}")
                st.rerun()

    with tab2:
        with st.form("new_raw_material_form"):
            r_cat = st.selectbox("Material Category", ["Copper", "PVC", "XLPE", "Aluminium", "Other"])
            r_name = st.text_input("Material Description / Name")
            r_grade = st.text_input("Grade (e.g., Pure, Loose Gauge, China, Insulation, Sheathing)")
            r_unit = st.text_input("Unit of Measure", "Kg")
            r_stock = st.number_input("Initial Stock Qty", min_value=0.0, step=10.0)
            r_thresh = st.number_input("Min Alert Threshold", min_value=0.0, step=10.0)

            if st.form_submit_button("Save Raw Material"):
                if not r_name:
                    st.error("Material Name is required.")
                else:
                    try:
                        execute_query(
                            "INSERT INTO raw_materials (category, item_name, grade, unit, stock_qty, min_threshold) VALUES (?,?,?,?,?,?)",
                            (r_cat, r_name, r_grade, r_unit, r_stock, r_thresh)
                        )
                        st.success("New raw material added!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error: {e}")

# ------------------------------------------------------------------------------------
# MODULE 5: SCRAP TRACKING
# ------------------------------------------------------------------------------------
elif nav_choice == "♻️ Scrap Tracking":
    st.title("♻️ Condemned & Scrap Material Inventory")

    df_scrap = load_data("SELECT * FROM scrap_inventory")
    
    col_c, col_p = st.columns(2)
    if not df_scrap.empty:
        for idx, row in df_scrap.iterrows():
            if "Copper" in row["material_type"]:
                with col_c:
                    st.metric(label="🟤 Copper Scrap Balance", value=f"{row['stock_kg']:,.2f} Kg")
            else:
                with col_p:
                    st.metric(label="⚪ PVC Scrap Balance", value=f"{row['stock_kg']:,.2f} Kg")

    st.subheader("Manage Scrap Stock")
    tab1, tab2 = st.tabs(["📝 Manual Adjustment", "💵 Sell / Dispose Scrap"])

    with tab1:
        if not df_scrap.empty:
            col1, col2 = st.columns(2)
            with col1:
                scrap_sel = st.selectbox("Select Scrap Material", df_scrap["material_type"].tolist())
            with col2:
                adj_qty = st.number_input("Quantity Difference (+ or - Kg)", step=1.0)

            if st.button("Apply Adjustment"):
                execute_query("UPDATE scrap_inventory SET stock_kg = stock_kg + ? WHERE material_type = ?", (adj_qty, scrap_sel))
                st.success("Scrap balance updated!")
                st.rerun()

    with tab2:
        if not df_scrap.empty:
            col_s1, col_s2, col_s3 = st.columns(3)
            with col_s1:
                sell_type = st.selectbox("Scrap Type Sold", df_scrap["material_type"].tolist(), key="sell_type")
            with col_s2:
                sell_qty = st.number_input("Quantity Sold (Kg)", min_value=0.0, step=1.0)
            with col_s3:
                sell_rate = st.number_input(f"Rate Per Kg ({CURRENCY})", min_value=0.0, step=10.0)

            if st.button("Record Scrap Sale"):
                curr_stock = df_scrap[df_scrap["material_type"] == sell_type].iloc[0]["stock_kg"]
                if sell_qty > curr_stock:
                    st.error("Sale quantity exceeds existing scrap balance.")
                else:
                    total_rev = sell_qty * sell_rate
                    execute_query("UPDATE scrap_inventory SET stock_kg = stock_kg - ? WHERE material_type = ?", (sell_qty, sell_type))
                    st.success(f"Recorded sale of {sell_qty} Kg {sell_type} for total {CURRENCY} {total_rev:,.2f}!")
                    st.rerun()

# ------------------------------------------------------------------------------------
# MODULE 6: SALES ORDERS MANAGEMENT
# ------------------------------------------------------------------------------------
elif nav_choice == "📋 Sales Orders":
    st.title("📋 Order Management & Sales Registry")

    df_orders = load_data("SELECT * FROM orders ORDER BY id DESC")

    if df_orders.empty:
        st.info("No orders placed yet.")
    else:
        st.dataframe(df_orders, use_container_width=True)

        st.subheader("Update Order Status")
        col_o1, col_o2 = st.columns(2)
        with col_o1:
            o_id = st.selectbox("Select Order ID", df_orders["id"].tolist())
        with col_o2:
            o_status = st.selectbox("New Status", ["Pending", "Processing", "Dispatched", "Delivered", "Cancelled"])

        if st.button("Update Order Status"):
            execute_query("UPDATE orders SET status = ? WHERE id = ?", (o_status, o_id))
            st.success(f"Order #{o_id} status updated to {o_status}!")
            st.rerun()

# ------------------------------------------------------------------------------------
# MODULE 7: CUSTOMER DIRECTORY
# ------------------------------------------------------------------------------------
elif nav_choice == "👥 Customer Directory":
    st.title("👥 Customer Directory")

    df_cust = load_data("SELECT * FROM customers")
    if df_cust.empty:
        st.info("No customer records found.")
    else:
        st.dataframe(df_cust, use_container_width=True)

    with st.expander("➕ Add New Customer Manually"):
        with st.form("manual_cust_form"):
            mc_name = st.text_input("Customer Name")
            mc_phone = st.text_input("Phone")
            mc_city = st.text_input("City")
            mc_address = st.text_area("Address")

            if st.form_submit_button("Save Customer"):
                if not mc_name or not mc_phone:
                    st.error("Name and Phone are mandatory.")
                else:
                    try:
                        execute_query("INSERT INTO customers (name, phone, city, address) VALUES (?,?,?,?)",
                                      (mc_name, mc_phone, mc_city, mc_address))
                        st.success("Customer saved successfully!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error saving customer: {e}")

# ------------------------------------------------------------------------------------
# MODULE 8: STAFF MANAGEMENT
# ------------------------------------------------------------------------------------
elif nav_choice == "👷 Staff Management":
    st.title("👷 Staff Records")

    df_staff = load_data("SELECT * FROM staff")
    if not df_staff.empty:
        st.dataframe(df_staff, use_container_width=True)
    else:
        st.info("No staff records added yet.")

    st.subheader("➕ Add New Staff Member")
    with st.form("add_staff_form"):
        s_name = st.text_input("Full Name")
        s_role = st.selectbox("Role", ["Labour", "Supervisor", "Production Manager", "Administrative"])
        s_phone = st.text_input("Contact Phone")

        if st.form_submit_button("Register Staff Member"):
            if not s_name:
                st.error("Please enter staff name.")
            else:
                execute_query("INSERT INTO staff (name, role, phone) VALUES (?,?,?)", (s_name, s_role, s_phone))
                st.success(f"{s_name} registered as {s_role}!")
                st.rerun()

# ------------------------------------------------------------------------------------
# MODULE 9: REPORTS & ANALYTICS
# ------------------------------------------------------------------------------------
elif nav_choice == "📊 Reports & Analytics":
    st.title("📊 Business Reports & Consolidated Analytics")

    df_orders = load_data("SELECT * FROM orders")
    df_prod = load_data("SELECT * FROM production_log")
    df_raw = load_data("SELECT * FROM raw_materials")

    tot_sales = df_orders[df_orders["status"] != "Cancelled"]["total_amount"].sum() if not df_orders.empty and "total_amount" in df_orders.columns else 0.0
    tot_orders = len(df_orders) if not df_orders.empty else 0
    tot_prod_qty = df_prod["output_qty"].sum() if not df_prod.empty and "output_qty" in df_prod.columns else 0.0

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("💰 Total Revenue", f"{CURRENCY} {tot_sales:,.2f}")
    kpi2.metric("📦 Total Orders", f"{tot_orders}")
    kpi3.metric("🏭 Total Production Output", f"{tot_prod_qty:,.1f} Units")

    st.divider()

    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        st.subheader("Sales by Product")
        if not df_orders.empty and "product_name" in df_orders.columns:
            sales_by_p = df_orders.groupby("product_name")["total_amount"].sum().reset_index()
            fig1 = px.bar(sales_by_p, x="product_name", y="total_amount", labels={"total_amount": f"Revenue ({CURRENCY})", "product_name": "Product"}, title="Revenue per Product")
            st.plotly_chart(fig1, use_container_width=True)
        else:
            st.info("No sales data to generate charts.")

    with col_chart2:
        st.subheader("Raw Material Stock Levels")
        if not df_raw.empty and "item_name" in df_raw.columns:
            fig2 = px.pie(df_raw, names="item_name", values="stock_qty", title="Raw Material Composition")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No raw material data to generate charts.")

    st.divider()

    st.subheader("📥 Export Complete Business Report")
    excel_file = generate_excel_report()
    st.download_button(
        label="Download ERP Excel Report",
        data=excel_file,
        file_name=f"Wire_Cable_ERP_Report_{date.today().strftime('%Y_%m_%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
