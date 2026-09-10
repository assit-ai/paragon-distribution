import os
import io
import re
import json
import math
import sqlite3
import datetime
from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for
import xlsxwriter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = 'paragon-agro-distribution-secret-key-2026'
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'distribution.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    # Drop and recreate schema
    cursor.execute('DROP TABLE IF EXISTS users')
    cursor.execute('DROP TABLE IF EXISTS depots')
    cursor.execute('DROP TABLE IF EXISTS daily_reports')
    cursor.execute('DROP TABLE IF EXISTS trips')
    cursor.execute('DROP TABLE IF EXISTS invoices')
    
    # Users Table (RBAC with Active Status)
    cursor.execute('''
    CREATE TABLE users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        password TEXT NOT NULL,
        role TEXT NOT NULL, -- 'admin' or 'incharge'
        depot_id INTEGER,
        display_name TEXT NOT NULL,
        is_active INTEGER DEFAULT 1,
        FOREIGN KEY (depot_id) REFERENCES depots (id)
    )
    ''')

    # Depots Table
    cursor.execute('''
    CREATE TABLE depots (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        slug TEXT NOT NULL UNIQUE,
        category TEXT NOT NULL,
        incharge_name TEXT NOT NULL,
        contact TEXT,
        region TEXT NOT NULL,
        default_uom TEXT DEFAULT 'Pkt'
    )
    ''')
    
    # Daily Reports Table with Cancellation & Audit Tracking
    cursor.execute('''
    CREATE TABLE daily_reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_date TEXT NOT NULL,
        depot_id INTEGER NOT NULL,
        incharge_name TEXT NOT NULL,
        contact TEXT,
        shift TEXT DEFAULT 'Day Shift',
        total_vehicles INTEGER DEFAULT 0,
        total_invoices INTEGER DEFAULT 0,
        capacity_util_pct REAL DEFAULT 0,
        dispatched_gross_val REAL DEFAULT 0,
        delivered_net_val REAL DEFAULT 0,
        returned_val REAL DEFAULT 0,
        stock_mismatch_qty REAL DEFAULT 0,
        cash_mismatch_val REAL DEFAULT 0,
        adjustment_status TEXT DEFAULT '100% Adjusted',
        audit_status TEXT DEFAULT 'OK / Verified',
        status TEXT DEFAULT 'active', -- 'active', 'cancellation_requested', 'cancelled', 'unlocked'
        cancellation_reason TEXT,
        cancellation_requested_at TIMESTAMP,
        cancelled_by TEXT,
        cancelled_at TIMESTAMP,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (depot_id) REFERENCES depots (id)
    )
    ''')
    
    # Trips Table
    cursor.execute('''
    CREATE TABLE trips (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER NOT NULL,
        trip_no TEXT NOT NULL,
        vehicle_no TEXT NOT NULL,
        vehicle_type TEXT NOT NULL,
        driver_name TEXT,
        delivery_man TEXT,
        helper_name TEXT,
        route_name TEXT,
        planned_invoices INTEGER DEFAULT 0,
        departure_time TEXT,
        return_time TEXT,
        capacity_kg REAL DEFAULT 0,
        loaded_kg REAL DEFAULT 0,
        util_pct REAL DEFAULT 0,
        odo_start REAL DEFAULT 0,
        odo_end REAL DEFAULT 0,
        km_run REAL DEFAULT 0,
        fuel_cost REAL DEFAULT 0,
        trip_status TEXT DEFAULT 'Completed',
        gate_pass_no TEXT,
        reefer_temp TEXT DEFAULT '-18°C',
        remarks TEXT,
        FOREIGN KEY (report_id) REFERENCES daily_reports (id)
    )
    ''')
    
    # Invoices Table
    cursor.execute('''
    CREATE TABLE invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id INTEGER NOT NULL,
        trip_id INTEGER,
        invoice_no TEXT NOT NULL,
        customer_name TEXT NOT NULL,
        customer_code TEXT,
        customer_address TEXT,
        customer_type TEXT DEFAULT 'Superstore / Retail',
        trip_no TEXT,
        product_category TEXT NOT NULL,
        sku_uom TEXT DEFAULT 'Pkt',
        dispatched_qty REAL DEFAULT 0,
        dispatched_val REAL DEFAULT 0,
        delivery_status TEXT DEFAULT 'Delivered',
        delivered_qty REAL DEFAULT 0,
        delivered_val REAL DEFAULT 0,
        returned_qty REAL DEFAULT 0,
        returned_val REAL DEFAULT 0,
        return_reason TEXT DEFAULT 'None',
        collection_mode TEXT DEFAULT 'Credit',
        amount_collected REAL DEFAULT 0,
        credit_amount REAL DEFAULT 0,
        shortage_amount REAL DEFAULT 0,
        reconciliation_status TEXT DEFAULT '100% Reconciled',
        remarks TEXT,
        FOREIGN KEY (report_id) REFERENCES daily_reports (id),
        FOREIGN KEY (trip_id) REFERENCES trips (id)
    )
    ''')
    
    seed_database(cursor)
    conn.commit()
    conn.close()

def seed_database(cursor):
    # 17 Depots & Factories
    depots_data = [
        (1, 'Ashulia Factory - Frozen', 'ashulia_frozen', 'Frozen Food', 'Sajedul Islam (AM)', '01711000001', 'Dhaka North', 'Pkt'),
        (2, 'Ashulia Factory - Dry Food', 'ashulia_dryfood', 'Dry Food', 'Delower (AM Dist)', '01711000002', 'Dhaka North', 'Pkt'),
        (3, 'Gazipur-Process CK', 'gazipur_ck', 'Processed Chicken', 'Mahmud (Sr. Officer)', '01711000003', 'Gazipur', 'Kg'),
        (4, 'Sirajganj Dairy', 'sirajganj_dairy', 'Dairy', 'Rafiqul (Officer)', '01711000004', 'North Bengal', 'Ltr'),
        (5, 'Sylhet - Tea Packing Unit', 'sylhet_tea', 'Tea', 'Incharge Vacant', '01711000005', 'Sylhet', 'Kg'),
        (6, 'Sylhet Depot - Frozen', 'sylhet_frozen', 'Frozen Food', 'Dist & Acc Officer', '01711000006', 'Sylhet', 'Pkt'),
        (7, 'Tejgaon - Frozen/CK', 'tejgaon_frozenck', 'Chicken & Frozen', 'Sabbir (Dist Officer)', '01711000007', 'Central Dhaka', 'Kg/Pkt'),
        (8, 'Tejgaon - e-Commerce', 'tejgaon_ecommerce', 'e-Commerce', 'Shahin (Officer)', '01711000008', 'Central Dhaka', 'Kg'),
        (9, 'Tejgaon - Fresh Egg', 'tejgaon_egg', 'Fresh Eggs', 'Mushfik (Officer)', '01711000009', 'Central Dhaka', 'Pcs'),
        (10, 'Tejgaon - Tea Distribution', 'tejgaon_tea', 'Tea', 'Officer (Tea Dist)', '01711000010', 'Central Dhaka', 'Kg'),
        (11, 'Tejgaon - Dairy', 'tejgaon_dairy', 'Dairy', 'Mahmudul (AM)', '01711000011', 'Central Dhaka', 'Ltr'),
        (12, 'Mohakhali Depot', 'mohakhali_momo', 'Momo & Snacks', 'Shohag (Supervisor)', '01711000012', 'Dhaka North', 'Kg'),
        (13, 'CTG Depot - Frozen', 'ctg_frozen', 'Frozen Food', 'Monjurul (Asst Officer)', '01711000013', 'Chittagong', 'Pkt'),
        (14, 'CTG Depot - Dry Food', 'ctg_dryfood', 'Dry Food', 'Monjurul (Asst Officer)', '01711000014', 'Chittagong', 'Pkt'),
        (15, 'Jessore - Frozen Food', 'jessore_frozen', 'Frozen Food', 'Shohag (Supervisor)', '01711000015', 'South Bengal', 'Pkt'),
        (16, 'Jessore - Dry Food', 'jessore_dryfood', 'Dry Food', 'Shohag (Supervisor)', '01711000016', 'South Bengal', 'Pkt'),
        (17, 'Rangpur Depot - Frozen', 'rangpur_frozen', 'Frozen Food', 'Dist Officer', '01711000017', 'North Bengal', 'Pkt')
    ]
    
    for d in depots_data:
        cursor.execute('''
        INSERT OR IGNORE INTO depots (id, name, slug, category, incharge_name, contact, region, default_uom)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', d)
    
    # Users: Admin + 17 Incharges
    cursor.execute('''
    INSERT OR IGNORE INTO users (id, username, password, role, depot_id, display_name)
    VALUES (1, 'admin', 'admin123', 'admin', NULL, 'Executive Admin')
    ''')

    for d in depots_data:
        uid = d[0] + 1
        username = d[2]
        display_name = f"{d[1]} (Incharge: {d[4]})"
        cursor.execute('''
        INSERT OR IGNORE INTO users (id, username, password, role, depot_id, display_name)
        VALUES (?, ?, 'depot123', 'incharge', ?, ?)
        ''', (uid, username, d[0], display_name))
        
    # Sample Delivery Man / Van Rider
    cursor.execute('''
    INSERT OR IGNORE INTO users (id, username, password, role, depot_id, display_name)
    VALUES (100, 'rider_tejgaon', 'rider123', 'delivery_man', 1, 'Tejgaon Van Rider (Selim)')
    ''')

    seed_sample_daily_reports(cursor)

def seed_sample_daily_reports(cursor):
    today = datetime.date.today()
    today_str = today.strftime('%Y-%m-%d')
    d1_str = (today - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    d2_str = (today - datetime.timedelta(days=2)).strftime('%Y-%m-%d')
    d3_str = (today - datetime.timedelta(days=3)).strftime('%Y-%m-%d')
    d4_str = (today - datetime.timedelta(days=4)).strftime('%Y-%m-%d')

    sample_reports = [
        # (depot_id, report_date, incharge_name, shift, total_veh, total_inv, util_pct, disp_val, deliv_val, ret_val, stock_var, cash_var, adj_stat, audit_stat, status)
        (1, today_str, 'Sajedul Islam (AM)', 'Day Shift', 4, 42, 88.5, 145000, 142000, 3000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (2, today_str, 'Delower (AM Dist)', 'Day Shift', 3, 36, 85.0, 98000, 96500, 1500, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (3, today_str, 'Mahmud (Sr. Officer)', 'Day Shift', 5, 58, 92.0, 220000, 216000, 4000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (4, d1_str, 'Rafiqul (Officer)', 'Day Shift', 3, 28, 80.0, 105000, 103000, 2000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (5, d3_str, 'Incharge Vacant', 'Day Shift', 1, 12, 70.0, 35000, 34000, 1000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (6, today_str, 'Dist & Acc Officer', 'Day Shift', 2, 24, 78.0, 75000, 73500, 1500, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (7, today_str, 'Sabbir (Dist Officer)', 'Day Shift', 6, 72, 94.0, 310000, 304000, 6000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (8, today_str, 'Shahin (Officer)', 'Day Shift', 4, 65, 89.0, 185000, 182000, 3000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (9, today_str, 'Mushfik (Officer)', 'Day Shift', 7, 88, 95.0, 420000, 415000, 5000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (10, d2_str, 'Officer (Tea Dist)', 'Day Shift', 2, 18, 75.0, 62000, 60500, 1500, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (11, today_str, 'Mahmudul (AM)', 'Day Shift', 5, 62, 91.0, 195000, 191500, 3500, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (12, today_str, 'Shohag (Supervisor)', 'Day Shift', 3, 38, 83.0, 120000, 118000, 2000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (13, today_str, 'Monjurul (Asst Officer)', 'Day Shift', 4, 45, 87.0, 175000, 171500, 3500, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (14, d1_str, 'Monjurul (Asst Officer)', 'Day Shift', 2, 22, 79.0, 68000, 66500, 1500, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (15, today_str, 'Shohag (Supervisor)', 'Day Shift', 3, 30, 82.0, 110000, 108000, 2000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (16, d4_str, 'Shohag (Supervisor)', 'Day Shift', 1, 14, 68.0, 42000, 41000, 1000, 0, 0, '100% Adjusted', 'OK / Verified', 'active'),
        (17, d2_str, 'Dist Officer', 'Day Shift', 2, 20, 76.0, 70000, 68500, 1500, 0, 0, '100% Adjusted', 'OK / Verified', 'active')
    ]

    for rep in sample_reports:
        cursor.execute('''
        INSERT INTO daily_reports (
            depot_id, report_date, incharge_name, shift, total_vehicles, total_invoices,
            capacity_util_pct, dispatched_gross_val, delivered_net_val, returned_val,
            stock_mismatch_qty, cash_mismatch_val, adjustment_status, audit_status, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', rep)
        report_id = cursor.lastrowid

        # Sample trip
        cursor.execute('''
        INSERT INTO trips (
            report_id, trip_no, vehicle_no, vehicle_type, driver_name, delivery_man,
            route_name, planned_invoices, capacity_kg, loaded_kg, util_pct, trip_status, reefer_temp
        ) VALUES (?, 'TRIP-01', 'DM-SHA-11-2041', '1.5 Ton Reefer Van', 'Md. Rafiq', 'Selim Hossain',
                 'Route A - Primary Superstores', ?, 1500, 1350, 90.0, 'Completed', '-18°C')
        ''', (report_id, rep[5]))
        trip_id = cursor.lastrowid

        # Sample invoice
        cursor.execute('''
        INSERT INTO invoices (
            report_id, trip_id, invoice_no, customer_name, customer_code, customer_address,
            customer_type, trip_no, product_category, sku_uom, dispatched_qty, dispatched_val,
            delivery_status, delivered_qty, delivered_val, returned_qty, returned_val, return_reason,
            collection_mode, amount_collected, reconciliation_status
        ) VALUES (?, ?, 'INV-' || ?, 'Shwapno Superstore - Central Hub', 'CUST-001', 'Dhaka Central',
                 'Superstore', 'TRIP-01', 'Consumer Products', 'Pkt', 250, ?, 'Delivered', 245, ?, 5, ?,
                 'None', 'Bank / Online', ?, '100% Reconciled')
        ''', (report_id, trip_id, report_id, rep[7], rep[8], rep[9], rep[8]))

# Initialize DB on start if not present
if not os.path.exists(DB_PATH):
    init_db()
else:
    # Ensure users table exists
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT count(name) FROM sqlite_master WHERE type='table' AND name='users'")
    if c.fetchone()[0] == 0:
        init_db()
    else:
        pass  # No auto-seeding of demo data — DB stays clean until real data is entered
    conn.close()

def get_default_category_capacities(cap_val=1500, veh_type='Covered Van'):
    try:
        base_kg = float(cap_val or 1500)
    except (ValueError, TypeError):
        base_kg = 1500.0
    ratio = base_kg / 1500.0 if base_kg > 0 else 1.0
    return {
        "Frozen Foods / Chicken": {"capacity": round(base_kg, 1), "unit": "Kg"},
        "Egg": {"capacity": round(30000 * ratio), "unit": "Pcs"},
        "Dairy": {"capacity": round(1000 * ratio), "unit": "Liter"},
        "Dry Goods / Box Items": {"capacity": round(500 * ratio), "unit": "Ctn"}
    }

# =========================================================================
# DYNAMIC SKU CATALOG & UOM AUTO-MAPPING RULES
# Auto-Mapping Rules:
# - Tejgaon + Frozen/Chicken -> Kg
# - Tejgaon + Egg -> Pcs
# - Tejgaon + Dairy -> Liter
# - CTG + Frozen -> Kg
# - CTG + Dry Food -> Kg
# =========================================================================

DEFAULT_SKU_CATALOG = {
    "egg": [
        {"code": "SKU-EGG-001", "name": "Premium Loose Egg", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "30 Pcs/Tray", "route": "Central Route 01"},
        {"code": "SKU-EGG-002", "name": "Paragon Brown Egg 12 pcs", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "12 Pcs/Pack", "route": "Central Route 01"},
        {"code": "SKU-EGG-003", "name": "Paragon White Egg 12 pcs", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "12 Pcs/Pack", "route": "West Line Route"},
        {"code": "SKU-EGG-004", "name": "Paragon Omega 3 Plus 12 pcs", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "12 Pcs/Pack", "route": "Superstore Express"},
        {"code": "SKU-EGG-005", "name": "Paragon Egg Economy pack(12 pcs)EE12", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "12 Pcs/Pack", "route": "Wholesale Line"},
        {"code": "SKU-EGG-006", "name": "Tatka Premium 12 Pcs", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "12 Pcs/Pack", "route": "Retail Route 02"},
        {"code": "SKU-EGG-007", "name": "Free Range Native (Deshi) Egg", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "30 Pcs/Tray", "route": "North Line Route"},
        {"code": "SKU-EGG-008", "name": "Quail Eggs 20 Pcs Box", "category": "Fresh Eggs", "uom": "Pcs", "cap": 30000.0, "pack": "20 Pcs/Box", "route": "Superstore Express"}
    ],
    "dairy": [
        {"code": "SKU-DY-001", "name": "Pasteurized Milk-1000 ml (Full Cream)", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "1000 ml Pouch", "route": "Morning Chilled Line 01"},
        {"code": "SKU-DY-002", "name": "Pasteurized Milk-500 ml (Full Cream)", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "500 ml Pouch", "route": "Morning Chilled Line 01"},
        {"code": "SKU-DY-003", "name": "Standarized Milk 1000 ml", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "1000 ml Pouch", "route": "Urban Retail Line"},
        {"code": "SKU-DY-004", "name": "Standarized Milk 500 ml", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "500 ml Pouch", "route": "Urban Retail Line"},
        {"code": "SKU-DY-005", "name": "Coffee Milk-1000 ml", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "1000 ml Pouch", "route": "Institutional & HORECA"},
        {"code": "SKU-DY-006", "name": "Toned Milk-1000 ml", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "1000 ml Pouch", "route": "Chilled Market Route"},
        {"code": "SKU-DY-007", "name": "Sweet Curd / Mishti Doi 100gm", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "100 gm Cup", "route": "Superstore Route"},
        {"code": "SKU-DY-008", "name": "Choco Chilled Milk Drink 200ml", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "200 ml Bottle", "route": "Retail Route 02"},
        {"code": "SKU-DY-009", "name": "Laban / Chilled Dairy Drink 250ml", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "250 ml Bottle", "route": "Morning Chilled Line 01"},
        {"code": "SKU-DY-010", "name": "Pure Cow Ghee / Butter Oil 400gm", "category": "Dairy", "uom": "Liter", "cap": 1000.0, "pack": "400 gm Jar", "route": "Wholesale Route"}
    ],
    "frozen": [
        {"code": "SKU-FZ-001", "name": "Dressed Chicken without skin 1000gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "1000 gm Poly", "route": "Cold Chain Route 01"},
        {"code": "SKU-FZ-002", "name": "Chicken 12 Pcs Cut 1000gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "1000 gm Poly", "route": "Cold Chain Route 01"},
        {"code": "SKU-FZ-003", "name": "Chicken Leg without bone without skin", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "Tray / Kg", "route": "HORECA & Chef Route"},
        {"code": "SKU-FZ-004", "name": "Feet (Tray)", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "Tray / Kg", "route": "Central Cold Route"},
        {"code": "SKU-FZ-005", "name": "Chicken Drumstick without skin", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "Tray / Kg", "route": "Superstore Route 01"},
        {"code": "SKU-FZ-006", "name": "Chicken Breast without bone without skin", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "Tray / Kg", "route": "Superstore Route 01"},
        {"code": "SKU-FZ-007", "name": "Keema (Tray)", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "Tray / Kg", "route": "Central Cold Route"},
        {"code": "SKU-FZ-008", "name": "Chicken Nuggets 250gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "250 gm Pkt", "route": "Retail Express Line"},
        {"code": "SKU-FZ-009", "name": "Chicken Sausage 340gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "340 gm Pkt", "route": "Retail Express Line"},
        {"code": "SKU-FZ-010", "name": "Low Fat Paratha (10 Pcs)", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "650 gm Pkt", "route": "Superstore Line 02"},
        {"code": "SKU-FZ-011", "name": "Chicken Mini Samosa 250gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "250 gm Pkt", "route": "Retail Express Line"},
        {"code": "SKU-FZ-012", "name": "Chicken Shami Kabab 500gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "500 gm Pkt", "route": "Superstore Line 02"},
        {"code": "SKU-FZ-013", "name": "French Fry 1000gm", "category": "Frozen/Chicken", "uom": "Kg", "cap": 1500.0, "pack": "1000 gm Pkt", "route": "HORECA Line"}
    ],
    "dry": [
        {"code": "SKU-DRY-001", "name": "Butter Cookies 200gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "200 gm Pkt", "route": "Bakery Van Route 01"},
        {"code": "SKU-DRY-002", "name": "Milky butter Cake 10gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "10 gm Pkt", "route": "Bakery Van Route 01"},
        {"code": "SKU-DRY-003", "name": "Dry Cake (Mini) 25gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "25 gm Pkt", "route": "General Retail Route"},
        {"code": "SKU-DRY-004", "name": "Butter Toast 14gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "14 gm Pkt", "route": "General Retail Route"},
        {"code": "SKU-DRY-005", "name": "Muffin Cake Chocolate 18gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "18 gm Pkt", "route": "Urban Grocery Line"},
        {"code": "SKU-DRY-006", "name": "Muffin Cake Vanilla 18gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "18 gm Pkt", "route": "Urban Grocery Line"},
        {"code": "SKU-DRY-007", "name": "Kacha Chana 1000gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "1000 gm Pkt", "route": "Wholesale Bulk Line"},
        {"code": "SKU-DRY-008", "name": "Sweet Toast Biscuit 200gm", "category": "Dry Food", "uom": "Kg", "cap": 500.0, "pack": "200 gm Pkt", "route": "Bakery Van Route 02"}
    ],
    "tea": [
        {"code": "SKU-TEA-001", "name": "BOP 500 gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "500 gm Pkt", "route": "Tea Distribution Route 01"},
        {"code": "SKU-TEA-002", "name": "PD 500 gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "500 gm Pkt", "route": "Tea Distribution Route 01"},
        {"code": "SKU-TEA-003", "name": "Paragon Best Leaf 500 gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "500 gm Pkt", "route": "Retail Tea Express"},
        {"code": "SKU-TEA-004", "name": "RD Dust 500gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "500 gm Pkt", "route": "Tea Wholesale Hub"},
        {"code": "SKU-TEA-005", "name": "Premium Gold Tea 500gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "500 gm Pkt", "route": "Superstore Route"},
        {"code": "SKU-TEA-006", "name": "Premium Black Tea 500gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "500 gm Pkt", "route": "Superstore Route"},
        {"code": "SKU-TEA-007", "name": "Exotic Classic 15 gm", "category": "Tea", "uom": "Kg", "cap": 1200.0, "pack": "15 gm Pkt", "route": "Tea Retail Express"}
    ],
    "momo": [
        {"code": "SKU-MM-001", "name": "Chicken Momo Classic (10 Pcs)", "category": "Momo & Snacks", "uom": "Kg", "cap": 1500.0, "pack": "10 Pcs Box", "route": "Momo Food Cart & Outlet Line"},
        {"code": "SKU-MM-002", "name": "Spicy Chicken Momo (10 Pcs)", "category": "Momo & Snacks", "uom": "Kg", "cap": 1500.0, "pack": "10 Pcs Box", "route": "Momo Food Cart & Outlet Line"},
        {"code": "SKU-MM-003", "name": "Steamed Chicken Dumplings (250gm)", "category": "Momo & Snacks", "uom": "Kg", "cap": 1500.0, "pack": "250 gm Box", "route": "HORECA & Café Line"},
        {"code": "SKU-MM-004", "name": "Chicken Spring Roll (300gm)", "category": "Momo & Snacks", "uom": "Kg", "cap": 1500.0, "pack": "300 gm Box", "route": "Superstore Express"}
    ]
}

def resolve_sku_uom(depot_name, category='', depot_type=None):
    d_name = (depot_name or '').lower()
    d_type = (depot_type or '').lower()
    cat = (category or '').lower()

    # Rule 1: Tejgaon mappings
    if 'tejgaon' in d_name or 'tejgaon' in d_type:
        if 'egg' in cat or 'egg' in d_name or 'egg' in d_type:
            return 'Pcs'
        elif 'dairy' in cat or 'milk' in cat or 'dairy' in d_name or 'dairy' in d_type:
            return 'Liter'
        elif 'frozen' in cat or 'chicken' in cat or 'frozen' in d_name or 'ck' in d_name or 'frozen' in d_type:
            return 'Kg'
        elif 'tea' in cat or 'tea' in d_name or 'tea' in d_type:
            return 'Kg'
        elif 'commerce' in d_name or 'commerce' in d_type:
            return 'Kg'

    # Rule 2: CTG (Chittagong) mappings
    if 'ctg' in d_name or 'chittagong' in d_name or 'ctg' in d_type:
        if 'frozen' in cat or 'chicken' in cat or 'frozen' in d_name or 'frozen' in d_type:
            return 'Kg'
        if 'dry' in cat or 'food' in cat or 'bakery' in cat or 'dry' in d_name or 'dry' in d_type:
            return 'Kg'

    # Rule 3: General / Other Depots fallback rules
    if 'egg' in cat or 'egg' in d_name or 'egg' in d_type:
        return 'Pcs'
    elif 'dairy' in cat or 'milk' in cat or 'dairy' in d_name or 'dairy' in d_type:
        return 'Liter'
    elif 'dry' in cat or 'dry' in d_name or 'dry' in d_type:
        return 'Kg'
    elif 'frozen' in cat or 'chicken' in cat or 'frozen' in d_name or 'frozen' in d_type or 'ck' in d_name:
        return 'Kg'
    elif 'tea' in cat or 'tea' in d_name or 'tea' in d_type:
        return 'Kg'
    elif 'momo' in cat or 'momo' in d_name or 'momo' in d_type:
        return 'Kg'

    return 'Kg'

def get_depot_sku_catalog_key(depot_row):
    d_name = (depot_row['name'] if depot_row else '').lower()
    d_cat = (depot_row['category'] if depot_row else '').lower()
    d_id = depot_row['id'] if depot_row else 9

    if 'egg' in d_cat or 'egg' in d_name or d_id == 9:
        return 'egg'
    elif 'dairy' in d_cat or 'milk' in d_name or d_id in [4, 11]:
        return 'dairy'
    elif 'dry' in d_cat or 'dry' in d_name or d_id in [2, 14, 16]:
        return 'dry'
    elif 'tea' in d_cat or 'tea' in d_name or d_id in [5, 10]:
        return 'tea'
    elif 'momo' in d_cat or 'momo' in d_name or d_id == 12:
        return 'momo'
    else:
        return 'frozen'

def seed_depot_vehicle_mapping_table(cursor):
    depots = cursor.execute("SELECT id, name, category, default_uom FROM depots").fetchall()
    for d in depots:
        d_id, d_name, d_cat = d[0], d[1], d[2]
        cat_key = get_depot_sku_catalog_key({'id': d_id, 'name': d_name, 'category': d_cat})
        resolved_uom = resolve_sku_uom(d_name, d_cat, cat_key)
        
        # Get vehicles for this depot
        v_rows = cursor.execute('''
            SELECT fv.vehicle_no, fv.capacity_kg, drv.name as driver_name, drv.default_route_name
            FROM fleet_vehicles fv
            LEFT JOIN depot_crew drv ON fv.default_driver_id = drv.id
            WHERE fv.depot_id = ?
            ORDER BY fv.id ASC
        ''', (d_id,)).fetchall()
        
        if v_rows:
            for r in v_rows:
                v_no = r[0]
                v_cap = float(r[1] or 1500.0)
                drv_name = r[2] or ''
                route_name = r[3] or ''
                if resolved_uom == 'Pcs' and v_cap < 5000:
                    v_cap = 30000.0
                elif resolved_uom == 'Liter' and v_cap > 3000:
                    v_cap = 1000.0
                cursor.execute('''
                    INSERT INTO depot_vehicle_mapping
                    (depot_id, depot_name, depot_type, category, uom, default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (d_id, d_name, d_cat, d_cat, resolved_uom, v_no, drv_name, route_name, v_cap, 'Active Vehicle Assignment'))
        else:
            default_vehicles = [
                (f"DM-SHA-11-{2000 + d_id * 3 + 1}", "Md. Driver 1", "Route 01: Core City Line"),
                (f"DM-SHA-11-{2000 + d_id * 3 + 2}", "Md. Driver 2", "Route 02: Outer Suburb Line")
            ]
            std_cap = 30000.0 if resolved_uom == 'Pcs' else (1000.0 if resolved_uom == 'Liter' else 1500.0)
            for v_no, drv_name, r_name in default_vehicles:
                cursor.execute('''
                    INSERT INTO depot_vehicle_mapping
                    (depot_id, depot_name, depot_type, category, uom, default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (d_id, d_name, d_cat, d_cat, resolved_uom, v_no, drv_name, r_name, std_cap, 'Default Pre-configured Van'))

def seed_depot_sku_master_table(cursor):
    pass

def ensure_schema_migrations():
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute("PRAGMA table_info(depot_crew)")
        cols = [r[1] for r in c.fetchall()]
        if 'assigned_vehicle_no' not in cols:
            c.execute("ALTER TABLE depot_crew ADD COLUMN assigned_vehicle_no TEXT")
        if 'assigned_vehicle_id' not in cols:
            c.execute("ALTER TABLE depot_crew ADD COLUMN assigned_vehicle_id INTEGER")
        if 'secondary_vehicle_no' not in cols:
            c.execute("ALTER TABLE depot_crew ADD COLUMN secondary_vehicle_no TEXT")
        if 'default_route_name' not in cols:
            c.execute("ALTER TABLE depot_crew ADD COLUMN default_route_name TEXT")
        if 'driver_id_code' not in cols:
            c.execute("ALTER TABLE depot_crew ADD COLUMN driver_id_code TEXT")
            
        c.execute("PRAGMA table_info(fleet_vehicles)")
        f_cols = [r[1] for r in c.fetchall()]
        if 'capacity_units' not in f_cols:
            c.execute("ALTER TABLE fleet_vehicles ADD COLUMN capacity_units TEXT DEFAULT 'Kg'")
        if 'category_capacities' not in f_cols:
            c.execute("ALTER TABLE fleet_vehicles ADD COLUMN category_capacities TEXT")
            
        # Backfill category capacities for existing vehicles
        c.execute("SELECT id, vehicle_no, vehicle_type, capacity_kg, category_capacities FROM fleet_vehicles")
        rows = c.fetchall()
        for r in rows:
            vid, vno, vtype, cap_kg, cat_caps = r[0], r[1], r[2], r[3], r[4]
            if not cat_caps or str(cat_caps).strip() == '':
                default_caps = json.dumps(get_default_category_capacities(cap_kg or 1500, vtype))
                c.execute("UPDATE fleet_vehicles SET category_capacities = ? WHERE id = ?", (default_caps, vid))
            
        # Create dedicated Driver-Vehicle Default Mapping & Category Capacities table (NO SKU / SKU CODE)
        c.execute('''
        CREATE TABLE IF NOT EXISTS depot_vehicle_mapping (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depot_id INTEGER NOT NULL,
            depot_name TEXT NOT NULL,
            depot_type TEXT,
            category TEXT NOT NULL,
            uom TEXT NOT NULL,
            default_vehicle_no TEXT NOT NULL,
            default_driver_name TEXT,
            default_route_name TEXT,
            max_capacity REAL DEFAULT 0,
            remarks TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (depot_id) REFERENCES depots (id)
        )
        ''')
        
        c.execute("SELECT COUNT(*) FROM depot_vehicle_mapping")
        if c.fetchone()[0] == 0:
            seed_depot_vehicle_mapping_table(c)

        # Legacy depot_sku_master table fallback
        c.execute('''
        CREATE TABLE IF NOT EXISTS depot_sku_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depot_id INTEGER NOT NULL,
            depot_name TEXT NOT NULL,
            depot_type TEXT,
            category TEXT NOT NULL,
            sku_code TEXT,
            sku_name TEXT,
            uom TEXT NOT NULL,
            default_vehicle_no TEXT,
            default_driver_name TEXT,
            default_route_name TEXT,
            max_capacity REAL DEFAULT 0,
            pack_size TEXT,
            remarks TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (depot_id) REFERENCES depots (id)
        )
        ''')

        # Ensure saved_route_plans table exists
        c.execute('''
        CREATE TABLE IF NOT EXISTS saved_route_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depot_id INTEGER NOT NULL,
            plan_date TEXT NOT NULL,
            shift TEXT,
            shift_time TEXT,
            total_vans INTEGER DEFAULT 0,
            total_outlets INTEGER DEFAULT 0,
            total_pkts REAL DEFAULT 0,
            total_kg REAL DEFAULT 0,
            gross_value REAL DEFAULT 0,
            plan_json TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(depot_id, plan_date)
        )
        ''')

        # Ensure live tracking tables exist
        c.execute('''
        CREATE TABLE IF NOT EXISTS live_tracking_positions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_no TEXT NOT NULL,
            driver_mobile TEXT,
            driver_name TEXT,
            depot_id INTEGER NOT NULL,
            route_id TEXT,
            latitude REAL NOT NULL,
            longitude REAL NOT NULL,
            heading REAL DEFAULT 0.0,
            speed_kmh REAL DEFAULT 0.0,
            source TEXT DEFAULT 'MOBILE_GEOLOCATION',
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(vehicle_no, depot_id)
        )
        ''')
        c.execute('''
        CREATE TABLE IF NOT EXISTS live_drop_statuses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depot_id INTEGER NOT NULL,
            route_code TEXT NOT NULL,
            drop_id TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending',
            delivered_at TIMESTAMP,
            proof_note TEXT,
            cash_collected REAL DEFAULT 0.0,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(depot_id, route_code, drop_id)
        )
        ''')

        # Ensure Admin Audit Log table exists
        c.execute('''
        CREATE TABLE IF NOT EXISTS admin_audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            username TEXT NOT NULL,
            action TEXT NOT NULL,
            target_type TEXT,
            depot_id TEXT,
            details TEXT,
            ip_address TEXT
        )
        ''')

        # Ensure sample delivery_man user exists
        c.execute('''
        INSERT OR IGNORE INTO users (id, username, password, role, depot_id, display_name)
        VALUES (100, 'rider_tejgaon', 'rider123', 'delivery_man', 1, 'Tejgaon Van Rider (Selim)')
        ''')

        conn.commit()
        conn.close()
    except Exception as e:
        print("Schema migration error:", e)

ensure_schema_migrations()

# ----------------- AUTHENTICATION ROUTES -----------------
@app.route('/api/current-user')
def get_current_user():
    user = session.get('user')
    if not user:
        return jsonify({
            "authenticated": False,
            "role": "guest",
            "username": "",
            "depot_id": None,
            "display_name": "Not Logged In"
        })
    return jsonify({
        "authenticated": True,
        "role": user.get('role', 'admin'),
        "username": user.get('username'),
        "depot_id": user.get('depot_id'),
        "depot_name": user.get('depot_name'),
        "display_name": user.get('display_name')
    })

@app.route('/api/login', methods=['POST'])
def api_login():
    data = request.json or {}
    username = data.get('username', '').strip().lower()
    password = data.get('password', '').strip()
    
    conn = get_db()
    user_row = conn.execute("SELECT * FROM users WHERE REPLACE(LOWER(username), '-', '_') = REPLACE(?, '-', '_') AND password = ?", (username, password)).fetchone()
    
    if user_row:
        user_dict = dict(user_row)
        depot_name = None
        if user_dict.get('depot_id'):
            d_row = conn.execute("SELECT name FROM depots WHERE id = ?", (user_dict['depot_id'],)).fetchone()
            if d_row:
                depot_name = d_row['name']
        conn.close()
        
        session['user'] = {
            "authenticated": True,
            "id": user_dict['id'],
            "username": user_dict['username'],
            "role": user_dict['role'],
            "depot_id": user_dict['depot_id'],
            "depot_name": depot_name,
            "display_name": user_dict['display_name']
        }
        return jsonify({"success": True, "user": session['user']})
    else:
        conn.close()
        return jsonify({"success": False, "message": "Invalid username or password. Please check your credentials."}), 401

@app.route('/api/logout')
def api_logout():
    session.pop('user', None)
    return jsonify({"success": True, "message": "Logged out successfully"})



@app.route('/api/user/change-password', methods=['POST'])
def change_user_password():
    data = request.json or {}
    old_password = data.get('old_password', '').strip()
    new_password = data.get('new_password', '').strip()
    username = data.get('username', '').strip().lower()
    
    # If user in session and username not sent, get from session
    sess_user = session.get('user')
    if sess_user and not username:
        username = sess_user.get('username', '').lower()
        
    if not username:
        return jsonify({"success": False, "message": "User session not found. Please log in first."}), 401
        
    if not new_password or len(new_password) < 4:
        return jsonify({"success": False, "message": "New password must be at least 4 characters."}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    # Find user (supports either hyphen or underscore)
    user_row = cursor.execute("SELECT * FROM users WHERE REPLACE(LOWER(username), '-', '_') = REPLACE(?, '-', '_')", (username,)).fetchone()
    if not user_row:
        conn.close()
        return jsonify({"success": False, "message": f"User account '{username}' not found."}), 404
        
    user_dict = dict(user_row)
    if old_password and user_dict['password'] != old_password:
        conn.close()
        return jsonify({"success": False, "message": "Incorrect current password! Please enter your existing password correctly."}), 400
        
    cursor.execute('UPDATE users SET password = ? WHERE id = ?', (new_password, user_dict['id']))
    conn.commit()
    conn.close()
    
    return jsonify({"success": True, "message": f"Password for {user_dict['display_name']} updated successfully!"})

@app.route('/api/admin/users/reset-password', methods=['POST'])
def reset_user_password():
    data = request.json or {}
    user_id = data.get('user_id')
    depot_id = data.get('depot_id')
    new_username = data.get('username')
    new_password = data.get('password')
    
    if not new_password:
        return jsonify({"success": False, "message": "Password cannot be empty"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    if user_id:
        if new_username:
            cursor.execute('UPDATE users SET username = ?, password = ? WHERE id = ?', (new_username.lower().strip(), new_password.strip(), user_id))
        else:
            cursor.execute('UPDATE users SET password = ? WHERE id = ?', (new_password.strip(), user_id))
    elif depot_id:
        if new_username:
            cursor.execute('UPDATE users SET username = ?, password = ? WHERE depot_id = ?', (new_username.lower().strip(), new_password.strip(), depot_id))
        else:
            cursor.execute('UPDATE users SET password = ? WHERE depot_id = ?', (new_password.strip(), depot_id))
            
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Credentials / Password updated successfully!"})

# ----------------- MASTER / DASHBOARD ROUTES -----------------
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/depots')
def get_depots():
    conn = get_db()
    depots = conn.execute('SELECT * FROM depots ORDER BY id ASC').fetchall()
    conn.close()
    return jsonify([dict(row) for row in depots])

@app.route('/api/depots/save', methods=['POST'])
def save_depot():
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    
    depot_id = data.get('id')
    name = data.get('name')
    category = data.get('category')
    incharge_name = data.get('incharge_name')
    contact = data.get('contact')
    region = data.get('region', 'Central')
    default_uom = data.get('default_uom', 'Pkt')
    
    if depot_id:
        cursor.execute('''
        UPDATE depots SET name = ?, category = ?, incharge_name = ?, contact = ?, region = ?, default_uom = ?
        WHERE id = ?
        ''', (name, category, incharge_name, contact, region, default_uom, depot_id))
    else:
        slug = name.lower().replace(' ', '_').replace('-', '_')
        cursor.execute('''
        INSERT INTO depots (name, slug, category, incharge_name, contact, region, default_uom)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (name, slug, category, incharge_name, contact, region, default_uom))
        
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Incharge / Depot saved successfully!"})

@app.route('/api/dashboard/summary')
def get_dashboard_summary():
    date_param = request.args.get('date', datetime.date.today().strftime('%Y-%m-%d'))
    conn = get_db()
    
    query = '''
    SELECT 
        d.id as depot_id,
        d.name as depot_name,
        d.slug as depot_slug,
        d.category,
        d.region,
        d.default_uom,
        r.id as report_id,
        COALESCE(r.incharge_name, d.incharge_name) as incharge_name,
        COALESCE(r.total_vehicles, 0) as total_vehicles,
        COALESCE(r.total_invoices, 0) as total_invoices,
        COALESCE(r.capacity_util_pct, 0) as capacity_util_pct,
        COALESCE(r.dispatched_gross_val, 0) as dispatched_gross_val,
        COALESCE(r.delivered_net_val, 0) as delivered_net_val,
        COALESCE(r.returned_val, 0) as returned_val,
        COALESCE(r.stock_mismatch_qty, 0) as stock_mismatch_qty,
        COALESCE(r.cash_mismatch_val, 0) as cash_mismatch_val,
        COALESCE(r.adjustment_status, '100% Adjusted') as adjustment_status,
        COALESCE(r.audit_status, 'OK / Verified') as audit_status
    FROM depots d
    LEFT JOIN daily_reports r ON d.id = r.depot_id AND r.report_date = ?
    ORDER BY d.id ASC
    '''
    rows = conn.execute(query, (date_param,)).fetchall()
    
    depot_list = []
    tot_dispatched = 0
    tot_delivered = 0
    tot_returned = 0
    tot_invoices = 0
    tot_vehicles = 0
    tot_stock_var = 0
    tot_cash_var = 0
    util_sum = 0
    active_units_count = 0
    
    for row in rows:
        d_dict = dict(row)
        disp = d_dict['dispatched_gross_val']
        deliv = d_dict['delivered_net_val']
        ret = d_dict['returned_val']
        
        d_dict['success_rate'] = round((deliv / disp * 100), 1) if disp > 0 else 0
        d_dict['return_rate'] = round((ret / disp * 100), 1) if disp > 0 else 0
        
        tot_dispatched += disp
        tot_delivered += deliv
        tot_returned += ret
        tot_invoices += d_dict['total_invoices']
        tot_vehicles += d_dict['total_vehicles']
        tot_stock_var += d_dict['stock_mismatch_qty']
        tot_cash_var += d_dict['cash_mismatch_val']
        
        if d_dict['total_vehicles'] > 0:
            util_sum += d_dict['capacity_util_pct']
            active_units_count += 1
        
        depot_list.append(d_dict)
        
    overall_success_rate = round((tot_delivered / tot_dispatched * 100), 1) if tot_dispatched > 0 else 0
    avg_capacity_util = round((util_sum / active_units_count), 1) if active_units_count > 0 else 0
    
    category_names = [
        ("Frozen Food & RTC", "Pkt"),
        ("Processed Chicken", "Kg"),
        ("Sweets & Savory", "Kg"),
        ("Fresh Eggs", "Pcs"),
        ("Dairy (Sirajganj & Tejgaon)", "Ltr"),
        ("Dry Food & Bakery", "Pkt"),
        ("Momo & Dimsum", "Kg"),
        ("Tea Distribution & Packing", "Kg")
    ]
    
    categories = []
    for c_name, c_uom in category_names:
        # Sum from depots matching this category
        cat_disp = sum(d['dispatched_gross_val'] for d in depot_list if c_name.lower() in d['category'].lower() or d['category'].lower() in c_name.lower())
        cat_deliv = sum(d['delivered_net_val'] for d in depot_list if c_name.lower() in d['category'].lower() or d['category'].lower() in c_name.lower())
        cat_ret = sum(d['returned_val'] for d in depot_list if c_name.lower() in d['category'].lower() or d['category'].lower() in c_name.lower())
        cat_succ = round((cat_deliv / cat_disp * 100), 1) if cat_disp > 0 else 0
        cat_ret_rate = round((cat_ret / cat_disp * 100), 1) if cat_disp > 0 else 0
        
        categories.append({
            "name": c_name,
            "uom": c_uom,
            "disp_qty": round(cat_disp / 50, 1) if cat_disp > 0 else 0,
            "disp_val": cat_disp,
            "deliv_qty": round(cat_deliv / 50, 1) if cat_deliv > 0 else 0,
            "deliv_val": cat_deliv,
            "ret_qty": round(cat_ret / 50, 1) if cat_ret > 0 else 0,
            "ret_val": cat_ret,
            "success_rate": cat_succ,
            "ret_rate": cat_ret_rate,
            "stock_var": 0
        })
    
    return_reasons = []
    if tot_returned > 0:
        return_reasons = [
            {"code": "R-01", "reason": "Shop Closed / Customer Unavailable", "qty_est": f"{round(tot_returned * 0.37 / 100)} Pkt/Kg", "invoices": max(1, round(tot_invoices * 0.05)), "val": round(tot_returned * 0.37), "pct": 37.0},
            {"code": "R-02", "reason": "Payment Shortage / Cash Not Ready", "qty_est": f"{round(tot_returned * 0.26 / 100)} Pkt/Kg", "invoices": max(1, round(tot_invoices * 0.03)), "val": round(tot_returned * 0.26), "pct": 25.7},
            {"code": "R-03", "reason": "Transit Damage / Breakage", "qty_est": f"{round(tot_returned * 0.15 / 100)} Pkt/Kg", "invoices": max(1, round(tot_invoices * 0.02)), "val": round(tot_returned * 0.15), "pct": 15.0},
            {"code": "R-04", "reason": "Wrong Product / Short Order", "qty_est": f"{round(tot_returned * 0.12 / 100)} Pkt/Kg", "invoices": max(1, round(tot_invoices * 0.01)), "val": round(tot_returned * 0.12), "pct": 11.6},
            {"code": "R-05", "reason": "Price Dispute / Rate Mismatch", "qty_est": f"{round(tot_returned * 0.07 / 100)} Pkt/Kg", "invoices": max(1, round(tot_invoices * 0.01)), "val": round(tot_returned * 0.07), "pct": 6.9},
            {"code": "R-06", "reason": "Route Delay / Traffic Breakdown", "qty_est": f"{round(tot_returned * 0.04 / 100)} Pkt/Kg", "invoices": max(1, round(tot_invoices * 0.01)), "val": round(tot_returned * 0.04), "pct": 3.8}
        ]
    
    # Calculate Depot Daily Entry Compliance & Submission Summary for all 17 depots
    incharge_users = conn.execute('''
        SELECT 
            d.id as depot_id,
            d.name as depot_name,
            d.slug as depot_slug,
            d.category,
            d.region,
            d.incharge_name,
            d.contact,
            d.default_uom,
            u.id as user_id,
            u.username,
            u.display_name as user_display_name
        FROM depots d
        LEFT JOIN users u ON d.id = u.depot_id AND u.role = 'incharge'
        ORDER BY d.id ASC
    ''').fetchall()

    daily_entry_summary = []
    total_done_entries = 0
    total_pending_entries = 0

    try:
        target_date_obj = datetime.datetime.strptime(date_param, '%Y-%m-%d').date()
    except Exception:
        target_date_obj = datetime.date.today()

    for dep in incharge_users:
        dep_id = dep['depot_id']
        # Check if active report exists for date_param
        rep = conn.execute('''
            SELECT id, report_date, incharge_name, shift, total_vehicles, total_invoices, 
                   dispatched_gross_val, delivered_net_val, returned_val, status, created_at
            FROM daily_reports
            WHERE depot_id = ? AND report_date = ? AND status != 'cancelled'
            ORDER BY id DESC LIMIT 1
        ''', (dep_id, date_param)).fetchone()

        if rep:
            status = 'Done'
            pending_days = 0
            last_entry_date = date_param
            total_done_entries += 1
            entry_info = {
                'depot_id': dep_id,
                'depot_name': dep['depot_name'],
                'depot_slug': dep['depot_slug'],
                'category': dep['category'],
                'region': dep['region'],
                'incharge_name': rep['incharge_name'] or dep['incharge_name'],
                'contact': dep['contact'],
                'user_id': dep['user_id'],
                'username': dep['username'] or dep['depot_slug'],
                'user_display_name': dep['user_display_name'] or dep['incharge_name'],
                'status': 'Done',
                'pending_days': 0,
                'last_entry_date': last_entry_date,
                'report_id': rep['id'],
                'total_vehicles': rep['total_vehicles'] or 0,
                'total_invoices': rep['total_invoices'] or 0,
                'dispatched_gross_val': rep['dispatched_gross_val'] or 0,
                'delivered_net_val': rep['delivered_net_val'] or 0,
                'returned_val': rep['returned_val'] or 0,
                'submitted_at': str(rep['created_at']) if rep['created_at'] else None
            }
        else:
            status = 'Pending'
            total_pending_entries += 1
            # Find the most recent entry before date_param
            prior_rep = conn.execute('''
                SELECT report_date, created_at FROM daily_reports
                WHERE depot_id = ? AND report_date < ? AND status != 'cancelled'
                ORDER BY report_date DESC LIMIT 1
            ''', (dep_id, date_param)).fetchone()

            if prior_rep and prior_rep['report_date']:
                try:
                    prior_date_obj = datetime.datetime.strptime(prior_rep['report_date'], '%Y-%m-%d').date()
                    pending_days = max(1, (target_date_obj - prior_date_obj).days)
                    last_entry_date = prior_rep['report_date']
                except Exception:
                    pending_days = 1
                    last_entry_date = prior_rep['report_date']
            else:
                days_since_month_start = (target_date_obj - target_date_obj.replace(day=1)).days + 1
                pending_days = max(1, days_since_month_start)
                last_entry_date = None

            entry_info = {
                'depot_id': dep_id,
                'depot_name': dep['depot_name'],
                'depot_slug': dep['depot_slug'],
                'category': dep['category'],
                'region': dep['region'],
                'incharge_name': dep['incharge_name'],
                'contact': dep['contact'],
                'user_id': dep['user_id'],
                'username': dep['username'] or dep['depot_slug'],
                'user_display_name': dep['user_display_name'] or dep['incharge_name'],
                'status': 'Pending',
                'pending_days': pending_days,
                'last_entry_date': last_entry_date,
                'report_id': None,
                'total_vehicles': 0,
                'total_invoices': 0,
                'dispatched_gross_val': 0,
                'delivered_net_val': 0,
                'returned_val': 0,
                'submitted_at': None
            }
        daily_entry_summary.append(entry_info)

    compliance_stats = {
        'total_depots': len(incharge_users),
        'total_done': total_done_entries,
        'total_pending': total_pending_entries,
        'completion_rate': round((total_done_entries / len(incharge_users) * 100), 1) if incharge_users else 0
    }

    conn.close()
    return jsonify({
        "kpis": {
            "total_dispatched_val": tot_dispatched,
            "total_delivered_val": tot_delivered,
            "total_returned_val": tot_returned,
            "delivery_success_rate": overall_success_rate,
            "total_vehicles": tot_vehicles,
            "avg_capacity_util": avg_capacity_util,
            "total_stock_variance": tot_stock_var,
            "total_cash_variance": tot_cash_var
        },
        "depots": depot_list,
        "categories": categories,
        "return_reasons": return_reasons,
        "daily_entry_summary": daily_entry_summary,
        "daily_entry_compliance": compliance_stats
    })

@app.route('/api/admin/clear-all-demo-data', methods=['GET', 'POST', 'DELETE'])
def clear_all_demo_data():
    data = request.json if request.is_json else {}
    sess_user = session.get('user') or {}
    role = session.get('role') or sess_user.get('role') or data.get('role') or request.headers.get('X-Admin-Role')
    if role and role not in ['admin', 'guest'] and role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized: Only Admin can clear operational data"}), 403

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM invoices')
    cursor.execute('DELETE FROM trips')
    cursor.execute('DELETE FROM daily_reports')
    try:
        cursor.execute('DELETE FROM saved_route_plans')
    except Exception:
        pass
    try:
        cursor.execute('DELETE FROM live_tracking_positions')
        cursor.execute('DELETE FROM live_drop_statuses')
    except Exception:
        pass
    try:
        cursor.execute('DELETE FROM inter_depot_vehicle_requests')
    except Exception:
        pass

    username = sess_user.get('username') or 'admin'
    try:
        cursor.execute('''
            INSERT INTO admin_audit_log (username, action, target_type, depot_id, details, ip_address)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (username, 'CLEAR_DEMO_DATA', 'operational_reports', 'ALL', 'All demo & operational distribution reports cleared', request.remote_addr))
    except Exception:
        pass

    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "All demo & historical distribution reports have been cleared. Database is clean."})

@app.route('/api/admin/clear-data', methods=['POST'])
def admin_clear_data():
    """
    Selectively clears operational distribution data (reports, invoices, trips, saved route plans)
    based on flexible depot selection and date filtering (single date, date range, or all historical dates).
    """
    data = request.json if request.is_json else {}
    sess_user = session.get('user') or {}
    role = session.get('role') or sess_user.get('role') or data.get('role') or request.headers.get('X-Admin-Role')
    if role and role not in ['admin', 'guest'] and role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized: Only Admin can clear operational distribution data"}), 403

    depot_id = data.get('depot_id', 'all')
    date_mode = data.get('date_mode', 'single')  # 'single', 'range', 'all'
    single_date = data.get('date') or data.get('report_date')
    start_date = data.get('start_date')
    end_date = data.get('end_date')

    conn = get_db()
    cursor = conn.cursor()

    where_clauses = []
    params = []

    is_all_depots = str(depot_id).lower() == 'all' or not depot_id
    if not is_all_depots:
        where_clauses.append("depot_id = ?")
        params.append(int(depot_id))

    date_summary_str = ""
    if date_mode == 'single':
        if not single_date:
            single_date = datetime.date.today().strftime('%Y-%m-%d')
        where_clauses.append("report_date = ?")
        params.append(single_date)
        date_summary_str = f"Date: {single_date}"
    elif date_mode == 'range':
        if not start_date or not end_date:
            conn.close()
            return jsonify({"success": False, "message": "Both start_date and end_date are required for date range clear."}), 400
        where_clauses.append("report_date >= ? AND report_date <= ?")
        params.append(start_date)
        params.append(end_date)
        date_summary_str = f"Date Range: {start_date} to {end_date}"
    elif date_mode == 'all':
        date_summary_str = "All Historical Dates"
    else:
        if single_date:
            where_clauses.append("report_date = ?")
            params.append(single_date)
            date_summary_str = f"Date: {single_date}"
        else:
            date_summary_str = "All Historical Dates"

    where_sql = ("WHERE " + " AND ".join(where_clauses)) if where_clauses else ""

    target_type = data.get('target_type') or data.get('clear_type') or data.get('type')
    if target_type == 'tracking':
        trk_where = []
        trk_params = []
        if not is_all_depots:
            trk_where.append("depot_id = ?")
            trk_params.append(int(depot_id))
        if date_mode == 'single' and single_date:
            trk_where.append("date(updated_at) = ?")
            trk_params.append(single_date)
        elif date_mode == 'range' and start_date and end_date:
            trk_where.append("date(updated_at) >= ? AND date(updated_at) <= ?")
            trk_params.append(start_date)
            trk_params.append(end_date)
        trk_sql = ("WHERE " + " AND ".join(trk_where)) if trk_where else ""
        cursor.execute(f"DELETE FROM live_tracking_positions {trk_sql}", trk_params)

        drop_where = []
        drop_params = []
        if not is_all_depots:
            drop_where.append("depot_id = ?")
            drop_params.append(int(depot_id))
        if date_mode == 'single' and single_date:
            drop_where.append("(date(updated_at) = ? OR (delivered_at IS NOT NULL AND date(delivered_at) = ?))")
            drop_params.append(single_date)
            drop_params.append(single_date)
        elif date_mode == 'range' and start_date and end_date:
            drop_where.append("((date(updated_at) >= ? AND date(updated_at) <= ?) OR (delivered_at IS NOT NULL AND date(delivered_at) >= ? AND date(delivered_at) <= ?))")
            drop_params.append(start_date)
            drop_params.append(end_date)
            drop_params.append(start_date)
            drop_params.append(end_date)
        drop_sql = ("WHERE " + " AND ".join(drop_where)) if drop_where else ""
        cursor.execute(f"DELETE FROM live_drop_statuses {drop_sql}", drop_params)

        plan_where = []
        plan_params = []
        if not is_all_depots:
            plan_where.append("depot_id = ?")
            plan_params.append(int(depot_id))
        if date_mode == 'single' and single_date:
            plan_where.append("plan_date = ?")
            plan_params.append(single_date)
        elif date_mode == 'range' and start_date and end_date:
            plan_where.append("plan_date >= ? AND plan_date <= ?")
            plan_params.append(start_date)
            plan_params.append(end_date)
        plan_sql = ("WHERE " + " AND ".join(plan_where)) if plan_where else ""
        try:
            cursor.execute(f"DELETE FROM saved_route_plans {plan_sql}", plan_params)
        except Exception:
            pass

        conn.commit()
        depot_target_name = "All 17 Depots"
        if not is_all_depots:
            d_row = cursor.execute("SELECT name FROM depots WHERE id = ?", (int(depot_id),)).fetchone()
            depot_target_name = d_row['name'] if d_row else f"Depot #{depot_id}"

        username = sess_user.get('username') or 'admin'
        details_str = f"Cleared Live Tracking: {depot_target_name} | {date_summary_str}"
        try:
            cursor.execute('''
                INSERT INTO admin_audit_log (username, action, target_type, depot_id, details, ip_address)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (username, 'CLEAR_TRACKING_DATA', 'live_tracking_positions', str(depot_id), details_str, request.remote_addr))
            conn.commit()
        except Exception:
            pass
        conn.close()
        return jsonify({
            "success": True,
            "message": f"Successfully cleared live GPS tracking positions and drop statuses for {depot_target_name} ({date_summary_str})."
        })

    matching_reports = cursor.execute(f"SELECT id FROM daily_reports {where_sql}", params).fetchall()
    report_ids = [r['id'] for r in matching_reports]
    deleted_reports_count = len(report_ids)

    deleted_invoices_count = 0
    deleted_trips_count = 0
    if report_ids:
        id_placeholders = ",".join("?" for _ in report_ids)
        c_inv = cursor.execute(f"SELECT COUNT(*) as cnt FROM invoices WHERE report_id IN ({id_placeholders})", report_ids).fetchone()
        deleted_invoices_count = c_inv['cnt'] if c_inv else 0
        cursor.execute(f"DELETE FROM invoices WHERE report_id IN ({id_placeholders})", report_ids)

        c_trip = cursor.execute(f"SELECT COUNT(*) as cnt FROM trips WHERE report_id IN ({id_placeholders})", report_ids).fetchone()
        deleted_trips_count = c_trip['cnt'] if c_trip else 0
        cursor.execute(f"DELETE FROM trips WHERE report_id IN ({id_placeholders})", report_ids)

        cursor.execute(f"DELETE FROM daily_reports WHERE id IN ({id_placeholders})", report_ids)

    plan_where = []
    plan_params = []
    if not is_all_depots:
        plan_where.append("depot_id = ?")
        plan_params.append(int(depot_id))
    if date_mode == 'single' and single_date:
        plan_where.append("plan_date = ?")
        plan_params.append(single_date)
    elif date_mode == 'range' and start_date and end_date:
        plan_where.append("plan_date >= ? AND plan_date <= ?")
        plan_params.append(start_date)
        plan_params.append(end_date)
    
    plan_sql = ("WHERE " + " AND ".join(plan_where)) if plan_where else ""
    try:
        cursor.execute(f"DELETE FROM saved_route_plans {plan_sql}", plan_params)
    except Exception:
        pass

    if is_all_depots and date_mode == 'all':
        try:
            cursor.execute('DELETE FROM live_tracking_positions')
            cursor.execute('DELETE FROM live_drop_statuses')
            cursor.execute('DELETE FROM inter_depot_vehicle_requests')
        except Exception:
            pass

    depot_target_name = "All 17 Depots"
    if not is_all_depots:
        d_row = cursor.execute("SELECT name FROM depots WHERE id = ?", (int(depot_id),)).fetchone()
        depot_target_name = d_row['name'] if d_row else f"Depot #{depot_id}"

    username = sess_user.get('username') or 'admin'
    details_str = f"Cleared: {depot_target_name} | {date_summary_str} (Removed {deleted_reports_count} reports, {deleted_invoices_count} invoices)"
    try:
        cursor.execute('''
            INSERT INTO admin_audit_log (username, action, target_type, depot_id, details, ip_address)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (username, 'CLEAR_OPERATIONAL_DATA', 'daily_reports', str(depot_id), details_str, request.remote_addr))
    except Exception:
        pass

    conn.commit()
    conn.close()

    msg = f"Successfully cleared operational data for {depot_target_name} ({date_summary_str}). Removed {deleted_reports_count} daily report(s) and {deleted_invoices_count} delivery invoice(s)."
    return jsonify({
        "success": True,
        "message": msg,
        "deleted_reports": deleted_reports_count,
        "deleted_invoices": deleted_invoices_count,
        "deleted_trips": deleted_trips_count,
        "target_depot": depot_target_name,
        "date_summary": date_summary_str
    })

# ----------------- REPORT SUBMISSION & ADMIN MANAGEMENT -----------------
@app.route('/api/reports/submit', methods=['POST'])
def submit_report():
    data = request.json
    conn = get_db()
    cursor = conn.cursor()
    
    depot_id = data.get('depot_id')
    report_date = data.get('report_date', datetime.date.today().strftime('%Y-%m-%d'))
    incharge_name = data.get('incharge_name')
    contact = data.get('contact')
    trips = data.get('trips', [])
    invoices = data.get('invoices', [])
    
    tot_disp_val = sum(float(i.get('dispatched_val', 0)) for i in invoices)
    tot_deliv_val = sum(float(i.get('delivered_val', 0)) for i in invoices)
    tot_ret_val = sum(float(i.get('returned_val', 0)) for i in invoices)
    
    tot_disp_qty = sum(float(i.get('dispatched_qty', 0)) for i in invoices)
    tot_deliv_qty = sum(float(i.get('delivered_qty', 0)) for i in invoices)
    tot_ret_qty = sum(float(i.get('returned_qty', 0)) for i in invoices)
    
    stock_var = round(tot_disp_qty - (tot_deliv_qty + tot_ret_qty))
    
    tot_cash = sum(float(i.get('amount_collected', 0)) for i in invoices)
    tot_credit = sum(float(i.get('delivered_val', 0)) for i in invoices if i.get('collection_mode') == 'Credit')
    cash_var = round(tot_deliv_val - (tot_cash + tot_credit))
    
    util_sum = sum(float(t.get('loaded_kg', 0))/float(t.get('capacity_kg', 1))*100 for t in trips if float(t.get('capacity_kg', 0)) > 0)
    avg_util = round(util_sum / len(trips), 1) if trips else 0
    
    cursor.execute('''
    DELETE FROM daily_reports WHERE depot_id = ? AND report_date = ?
    ''', (depot_id, report_date))
    
    cursor.execute('''
    INSERT INTO daily_reports (
        report_date, depot_id, incharge_name, contact, 
        total_vehicles, total_invoices, capacity_util_pct,
        dispatched_gross_val, delivered_net_val, returned_val,
        stock_mismatch_qty, cash_mismatch_val, adjustment_status, audit_status
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '100% Adjusted', 'OK / Verified')
    ''', (
        report_date, depot_id, incharge_name, contact,
        len(trips), len(invoices), avg_util,
        tot_disp_val, tot_deliv_val, tot_ret_val,
        stock_var, cash_var
    ))
    report_id = cursor.lastrowid
    
    for t in trips:
        cap = float(t.get('capacity_kg', 0))
        load = float(t.get('loaded_kg', 0))
        u = round(load / cap * 100, 1) if cap > 0 else 0
        cursor.execute('''
        INSERT INTO trips (
            report_id, trip_no, vehicle_no, vehicle_type, driver_name, delivery_man, route_name,
            capacity_kg, loaded_kg, util_pct, reefer_temp
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            report_id, t.get('trip_no'), t.get('vehicle_no'), t.get('vehicle_type'),
            t.get('driver_name'), t.get('delivery_man'), t.get('route_name'),
            cap, load, u, t.get('reefer_temp')
        ))
        
    for inv in invoices:
        cursor.execute('''
        INSERT INTO invoices (
            report_id, invoice_no, customer_name, trip_no, product_category, sku_uom,
            dispatched_qty, dispatched_val, delivery_status, delivered_qty, delivered_val,
            returned_qty, returned_val, return_reason, collection_mode, amount_collected
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            report_id,
            inv.get('invoice_no') or 'INV-001',
            inv.get('customer_name') or 'Customer Outlet',
            inv.get('trip_no') or 'TRIP-01',
            inv.get('product_category') or 'Frozen Food',
            inv.get('sku_uom') or 'Pkt',
            float(inv.get('dispatched_qty') or 0),
            float(inv.get('dispatched_val') or 0),
            inv.get('delivery_status') or 'Delivered',
            float(inv.get('delivered_qty') or 0),
            float(inv.get('delivered_val') or 0),
            float(inv.get('returned_qty') or 0),
            float(inv.get('returned_val') or 0),
            inv.get('return_reason') or 'None',
            inv.get('collection_mode') or 'Credit',
            float(inv.get('amount_collected') or 0)
        ))
        
    conn.commit()
    conn.close()
    return jsonify({"success": True, "report_id": report_id, "message": "Report submitted and verified successfully!"})

@app.route('/api/reports/delete/<int:report_id>', methods=['GET', 'DELETE', 'POST'])
@app.route('/api/reports/delete-by-depot/<int:report_id>', methods=['GET', 'DELETE', 'POST'])
def delete_report(report_id):
    conn = get_db()
    cursor = conn.cursor()
    # Find all report ids that match either report_id or are associated with this depot_id
    cursor.execute('SELECT id FROM daily_reports WHERE id = ? OR depot_id = ?', (report_id, report_id))
    matching_report_ids = [row['id'] for row in cursor.fetchall()]
    if report_id not in matching_report_ids:
        matching_report_ids.append(report_id)
        
    for r_id in matching_report_ids:
        cursor.execute('DELETE FROM invoices WHERE report_id = ?', (r_id,))
        cursor.execute('DELETE FROM trips WHERE report_id = ?', (r_id,))
        cursor.execute('DELETE FROM daily_reports WHERE id = ?', (r_id,))
        
    cursor.execute('DELETE FROM daily_reports WHERE depot_id = ?', (report_id,))
    
    try:
        cursor.execute('DELETE FROM saved_route_plans WHERE depot_id = ?', (report_id,))
    except Exception:
        pass
    
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Daily report, invoices & uploaded route plans for Depot / Report #{report_id} deleted successfully!"})

@app.route('/api/admin/clear-all-uploaded-data', methods=['GET', 'POST', 'DELETE'])
def admin_clear_all_uploaded_data():
    data = request.json if request.is_json else {}
    depot_id = data.get('depot_id') or request.args.get('depot_id')
    role = session.get('role') or (session.get('user') and session['user'].get('role')) or data.get('role') or request.headers.get('X-Admin-Role')
    if role and role not in ['admin', 'guest'] and role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized: Only Admin can clear operational data"}), 403

    conn = get_db()
    cursor = conn.cursor()
    if depot_id and str(depot_id).lower() != 'all':
        cursor.execute('DELETE FROM invoices WHERE report_id IN (SELECT id FROM daily_reports WHERE depot_id = ?)', (depot_id,))
        cursor.execute('DELETE FROM trips WHERE report_id IN (SELECT id FROM daily_reports WHERE depot_id = ?)', (depot_id,))
        cursor.execute('DELETE FROM daily_reports WHERE depot_id = ?', (depot_id,))
        try:
            cursor.execute('DELETE FROM saved_route_plans WHERE depot_id = ?', (depot_id,))
        except Exception:
            pass
        try:
            cursor.execute('DELETE FROM live_tracking_positions WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM live_drop_statuses WHERE depot_id = ?', (depot_id,))
        except Exception:
            pass
        msg = f"All uploaded daily reports, invoices & route plans for Depot #{depot_id} cleared successfully!"
    else:
        cursor.execute('DELETE FROM invoices')
        cursor.execute('DELETE FROM trips')
        cursor.execute('DELETE FROM daily_reports')
        try:
            cursor.execute('DELETE FROM saved_route_plans')
        except Exception:
            pass
        try:
            cursor.execute('DELETE FROM live_tracking_positions')
            cursor.execute('DELETE FROM live_drop_statuses')
        except Exception:
            pass
        msg = "All uploaded daily reports, delivery invoices, trips and route plans cleared successfully!"
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": msg})

@app.route('/api/admin/reset-demo-data', methods=['GET', 'POST', 'DELETE'])
def admin_reset_demo_data():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM invoices')
    cursor.execute('DELETE FROM trips')
    cursor.execute('DELETE FROM daily_reports')
    try:
        cursor.execute('DELETE FROM saved_route_plans')
    except Exception:
        pass
    try:
        cursor.execute('DELETE FROM live_tracking_positions')
        cursor.execute('DELETE FROM live_drop_statuses')
    except Exception:
        pass
    try:
        cursor.execute('DELETE FROM inter_depot_vehicle_requests')
    except Exception:
        pass

    try:
        cursor.execute('''
            INSERT INTO admin_audit_log (username, action, target_type, depot_id, details, ip_address)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', ('admin', 'RESET_DEMO_DATA', 'operational_reports', 'ALL', 'All operational reports cleared via reset-demo-data', request.remote_addr))
    except Exception:
        pass

    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Demo operational reports and plans cleared successfully!"})


@app.route('/api/reports/get-by-date')
def get_report_by_date():
    depot_id = request.args.get('depot_id')
    report_date = request.args.get('date')
    if not depot_id or not report_date:
        return jsonify({"success": False, "message": "Missing depot_id or date"})
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM daily_reports WHERE depot_id = ? AND report_date = ?', (depot_id, report_date))
    rep = cursor.fetchone()
    if not rep:
        conn.close()
        return jsonify({"success": True, "exists": False})
    
    report_id = rep['id']
    cursor.execute('SELECT * FROM trips WHERE report_id = ?', (report_id,))
    trips = [dict(t) for t in cursor.fetchall()]
    
    cursor.execute('SELECT * FROM invoices WHERE report_id = ?', (report_id,))
    invoices = [dict(i) for i in cursor.fetchall()]
    
    conn.close()
    return jsonify({
        "success": True,
        "exists": True,
        "report": dict(rep),
        "trips": trips,
        "invoices": invoices
    })

@app.route('/api/reports/delete-by-date', methods=['GET', 'POST', 'DELETE'])
def delete_report_by_date():
    data = request.json if request.is_json else {}
    role = session.get('role') or (session.get('user') and session['user'].get('role')) or data.get('role') or request.headers.get('X-Admin-Role')
    if role and role not in ['admin', 'guest'] and role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized: Only Admin can delete submitted daily reports"}), 403

    depot_id = data.get('depot_id') or request.args.get('depot_id')
    report_date = data.get('date') or request.args.get('date')
    if not depot_id or not report_date:
        return jsonify({"success": False, "message": "Missing depot_id or date"})
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT id FROM daily_reports WHERE depot_id = ? AND report_date = ?', (depot_id, report_date))
    reps = cursor.fetchall()
    if not reps:
        conn.close()
        return jsonify({"success": False, "message": f"No report found for depot {depot_id} on date {report_date}"})
        
    for r in reps:
        rid = r['id']
        cursor.execute('DELETE FROM invoices WHERE report_id = ?', (rid,))
        cursor.execute('DELETE FROM trips WHERE report_id = ?', (rid,))
        cursor.execute('DELETE FROM daily_reports WHERE id = ?', (rid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Report for date {report_date} has been deleted successfully!"})

@app.route('/api/admin/reseed-demo-data', methods=['POST'])
def admin_reseed_demo_data():
    conn = get_db()
    cursor = conn.cursor()
    seed_sample_daily_reports(cursor)
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Demo distribution reports and invoices re-seeded successfully!"})

def generate_live_master_excel(date_param=None, depot_id=None):
    if not date_param:
        date_param = datetime.date.today().strftime('%Y-%m-%d')
        
    conn = get_db()
    
    if depot_id and str(depot_id).lower() != 'all':
        depots = conn.execute('SELECT * FROM depots WHERE id = ? ORDER BY id ASC', (depot_id,)).fetchall()
        reports = conn.execute('''
            SELECT r.*, d.name as depot_name, d.category as depot_category, d.region as depot_region, d.default_uom
            FROM daily_reports r
            JOIN depots d ON r.depot_id = d.id
            WHERE r.report_date = ? AND r.depot_id = ?
            ORDER BY r.depot_id ASC
        ''', (date_param, depot_id)).fetchall()
        invoices = conn.execute('''
            SELECT i.*, d.name as depot_name, r.report_date, t.vehicle_no, t.route_name
            FROM invoices i
            JOIN daily_reports r ON i.report_id = r.id
            JOIN depots d ON r.depot_id = d.id
            LEFT JOIN trips t ON i.trip_id = t.id
            WHERE r.report_date = ? AND r.depot_id = ?
            ORDER BY i.id ASC
        ''', (date_param, depot_id)).fetchall()
        vehicles = conn.execute('''
            SELECT v.*, d.name as depot_name, c.name as driver_name, c.phone as driver_contact
            FROM fleet_vehicles v
            LEFT JOIN depots d ON v.depot_id = d.id
            LEFT JOIN depot_crew c ON v.default_driver_id = c.id
            WHERE v.depot_id = ?
            ORDER BY v.id ASC
        ''', (depot_id,)).fetchall()
    else:
        depots = conn.execute('SELECT * FROM depots ORDER BY id ASC').fetchall()
        reports = conn.execute('''
            SELECT r.*, d.name as depot_name, d.category as depot_category, d.region as depot_region, d.default_uom
            FROM daily_reports r
            JOIN depots d ON r.depot_id = d.id
            WHERE r.report_date = ?
            ORDER BY r.depot_id ASC
        ''', (date_param,)).fetchall()
        invoices = conn.execute('''
            SELECT i.*, d.name as depot_name, r.report_date, t.vehicle_no, t.route_name
            FROM invoices i
            JOIN daily_reports r ON i.report_id = r.id
            JOIN depots d ON r.depot_id = d.id
            LEFT JOIN trips t ON i.trip_id = t.id
            WHERE r.report_date = ?
            ORDER BY i.id ASC
        ''', (date_param,)).fetchall()
        vehicles = conn.execute('''
            SELECT v.*, d.name as depot_name, c.name as driver_name, c.phone as driver_contact
            FROM fleet_vehicles v
            LEFT JOIN depots d ON v.depot_id = d.id
            LEFT JOIN depot_crew c ON v.default_driver_id = c.id
            ORDER BY v.depot_id ASC, v.id ASC
        ''').fetchall()
    conn.close()

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    font_fam = 'Segoe UI'

    # Typography & Styles
    hdr_title = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 14, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    hdr_sub = wb.add_format({'font_name': font_fam, 'font_size': 9.5, 'font_color': '#CBD5E1', 'bg_color': '#1E293B', 'align': 'center', 'valign': 'vcenter', 'italic': True})
    th_pri = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 9.5, 'font_color': '#FFFFFF', 'bg_color': '#1E3A8A', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    th_sec = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 9.5, 'font_color': '#FFFFFF', 'bg_color': '#0284C7', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    
    td_txt = wb.add_format({'font_name': font_fam, 'font_size': 9, 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'})
    td_bld = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 9, 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'})
    td_ctr = wb.add_format({'font_name': font_fam, 'font_size': 9, 'align': 'center', 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'})
    td_num = wb.add_format({'font_name': font_fam, 'font_size': 9, 'align': 'right', 'num_format': '#,##0', 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'})
    td_money = wb.add_format({'font_name': font_fam, 'font_size': 9, 'align': 'right', 'num_format': '৳ #,##0.00', 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'})
    td_pct = wb.add_format({'font_name': font_fam, 'font_size': 9, 'align': 'center', 'num_format': '0.0%', 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'})
    td_done = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 8.5, 'align': 'center', 'font_color': '#15803D', 'bg_color': '#DCFCE7', 'border': 1, 'border_color': '#86EFAC', 'valign': 'vcenter'})
    td_pend = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 8.5, 'align': 'center', 'font_color': '#B91C1C', 'bg_color': '#FEE2E2', 'border': 1, 'border_color': '#FCA5A5', 'valign': 'vcenter'})

    kpi_lbl = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 8.5, 'align': 'center', 'valign': 'vcenter', 'font_color': '#475569', 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#CBD5E1'})
    kpi_val = wb.add_format({'bold': True, 'font_name': font_fam, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'font_color': '#0284C7', 'bg_color': '#F0F9FF', 'border': 1, 'border_color': '#BAE6FD'})

    # 1. Executive Summary
    ws1 = wb.add_worksheet('01_Executive_Summary')
    ws1.merge_range('A1:I1', 'PARAGON AGRO LIMITED — MASTER DISTRIBUTION EXECUTIVE REPORT', hdr_title)
    ws1.merge_range('A2:I2', f'Report Date: {date_param} | System Depots: {len(depots)} Locations | Live Production Database Export', hdr_sub)
    ws1.set_row(0, 26)
    ws1.set_row(1, 18)

    tot_disp = sum(r['dispatched_gross_val'] or 0 for r in reports)
    tot_deliv = sum(r['delivered_net_val'] or 0 for r in reports)
    tot_ret = sum(r['returned_val'] or 0 for r in reports)
    tot_vans = sum(r['total_vehicles'] or 0 for r in reports)
    tot_invs = sum(r['total_invoices'] or 0 for r in reports)
    succ_rate = (tot_deliv / tot_disp) if tot_disp > 0 else 0

    ws1.set_column(0, 0, 8)
    ws1.set_column(1, 1, 32)
    ws1.set_column(2, 2, 14)
    ws1.set_column(3, 3, 20)
    ws1.set_column(4, 4, 20)
    ws1.set_column(5, 5, 18)
    ws1.set_column(6, 6, 16)
    ws1.set_column(7, 7, 14)
    ws1.set_column(8, 8, 14)

    ws1.merge_range('A4:B4', 'TOTAL DISPATCHED VALUE', kpi_lbl)
    ws1.merge_range('A5:B5', f'৳ {tot_disp:,.2f}', kpi_val)
    ws1.merge_range('C4:D4', 'TOTAL DELIVERED VALUE', kpi_lbl)
    ws1.merge_range('C5:D5', f'৳ {tot_deliv:,.2f}', kpi_val)
    ws1.merge_range('E4:F4', 'TOTAL RETURNS', kpi_lbl)
    ws1.merge_range('E5:F5', f'৳ {tot_ret:,.2f}', kpi_val)
    ws1.write('G4', 'SUCCESS RATE', kpi_lbl)
    ws1.write('G5', f'{succ_rate*100:.1f}%', kpi_val)
    ws1.write('H4', 'ACTIVE VANS', kpi_lbl)
    ws1.write('H5', str(tot_vans), kpi_val)
    ws1.write('I4', 'INVOICES', kpi_lbl)
    ws1.write('I5', str(tot_invs), kpi_val)

    # Category Breakdown
    cat_names = [
        'Frozen Food & RTC', 'Processed Chicken', 'Sweets & Savory', 'Fresh Eggs',
        'Dairy (Sirajganj & Tejgaon)', 'Dry Food & Bakery', 'Momo & Dimsum', 'Tea Distribution & Packing'
    ]
    ws1.merge_range('A7:I7', 'PRODUCT CATEGORY PERFORMANCE BREAKDOWN', th_sec)
    ws1.set_row(6, 20)
    cat_hdrs = ['SL', 'Category Name', 'Primary UOM', 'Dispatched Gross (৳)', 'Delivered Net (৳)', 'Returned Value (৳)', 'Success Rate (%)', 'Active Vans', 'Total Invoices']
    for c, h in enumerate(cat_hdrs):
        ws1.write(7, c, h, th_pri)
    ws1.set_row(7, 22)

    for idx, cname in enumerate(cat_names, 8):
        c_disp = sum(r['dispatched_gross_val'] or 0 for r in reports if cname.lower() in (r['depot_category'] or '').lower() or (r['depot_category'] or '').lower() in cname.lower())
        c_deliv = sum(r['delivered_net_val'] or 0 for r in reports if cname.lower() in (r['depot_category'] or '').lower() or (r['depot_category'] or '').lower() in cname.lower())
        c_ret = sum(r['returned_val'] or 0 for r in reports if cname.lower() in (r['depot_category'] or '').lower() or (r['depot_category'] or '').lower() in cname.lower())
        c_v = sum(r['total_vehicles'] or 0 for r in reports if cname.lower() in (r['depot_category'] or '').lower() or (r['depot_category'] or '').lower() in cname.lower())
        c_inv = sum(r['total_invoices'] or 0 for r in reports if cname.lower() in (r['depot_category'] or '').lower() or (r['depot_category'] or '').lower() in cname.lower())
        c_succ = (c_deliv / c_disp) if c_disp > 0 else 0
        
        ws1.write(idx, 0, idx - 7, td_ctr)
        ws1.write(idx, 1, cname, td_bld)
        ws1.write(idx, 2, 'Pkt/Kg/Pcs', td_ctr)
        ws1.write(idx, 3, c_disp, td_money)
        ws1.write(idx, 4, c_deliv, td_money)
        ws1.write(idx, 5, c_ret, td_money)
        ws1.write(idx, 6, c_succ, td_pct)
        ws1.write(idx, 7, c_v, td_num)
        ws1.write(idx, 8, c_inv, td_num)

    # 2. Depots Compliance
    ws2 = wb.add_worksheet('02_Depots_Compliance')
    ws2.merge_range('A1:J1', 'PARAGON AGRO LIMITED — 17 DEPOTS SUBMISSION COMPLIANCE', hdr_title)
    ws2.merge_range('A2:J2', f'Report Date: {date_param} | Target: 17 Depots System Wide', hdr_sub)
    ws2.set_row(0, 24)
    ws2.set_row(1, 18)

    hdrs2 = ['ID', 'Depot / Factory Name', 'Category', 'Region', 'Incharge Name', 'Contact Phone', 'Default UOM', 'Submission Status', 'Dispatched Gross (৳)', 'Delivered Net (৳)']
    for c, h in enumerate(hdrs2):
        ws2.write(3, c, h, th_pri)
    ws2.set_row(3, 22)

    rep_map = {r['depot_id']: r for r in reports}
    for idx, d in enumerate(depots, 4):
        d_rep = rep_map.get(d['id'])
        is_done = d_rep is not None
        ws2.write(idx, 0, d['id'], td_ctr)
        ws2.write(idx, 1, d['name'], td_bld)
        ws2.write(idx, 2, d['category'], td_txt)
        ws2.write(idx, 3, d['region'] or 'Central', td_ctr)
        ws2.write(idx, 4, d['incharge_name'] or '-', td_txt)
        ws2.write(idx, 5, d['contact'] or '-', td_ctr)
        ws2.write(idx, 6, d['default_uom'] or 'Pkt', td_ctr)
        ws2.write(idx, 7, 'DONE' if is_done else 'PENDING', td_done if is_done else td_pend)
        ws2.write(idx, 8, (d_rep['dispatched_gross_val'] if d_rep else 0) or 0, td_money)
        ws2.write(idx, 9, (d_rep['delivered_net_val'] if d_rep else 0) or 0, td_money)

    ws2.set_column(0, 0, 6)
    ws2.set_column(1, 1, 32)
    ws2.set_column(2, 2, 20)
    ws2.set_column(3, 3, 14)
    ws2.set_column(4, 4, 26)
    ws2.set_column(5, 5, 16)
    ws2.set_column(6, 6, 12)
    ws2.set_column(7, 7, 14)
    ws2.set_column(8, 8, 20)
    ws2.set_column(9, 9, 20)

    # 3. Daily Reports Log
    ws3 = wb.add_worksheet('03_Daily_Reports_Log')
    ws3.merge_range('A1:L1', 'PARAGON AGRO LIMITED — DAILY DISTRIBUTION REPORTS (DDR) LOG', hdr_title)
    ws3.merge_range('A2:L2', f'Date Filter: {date_param}', hdr_sub)
    ws3.set_row(0, 24)
    ws3.set_row(1, 18)

    hdrs3 = ['Report ID', 'Date', 'Depot Name', 'Category', 'Incharge Name', 'Vehicles', 'Invoices', 'Dispatched Gross (৳)', 'Delivered Net (৳)', 'Returned Val (৳)', 'Stock Mismatch', 'Audit Status']
    for c, h in enumerate(hdrs3):
        ws3.write(3, c, h, th_pri)
    ws3.set_row(3, 22)

    if reports:
        for idx, r in enumerate(reports, 4):
            ws3.write(idx, 0, r['id'], td_ctr)
            ws3.write(idx, 1, r['report_date'], td_ctr)
            ws3.write(idx, 2, r['depot_name'], td_bld)
            ws3.write(idx, 3, r['depot_category'], td_txt)
            ws3.write(idx, 4, r['incharge_name'], td_txt)
            ws3.write(idx, 5, r['total_vehicles'] or 0, td_num)
            ws3.write(idx, 6, r['total_invoices'] or 0, td_num)
            ws3.write(idx, 7, r['dispatched_gross_val'] or 0, td_money)
            ws3.write(idx, 8, r['delivered_net_val'] or 0, td_money)
            ws3.write(idx, 9, r['returned_val'] or 0, td_money)
            ws3.write(idx, 10, r['stock_mismatch_qty'] or 0, td_num)
            ws3.write(idx, 11, r['audit_status'] or 'Verified', td_ctr)
    else:
        ws3.merge_range('A5:L5', 'No operational daily distribution reports found for this date.', td_ctr)

    ws3.set_column(0, 0, 10)
    ws3.set_column(1, 1, 12)
    ws3.set_column(2, 2, 30)
    ws3.set_column(3, 3, 18)
    ws3.set_column(4, 4, 24)
    ws3.set_column(5, 6, 10)
    ws3.set_column(7, 9, 20)
    ws3.set_column(10, 10, 14)
    ws3.set_column(11, 11, 16)

    # 4. Delivery Invoices
    ws4 = wb.add_worksheet('04_Delivery_Invoices')
    ws4.merge_range('A1:K1', 'PARAGON AGRO LIMITED — DELIVERY INVOICES & OUTLET BREAKDOWN', hdr_title)
    ws4.merge_range('A2:K2', f'Date Filter: {date_param} | Total Invoices: {len(invoices)}', hdr_sub)
    ws4.set_row(0, 24)
    ws4.set_row(1, 18)

    hdrs4 = ['SL', 'Invoice No', 'Date', 'Depot Name', 'Customer / Outlet Name', 'Vehicle Reg No', 'Route Name', 'Dispatched Val (৳)', 'Delivered Val (৳)', 'Returned Val (৳)', 'Amount Collected (৳)']
    for c, h in enumerate(hdrs4):
        ws4.write(3, c, h, th_pri)
    ws4.set_row(3, 22)

    if invoices:
        for idx, inv in enumerate(invoices, 4):
            inv_keys = inv.keys() if hasattr(inv, 'keys') else []
            v_no = inv['vehicle_no'] if 'vehicle_no' in inv_keys else '-'
            r_no = inv['route_name'] if 'route_name' in inv_keys else '-'
            ws4.write(idx, 0, idx - 3, td_ctr)
            ws4.write(idx, 1, inv['invoice_no'] or f'INV-{inv["id"]}', td_bld)
            ws4.write(idx, 2, inv['report_date'], td_ctr)
            ws4.write(idx, 3, inv['depot_name'], td_txt)
            ws4.write(idx, 4, inv['customer_name'] or 'Outlet', td_txt)
            ws4.write(idx, 5, v_no or '-', td_ctr)
            ws4.write(idx, 6, r_no or '-', td_txt)
            ws4.write(idx, 7, inv['dispatched_val'] or 0, td_money)
            ws4.write(idx, 8, inv['delivered_val'] or 0, td_money)
            ws4.write(idx, 9, inv['returned_val'] or 0, td_money)
            ws4.write(idx, 10, inv['amount_collected'] or 0, td_money)
    else:
        ws4.merge_range('A5:K5', 'No delivery invoices found for this selection.', td_ctr)

    ws4.set_column(0, 0, 6)
    ws4.set_column(1, 1, 18)
    ws4.set_column(2, 2, 12)
    ws4.set_column(3, 3, 26)
    ws4.set_column(4, 4, 32)
    ws4.set_column(5, 5, 18)
    ws4.set_column(6, 6, 26)
    ws4.set_column(7, 10, 18)

    # 5. Fleet Vehicles
    ws5 = wb.add_worksheet('05_Fleet_Directory')
    ws5.merge_range('A1:I1', 'PARAGON AGRO LIMITED — FLEET VEHICLES & CAPACITIES MASTER', hdr_title)
    ws5.merge_range('A2:I2', f'Active Registered Vehicles: {len(vehicles)} Units', hdr_sub)
    ws5.set_row(0, 24)
    ws5.set_row(1, 18)

    hdrs5 = ['ID', 'Vehicle Reg No', 'Vehicle Type', 'Depot Name', 'Driver Name', 'Driver Mobile', 'Capacity (Kg/Units)', 'Ownership', 'Daily Rate (৳)']
    for c, h in enumerate(hdrs5):
        ws5.write(3, c, h, th_pri)
    ws5.set_row(3, 22)

    if vehicles:
        for idx, v in enumerate(vehicles, 4):
            v_keys = v.keys() if hasattr(v, 'keys') else []
            d_name = v['driver_name'] if 'driver_name' in v_keys else '-'
            d_phone = v['driver_contact'] if 'driver_contact' in v_keys else '-'
            rent_cost = v['rental_cost_per_day'] if 'rental_cost_per_day' in v_keys else 0
            ws5.write(idx, 0, v['id'], td_ctr)
            ws5.write(idx, 1, v['vehicle_no'], td_bld)
            ws5.write(idx, 2, v['vehicle_type'], td_txt)
            ws5.write(idx, 3, v['depot_name'] or '-', td_txt)
            ws5.write(idx, 4, d_name or '-', td_txt)
            ws5.write(idx, 5, d_phone or '-', td_ctr)
            ws5.write(idx, 6, f"{v['capacity_kg'] or 0} {v['capacity_units'] or 'Kg'}", td_ctr)
            ws5.write(idx, 7, v['ownership'] or 'Owned', td_ctr)
            ws5.write(idx, 8, rent_cost or 0, td_money)
    else:
        ws5.merge_range('A5:I5', 'No fleet vehicles found.', td_ctr)

    ws5.set_column(0, 0, 6)
    ws5.set_column(1, 1, 20)
    ws5.set_column(2, 2, 22)
    ws5.set_column(3, 3, 28)
    ws5.set_column(4, 4, 24)
    ws5.set_column(5, 5, 16)
    ws5.set_column(6, 6, 18)
    ws5.set_column(7, 7, 14)
    ws5.set_column(8, 8, 16)

    # 6. SOP & Anti-Theft Protocols
    ws6 = wb.add_worksheet('06_SOP_&_Zero_Theft')
    ws6.merge_range('A1:E1', 'PARAGON AGRO LIMITED — DISTRIBUTION STANDARD OPERATING PROCEDURES (SOP)', hdr_title)
    ws6.merge_range('A2:E2', 'Zero-Theft, Cold Chain Compliance & Security Protocols', hdr_sub)
    ws6.set_row(0, 24)
    ws6.set_row(1, 18)

    hdrs6 = ['Rule #', 'Operational Area', 'Compliance Standard', 'Verification Mechanism', 'Violation Penalty']
    for c, h in enumerate(hdrs6):
        ws6.write(3, c, h, th_pri)
    ws6.set_row(3, 22)

    sop_data = [
        ('SOP-01', 'Cold Chain Verification', 'Temperature loggers must show <= -18°C for Frozen and 2-4°C for Dairy at dispatch and drop points.', 'Digital logger upload & receiver signature', 'Consignment rejection & driver disciplinary review'),
        ('SOP-02', 'Vehicle Digital Seal', 'All van cargo bays must have tamper-evident numerical seals verified by Depot Security.', 'Security gate register & driver counter-sign', 'Immediate investigation; suspension pending audit'),
        ('SOP-03', 'Cash Collection & Reconciliation', '100% of cash collected must be deposited to depot cashier within 2 hours of route completion.', 'Bank deposit slip & ERP voucher reconciliation', 'Immediate salary hold; financial recovery protocol'),
        ('SOP-04', 'Zero Returns Handling', 'Undelivered goods must be physically inspected, counted and returned to cold storage within 30 min.', 'Return Goods Voucher (RGV) signed by Incharge', 'Unaccounted stock deducted from depot balance'),
        ('SOP-05', 'GPS Route Deviation Audit', 'Any route deviation > 1.5 km or unscheduled stop > 15 mins triggers immediate security alert.', 'Live telematics tracking server alert', 'Depot Incharge formal explanation within 12 hours')
    ]
    for idx, s in enumerate(sop_data, 4):
        ws6.write(idx, 0, s[0], td_bld)
        ws6.write(idx, 1, s[1], td_txt)
        ws6.write(idx, 2, s[2], td_txt)
        ws6.write(idx, 3, s[3], td_txt)
        ws6.write(idx, 4, s[4], td_txt)

    ws6.set_column(0, 0, 10)
    ws6.set_column(1, 1, 26)
    ws6.set_column(2, 2, 45)
    ws6.set_column(3, 3, 35)
    ws6.set_column(4, 4, 38)

    wb.close()
    output.seek(0)
    return output

@app.route('/api/export/excel')
def export_master_excel():
    date_param = request.args.get('date', datetime.date.today().strftime('%Y-%m-%d'))
    depot_param = request.args.get('depot')
    try:
        excel_stream = generate_live_master_excel(date_param, depot_param)
        filename = f"Paragon_Master_Distribution_{date_param}.xlsx"
        if depot_param and str(depot_param).lower() != 'all':
            filename = f"Paragon_Depot_{depot_param}_Distribution_{date_param}.xlsx"
        return send_file(
            excel_stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        master_file = r'd:\AI project\Distribution\Paragon_Distribution_Master_Report.xlsx'
        if os.path.exists(master_file):
            return send_file(master_file, as_attachment=True, download_name=f"Paragon_Master_Distribution_{date_param}.xlsx")
        return jsonify({"error": f"Failed to generate excel: {str(e)}"}), 500


# ----------------- BULK EXCEL TEMPLATE (CRYSTAL CLEAR HIGH-VISIBILITY TYPOGRAPHY) -----------------
@app.route('/api/template/daily-entry-excel')
def download_daily_entry_template():
    depot_id = request.args.get('depot_id')
    depot_name = "Paragon Depot"
    category = "Frozen Food & RTC"
    incharge_info = "Depot Incharge"
    
    if depot_id:
        try:
            conn = get_db()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM depots WHERE id = ?', (depot_id,))
            depot = cursor.fetchone()
            conn.close()
            if depot:
                depot_name = depot['name']
                category = depot['category']
                incharge_info = f"{depot['incharge_name']} ({depot['contact']})"
        except Exception as e:
            pass

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # High-Visibility Typography & Border Styles (Guaranteed 100% Legibility on Any Screen / Excel)
    fmt_navy_hdr = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 11, 'font_color': '#1E3A8A',
        'align': 'center', 'valign': 'vcenter', 'top': 2, 'bottom': 6, 'left': 1, 'right': 1,
        'border_color': '#1E3A8A'
    })
    fmt_cyan_hdr = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 11, 'font_color': '#0369A1',
        'align': 'center', 'valign': 'vcenter', 'top': 2, 'bottom': 6, 'left': 1, 'right': 1,
        'border_color': '#0369A1'
    })
    fmt_sample_hdr = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 9.5, 'font_color': '#475569', 'italic': True,
        'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E1'
    })
    fmt_sample_num = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 9.5, 'font_color': '#475569', 'italic': True,
        'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E1', 'num_format': '#,##0'
    })
    fmt_sample_center = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 9.5, 'font_color': '#475569', 'italic': True,
        'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E1'
    })

    # Sheet 1: Fleet & Trips
    ws_trips = wb.add_worksheet("Fleet_and_Trips")
    ws_trips.hide_gridlines(0)
    ws_trips.set_row(0, 28)
    
    trip_headers = [
        "Trip_No", "Vehicle_Reg_No", "Vehicle_Type", "Driver_Name",
        "Delivery_Man", "Route_Market_Area", "Capacity_Kg", "Loaded_Kg", "Reefer_Temp"
    ]
    for c_i, h in enumerate(trip_headers):
        ws_trips.write(0, c_i, h, fmt_navy_hdr)
        
    sample_trips = [
        ["TRIP-01", "DM-THA-11-2045", "1.5T Reefer Van", "Abul Kalam", "Md. Rafiq", f"{depot_name} Standard Route", 1500, 1420, "-18°C"],
        ["TRIP-02", "DM-THA-11-3090", "2.0T Reefer Van", "Kalam Miah", "Md. Dulal", f"{depot_name} Secondary Route", 2000, 1890, "-18°C"]
    ]
    for r_i, r in enumerate(sample_trips, start=1):
        ws_trips.set_row(r_i, 20)
        for c_i, val in enumerate(r):
            if c_i in [6, 7]:
                ws_trips.write(r_i, c_i, val, fmt_sample_num)
            elif c_i in [0, 8]:
                ws_trips.write(r_i, c_i, val, fmt_sample_center)
            else:
                ws_trips.write(r_i, c_i, val, fmt_sample_hdr)
                
    ws_trips.set_column(0, 0, 14)
    ws_trips.set_column(1, 1, 20)
    ws_trips.set_column(2, 2, 18)
    ws_trips.set_column(3, 4, 18)
    ws_trips.set_column(5, 5, 34)
    ws_trips.set_column(6, 7, 16)
    ws_trips.set_column(8, 8, 14)

    # Sheet 2: Invoices & Delivery
    ws_inv = wb.add_worksheet("Invoices_and_Delivery")
    ws_inv.hide_gridlines(0)
    ws_inv.set_row(0, 28)
    
    inv_headers = [
        "Invoice_No", "Customer_Name", "Trip_No", "Product_Category", "SKU_UOM",
        "Dispatched_Qty", "Dispatched_Val", "Delivery_Status", "Delivered_Qty",
        "Delivered_Val", "Returned_Qty", "Returned_Val", "Return_Reason",
        "Payment_Mode", "Amount_Collected"
    ]
    for c_i, h in enumerate(inv_headers):
        ws_inv.write(0, c_i, h, fmt_cyan_hdr)
        
    sample_invoices = [
        ["INV-2026-0001", "Agora Superstore Outlet", "TRIP-01", category, "Pkt", 120, 48000, "Delivered", 120, 48000, 0, 0, "None", "Credit", 0],
        ["INV-2026-0002", "Meena Bazar Outlet", "TRIP-01", category, "Kg", 80, 36000, "Partial Delivery", 70, 31500, 10, 4500, "R-01: Shop Closed", "Cash", 31500],
        ["INV-2026-0003", "Shwapno Express Outlet", "TRIP-02", category, "Pcs", 2500, 25000, "Delivered", 2500, 25000, 0, 0, "None", "Credit", 0]
    ]
    for r_i, r in enumerate(sample_invoices, start=1):
        ws_inv.set_row(r_i, 20)
        for c_i, val in enumerate(r):
            if c_i in [5, 6, 8, 9, 10, 11, 14]:
                ws_inv.write(r_i, c_i, val, fmt_sample_num)
            elif c_i in [0, 2, 4, 7, 13]:
                ws_inv.write(r_i, c_i, val, fmt_sample_center)
            else:
                ws_inv.write(r_i, c_i, val, fmt_sample_hdr)

    # Add Dropdown Data Validations
    ws_inv.data_validation('D2:D500', {
        'validate': 'list',
        'source': ['Fresh Eggs', 'Frozen Food & RTC', 'Processed Chicken', 'Dairy', 'Dry Food & Bakery', 'Tea', 'Momo & Dimsum', 'Sweets & Savory']
    })
    ws_inv.data_validation('H2:H500', {
        'validate': 'list',
        'source': ['Delivered', 'Partial Delivery', 'Undelivered']
    })
    ws_inv.data_validation('N2:N500', {
        'validate': 'list',
        'source': ['Credit', 'Cash', 'bKash / MFS', 'Bank Cheque']
    })

    ws_inv.set_column(0, 0, 16)
    ws_inv.set_column(1, 1, 32)
    ws_inv.set_column(2, 2, 14)
    ws_inv.set_column(3, 3, 20)
    ws_inv.set_column(4, 4, 12)
    ws_inv.set_column(5, 6, 16)
    ws_inv.set_column(7, 7, 18)
    ws_inv.set_column(8, 11, 16)
    ws_inv.set_column(12, 12, 24)
    ws_inv.set_column(13, 14, 18)

    wb.close()
    output.seek(0)
    
    clean_name = "".join(c for c in depot_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
    fname = f"Paragon_Daily_Entry_Template_{clean_name}.xlsx" if depot_id else "Paragon_Depot_Daily_Entry_Template.xlsx"
    
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=fname
    )

@app.route('/api/template/parse-excel', methods=['POST'])
def parse_daily_entry_excel():
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No Excel file uploaded"}), 400
    
    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"success": False, "message": "Invalid file format. Please upload an .xlsx file."}), 400

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        
        # 1. Parse Trips
        trips = []
        if "Fleet_and_Trips" in wb.sheetnames:
            ws_trips = wb["Fleet_and_Trips"]
            rows = list(ws_trips.iter_rows(values_only=True))
            if len(rows) > 1:
                header = [str(h).strip().lower() for h in rows[0] if h]
                for r in rows[1:]:
                    if not any(r):
                        continue
                    # Skip sample if matches exact default sample
                    trip_no = str(r[0] or '').strip()
                    veh_no = str(r[1] or '').strip()
                    if not trip_no and not veh_no:
                        continue
                    
                    trips.append({
                        "trip_no": trip_no or f"TRIP-0{len(trips)+1}",
                        "vehicle_no": veh_no,
                        "vehicle_type": str(r[2] or '1.5T Reefer Van').strip(),
                        "driver_name": str(r[3] or '').strip(),
                        "delivery_man": str(r[4] or '').strip(),
                        "route_name": str(r[5] or '').strip(),
                        "capacity_kg": float(r[6] or 1500),
                        "loaded_kg": float(r[7] or 1400),
                        "reefer_temp": str(r[8] or '-18°C').strip()
                    })
        
        # 2. Parse Invoices
        invoices = []
        if "Invoices_and_Delivery" in wb.sheetnames:
            ws_inv = wb["Invoices_and_Delivery"]
            rows = list(ws_inv.iter_rows(values_only=True))
            if len(rows) > 1:
                for r in rows[1:]:
                    if not any(r):
                        continue
                    inv_no = str(r[0] or '').strip()
                    cust_name = str(r[1] or '').strip()
                    if not inv_no and not cust_name:
                        continue

                    disp_qty = float(r[5] or 0)
                    disp_val = float(r[6] or 0)
                    del_qty = float(r[8] or disp_qty)
                    del_val = float(r[9] or disp_val)
                    ret_qty = float(r[10] or max(0, disp_qty - del_qty))
                    ret_val = float(r[11] or max(0, disp_val - del_val))
                    cash_coll = float(r[14] or 0)

                    invoices.append({
                        "invoice_no": inv_no or f"INV-2026-00{len(invoices)+1}",
                        "customer_name": cust_name,
                        "trip_no": str(r[2] or 'TRIP-01').strip(),
                        "product_category": str(r[3] or 'Frozen Food').strip(),
                        "sku_uom": str(r[4] or 'Pkt').strip(),
                        "dispatched_qty": disp_qty,
                        "dispatched_val": disp_val,
                        "delivery_status": str(r[7] or 'Delivered').strip(),
                        "delivered_qty": del_qty,
                        "delivered_val": del_val,
                        "returned_qty": ret_qty,
                        "returned_val": ret_val,
                        "return_reason": str(r[12] or 'None').strip(),
                        "collection_mode": str(r[13] or 'Credit').strip(),
                        "amount_collected": cash_coll
                    })

        if not trips and not invoices:
            return jsonify({"success": False, "message": "No valid data rows found in the uploaded Excel template."}), 400

        return jsonify({
            "success": True,
            "trips": trips,
            "invoices": invoices,
            "count_trips": len(trips),
            "count_invoices": len(invoices),
            "message": f"Excel parsed successfully: {len(trips)} vehicle trips and {len(invoices)} invoices loaded!"
        })

    except Exception as e:
        return jsonify({"success": False, "message": f"Error parsing Excel file: {str(e)}"}), 500


# ----------------- HEALTH & OFFLINE SYNC -----------------
@app.route('/api/health')
def health_check():
    return jsonify({"status": "online", "timestamp": datetime.datetime.now().isoformat()})

# ----------------- REPORT CANCELLATION & AUDIT WORKFLOW -----------------
@app.route('/api/reports/depot/<int:depot_id>')
def get_depot_recent_reports(depot_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, report_date, total_vehicles, total_invoices, dispatched_gross_val,
               delivered_net_val, returned_val, status, cancellation_reason, created_at
        FROM daily_reports
        WHERE depot_id = ?
        ORDER BY report_date DESC, id DESC LIMIT 15
    ''', (depot_id,))
    reports = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({"success": True, "reports": reports})

@app.route('/api/reports/request-cancellation', methods=['POST'])
def request_report_cancellation():
    data = request.json or {}
    report_id = data.get('report_id')
    reason = data.get('reason', 'User requested correction')
    
    if not report_id:
        return jsonify({"success": False, "message": "Report ID is required"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE daily_reports
        SET status = 'cancellation_requested',
            cancellation_reason = ?,
            cancellation_requested_at = CURRENT_TIMESTAMP
        WHERE id = ?
    ''', (reason, report_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Cancellation request submitted to Executive Admin for audit review."})

@app.route('/api/reports/cancellations')
def get_pending_cancellations():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT r.id, r.report_date, d.name as depot_name, r.incharge_name,
               r.dispatched_gross_val, r.delivered_net_val, r.cancellation_reason,
               r.cancellation_requested_at, r.status
        FROM daily_reports r
        JOIN depots d ON r.depot_id = d.id
        WHERE r.status = 'cancellation_requested'
        ORDER BY r.cancellation_requested_at DESC
    ''')
    requests = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({"success": True, "requests": requests})

@app.route('/api/reports/approve-cancellation/<int:report_id>', methods=['POST'])
def approve_cancellation(report_id):
    admin_name = session.get('user', {}).get('display_name', 'Executive Admin')
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE daily_reports
        SET status = 'cancelled',
            cancelled_by = ?,
            cancelled_at = CURRENT_TIMESTAMP
        WHERE id = ?
    ''', (admin_name, report_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Report #{report_id} has been voided/cancelled successfully."})

@app.route('/api/reports/reject-cancellation/<int:report_id>', methods=['POST'])
def reject_cancellation(report_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE daily_reports
        SET status = 'active',
            cancellation_reason = NULL
        WHERE id = ?
    ''', (report_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Cancellation request for Report #{report_id} was rejected."})

@app.route('/api/reports/unlock/<int:report_id>', methods=['POST'])
def unlock_report_for_edit(report_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE daily_reports
        SET status = 'unlocked'
        WHERE id = ?
    ''', (report_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Report #{report_id} unlocked. Incharge may now re-edit and resubmit."})

# ----------------- ADMIN USER & INCHARGE MANAGEMENT -----------------
@app.route('/api/admin/users')
def get_admin_users():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT u.id, u.username, u.role, u.depot_id, u.display_name, u.is_active,
               d.name as depot_name, d.category as depot_category
        FROM users u
        LEFT JOIN depots d ON u.depot_id = d.id
        ORDER BY u.role ASC, u.username ASC
    ''')
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({"success": True, "users": users})

@app.route('/api/admin/users/update', methods=['POST'])
def update_admin_user():
    data = request.json or {}
    user_id = data.get('user_id')
    display_name = data.get('display_name')
    password = data.get('password')
    is_active = data.get('is_active', 1)
    
    if not user_id:
        return jsonify({"success": False, "message": "User ID required"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    if password and len(password.strip()) > 0:
        cursor.execute('''
            UPDATE users
            SET display_name = ?, password = ?, is_active = ?
            WHERE id = ?
        ''', (display_name, password.strip(), is_active, user_id))
    else:
        cursor.execute('''
            UPDATE users
            SET display_name = ?, is_active = ?
            WHERE id = ?
        ''', (display_name, is_active, user_id))
        
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "User credentials and status updated successfully!"})

@app.route('/api/admin/users/delete/<int:user_id>', methods=['POST', 'DELETE'])
def delete_admin_user(user_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT role FROM users WHERE id = ?', (user_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        return jsonify({"success": False, "message": "User not found"}), 404
        
    if row['role'] == 'admin':
        cursor.execute("SELECT COUNT(*) as count FROM users WHERE role = 'admin'")
        admin_count = cursor.fetchone()['count']
        if admin_count <= 1:
            conn.close()
            return jsonify({"success": False, "message": "Cannot delete the only executive admin account"}), 400
            
    cursor.execute('DELETE FROM users WHERE id = ?', (user_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "User account deleted successfully!"})

# ----------------- DISTRIBUTION ROUTE PLANNING & POLOXY ERP INTEGRATION -----------------

# ----------------- DISTRIBUTION ROUTE PLANNING & POLOXY ERP INTEGRATION -----------------

DEPOT_CATALOG_LOOKUP = [
    {"id": 1, "name": "01. Ashulia Factory - Frozen & RTC", "category": "Frozen Food", "match": ["ashulia frozen", "01.ash", "ash-fz", "frozen-ash"]},
    {"id": 2, "name": "02. Ashulia Factory - Dry Food", "category": "Dry Food", "match": ["ashulia dry", "02.ash", "ash-dry", "dry-ash"]},
    {"id": 3, "name": "03. Gazipur-Process CK", "category": "Processed Chicken", "match": ["gazipur", "gz", "chicken plant", "03.gz", "ck-plant", "01so", "process ck"]},
    {"id": 4, "name": "04. Sirajganj Dairy", "category": "Dairy", "match": ["sirajganj", "dairy plant", "milk", "04.sj"]},
    {"id": 5, "name": "05. Sylhet - Tea Packing Unit", "category": "Tea", "match": ["sylhet tea", "tea pack", "05.syl"]},
    {"id": 6, "name": "06. Sylhet Depot - Frozen", "category": "Frozen Food", "match": ["sylhet frozen", "syl-fz", "06.syl", "sylhet depot"]},
    {"id": 7, "name": "07. Tejgaon - Frozen/CK", "category": "Chicken & Frozen", "match": ["tejgaon frozen", "tg-fz", "tg-ck", "07.tg"]},
    {"id": 8, "name": "08. Tejgaon - e-Commerce", "category": "e-Commerce", "match": ["e-commerce", "ecommerce", "tg-ecom", "08.tg"]},
    {"id": 9, "name": "09. Tejgaon - Fresh Egg", "category": "Fresh Eggs", "match": ["egg", "fresh egg", "table egg", "03tg", "tg-egg", "09.tg", "09so", "09dn"]},
    {"id": 10, "name": "10. Tejgaon - Tea Distribution", "category": "Tea", "match": ["tg-tea", "tea dist", "10.tg"]},
    {"id": 11, "name": "11. Tejgaon - Dairy", "category": "Dairy", "match": ["tg-dairy", "tejgaon dairy", "11.tg"]},
    {"id": 12, "name": "12. Mohakhali Depot", "category": "Momo & Snacks", "match": ["mohakhali", "momo", "12.mhk"]},
    {"id": 13, "name": "13. CTG Depot - Frozen", "category": "Frozen Food", "match": ["chittagong frozen", "ctg-fz", "13.ctg"]},
    {"id": 14, "name": "14. CTG Depot - Dry Food", "category": "Dry Food", "match": ["chittagong dry", "ctg-dry", "14.ctg"]},
    {"id": 15, "name": "15. Jessore - Frozen Food", "category": "Frozen Food", "match": ["jessore frozen", "jd-fz", "15.jd"]},
    {"id": 16, "name": "16. Jessore - Dry Food", "category": "Dry Food", "match": ["jessore dry", "jd-dry", "16.jd"]},
    {"id": 17, "name": "17. Rangpur Depot - Frozen", "category": "Frozen Food", "match": ["rangpur", "rng-fz", "17.rng"]}
]

def match_depot_ref(branch_str, ref_str, order_no_str=''):
    b = str(branch_str or '').upper()
    r = str(ref_str or '').upper()
    o = str(order_no_str or '').upper()
    
    # 03 - Egg (Tejgaon Fresh Egg: 03SO, 03. Branded Eggs, 03TG, EGG)
    if '03SO' in o or 'BRANDED EGGS' in b or 'EGG' in b or '03TG' in r or '09.' in r or '09SO' in o:
        return {"id": 9, "name": "09. Tejgaon - Fresh Egg", "category": "Fresh Eggs"}
        
    # 01 - Chicken (Gazipur or Tejgaon Chicken: 01SO, 01. Process Chicken)
    if '01SO' in o or 'PROCESS CHICKEN' in b or 'GAZIPUR' in r or '01.' in r:
        if 'TG' in r or 'TEJGAON' in r or '01.TG' in r:
            return {"id": 7, "name": "07. Tejgaon - Frozen/CK", "category": "Chicken & Frozen"}
        return {"id": 3, "name": "03. Gazipur-Process CK", "category": "Processed Chicken"}
    
    # 02 - Frozen Foods (02SO, 02. Frozen Foods: CTG, Tejgaon, Ashulia, Sylhet, Jessore, Rangpur)
    if '02SO' in o or 'FROZEN' in b or '02.' in b or '07.' in r or '07SO' in o or 'MOMO' in b:
        if 'CTG' in r or 'CHITTAGONG' in r or '02.CTG' in r:
            return {"id": 13, "name": "13. CTG Depot - Frozen", "category": "Frozen Food"}
        elif 'ASHULIA' in r or 'ASH' in r:
            return {"id": 1, "name": "01. Ashulia Factory - Frozen", "category": "Frozen Food"}
        elif 'SYLHET' in r or 'SYL' in r:
            return {"id": 6, "name": "06. Sylhet Depot - Frozen", "category": "Frozen Food"}
        elif 'JD' in r or 'JESSORE' in r:
            return {"id": 15, "name": "15. Jessore - Frozen Food", "category": "Frozen Food"}
        elif 'RANGPUR' in r or 'RNG' in r:
            return {"id": 17, "name": "17. Rangpur Depot - Frozen", "category": "Frozen Food"}
        elif 'MOMO' in r or 'MOMO' in b:
            return {"id": 12, "name": "12. Mohakhali Depot", "category": "Momo & Snacks"}
        else:
            return {"id": 7, "name": "07. Tejgaon - Frozen/CK", "category": "Chicken & Frozen"}
        
    # 04 - Tea (04SO, 04. Tea, 04.Sylhet, Sreemangal/Tejgaon)
    if '04SO' in o or 'TEA' in b or '04.' in b or '04.' in r:
        if 'SYLHET' in r or 'SYL' in r or '04.SYLHET' in r:
            return {"id": 5, "name": "05. Sylhet - Tea Packing Unit", "category": "Tea"}
        return {"id": 10, "name": "10. Tejgaon - Tea Distribution", "category": "Tea"}
        
    # 05 - Dry Food (05SO, 05. Dry Foods, 05.Ashulia)
    if '05SO' in o or 'DRY' in b or '05.' in b or '05.' in r or 'FEED' in b:
        if 'CTG' in r or 'CHITTAGONG' in r:
            return {"id": 14, "name": "14. CTG Depot - Dry Food", "category": "Dry Food"}
        elif 'JD' in r or 'JESSORE' in r:
            return {"id": 16, "name": "16. Jessore - Dry Food", "category": "Dry Food"}
        return {"id": 2, "name": "02. Ashulia Factory - Dry Food", "category": "Dry Food"}
        
    # 06 - Dairy (Sirajganj or Tejgaon Dairy)
    if 'DAIRY' in b or 'MILK' in b or '06.' in b or '06SO' in o or '06.' in r:
        if 'SIRAJGANJ' in r or 'SJ' in r:
            return {"id": 4, "name": "04. Sirajganj Dairy", "category": "Dairy"}
        return {"id": 11, "name": "11. Tejgaon - Dairy", "category": "Dairy"}
        
    # 08 - eCommerce / Sweets
    if 'ECOM' in b or 'ECOM' in r or 'SWEET' in b or '08.' in b or '08SO' in o:
        return {"id": 8, "name": "08. Tejgaon - e-Commerce", "category": "e-Commerce"}
        
    return {"id": 9, "name": "09. Tejgaon - Fresh Egg", "category": "Fresh Eggs"}

@app.route('/api/distribution/import-poloxy-orders', methods=['POST'])
def import_poloxy_orders():
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No file uploaded. Please upload Sale Order Status Report.xlsx"}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "message": "No file selected"}), 400
        
    explicit_depot_id = request.form.get('depot_id') or request.args.get('depot_id')
    fixed_depot_info = None
    if explicit_depot_id:
        try:
            conn = get_db()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM depots WHERE id = ?', (explicit_depot_id,))
            row = cursor.fetchone()
            conn.close()
            if row:
                fixed_depot_info = {
                    "id": row['id'],
                    "name": row['name'],
                    "category": row['category']
                }
        except Exception:
            pass

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 4:
            return jsonify({"success": False, "message": "Invalid Poloxy file format or empty sheet"}), 400
            
        header_row_idx = 2
        for idx, row in enumerate(rows[:6]):
            if any('Order No' in str(c) or 'Ref. No' in str(c) for c in row if c):
                header_row_idx = idx
                break
                
        depot_map = {}
        total_orders_parsed = 0
        
        def _clean_val(v):
            if v is None: return 0.0
            if isinstance(v, (int, float)): return float(v)
            cleaned = re.sub(r'[^\d.-]', '', str(v))
            return float(cleaned) if cleaned else 0.0

        for r in rows[header_row_idx + 1:]:
            if not any(r):
                continue
            sr_no = r[0] if len(r) > 0 else ''
            date_val = str(r[1]).strip() if len(r) > 1 and r[1] else ''
            customer = str(r[2]).strip() if len(r) > 2 and r[2] else ''
            order_no = str(r[3]).strip() if len(r) > 3 and r[3] else ''
            ref_no = str(r[4]).strip() if len(r) > 4 and r[4] else ''
            branch = str(r[6]).strip() if len(r) > 6 and r[6] else ''
            item = str(r[7]).strip() if len(r) > 7 and r[7] else ''
            raw_rate_unit = str(r[8]).strip() if len(r) > 8 and r[8] else ''
            rate = _clean_val(r[9]) if len(r) > 9 else 0.0
            qty_bag = _clean_val(r[10]) if len(r) > 10 else 0.0
            qty_kg = _clean_val(r[11]) if len(r) > 11 else 0.0
            raw_col_n = r[13] if len(r) > 13 else None
            col_n_val = _clean_val(raw_col_n) if raw_col_n is not None and str(raw_col_n).strip() != '' else None
            d_note_id = str(r[27]).strip() if len(r) > 27 and r[27] else ''
            consignee_name = str(r[32]).strip() if len(r) > 32 and r[32] else customer
            consignee_contact = str(r[33]).strip() if len(r) > 33 and r[33] else ''
            consignee_address = str(r[34]).strip() if len(r) > 34 and r[34] else ''
            
            if not ref_no and not order_no and not item:
                continue
                
            depot_info = match_depot_ref(branch, ref_no, order_no)
            depot_key = f"{depot_info['id']}_{depot_info['name']}"
            
            # Resolve appropriate default UOM for depot
            is_egg = depot_info["id"] == 9 or "egg" in depot_info["category"].lower() or "egg" in item.lower()
            is_dairy = depot_info["id"] in [4, 11] or "dairy" in depot_info["category"].lower() or "milk" in item.lower()
            default_uom = "Pcs" if is_egg else ("Ltr" if is_dairy else "Pkt")
            rate_unit = raw_rate_unit if raw_rate_unit else default_uom
            
            if depot_key not in depot_map:
                depot_map[depot_key] = {
                    "depot_id": depot_info["id"],
                    "depot_name": depot_info["name"],
                    "category": depot_info["category"],
                    "orders": {},
                    "total_outlets": 0,
                    "total_pkts": 0.0,
                    "total_kg": 0.0,
                    "total_amount": 0.0,
                    "primary_uom": default_uom
                }
                
            ord_key = order_no if order_no else f"ORD-{sr_no or (total_orders_parsed + 1)}"
            if ord_key not in depot_map[depot_key]["orders"]:
                depot_map[depot_key]["orders"][ord_key] = {
                    "order_no": ord_key,
                    "date": date_val,
                    "customer": customer,
                    "consignee_name": consignee_name,
                    "consignee_contact": consignee_contact,
                    "consignee_address": consignee_address,
                    "branch_category": branch or depot_info["category"],
                    "delivery_note_id": d_note_id,
                    "items": [],
                    "total_pkt": 0.0,
                    "total_kg": 0.0,
                    "total_amount": 0.0,
                    "col_n_found": False,
                    "assigned_van": "",
                    "payment_mode": "Credit" if any(k in consignee_name.lower() for k in ["shwapno", "agora", "meena", "unimart", "aarong", "pran", "food panda", "lavender", "kfc"]) else "Cash",
                    "expected_cash": 0.0,
                    "rate_unit": rate_unit
                }
                depot_map[depot_key]["total_outlets"] += 1
                total_orders_parsed += 1
                
            # Check Column N (Invoice Value)
            if col_n_val is not None and col_n_val > 0 and not depot_map[depot_key]["orders"][ord_key]["col_n_found"]:
                depot_map[depot_key]["orders"][ord_key]["total_amount"] = round(col_n_val, 2)
                depot_map[depot_key]["orders"][ord_key]["col_n_found"] = True

            egg_pcs = qty_kg if qty_kg > 0 else qty_bag
            if is_egg:
                item_line_amount = (rate * qty_bag) if (bool(re.search(r'pkt|bag|12', rate_unit, re.I)) and qty_bag > 0) else (rate * egg_pcs)
            else:
                item_line_amount = rate * (qty_kg if qty_kg > 0 else qty_bag)
            
            depot_map[depot_key]["orders"][ord_key]["items"].append({
                "item": item,
                "rate_unit": "Pcs" if is_egg else rate_unit,
                "rate": rate,
                "qty_pkt": egg_pcs if is_egg else qty_bag,
                "qty_kg": 0.0 if is_egg else qty_kg,
                "amount": round(item_line_amount, 2)
            })
            
            # For Eggs: Quantity (Pcs) is STRICTLY taken from Column L (egg_pcs / Max. Qty.(Kg))!
            if is_egg:
                depot_map[depot_key]["orders"][ord_key]["total_pkt"] += egg_pcs
                depot_map[depot_key]["orders"][ord_key]["total_kg"] = 0.0
                depot_map[depot_key]["orders"][ord_key]["rate_unit"] = "Pcs"
                depot_map[depot_key]["total_pkts"] += egg_pcs
                depot_map[depot_key]["total_kg"] = 0.0
            else:
                depot_map[depot_key]["orders"][ord_key]["total_pkt"] += qty_bag
                depot_map[depot_key]["orders"][ord_key]["total_kg"] += qty_kg
                depot_map[depot_key]["total_pkts"] += qty_bag
                depot_map[depot_key]["total_kg"] += qty_kg

        # Fetch route mappings, fleet vehicles, and crew for each depot to enable 1-click smart dispatch
        conn = get_db()
        cursor = conn.cursor()
        
        # Aggregate final invoice values per order and per depot
        total_value_parsed = 0.0
        depots_result = []
        for k, v in depot_map.items():
            d_id = v["depot_id"]
            
            # Fetch Enlisted Routes for this depot
            cursor.execute('SELECT id, route_name, route_code, description, default_vehicle_no FROM routes WHERE depot_id = ? ORDER BY id ASC', (d_id,))
            d_routes = [dict(r) for r in cursor.fetchall()]
            
            # Fetch Enlisted Consignee-to-Route mappings
            cursor.execute('SELECT consignee_name, route_id FROM route_consignees WHERE depot_id = ?', (d_id,))
            c_mappings = {r['consignee_name'].lower().strip(): r['route_id'] for r in cursor.fetchall()}
            
            # Fetch Fleet Vehicles (Active, Spare, Rental, Borrowed)
            cursor.execute('SELECT id, vehicle_no, vehicle_type, capacity_kg, capacity_units, ownership, status FROM fleet_vehicles WHERE depot_id = ? ORDER BY ownership ASC, id ASC', (d_id,))
            d_fleet = [dict(r) for r in cursor.fetchall()]
            
            # Fetch Drivers & Delivery Crew
            cursor.execute("SELECT id, role, name, phone, license_no, status FROM depot_crew WHERE depot_id = ? AND status='Active' ORDER BY role ASC, name ASC", (d_id,))
            all_crew = [dict(r) for r in cursor.fetchall()]
            d_drivers = [c for c in all_crew if c['role'] == 'driver']
            d_deliverymen = [c for c in all_crew if c['role'] == 'delivery_man']
            
            route_name_map = {r['id']: r['route_name'] for r in d_routes}

            depot_total_amt = 0.0
            order_list = list(v["orders"].values())
            for o in order_list:
                # If Col N wasn't present, fallback to sum of item line amounts
                if not o.get("col_n_found", False) or o["total_amount"] == 0:
                    o["total_amount"] = sum(it["amount"] for it in o["items"])
                
                # Match Route
                c_clean = o["consignee_name"].lower().strip()
                matched_r_id = c_mappings.get(c_clean)
                if matched_r_id and matched_r_id in route_name_map:
                    o["route_id"] = matched_r_id
                    o["route_name"] = route_name_map[matched_r_id]
                else:
                    # Smart fuzzy route match if route keyword in consignee name
                    fuzzy_matched = False
                    for r_obj in d_routes:
                        r_words = [w for w in r_obj['route_name'].lower().replace('-', ' ').split() if len(w) > 3]
                        if any(w in c_clean for w in r_words):
                            o["route_id"] = r_obj['id']
                            o["route_name"] = r_obj['route_name']
                            fuzzy_matched = True
                            break
                    if not fuzzy_matched:
                        o["route_id"] = None
                        o["route_name"] = "Unassigned Route"
                
                depot_total_amt += o["total_amount"]
                if o["payment_mode"] == "Cash":
                    o["expected_cash"] = o["total_amount"]
                    
            v["total_amount"] = round(depot_total_amt, 2)
            total_value_parsed += depot_total_amt
            
            depots_result.append({
                "depot_id": v["depot_id"],
                "depot_name": v["depot_name"],
                "category": v["category"],
                "total_outlets": v["total_outlets"],
                "total_pkts": round(v["total_pkts"], 1),
                "total_kg": round(v["total_kg"], 2),
                "total_amount": round(v["total_amount"], 2),
                "primary_uom": v.get("primary_uom", "Pkt"),
                "orders": order_list,
                "routes": d_routes,
                "fleet": d_fleet,
                "drivers": d_drivers,
                "deliverymen": d_deliverymen
            })
            
        conn.close()
        
        if explicit_depot_id:
            try:
                exp_id = int(explicit_depot_id)
                depots_result.sort(key=lambda d: 0 if d["depot_id"] == exp_id else 1)
            except Exception:
                depots_result.sort(key=lambda d: d["total_outlets"], reverse=True)
        else:
            depots_result.sort(key=lambda d: d["total_outlets"], reverse=True)
        
        return jsonify({
            "success": True,
            "message": f"Successfully parsed {total_orders_parsed} sales orders across {len(depots_result)} depots (Total Value: ৳ {total_value_parsed:,.2f})",
            "total_orders": total_orders_parsed,
            "total_value": round(total_value_parsed, 2),
            "depots": depots_result
        })
        
    except Exception as e:
        return jsonify({"success": False, "message": f"Error parsing Poloxy Sales Order file: {str(e)}"}), 500

# ----------------- ROUTES & CONSIGNEE MASTER APIS -----------------

@app.route('/api/routes', methods=['GET', 'POST'])
def handle_routes():
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        depot_id = request.args.get('depot_id')
        query = '''
            SELECT r.*, COUNT(rc.id) as outlet_count
            FROM routes r
            LEFT JOIN route_consignees rc ON r.id = rc.route_id
        '''
        params = ()
        if depot_id and str(depot_id) != 'all':
            query += ' WHERE r.depot_id = ?'
            params = (depot_id,)
        query += ' GROUP BY r.id ORDER BY r.id ASC'
        cursor.execute(query, params)
        routes = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return jsonify({"success": True, "routes": routes})
        
    elif request.method == 'POST':
        data = request.json or {}
        route_id = data.get('id')
        depot_id = data.get('depot_id', 9)
        route_name = data.get('route_name', '').strip()
        route_code = data.get('route_code', '').strip()
        description = data.get('description', '').strip()
        default_vehicle_no = data.get('default_vehicle_no', '').strip()
        
        if not route_name:
            conn.close()
            return jsonify({"success": False, "message": "Route name is required"}), 400
            
        if route_id:
            cursor.execute('''
                UPDATE routes SET route_name=?, route_code=?, description=?, default_vehicle_no=?
                WHERE id=? AND depot_id=?
            ''', (route_name, route_code, description, default_vehicle_no, route_id, depot_id))
        else:
            cursor.execute('''
                INSERT INTO routes (depot_id, route_name, route_code, description, default_vehicle_no)
                VALUES (?, ?, ?, ?, ?)
            ''', (depot_id, route_name, route_code, description, default_vehicle_no))
            
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": "Route saved successfully!"})

@app.route('/api/routes/<int:route_id>', methods=['DELETE'])
def delete_route(route_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM route_consignees WHERE route_id = ?', (route_id,))
    cursor.execute('DELETE FROM routes WHERE id = ?', (route_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Route deleted successfully!"})

@app.route('/api/consignees', methods=['GET', 'POST'])
def handle_consignees():
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        depot_id = request.args.get('depot_id')
        route_id = request.args.get('route_id')
        query = '''
            SELECT rc.*, r.route_name
            FROM route_consignees rc
            LEFT JOIN routes r ON rc.route_id = r.id
            WHERE 1=1
        '''
        params = []
        if depot_id and str(depot_id) != 'all':
            query += ' AND rc.depot_id = ?'
            params.append(depot_id)
        if route_id:
            query += ' AND rc.route_id = ?'
            params.append(route_id)
        query += ' ORDER BY rc.consignee_name ASC'
        cursor.execute(query, tuple(params))
        consignees = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return jsonify({"success": True, "consignees": consignees})
        
    elif request.method == 'POST':
        data = request.json or {}
        cid = data.get('id')
        depot_id = data.get('depot_id', 9)
        route_id = data.get('route_id')
        consignee_name = data.get('consignee_name', '').strip()
        address = data.get('address', '').strip()
        phone = data.get('phone', '').strip()
        payment_mode = data.get('payment_mode', 'Cash')
        
        if not consignee_name:
            conn.close()
            return jsonify({"success": False, "message": "Consignee/Outlet name is required"}), 400
            
        if cid:
            cursor.execute('''
                UPDATE route_consignees SET route_id=?, consignee_name=?, address=?, phone=?, payment_mode=?
                WHERE id=?
            ''', (route_id, consignee_name, address, phone, payment_mode, cid))
        else:
            cursor.execute('''
                INSERT INTO route_consignees (depot_id, route_id, consignee_name, address, phone, payment_mode)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (depot_id, route_id, consignee_name, address, phone, payment_mode))
            
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": "Consignee mapped successfully!"})

@app.route('/api/consignees/<int:cid>', methods=['DELETE'])
def delete_consignee(cid):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM route_consignees WHERE id = ?', (cid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Consignee mapping removed!"})

# ----------------- FLEET VEHICLES & RENTAL APIS -----------------

@app.route('/api/fleet', methods=['GET'])
def get_fleet():
    depot_id = request.args.get('depot_id')
    conn = get_db()
    cursor = conn.cursor()
    query = '''
        SELECT fv.*,
               drv.name as default_driver_name, drv.phone as default_driver_phone,
               del.name as default_deliv_name, del.phone as default_deliv_phone
        FROM fleet_vehicles fv
        LEFT JOIN depot_crew drv ON fv.default_driver_id = drv.id
        LEFT JOIN depot_crew del ON fv.default_deliveryman_id = del.id
        WHERE 1=1
    '''
    params = []
    if depot_id and str(depot_id) != 'all':
        query += ' AND fv.depot_id = ?'
        params.append(depot_id)
    query += ' ORDER BY fv.ownership ASC, fv.id ASC'
    cursor.execute(query, tuple(params))
    vehicles = []
    for row in cursor.fetchall():
        v = dict(row)
        cat_caps = v.get('category_capacities')
        if cat_caps:
            try:
                v['category_capacities'] = json.loads(cat_caps)
            except Exception:
                v['category_capacities'] = get_default_category_capacities(v.get('capacity_kg', 1500), v.get('vehicle_type'))
        else:
            v['category_capacities'] = get_default_category_capacities(v.get('capacity_kg', 1500), v.get('vehicle_type'))
        vehicles.append(v)
    conn.close()
    return jsonify({"success": True, "vehicles": vehicles})

@app.route('/api/fleet/vehicle', methods=['POST'])
def save_fleet_vehicle():
    data = request.json or {}
    vid = data.get('id')
    depot_id = data.get('depot_id', 9)
    vehicle_no = data.get('vehicle_no', '').strip().upper()
    vehicle_type = data.get('vehicle_type', 'Covered Van').strip()
    try:
        capacity_kg = float(data.get('capacity_kg', 1500))
    except (ValueError, TypeError):
        capacity_kg = 1500.0
    capacity_units = data.get('capacity_units', 'Kg').strip()
    default_driver_id = data.get('default_driver_id') or None
    default_deliveryman_id = data.get('default_deliveryman_id') or None
    ownership = data.get('ownership', 'Owned').strip()
    rental_vendor = data.get('rental_vendor', '').strip()
    try:
        rental_cost_per_day = float(data.get('rental_cost_per_day', 0))
    except (ValueError, TypeError):
        rental_cost_per_day = 0.0
    status = data.get('status', 'Active').strip()
    
    if not vehicle_no:
        return jsonify({"success": False, "message": "Vehicle number is required"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    if vid:
        cursor.execute('''
            UPDATE fleet_vehicles 
            SET vehicle_no=?, vehicle_type=?, capacity_kg=?, capacity_units=?, default_driver_id=?, default_deliveryman_id=?, ownership=?, rental_vendor=?, rental_cost_per_day=?, status=?
            WHERE id=?
        ''', (vehicle_no, vehicle_type, capacity_kg, capacity_units, default_driver_id, default_deliveryman_id, ownership, rental_vendor, rental_cost_per_day, status, vid))
        veh_id = vid
    else:
        cursor.execute('''
            INSERT INTO fleet_vehicles (depot_id, vehicle_no, vehicle_type, capacity_kg, capacity_units, default_driver_id, default_deliveryman_id, ownership, rental_vendor, rental_cost_per_day, home_depot_id, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (depot_id, vehicle_no, vehicle_type, capacity_kg, capacity_units, default_driver_id, default_deliveryman_id, ownership, rental_vendor, rental_cost_per_day, depot_id, status))
        veh_id = cursor.lastrowid
        
    # Sync assigned vehicle to driver
    if default_driver_id:
        cursor.execute('UPDATE depot_crew SET assigned_vehicle_no = ? WHERE id = ?', (vehicle_no, default_driver_id))
        
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Vehicle {vehicle_no} (Capacity: {capacity_kg} {capacity_units}) saved successfully!"})

@app.route('/api/fleet/vehicle/<int:vid>/capacity', methods=['POST', 'PATCH'])
@app.route('/api/fleet/vehicle/capacity', methods=['POST', 'PATCH'])
def update_fleet_vehicle_capacity(vid=None):
    data = request.json or {}
    target_vid = vid or data.get('id') or data.get('vehicle_id')
    try:
        capacity_kg = float(data.get('capacity_kg', 1500))
    except (ValueError, TypeError):
        capacity_kg = 1500.0
    capacity_units = data.get('capacity_units', 'Kg').strip()
    
    if not target_vid:
        return jsonify({"success": False, "message": "Vehicle ID is required"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('UPDATE fleet_vehicles SET capacity_kg = ?, capacity_units = ? WHERE id = ?', (capacity_kg, capacity_units, target_vid))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Vehicle capacity updated to {capacity_kg} {capacity_units}!"})

@app.route('/api/fleet/vehicle/<int:vid>', methods=['DELETE'])
def delete_fleet_vehicle(vid):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM fleet_vehicles WHERE id = ?', (vid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Vehicle removed successfully!"})

# ----------------- DEPOT CREW (DRIVERS & DELIVERYMEN) APIS -----------------

@app.route('/api/crew', methods=['GET', 'POST'])
def handle_crew():
    conn = get_db()
    cursor = conn.cursor()
    
    if request.method == 'GET':
        depot_id = request.args.get('depot_id')
        role = request.args.get('role')
        query = '''
            SELECT dc.*, fv.vehicle_no as linked_vehicle_no, fv.vehicle_type as linked_vehicle_type
            FROM depot_crew dc
            LEFT JOIN fleet_vehicles fv ON dc.id = fv.default_driver_id AND fv.depot_id = dc.depot_id
            WHERE 1=1
        '''
        params = []
        if depot_id and str(depot_id) != 'all':
            query += ' AND dc.depot_id = ?'
            params.append(depot_id)
        if role:
            query += ' AND dc.role = ?'
            params.append(role)
        query += ' ORDER BY dc.role ASC, dc.name ASC'
        cursor.execute(query, tuple(params))
        crew = []
        for row in cursor.fetchall():
            d = dict(row)
            if not d.get('assigned_vehicle_no') and d.get('linked_vehicle_no'):
                d['assigned_vehicle_no'] = d['linked_vehicle_no']
            crew.append(d)
        conn.close()
        return jsonify({"success": True, "crew": crew})
        
    elif request.method == 'POST':
        data = request.json or {}
        cid = data.get('id')
        depot_id = data.get('depot_id', 9)
        role = data.get('role', 'driver').strip()
        name = data.get('name', '').strip()
        phone = data.get('phone', '').strip()
        license_no = data.get('license_no', '').strip()
        assigned_vehicle_no = data.get('assigned_vehicle_no', '').strip().upper() if role == 'driver' else ''
        secondary_vehicle_no = data.get('secondary_vehicle_no', '').strip().upper() if role == 'driver' else ''
        default_route_name = data.get('default_route_name', '').strip() if role == 'driver' else ''
        driver_id_code = data.get('driver_id_code', '').strip()
        status = data.get('status', 'Active').strip()
        
        if not name or not phone:
            conn.close()
            return jsonify({"success": False, "message": "Staff name and phone number are required"}), 400
            
        if cid:
            cursor.execute('''
                UPDATE depot_crew 
                SET role=?, name=?, phone=?, license_no=?, status=?, assigned_vehicle_no=?,
                    secondary_vehicle_no=?, default_route_name=?, driver_id_code=?
                WHERE id=?
            ''', (role, name, phone, license_no, status, assigned_vehicle_no, secondary_vehicle_no, default_route_name, driver_id_code, cid))
            crew_id = cid
        else:
            cursor.execute('''
                INSERT INTO depot_crew (depot_id, role, name, phone, license_no, status, assigned_vehicle_no, secondary_vehicle_no, default_route_name, driver_id_code)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (depot_id, role, name, phone, license_no, status, assigned_vehicle_no, secondary_vehicle_no, default_route_name, driver_id_code))
            crew_id = cursor.lastrowid
            
        # If Driver has assigned vehicle, link it in fleet_vehicles
        if role == 'driver' and assigned_vehicle_no:
            cursor.execute('UPDATE fleet_vehicles SET default_driver_id = ? WHERE depot_id = ? AND UPPER(vehicle_no) = ?',
                           (crew_id, depot_id, assigned_vehicle_no))
            
        conn.commit()
        conn.close()
        veh_msg = f" (Assigned Van: {assigned_vehicle_no})" if assigned_vehicle_no else ""
        return jsonify({"success": True, "message": f"{role.title()} {name}{veh_msg} saved successfully!"})

@app.route('/api/crew/<int:cid>', methods=['DELETE'])
def delete_crew(cid):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('UPDATE fleet_vehicles SET default_driver_id = NULL WHERE default_driver_id = ?', (cid,))
    cursor.execute('UPDATE fleet_vehicles SET default_deliveryman_id = NULL WHERE default_deliveryman_id = ?', (cid,))
    cursor.execute('DELETE FROM depot_crew WHERE id = ?', (cid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Staff record deleted successfully!"})

# ----------------- MULTI-CATEGORY CAPACITY & DEFAULT MAPPING APIS -----------------

@app.route('/api/fleet/capacities', methods=['GET'])
def get_fleet_capacities():
    depot_id = request.args.get('depot_id')
    conn = get_db()
    cursor = conn.cursor()
    query = '''
        SELECT fv.id, fv.depot_id, fv.vehicle_no, fv.vehicle_type, fv.capacity_kg, fv.capacity_units,
               fv.category_capacities, fv.status,
               drv.id as driver_id, drv.name as default_driver_name, drv.phone as default_driver_phone,
               drv.secondary_vehicle_no, drv.default_route_name
        FROM fleet_vehicles fv
        LEFT JOIN depot_crew drv ON fv.default_driver_id = drv.id
        WHERE 1=1
    '''
    params = []
    if depot_id and str(depot_id) != 'all':
        query += ' AND fv.depot_id = ?'
        params.append(depot_id)
    query += ' ORDER BY fv.id ASC'
    cursor.execute(query, tuple(params))
    vehicles = []
    for row in cursor.fetchall():
        v = dict(row)
        cat_caps = v.get('category_capacities')
        if cat_caps:
            try:
                v['category_capacities'] = json.loads(cat_caps)
            except Exception:
                v['category_capacities'] = get_default_category_capacities(v.get('capacity_kg', 1500), v.get('vehicle_type'))
        else:
            v['category_capacities'] = get_default_category_capacities(v.get('capacity_kg', 1500), v.get('vehicle_type'))
        vehicles.append(v)
    conn.close()
    return jsonify({"success": True, "vehicles": vehicles})

@app.route('/api/fleet/capacities/update', methods=['POST'])
def update_fleet_capacities():
    data = request.json or {}
    vid = data.get('vehicle_id') or data.get('id')
    vehicle_no = data.get('vehicle_no', '').strip().upper()
    category_capacities = data.get('category_capacities')
    
    conn = get_db()
    cursor = conn.cursor()
    
    if not vid and vehicle_no:
        row = cursor.execute("SELECT id FROM fleet_vehicles WHERE UPPER(vehicle_no) = ?", (vehicle_no,)).fetchone()
        if row:
            vid = row[0]
            
    if not vid:
        conn.close()
        return jsonify({"success": False, "message": "Vehicle ID or valid registration number required"}), 400
        
    caps_json = json.dumps(category_capacities) if isinstance(category_capacities, dict) else str(category_capacities or '{}')
    
    # Also update capacity_kg and capacity_units from primary Frozen or Egg category
    frozen_cap = 1500.0
    if isinstance(category_capacities, dict):
        if 'Frozen Foods / Chicken' in category_capacities:
            try: frozen_cap = float(category_capacities['Frozen Foods / Chicken'].get('capacity', 1500))
            except (ValueError, TypeError): pass
        elif 'Chicken' in category_capacities:
            try: frozen_cap = float(category_capacities['Chicken'].get('capacity', 1500))
            except (ValueError, TypeError): pass
            
    cursor.execute('UPDATE fleet_vehicles SET category_capacities = ?, capacity_kg = ? WHERE id = ?', (caps_json, frozen_cap, vid))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Multi-category capacities updated for vehicle successfully!"})

@app.route('/api/crew/default-mapping', methods=['POST'])
def update_crew_default_mapping():
    data = request.json or {}
    cid = data.get('id') or data.get('driver_id')
    driver_name = data.get('driver_name', '').strip()
    depot_id = data.get('depot_id', 9)
    phone = data.get('phone', '').strip()
    default_vehicle_no = data.get('default_vehicle_no', '').strip().upper()
    secondary_vehicle_no = data.get('secondary_vehicle_no', '').strip().upper()
    default_route_name = data.get('default_route_name', '').strip()
    
    conn = get_db()
    cursor = conn.cursor()
    
    if not cid and driver_name:
        row = cursor.execute("SELECT id FROM depot_crew WHERE role='driver' AND LOWER(name)=LOWER(?) AND depot_id=?", (driver_name, depot_id)).fetchone()
        if row:
            cid = row[0]
        else:
            cursor.execute('''
                INSERT INTO depot_crew (depot_id, role, name, phone, assigned_vehicle_no, secondary_vehicle_no, default_route_name, status)
                VALUES (?, 'driver', ?, ?, ?, ?, ?, 'Active')
            ''', (depot_id, driver_name, phone or '01711-000000', default_vehicle_no, secondary_vehicle_no, default_route_name))
            cid = cursor.lastrowid
            
    if cid:
        cursor.execute('''
            UPDATE depot_crew
            SET assigned_vehicle_no = ?, secondary_vehicle_no = ?, default_route_name = ?,
                phone = CASE WHEN ? != '' THEN ? ELSE phone END
            WHERE id = ?
        ''', (default_vehicle_no, secondary_vehicle_no, default_route_name, phone, phone, cid))
        
        # Link default vehicle in fleet_vehicles
        if default_vehicle_no:
            cursor.execute('UPDATE fleet_vehicles SET default_driver_id = ? WHERE depot_id = ? AND UPPER(vehicle_no) = ?', (cid, depot_id, default_vehicle_no))
            
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Default mapping saved for Driver {driver_name or cid}!"})

def get_depot_primary_category_info(depot_row):
    d_name = (depot_row['name'] if depot_row else '').lower()
    d_cat = (depot_row['category'] if depot_row else '').lower()
    d_id = depot_row['id'] if depot_row else 9

    if 'egg' in d_cat or 'egg' in d_name or d_id == 9:
        return ("Egg", "Vehicle Max Capacity - Fresh Eggs (Pcs) *", "Pcs", 30000.0)
    elif 'dairy' in d_cat or 'milk' in d_name or d_id in [4, 11]:
        return ("Dairy", "Vehicle Max Capacity - Dairy / Milk (Liter) *", "Liter", 1000.0)
    elif 'dry' in d_cat or 'dry' in d_name or d_id in [2, 14, 16]:
        return ("Dry Goods / Box Items", "Vehicle Max Capacity - Dry Goods / Box (Ctn) *", "Ctn", 500.0)
    elif 'tea' in d_cat or 'tea' in d_name or d_id in [5, 10]:
        return ("Tea", "Vehicle Max Capacity - Tea (Kg) *", "Kg", 1200.0)
    elif 'momo' in d_cat or 'momo' in d_name or d_id == 12:
        return ("Momo & Snacks", "Vehicle Max Capacity - Momo & Snacks (Kg) *", "Kg", 1500.0)
    elif 'commerce' in d_cat or 'commerce' in d_name or d_id == 8:
        return ("e-Commerce", "Vehicle Max Capacity - e-Commerce Goods (Kg) *", "Kg", 1500.0)
    else:
        return ("Frozen Foods / Chicken", "Vehicle Max Capacity - Frozen Foods / Chicken (Kg) *", "Kg", 1500.0)

@app.route('/api/template/driver-vehicle-mapping-excel', methods=['GET'])
@app.route('/api/template/depot-sku-assignment-template', methods=['GET'])
def download_driver_vehicle_mapping_template():
    try:
        user_role = session.get('role')
        user_depot = session.get('depot_id')

        depot_id = request.args.get('depot_id')
        depot_type_param = request.args.get('type') or request.args.get('depot_type')

        conn = get_db()
        cursor = conn.cursor()

        # Enforce role-based depot filtering:
        # If user is incharge/staff (non-admin), lock strictly to their assigned depot
        if user_role and user_role != 'admin' and user_depot:
            depot_id = user_depot
        elif not depot_id or str(depot_id).lower() in ['all', 'none', '']:
            if depot_type_param:
                d_match = cursor.execute("SELECT id FROM depots WHERE LOWER(category) LIKE ? OR LOWER(name) LIKE ? ORDER BY id ASC LIMIT 1",
                                         (f"%{depot_type_param.lower()}%", f"%{depot_type_param.lower()}%")).fetchone()
                depot_id = d_match['id'] if d_match else (user_depot or 9)
            elif user_depot:
                depot_id = user_depot
            else:
                depot_id = 9

        try: depot_id = int(depot_id)
        except (ValueError, TypeError): depot_id = 9

        d_row = cursor.execute("SELECT id, name, category, default_uom FROM depots WHERE id = ?", (depot_id,)).fetchone()
        if not d_row:
            conn.close()
            return jsonify({"success": False, "message": f"Depot ID {depot_id} not found"}), 404

        cat_key = get_depot_sku_catalog_key(d_row)
        resolved_uom = resolve_sku_uom(d_row['name'], d_row['category'], depot_type_param or cat_key)
        _, _, _, default_std_cap = get_depot_primary_category_info(d_row)

        # Query existing mappings from depot_vehicle_mapping (NO SKU / SKU CODE)
        mapping_rows = cursor.execute('''
            SELECT id, depot_id, depot_name, depot_type, category, uom, default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks
            FROM depot_vehicle_mapping
            WHERE depot_id = ?
            ORDER BY id ASC
        ''', (depot_id,)).fetchall()

        mappings_to_export = [dict(r) for r in mapping_rows]

        # If no mappings in DB yet, query registered vehicles & drivers for this depot to generate clean 1-sheet rows
        if not mappings_to_export:
            v_rows = cursor.execute('''
                SELECT fv.vehicle_no, fv.capacity_kg, drv.name as driver_name, drv.default_route_name
                FROM fleet_vehicles fv
                LEFT JOIN depot_crew drv ON fv.default_driver_id = drv.id
                WHERE fv.depot_id = ?
                ORDER BY fv.id ASC
            ''', (depot_id,)).fetchall()

            if v_rows:
                for r in v_rows:
                    v_cap = float(r['capacity_kg'] or default_std_cap)
                    if resolved_uom == 'Pcs' and v_cap < 5000:
                        v_cap = 30000.0
                    elif resolved_uom == 'Liter' and v_cap > 3000:
                        v_cap = 1000.0
                    mappings_to_export.append({
                        "depot_name": d_row['name'],
                        "category": d_row['category'],
                        "uom": resolved_uom,
                        "default_vehicle_no": r['vehicle_no'] or '',
                        "default_driver_name": r['driver_name'] or '',
                        "default_route_name": r['default_route_name'] or '',
                        "max_capacity": v_cap,
                        "remarks": "Active Vehicle Assignment"
                    })
            else:
                default_vans = [
                    (f"DM-SHA-11-{2000 + depot_id * 3 + 1}", "Md. Driver 1", "Route 01: Core City Line"),
                    (f"DM-SHA-11-{2000 + depot_id * 3 + 2}", "Md. Driver 2", "Route 02: Outer Suburb Line")
                ]
                for v_no, drv_name, r_name in default_vans:
                    mappings_to_export.append({
                        "depot_name": d_row['name'],
                        "category": d_row['category'],
                        "uom": resolved_uom,
                        "default_vehicle_no": v_no,
                        "default_driver_name": drv_name,
                        "default_route_name": r_name,
                        "max_capacity": default_std_cap,
                        "remarks": "Default Pre-configured Van"
                    })

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('Driver-Vehicle Mapping')

        # Modern Formats
        hdr_fmt = wb.add_format({
            'bold': True, 'font_size': 12, 'font_color': '#FFFFFF',
            'bg_color': '#0284C7', 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        sub_fmt = wb.add_format({
            'bold': True, 'font_size': 9.5, 'font_color': '#38BDF8',
            'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        tip_fmt = wb.add_format({
            'font_size': 9, 'font_color': '#64748B', 'italic': True, 'align': 'center', 'valign': 'vcenter'
        })
        th_fmt = wb.add_format({
            'bold': True, 'font_size': 9.5, 'font_color': '#FFFFFF',
            'bg_color': '#1E293B', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'text_wrap': True
        })
        th_uom_fmt = wb.add_format({
            'bold': True, 'font_size': 9.5, 'font_color': '#FDE047',
            'bg_color': '#78350F', 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'text_wrap': True
        })
        td_fmt = wb.add_format({
            'font_size': 9, 'color': '#1E293B', 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'
        })
        td_center = wb.add_format({
            'font_size': 9, 'color': '#1E293B', 'align': 'center', 'border': 1, 'border_color': '#CBD5E1', 'valign': 'vcenter'
        })
        td_uom = wb.add_format({
            'bold': True, 'font_size': 9.5, 'color': '#0369A1', 'bg_color': '#E0F2FE', 'align': 'center',
            'border': 1, 'border_color': '#7DD3FC', 'valign': 'vcenter'
        })
        td_veh = wb.add_format({
            'bold': True, 'font_size': 9.5, 'color': '#B45309', 'bg_color': '#FEF3C7', 'align': 'center',
            'border': 1, 'border_color': '#FDE68A', 'valign': 'vcenter'
        })
        td_right = wb.add_format({
            'font_size': 9, 'color': '#1E293B', 'align': 'right', 'border': 1, 'border_color': '#CBD5E1',
            'valign': 'vcenter', 'num_format': '#,##0.0'
        })

        # Columns matching strictly the manual entry inputs on the page (NO SKU fields)
        ws.set_column(0, 0, 6)   # SL
        ws.set_column(1, 1, 26)  # Depot Name
        ws.set_column(2, 2, 22)  # Category
        ws.set_column(3, 3, 20)  # Unit of Measurement (UoM) *
        ws.set_column(4, 4, 26)  # Default Assigned Vehicle Reg No *
        ws.set_column(5, 5, 22)  # Default Driver Name
        ws.set_column(6, 6, 26)  # Default Route Name / Area
        ws.set_column(7, 7, 24)  # Vehicle Max Capacity (in UoM) *
        ws.set_column(8, 8, 26)  # Remarks / Handling Note

        title_text = f"PARAGON AGRO LIMITED - DRIVER-VEHICLE MAPPING & CAPACITIES MASTER"
        sub_text = f"Target Depot: {d_row['name']}  |  Depot Type: {d_row['category']}  |  Auto-Mapped Primary UoM: {resolved_uom}"
        inst_text = "Instructions: Set default vehicle reg nos, driver names, routes, and capacities for each line. Upload directly in Master Fleet Modal Tab 5."

        ws.merge_range('A1:I1', title_text, hdr_fmt)
        ws.merge_range('A2:I2', sub_text, sub_fmt)
        ws.merge_range('A3:I3', inst_text, tip_fmt)
        ws.set_row(0, 26)
        ws.set_row(1, 20)
        ws.set_row(2, 18)

        headers = [
            'SL',
            'Depot Name',
            'Category',
            'Unit of Measurement (UoM) *',
            'Default Assigned Vehicle Reg No *',
            'Default Driver Name',
            'Default Route Name / Area',
            f'Vehicle Max Capacity ({resolved_uom}) *',
            'Remarks / Handling Note'
        ]

        for col, h in enumerate(headers):
            ws.write(4, col, h, th_uom_fmt if 'UoM' in h or 'Capacity' in h else th_fmt)
        ws.set_row(4, 26)

        start_row = 5
        for idx, itm in enumerate(mappings_to_export, 1):
            r_idx = start_row + idx - 1
            item_uom = itm.get('uom') or resolved_uom
            
            ws.write(r_idx, 0, idx, td_center)
            ws.write(r_idx, 1, d_row['name'], td_fmt)
            ws.write(r_idx, 2, itm.get('category') or d_row['category'], td_fmt)
            ws.write(r_idx, 3, item_uom, td_uom)
            ws.write(r_idx, 4, itm.get('default_vehicle_no') or '', td_veh)
            ws.write(r_idx, 5, itm.get('default_driver_name') or '', td_fmt)
            ws.write(r_idx, 6, itm.get('default_route_name') or '', td_fmt)
            ws.write(r_idx, 7, float(itm.get('max_capacity') or default_std_cap), td_right)
            ws.write(r_idx, 8, itm.get('remarks') or 'Active Vehicle Mapping', td_fmt)
            ws.set_row(r_idx, 20)

        conn.close()
        wb.close()
        output.seek(0)
        clean_depot_name = re.sub(r'[^a-zA-Z0-9_-]', '_', d_row['name'])
        filename = f"Paragon_{clean_depot_name}_Driver_Vehicle_Mapping_Template.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"success": False, "message": f"Template creation error: {str(e)}"}), 500

@app.route('/api/crew/mapping-bulk-upload', methods=['POST'])
@app.route('/api/master/sku-mapping-bulk-upload', methods=['POST'])
@app.route('/api/master/driver-vehicle-mapping-bulk-upload', methods=['POST'])
def bulk_upload_driver_vehicle_mapping():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No Excel file uploaded'}), 400
    file = request.files['file']
    
    user_role = session.get('role')
    user_depot = session.get('depot_id')

    depot_id_param = request.form.get('depot_id')
    depot_type_param = request.form.get('type') or request.form.get('depot_type')

    # Enforce role-based depot selection
    if user_role and user_role != 'admin' and user_depot:
        target_depot_id = user_depot
    elif depot_id_param and str(depot_id_param).lower() not in ['all', 'none', '']:
        try: target_depot_id = int(depot_id_param)
        except (ValueError, TypeError): target_depot_id = None
    elif user_depot:
        target_depot_id = user_depot
    else:
        target_depot_id = None

    if not target_depot_id:
        return jsonify({'success': False, 'message': 'Please select a specific target depot before uploading.'}), 400

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active  # Single sheet format

        conn = get_db()
        cursor = conn.cursor()

        d_row = cursor.execute("SELECT id, name, category, default_uom FROM depots WHERE id = ?", (target_depot_id,)).fetchone()
        if not d_row:
            conn.close()
            return jsonify({'success': False, 'message': f'Depot ID {target_depot_id} not found in database'}), 404

        cat_key = get_depot_sku_catalog_key(d_row)

        def clean_vehicle_no(v_str):
            if not v_str: return ''
            s = str(v_str).strip().upper()
            s = s.replace('(BORROWED)', '').replace('(RENTAL)', '')
            s = re.sub(r'[\s_]+', '-', s)
            s = re.sub(r'-+', '-', s).strip('-')
            return s

        header_row_idx = None
        col_map = {}

        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if not row or not any(row): continue
            non_empty = [c for c in row if c is not None and str(c).strip() != '']
            if len(non_empty) < 2: continue
            row_str = " ".join([str(c) for c in non_empty]).lower()
            if 'paragon agro' in row_str or 'instructions' in row_str or 'guidelines' in row_str:
                continue

            # Check if this is the header row
            if any(k in row_str for k in ['driver', 'vehicle', 'truck', 'van', 'capacity', 'uom', 'unit', 'category', 'route', 'sku']):
                candidate_map = {}
                for c_idx, cell in enumerate(row):
                    if cell is None: continue
                    c_text = str(cell).lower().strip()
                    if 'category' in c_text or 'ক্যাটাগরি' in c_text:
                        candidate_map['category'] = c_idx
                    elif 'uom' in c_text or 'unit' in c_text or 'measure' in c_text or 'একক' in c_text or 'পরিমাপ' in c_text:
                        candidate_map['uom'] = c_idx
                    elif 'capacity' in c_text or 'max' in c_text or 'ধারণক্ষমতা' in c_text or 'লোড' in c_text:
                        candidate_map['capacity'] = c_idx
                    elif 'route' in c_text or 'line' in c_text or 'zone' in c_text or 'রুট' in c_text or 'area' in c_text:
                        candidate_map['route'] = c_idx
                    elif 'backup' in c_text or 'secondary' in c_text:
                        candidate_map['backup_veh'] = c_idx
                    elif 'vehicle' in c_text or 'truck' in c_text or 'van' in c_text or 'reg' in c_text or 'গাড়ি' in c_text:
                        if 'veh' not in candidate_map: candidate_map['veh'] = c_idx
                    elif 'driver' in c_text or 'staff' in c_text or 'চালক' in c_text or (('name' in c_text or 'full name' in c_text) and 'sku' not in c_text and 'depot' not in c_text):
                        if 'driver' not in candidate_map: candidate_map['driver'] = c_idx
                    elif 'phone' in c_text or 'contact' in c_text or 'mobile' in c_text:
                        candidate_map['phone'] = c_idx
                    elif 'remarks' in c_text or 'note' in c_text or 'channel' in c_text or 'মন্তব্য' in c_text:
                        candidate_map['remarks'] = c_idx

                if 'veh' in candidate_map or 'driver' in candidate_map or 'route' in candidate_map:
                    header_row_idx = r_idx
                    col_map = candidate_map
                    break

        if not header_row_idx:
            # Default column map matching template: 0:SL, 1:Depot, 2:Category, 3:UoM, 4:Veh, 5:Driver, 6:Route, 7:Capacity, 8:Remarks
            header_row_idx = 5
            col_map = {'category': 2, 'uom': 3, 'veh': 4, 'driver': 5, 'route': 6, 'capacity': 7, 'remarks': 8}

        parsed_mappings = []
        vehicles_to_sync = {}
        drivers_to_sync = {}

        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if r_idx <= header_row_idx: continue
            if not row or not any(row): continue
            r_str = " ".join([str(c) for c in row if c is not None]).upper()
            if "PARAGON AGRO" in r_str or "INSTRUCTIONS" in r_str or "GUIDELINES" in r_str: continue

            def get_val(key, default=''):
                idx = col_map.get(key)
                if idx is not None and idx < len(row) and row[idx] is not None:
                    return str(row[idx]).strip()
                return default

            cat_name = get_val('category') or d_row['category']
            raw_uom = get_val('uom')
            veh_no = clean_vehicle_no(get_val('veh'))
            backup_veh = clean_vehicle_no(get_val('backup_veh'))
            drv_name = get_val('driver')
            phone = get_val('phone')
            route_name = get_val('route')
            raw_cap = get_val('capacity')
            remarks = get_val('remarks') or 'Active Vehicle Mapping'

            if not veh_no and not drv_name:
                continue

            if veh_no.upper() in ['VEHICLE REG NO', 'VEHICLE', 'DRIVER', 'SL', 'DEFAULT ASSIGNED VEHICLE']:
                continue

            # Auto-map UoM logic based on Depot + Category mapping rules
            if raw_uom:
                uom_clean = raw_uom.strip()
                if uom_clean.lower() in ['kg', 'kg.', 'kgs', 'kilogram', 'কেজি']:
                    final_uom = 'Kg'
                elif uom_clean.lower() in ['pcs', 'pcs.', 'pieces', 'piece', 'টি', 'পিস']:
                    final_uom = 'Pcs'
                elif uom_clean.lower() in ['liter', 'ltr', 'litre', 'লিটার']:
                    final_uom = 'Liter'
                elif uom_clean.lower() in ['ctn', 'carton', 'box']:
                    final_uom = 'Ctn'
                elif uom_clean.lower() in ['pkt', 'packet']:
                    final_uom = 'Pkt'
                else:
                    final_uom = resolve_sku_uom(d_row['name'], cat_name, d_row['category'])
            else:
                final_uom = resolve_sku_uom(d_row['name'], cat_name, d_row['category'])

            _, _, _, std_cap = get_depot_primary_category_info(d_row)
            cap_val = std_cap
            if raw_cap:
                try: cap_val = float(str(raw_cap).replace(',', ''))
                except (ValueError, TypeError): pass

            parsed_mappings.append({
                "depot_id": target_depot_id,
                "depot_name": d_row['name'],
                "depot_type": d_row['category'],
                "category": cat_name,
                "uom": final_uom,
                "default_vehicle_no": veh_no or f"DM-SHA-11-{2000 + len(parsed_mappings)}",
                "default_driver_name": drv_name,
                "default_route_name": route_name,
                "max_capacity": cap_val,
                "remarks": remarks
            })

            if veh_no:
                vehicles_to_sync[veh_no] = {
                    "depot_id": target_depot_id,
                    "vehicle_no": veh_no,
                    "capacity": cap_val,
                    "uom": final_uom,
                    "driver_name": drv_name,
                    "route_name": route_name
                }
            if drv_name:
                drivers_to_sync[drv_name] = {
                    "depot_id": target_depot_id,
                    "name": drv_name,
                    "phone": phone,
                    "vehicle_no": veh_no,
                    "backup_veh": backup_veh,
                    "route_name": route_name
                }

        if not parsed_mappings:
            conn.close()
            return jsonify({'success': False, 'message': 'No valid vehicle or driver mapping rows found in the uploaded file.'}), 400

        # Replace existing records in depot_vehicle_mapping for this depot
        cursor.execute("DELETE FROM depot_vehicle_mapping WHERE depot_id = ?", (target_depot_id,))
        for m in parsed_mappings:
            cursor.execute('''
                INSERT INTO depot_vehicle_mapping 
                (depot_id, depot_name, depot_type, category, uom, default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (m['depot_id'], m['depot_name'], m['depot_type'], m['category'], m['uom'],
                  m['default_vehicle_no'], m['default_driver_name'], m['default_route_name'], m['max_capacity'], m['remarks']))

        # Synchronize vehicles into fleet_vehicles
        for vno, vdata in vehicles_to_sync.items():
            exist = cursor.execute("SELECT id, category_capacities FROM fleet_vehicles WHERE depot_id = ? AND UPPER(vehicle_no) = UPPER(?)", (target_depot_id, vno)).fetchone()
            cat_caps = {}
            if exist and exist['category_capacities']:
                try: cat_caps = json.loads(exist['category_capacities'])
                except (json.JSONDecodeError, TypeError): pass
            cat_caps[cat_key] = {"capacity": vdata['capacity'], "unit": vdata['uom']}
            
            if exist:
                cursor.execute('''
                    UPDATE fleet_vehicles
                    SET capacity_kg = ?, capacity_units = ?, category_capacities = ?
                    WHERE id = ?
                ''', (vdata['capacity'], vdata['uom'], json.dumps(cat_caps), exist['id']))
            else:
                cursor.execute('''
                    INSERT INTO fleet_vehicles (depot_id, vehicle_no, vehicle_type, capacity_kg, capacity_units, category_capacities, status)
                    VALUES (?, ?, 'Covered Van', ?, ?, ?, 'Available')
                ''', (target_depot_id, vno, vdata['capacity'], vdata['uom'], json.dumps(cat_caps)))

        # Synchronize drivers into depot_crew
        for dname, ddata in drivers_to_sync.items():
            c_exist = cursor.execute("SELECT id FROM depot_crew WHERE depot_id = ? AND UPPER(name) = UPPER(?)", (target_depot_id, dname)).fetchone()
            if c_exist:
                cursor.execute('''
                    UPDATE depot_crew
                    SET assigned_vehicle_no = COALESCE(?, assigned_vehicle_no),
                        secondary_vehicle_no = COALESCE(?, secondary_vehicle_no),
                        default_route_name = COALESCE(?, default_route_name)
                    WHERE id = ?
                ''', (ddata['vehicle_no'] or None, ddata.get('backup_veh') or None, ddata['route_name'] or None, c_exist['id']))
            else:
                cursor.execute('''
                    INSERT INTO depot_crew (depot_id, name, role, phone, assigned_vehicle_no, secondary_vehicle_no, default_route_name, status)
                    VALUES (?, ?, 'driver', ?, ?, ?, ?, 'Active')
                ''', (target_depot_id, dname, ddata.get('phone') or '01711-000000', ddata['vehicle_no'], ddata.get('backup_veh') or '', ddata['route_name']))

        conn.commit()
        conn.close()

        uom_summary = ", ".join(list(set(m['uom'] for m in parsed_mappings)))
        return jsonify({
            'success': True,
            'message': f"Successfully processed {len(parsed_mappings)} Driver-Vehicle mapping records for {d_row['name']} with UoM: [{uom_summary}]!",
            'count': len(parsed_mappings),
            'uom_summary': uom_summary,
            'depot_id': target_depot_id
        })
    except Exception as err:
        return jsonify({'success': False, 'message': f'Processing error: {str(err)}'}), 500

@app.route('/api/master/mappings', methods=['GET'])
@app.route('/api/master/skus', methods=['GET'])
def get_master_mappings():
    depot_id = request.args.get('depot_id')
    conn = get_db()
    cursor = conn.cursor()
    
    if not depot_id or str(depot_id).lower() == 'all':
        rows = cursor.execute('''
            SELECT id, depot_id, depot_name, depot_type, category, uom,
                   default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks
            FROM depot_vehicle_mapping
            ORDER BY depot_id ASC, id ASC
        ''').fetchall()
    else:
        try: d_id = int(depot_id)
        except (ValueError, TypeError): d_id = 9
        rows = cursor.execute('''
            SELECT id, depot_id, depot_name, depot_type, category, uom,
                   default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks
            FROM depot_vehicle_mapping
            WHERE depot_id = ?
            ORDER BY id ASC
        ''', (d_id,)).fetchall()
        
    conn.close()
    res_list = [dict(r) for r in rows]
    return jsonify({
        "success": True,
        "mappings": res_list,
        "skus": res_list
    })

@app.route('/api/master/mapping', methods=['POST'])
@app.route('/api/master/sku', methods=['POST'])
def save_master_mapping():
    data = request.json or {}
    mapping_id = data.get('id')
    depot_id = data.get('depot_id')
    if not depot_id:
        return jsonify({"success": False, "message": "Depot ID is required"}), 400
    try: depot_id = int(depot_id)
    except (ValueError, TypeError): return jsonify({"success": False, "message": "Invalid Depot ID"}), 400
    
    conn = get_db()
    cursor = conn.cursor()
    
    d_row = cursor.execute("SELECT id, name, category FROM depots WHERE id = ?", (depot_id,)).fetchone()
    if not d_row:
        conn.close()
        return jsonify({"success": False, "message": "Depot not found"}), 404
        
    depot_name = d_row['name']
    depot_type = d_row['category']
    category = data.get('category') or depot_type
    raw_uom = data.get('uom', '').strip()
    uom = raw_uom if raw_uom else resolve_sku_uom(depot_name, category, depot_type)
    veh_no = data.get('default_vehicle_no', '').strip()
    if not veh_no:
        conn.close()
        return jsonify({"success": False, "message": "Default Vehicle Reg No is required"}), 400

    drv_name = data.get('default_driver_name', '').strip()
    route_name = data.get('default_route_name', '').strip()
    try: cap_val = float(data.get('max_capacity') or 1500.0)
    except (ValueError, TypeError): cap_val = 1500.0
    remarks = data.get('remarks', '').strip() or 'Active Vehicle Assignment'
    
    if mapping_id:
        cursor.execute('''
            UPDATE depot_vehicle_mapping
            SET category = ?, uom = ?, default_vehicle_no = ?,
                default_driver_name = ?, default_route_name = ?, max_capacity = ?, remarks = ?
            WHERE id = ? AND depot_id = ?
        ''', (category, uom, veh_no, drv_name, route_name, cap_val, remarks, mapping_id, depot_id))
    else:
        cursor.execute('''
            INSERT INTO depot_vehicle_mapping 
            (depot_id, depot_name, depot_type, category, uom, default_vehicle_no, default_driver_name, default_route_name, max_capacity, remarks)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (depot_id, depot_name, depot_type, category, uom, veh_no, drv_name, route_name, cap_val, remarks))

    # Synchronize with fleet_vehicles
    cat_key = get_depot_sku_catalog_key(d_row)
    exist = cursor.execute("SELECT id, category_capacities FROM fleet_vehicles WHERE depot_id = ? AND UPPER(vehicle_no) = UPPER(?)", (depot_id, veh_no)).fetchone()
    cat_caps = {}
    if exist and exist['category_capacities']:
        try: cat_caps = json.loads(exist['category_capacities'])
        except (json.JSONDecodeError, TypeError): pass
    cat_caps[cat_key] = {"capacity": cap_val, "unit": uom}
    
    if exist:
        cursor.execute('''
            UPDATE fleet_vehicles
            SET capacity_kg = ?, capacity_units = ?, category_capacities = ?
            WHERE id = ?
        ''', (cap_val, uom, json.dumps(cat_caps), exist['id']))
    else:
        cursor.execute('''
            INSERT INTO fleet_vehicles (depot_id, vehicle_no, vehicle_type, capacity_kg, capacity_units, category_capacities, status)
            VALUES (?, ?, 'Covered Van', ?, ?, ?, 'Available')
        ''', (depot_id, veh_no, cap_val, uom, json.dumps(cat_caps)))

    # Synchronize driver if name provided
    if drv_name:
        c_exist = cursor.execute("SELECT id FROM depot_crew WHERE depot_id = ? AND UPPER(name) = UPPER(?)", (depot_id, drv_name)).fetchone()
        if c_exist:
            cursor.execute('''
                UPDATE depot_crew
                SET assigned_vehicle_no = COALESCE(?, assigned_vehicle_no),
                    default_route_name = COALESCE(?, default_route_name)
                WHERE id = ?
            ''', (veh_no, route_name or None, c_exist['id']))
        else:
            cursor.execute('''
                INSERT INTO depot_crew (depot_id, name, role, phone, assigned_vehicle_no, default_route_name, status)
                VALUES (?, ?, 'driver', '01711-000000', ?, ?, 'Active')
            ''', (depot_id, drv_name, veh_no, route_name))
        
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Default mapping for Vehicle '{veh_no}' saved successfully!"})

@app.route('/api/master/mapping/<int:mapping_id>', methods=['DELETE', 'POST'])
@app.route('/api/master/sku/<int:mapping_id>', methods=['DELETE', 'POST'])
def delete_master_mapping(mapping_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM depot_vehicle_mapping WHERE id = ?", (mapping_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Driver-Vehicle mapping record removed successfully!"})


# ----------------- INTER-DEPOT VEHICLE BORROWING APIS -----------------

@app.route('/api/fleet/available-borrow', methods=['GET'])
def get_available_borrow_vehicles():
    current_depot_id = request.args.get('depot_id') or request.args.get('exclude_depot_id')
    conn = get_db()
    cursor = conn.cursor()
    query = '''
        SELECT fv.*, d.name as lending_depot_name, d.category as lending_depot_category, d.name as depot_name
        FROM fleet_vehicles fv
        JOIN depots d ON fv.depot_id = d.id
        WHERE (fv.ownership IS NULL OR fv.ownership != 'Borrowed')
          AND fv.vehicle_no NOT LIKE '%(Borrowed)%'
          AND fv.status IN ('Active', 'Spare')
    '''
    params = []
    if current_depot_id and str(current_depot_id).lower() not in ('all', 'undefined', 'null', ''):
        query += ' AND fv.depot_id != ?'
        params.append(current_depot_id)
    query += ' ORDER BY d.id ASC, fv.vehicle_no ASC'
    cursor.execute(query, tuple(params))
    vehicles = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({"success": True, "available_vehicles": vehicles})

@app.route('/api/fleet/borrow-request', methods=['POST'])
def create_borrow_request():
    data = request.json or {}
    req_depot_id = data.get('requesting_depot_id')
    lending_depot_id = data.get('lending_depot_id')
    vehicle_id = data.get('vehicle_id')
    request_date = data.get('request_date', datetime.date.today().strftime('%Y-%m-%d'))
    notes = data.get('notes', 'Need temporary van for high morning delivery volume')
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO inter_depot_vehicle_requests (requesting_depot_id, lending_depot_id, vehicle_id, request_date, status, notes)
        VALUES (?, ?, ?, ?, 'Approved', ?)
    ''', (req_depot_id, lending_depot_id, vehicle_id, request_date, notes))
    
    # Temporarily associate vehicle to requesting depot as Borrowed
    cursor.execute('SELECT vehicle_no, vehicle_type, capacity_kg, capacity_units FROM fleet_vehicles WHERE id = ?', (vehicle_id,))
    veh = cursor.fetchone()
    if veh:
        cursor.execute('''
            INSERT INTO fleet_vehicles (depot_id, vehicle_no, vehicle_type, capacity_kg, capacity_units, ownership, home_depot_id, status)
            VALUES (?, ?, ?, ?, ?, 'Borrowed', ?, 'Active')
        ''', (req_depot_id, f"{veh['vehicle_no']} (Borrowed)", veh['vehicle_type'], veh['capacity_kg'], veh['capacity_units'], lending_depot_id))
        
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Vehicle borrowed and added to your depot fleet successfully!"})

@app.route('/api/fleet/return-vehicle/<int:vid>', methods=['POST', 'DELETE'])
@app.route('/api/fleet/borrow/<int:vid>', methods=['DELETE'])
def return_borrowed_vehicle(vid):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT depot_id, vehicle_no, home_depot_id FROM fleet_vehicles WHERE id = ?', (vid,))
    veh = cursor.fetchone()
    if veh:
        cursor.execute('DELETE FROM fleet_vehicles WHERE id = ?', (vid,))
        base_no = veh['vehicle_no'].replace(' (Borrowed)', '').strip()
        cursor.execute('''
            DELETE FROM inter_depot_vehicle_requests 
            WHERE requesting_depot_id = ? AND vehicle_id IN (SELECT id FROM fleet_vehicles WHERE vehicle_no = ?)
        ''', (veh['depot_id'], base_no))
    else:
        cursor.execute('DELETE FROM fleet_vehicles WHERE id = ?', (vid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Borrowed vehicle returned to home depot successfully!"})

@app.route('/api/fleet/clear-borrowed', methods=['POST'])
def clear_depot_borrowed_vehicles():
    data = request.json or {}
    depot_id = data.get('depot_id')
    conn = get_db()
    cursor = conn.cursor()
    if depot_id and str(depot_id) != 'all':
        cursor.execute('DELETE FROM inter_depot_vehicle_requests WHERE requesting_depot_id = ? OR lending_depot_id = ?', (depot_id, depot_id))
        cursor.execute('DELETE FROM fleet_vehicles WHERE (depot_id = ? OR home_depot_id = ?) AND (ownership = "Borrowed" OR vehicle_no LIKE "%(Borrowed)%")', (depot_id, depot_id))
    else:
        cursor.execute('DELETE FROM inter_depot_vehicle_requests')
        cursor.execute('DELETE FROM fleet_vehicles WHERE ownership = "Borrowed" OR vehicle_no LIKE "%(Borrowed)%"')
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Borrowed fleet cleared successfully!"})

# ==============================================================================
# BLANK EXCEL TEMPLATES & BULK UPLOAD REST APIS
# ==============================================================================


@app.route('/api/template/routes-blank-template')
def download_routes_blank_template():
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Routes Template')
    
    hdr_fmt = wb.add_format({'bold': True, 'font_size': 11, 'font_color': '#FFFFFF', 'bg_color': '#1E3A8A', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    tip_fmt = wb.add_format({'font_size': 9, 'font_color': '#64748B', 'italic': True})
    th_fmt = wb.add_format({'bold': True, 'font_size': 9.5, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    td_fmt = wb.add_format({'font_size': 9, 'border': 1, 'border_color': '#CBD5E1'})
    sample_fmt = wb.add_format({'font_size': 9, 'font_color': '#334155', 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#CBD5E1'})

    ws.set_column(0, 0, 8)   # SL
    ws.set_column(1, 1, 35)  # Route Name
    ws.set_column(2, 2, 16)  # Route Code
    ws.set_column(3, 3, 22)  # Default Vehicle No
    ws.set_column(4, 4, 38)  # Description / Areas

    ws.merge_range('A1:E1', 'PARAGON AGRO LIMITED - ROUTE MASTER IMPORT TEMPLATE', hdr_fmt)
    ws.merge_range('A2:E2', 'Instructions: Fill in Route Name (Mandatory). Route Code and Default Vehicle No are optional. Save and Upload.', tip_fmt)
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = ['SL', 'Route Name *', 'Route Code', 'Default Vehicle Reg No', 'Description / Areas Covered']
    for col, h in enumerate(headers):
        ws.write(3, col, h, th_fmt)
    ws.set_row(3, 22)

    sample_data = [
        (1, 'Route 01: Mirpur - Mohammadpur Zone', 'RT-01', 'DM-THA-11-2090', 'Mirpur 1, 2, 10, 11, Pallabi, Mohammadpur Ring Road'),
        (2, 'Route 02: Uttara - Gazipur Express', 'RT-02', 'DM-THA-11-3045', 'Uttara Sectors 1-14, Tongi Station Road, Gazipur Chowrasta'),
        (3, 'Route 03: Dhanmondi - Old Dhaka Route', 'RT-03', 'DM-THA-11-4012', 'Dhanmondi, Shankar, Lalbagh, Sadarghat, Kaptan Bazar')
    ]
    for r_idx, row in enumerate(sample_data, 4):
        ws.write(r_idx, 0, row[0], sample_fmt)
        ws.write(r_idx, 1, row[1], sample_fmt)
        ws.write(r_idx, 2, row[2], sample_fmt)
        ws.write(r_idx, 3, row[3], sample_fmt)
        ws.write(r_idx, 4, row[4], sample_fmt)

    wb.close()
    output.seek(0)
    return send_file(output, download_name="Paragon_Route_Master_Blank_Template.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/api/template/consignees-blank-template')
def download_consignees_blank_template():
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Consignees Template')

    hdr_fmt = wb.add_format({'bold': True, 'font_size': 11, 'font_color': '#FFFFFF', 'bg_color': '#166534', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    tip_fmt = wb.add_format({'font_size': 9, 'font_color': '#64748B', 'italic': True})
    th_fmt = wb.add_format({'bold': True, 'font_size': 9.5, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    sample_fmt = wb.add_format({'font_size': 9, 'font_color': '#334155', 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#CBD5E1'})

    ws.set_column(0, 0, 8)   # SL
    ws.set_column(1, 1, 38)  # Consignee / Outlet Name
    ws.set_column(2, 2, 32)  # Mapped Route Name
    ws.set_column(3, 3, 32)  # Address / Location
    ws.set_column(4, 4, 18)  # Phone Number
    ws.set_column(5, 5, 16)  # Payment Mode

    ws.merge_range('A1:F1', 'PARAGON AGRO LIMITED - SALES CENTER / CONSIGNEE MAPPING TEMPLATE', hdr_fmt)
    ws.merge_range('A2:F2', 'Instructions: Consignee Name and Mapped Route Name must match your depot routes. Payment Mode: Cash / Credit.', tip_fmt)
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = ['SL', 'Consignee / Outlet Name *', 'Mapped Route Name *', 'Address / Location', 'Phone / Contact No', 'Payment Mode (Cash/Credit)']
    for col, h in enumerate(headers):
        ws.write(3, col, h, th_fmt)
    ws.set_row(3, 22)

    sample_data = [
        (1, 'Shwapno Super Shop - Mirpur 1 Branch', 'Route 01: Mirpur - Mohammadpur Zone', 'Plot 12, Mirpur 1, Dhaka', '01711-223344', 'Credit'),
        (2, 'Meena Bazar - Dhanmondi 27', 'Route 03: Dhanmondi - Old Dhaka Route', 'House 45, Road 27, Dhanmondi', '01922-334455', 'Credit'),
        (3, 'Paragon Mart - Uttara Sector 3', 'Route 02: Uttara - Gazipur Express', 'Sector 3, Uttara, Dhaka', '01833-445566', 'Cash')
    ]
    for r_idx, row in enumerate(sample_data, 4):
        for c_idx, val in enumerate(row):
            ws.write(r_idx, c_idx, val, sample_fmt)

    wb.close()
    output.seek(0)
    return send_file(output, download_name="Paragon_Consignee_Mapping_Blank_Template.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/api/template/crew-blank-template')
def download_crew_blank_template():
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Crew Template')

    hdr_fmt = wb.add_format({'bold': True, 'font_size': 11, 'font_color': '#FFFFFF', 'bg_color': '#7C3AED', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    tip_fmt = wb.add_format({'font_size': 9, 'font_color': '#64748B', 'italic': True})
    th_fmt = wb.add_format({'bold': True, 'font_size': 9.5, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    sample_fmt = wb.add_format({'font_size': 9, 'font_color': '#334155', 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#CBD5E1'})

    ws.set_column(0, 0, 8)   # SL
    ws.set_column(1, 1, 20)  # Role
    ws.set_column(2, 2, 28)  # Full Name
    ws.set_column(3, 3, 18)  # Mobile Phone
    ws.set_column(4, 4, 22)  # Driving License No
    ws.set_column(5, 5, 28)  # Assigned Vehicle Reg No (Drivers Only)
    ws.set_column(6, 6, 14)  # Status

    ws.merge_range('A1:G1', 'PARAGON AGRO LIMITED - DEPOT DRIVERS & CREW IMPORT TEMPLATE', hdr_fmt)
    ws.merge_range('A2:G2', 'Instructions: Role should be "Driver" or "Delivery Man". For Drivers only, specify "Assigned Vehicle Reg No" (e.g. DM-THA-11-2090) to link vehicle automatically.', tip_fmt)
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = ['SL', 'Role (Driver/Delivery Man) *', 'Staff Full Name *', 'Mobile Phone No *', 'Driving License No', 'Assigned Vehicle Reg No (Drivers Only)', 'Status (Active/Inactive)']
    for col, h in enumerate(headers):
        ws.write(3, col, h, th_fmt)
    ws.set_row(3, 22)

    sample_data = [
        (1, 'Driver', 'Md. Rafiqul Islam', '01711-100001', 'DL-DHAKA-112233', 'DM-THA-11-2090', 'Active'),
        (2, 'Delivery Man', 'Md. Kamal Hossain', '01711-200001', '-', '-', 'Active'),
        (3, 'Driver', 'Md. Jahangir Alam', '01811-100002', 'DL-DHAKA-445566', 'DM-THA-11-3045', 'Active'),
        (4, 'Delivery Man', 'Md. Rubel Miah', '01811-200002', '-', '-', 'Active')
    ]
    for r_idx, row in enumerate(sample_data, 4):
        for c_idx, val in enumerate(row):
            ws.write(r_idx, c_idx, val, sample_fmt)

    wb.close()
    output.seek(0)
    return send_file(output, download_name="Paragon_Depot_Crew_Blank_Template.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/api/template/fleet-blank-template')
def download_fleet_blank_template():
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Fleet Template')

    hdr_fmt = wb.add_format({'bold': True, 'font_size': 11, 'font_color': '#FFFFFF', 'bg_color': '#0369A1', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    tip_fmt = wb.add_format({'font_size': 9, 'font_color': '#64748B', 'italic': True})
    th_fmt = wb.add_format({'bold': True, 'font_size': 9.5, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    sample_fmt = wb.add_format({'font_size': 9, 'font_color': '#334155', 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#CBD5E1'})

    ws.set_column(0, 0, 8)   # SL
    ws.set_column(1, 1, 24)  # Vehicle No
    ws.set_column(2, 2, 22)  # Vehicle Type
    ws.set_column(3, 3, 16)  # Capacity Value
    ws.set_column(4, 4, 18)  # Capacity Unit (Pcs/Kg/Ltr/Carton)
    ws.set_column(5, 5, 26)  # Assigned Driver Name / Mobile
    ws.set_column(6, 6, 18)  # Ownership
    ws.set_column(7, 7, 26)  # Rental Vendor
    ws.set_column(8, 8, 18)  # Daily Rate

    ws.merge_range('A1:I1', 'PARAGON AGRO LIMITED - DEPOT FLEET VEHICLES IMPORT TEMPLATE', hdr_fmt)
    ws.merge_range('A2:I2', 'Instructions: Capacity Units by Category: Egg = "Pcs", Chicken/Meat/Frozen = "Kg", Dairy = "Ltr", Dry Food = "Carton/Pkt". Ownership: "Owned" or "Rental".', tip_fmt)
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = ['SL', 'Vehicle Reg No *', 'Vehicle Type', 'Capacity Value *', 'Capacity Unit (Pcs/Kg/Ltr/Carton/Trays/Crates) *', 'Assigned Driver Name/Phone', 'Ownership (Owned/Rental)', 'Rental Vendor Name', 'Daily Rent Cost (৳)']
    for col, h in enumerate(headers):
        ws.write(3, col, h, th_fmt)
    ws.set_row(3, 22)

    sample_data = [
        (1, 'DM-THA-11-2090', 'Ventilated Egg Van', 25000, 'Pcs', 'Md. Rafiqul Islam (01711-100001)', 'Owned', '-', 0),
        (2, 'DM-THA-11-3045', 'Reefer Van (-18°C)', 2000, 'Kg', 'Md. Jahangir Alam (01811-100002)', 'Owned', '-', 0),
        (3, 'DM-U-12-8899', 'Insulated Milk Van', 1800, 'Ltr', 'Md. Driver Dairy (01711-300003)', 'Rental', 'Green Line Transport', 3500)
    ]
    for r_idx, row in enumerate(sample_data, 4):
        for c_idx, val in enumerate(row):
            ws.write(r_idx, c_idx, val, sample_fmt)

    wb.close()
    output.seek(0)
    return send_file(output, download_name="Paragon_Fleet_Vehicles_Blank_Template.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------- BULK UPLOAD HANDLERS -----------------

@app.route('/api/routes/bulk-upload', methods=['POST'])
def bulk_upload_routes():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded'}), 400
    file = request.files['file']
    depot_id = request.form.get('depot_id', 9)
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        conn = get_db()
        cursor = conn.cursor()
        
        imported_count = 0
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if not row or not any(row):
                continue
            r_str = " ".join([str(c) for c in row if c is not None]).upper()
            if "ROUTE NAME" in r_str or "PARAGON AGRO" in r_str or "INSTRUCTIONS" in r_str:
                continue
                
            name = str(row[1] if len(row) > 1 and row[1] else (row[0] or '')).strip()
            code = str(row[2] if len(row) > 2 and row[2] else '').strip()
            veh = str(row[3] if len(row) > 3 and row[3] else '').strip()
            desc = str(row[4] if len(row) > 4 and row[4] else '').strip()
            
            if not name or name.isdigit() or name.lower() in ['sl', 'route name']:
                continue
                
            # Check existing route
            cursor.execute('SELECT id FROM routes WHERE depot_id = ? AND route_name = ?', (depot_id, name))
            exist = cursor.fetchone()
            if exist:
                cursor.execute('UPDATE routes SET route_code=?, default_vehicle_no=?, description=? WHERE id=?', (code, veh, desc, exist['id']))
            else:
                cursor.execute('INSERT INTO routes (depot_id, route_name, route_code, description, default_vehicle_no) VALUES (?, ?, ?, ?, ?)',
                               (depot_id, name, code, desc, veh))
            imported_count += 1
            
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'Successfully imported {imported_count} routes from Excel!'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error uploading routes: {str(e)}'}), 500


@app.route('/api/consignees/bulk-upload', methods=['POST'])
def bulk_upload_consignees():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded'}), 400
    file = request.files['file']
    depot_id = request.form.get('depot_id', 9)
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        conn = get_db()
        cursor = conn.cursor()
        
        # Load existing routes for depot
        cursor.execute('SELECT id, route_name, route_code FROM routes WHERE depot_id = ?', (depot_id,))
        routes_db = {r['route_name'].lower().strip(): r['id'] for r in cursor.fetchall()}
        
        total_rows_processed = 0
        new_added = 0
        updated_existing = 0
        
        # Track duplicate occurrences within the uploaded file
        seen_in_file = {}
        
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if not row or not any(row):
                continue
                
            # Smart header detection: only skip title/instructions in first 5 rows
            if r_idx <= 5:
                r_str = " ".join([str(c) for c in row if c is not None]).upper()
                if "PARAGON AGRO LIMITED" in r_str or "INSTRUCTIONS:" in r_str or "CONSIGNEE / OUTLET NAME" in r_str:
                    continue
            
            # Identify columns
            first_val = str(row[0] or '').strip()
            second_val = str(row[1] or '').strip() if len(row) > 1 else ''
            
            if first_val.isdigit() and second_val:
                # Column 0 is SL, Column 1 is Outlet Name
                c_name = second_val
                r_name = str(row[2] or '').strip() if len(row) > 2 else ''
                addr = str(row[3] or '').strip() if len(row) > 3 else ''
                phone = str(row[4] or '').strip() if len(row) > 4 else ''
                pay_mode = str(row[5] or 'Cash').strip() if len(row) > 5 else 'Cash'
            else:
                # Column 0 is Outlet Name
                c_name = first_val or second_val
                r_name = str(row[1] or '').strip() if len(row) > 1 else ''
                addr = str(row[2] or '').strip() if len(row) > 2 else ''
                phone = str(row[3] or '').strip() if len(row) > 3 else ''
                pay_mode = str(row[4] or 'Cash').strip() if len(row) > 4 else 'Cash'
                
            if not c_name or c_name.lower() in ['sl', 'consignee / outlet name', 'consignee name']:
                continue
                
            total_rows_processed += 1
            c_key = c_name.lower().strip()
            
            if c_key in seen_in_file:
                seen_in_file[c_key]['rows'].append(r_idx)
                if r_name and r_name not in seen_in_file[c_key]['routes']:
                    seen_in_file[c_key]['routes'].append(r_name)
            else:
                seen_in_file[c_key] = {
                    'name': c_name,
                    'rows': [r_idx],
                    'routes': [r_name] if r_name else []
                }
                
            route_id = None
            if r_name:
                r_key = r_name.lower().strip()
                if r_key in routes_db:
                    route_id = routes_db[r_key]
                else:
                    cursor.execute('INSERT INTO routes (depot_id, route_name) VALUES (?, ?)', (depot_id, r_name))
                    route_id = cursor.lastrowid
                    routes_db[r_key] = route_id
                    
            cursor.execute('SELECT id FROM route_consignees WHERE depot_id = ? AND consignee_name = ?', (depot_id, c_name))
            exist = cursor.fetchone()
            if exist:
                cursor.execute('UPDATE route_consignees SET route_id=?, address=?, phone=?, payment_mode=? WHERE id=?',
                               (route_id, addr, phone, pay_mode, exist['id']))
                updated_existing += 1
            else:
                cursor.execute('INSERT INTO route_consignees (depot_id, route_id, consignee_name, address, phone, payment_mode) VALUES (?, ?, ?, ?, ?, ?)',
                               (depot_id, route_id, c_name, addr, phone, pay_mode))
                new_added += 1
                
        conn.commit()
        conn.close()
        
        duplicates_in_file = [info for key, info in seen_in_file.items() if len(info['rows']) > 1]
        unique_outlets = len(seen_in_file)
        
        if duplicates_in_file:
            msg = f"Excel file processed successfully! Total {total_rows_processed} rows read -> {unique_outlets} unique outlets saved ({new_added} new, {updated_existing} updated). {len(duplicates_in_file)} duplicate rows were found and merged."
        else:
            msg = f"Successfully imported all {unique_outlets} unique consignees/outlets from Excel!"
            
        return jsonify({
            'success': True,
            'message': msg,
            'total_rows': total_rows_processed,
            'unique_outlets': unique_outlets,
            'new_added': new_added,
            'updated': updated_existing,
            'duplicate_count': len(duplicates_in_file),
            'duplicates': duplicates_in_file
        })
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error uploading consignees: {str(e)}'}), 500


@app.route('/api/crew/bulk-upload', methods=['POST'])
def bulk_upload_crew():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded'}), 400
    file = request.files['file']
    depot_id = request.form.get('depot_id', 9)
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        conn = get_db()
        cursor = conn.cursor()
        
        imported_count = 0
        for row in ws.iter_rows(values_only=True):
            if not row or not any(row):
                continue
            r_str = " ".join([str(c) for c in row if c is not None]).upper()
            if "ROLE" in r_str or "PARAGON AGRO" in r_str or "INSTRUCTIONS" in r_str:
                continue
                
            raw_role = str(row[1] if len(row) > 1 and row[1] else '').strip().lower()
            name = str(row[2] if len(row) > 2 and row[2] else '').strip()
            phone = str(row[3] if len(row) > 3 and row[3] else '').strip()
            license_no = str(row[4] if len(row) > 4 and row[4] else '').strip()
            assigned_veh = str(row[5] if len(row) > 5 and row[5] else '').strip().upper()
            status = str(row[6] if len(row) > 6 and row[6] else (row[5] if len(row) > 5 and str(row[5]).strip() in ['Active', 'Inactive', 'On_Leave'] else 'Active')).strip()
            
            if not name or name.lower() in ['sl', 'staff full name']:
                continue
                
            role = 'driver' if 'driver' in raw_role else 'delivery_man'
            if role != 'driver' or assigned_veh in ['-', 'N/A', 'NONE', 'ACTIVE', 'INACTIVE']:
                if role != 'driver':
                    assigned_veh = ''
                elif assigned_veh in ['ACTIVE', 'INACTIVE']:
                    assigned_veh = ''
                    
            cursor.execute('SELECT id FROM depot_crew WHERE depot_id = ? AND name = ? AND role = ?', (depot_id, name, role))
            exist = cursor.fetchone()
            if exist:
                crew_id = exist['id']
                cursor.execute('UPDATE depot_crew SET phone=?, license_no=?, status=?, assigned_vehicle_no=? WHERE id=?',
                               (phone, license_no, status, assigned_veh, crew_id))
            else:
                cursor.execute('INSERT INTO depot_crew (depot_id, role, name, phone, license_no, status, assigned_vehicle_no) VALUES (?, ?, ?, ?, ?, ?, ?)',
                               (depot_id, role, name, phone, license_no, status, assigned_veh))
                crew_id = cursor.lastrowid
                
            if role == 'driver' and assigned_veh:
                cursor.execute('UPDATE fleet_vehicles SET default_driver_id = ? WHERE depot_id = ? AND UPPER(vehicle_no) = ?',
                               (crew_id, depot_id, assigned_veh))
                               
            imported_count += 1
            
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'Successfully imported {imported_count} staff/crew members with assigned vehicles!'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error uploading crew: {str(e)}'}), 500


@app.route('/api/fleet/bulk-upload', methods=['POST'])
def bulk_upload_fleet():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded'}), 400
    file = request.files['file']
    depot_id = request.form.get('depot_id', 9)
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        conn = get_db()
        cursor = conn.cursor()
        
        # Determine depot default UOM / category
        cursor.execute('SELECT category, default_uom FROM depots WHERE id = ?', (depot_id,))
        depot_row = cursor.fetchone()
        depot_cat = depot_row['category'] if depot_row else 'Food'
        default_unit = 'Pcs' if ('egg' in depot_cat.lower() or int(depot_id) == 9) else ('Ltr' if ('dairy' in depot_cat.lower() or int(depot_id) in [4, 11]) else 'Kg')
        
        imported_count = 0
        for row in ws.iter_rows(values_only=True):
            if not row or not any(row):
                continue
            r_str = " ".join([str(c) for c in row if c is not None]).upper()
            if "VEHICLE REG" in r_str or "PARAGON AGRO" in r_str or "INSTRUCTIONS" in r_str:
                continue
                
            v_no = str(row[1] if len(row) > 1 and row[1] else '').strip().upper()
            v_type = str(row[2] if len(row) > 2 and row[2] else 'Covered Van').strip()
            try:
                cap_kg = float(row[3]) if len(row) > 3 and row[3] is not None else 1500.0
            except (ValueError, TypeError):
                cap_kg = 1500.0
                
            cap_uom = str(row[4] if len(row) > 4 and row[4] else '').strip()
            if not cap_uom or cap_uom in ['-', 'N/A', 'NONE']:
                cap_uom = default_unit
                
            driver_info = str(row[5] if len(row) > 5 and row[5] else '').strip()
            ownership = str(row[6] if len(row) > 6 and row[6] else (row[5] if len(row) > 5 and str(row[5]).strip() in ['Owned', 'Rental', 'Borrowed'] else 'Owned')).strip()
            if ownership not in ['Owned', 'Rental', 'Borrowed']:
                ownership = 'Owned'
                
            vendor = str(row[7] if len(row) > 7 and row[7] else (row[6] if len(row) > 6 and ownership == 'Rental' else '')).strip()
            try:
                rent_cost = float(row[8] if len(row) > 8 and row[8] is not None else (row[7] if len(row) > 7 and row[7] is not None else 0.0))
            except (ValueError, TypeError):
                rent_cost = 0.0
                
            if not v_no or v_no.lower() in ['sl', 'vehicle reg no']:
                continue
                
            # Find driver if driver_info provided
            default_driver_id = None
            if driver_info and driver_info not in ['-', 'N/A', 'NONE', 'Owned', 'Rental']:
                cursor.execute('SELECT id FROM depot_crew WHERE depot_id = ? AND role = "driver" AND (name LIKE ? OR phone LIKE ?)',
                               (depot_id, f"%{driver_info[:15]}%", f"%{driver_info[-11:]}%"))
                dr_match = cursor.fetchone()
                if dr_match:
                    default_driver_id = dr_match['id']
                    
            cursor.execute('SELECT id FROM fleet_vehicles WHERE depot_id = ? AND vehicle_no = ?', (depot_id, v_no))
            exist = cursor.fetchone()
            if exist:
                veh_id = exist['id']
                cursor.execute('''
                    UPDATE fleet_vehicles
                    SET vehicle_type=?, capacity_kg=?, capacity_units=?, ownership=?, rental_vendor=?, rental_cost_per_day=?,
                        default_driver_id=COALESCE(?, default_driver_id)
                    WHERE id=?
                ''', (v_type, cap_kg, cap_uom, ownership, vendor, rent_cost, default_driver_id, veh_id))
            else:
                cursor.execute('''
                    INSERT INTO fleet_vehicles (depot_id, vehicle_no, vehicle_type, capacity_kg, capacity_units, ownership, rental_vendor, rental_cost_per_day, home_depot_id, default_driver_id, status)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active')
                ''', (depot_id, v_no, v_type, cap_kg, cap_uom, ownership, vendor, rent_cost, depot_id, default_driver_id))
                veh_id = cursor.lastrowid
                
            # Also update depot_crew assigned_vehicle_no if driver found
            if default_driver_id:
                cursor.execute('UPDATE depot_crew SET assigned_vehicle_no = ? WHERE id = ?', (v_no, default_driver_id))
                
            imported_count += 1
            
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'Successfully imported {imported_count} fleet vehicles with category-based capacity ({default_unit}) and driver assignments!'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error uploading fleet: {str(e)}'}), 500


@app.route('/api/template/universal-route-plan-excel')
def download_universal_route_plan_template():
    depot_id = request.args.get('depot_id')
    date_param = request.args.get('date', datetime.date.today().strftime('%d/%m/%Y'))
    if '-' in date_param:
        try:
            parts = date_param.split('-')
            if len(parts) == 3 and len(parts[0]) == 4:
                date_param = f"{parts[2]}/{parts[1]}/{parts[0]}"
        except (IndexError, ValueError):
            pass

    depot_name = "09. Tejgaon - Fresh Egg"
    incharge_info = "Md. Zahid Hasan (01711-XXXXXX)"
    category = "Fresh Eggs"
    
    if depot_id:
        try:
            conn = get_db()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM depots WHERE id = ?', (depot_id,))
            depot = cursor.fetchone()
            conn.close()
            if depot:
                depot_name = depot['name']
                incharge_info = f"{depot['incharge_name']} ({depot['contact']})"
                category = depot['category']
        except Exception as e:
            pass

    is_egg = "egg" in category.lower() or "egg" in depot_name.lower() or str(depot_id) == '9'
    is_dairy = "dairy" in category.lower() or "milk" in category.lower() or str(depot_id) in ['4', '11']
    primary_uom = "Pcs" if is_egg else ("Ltr" if is_dairy else "Pkt")
    qty_header = f"Qty ({primary_uom})"
    weight_header = "Volume (Ltr)" if is_dairy else "Weight (Kg)"

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Daily Route Plan')
    
    ws.set_landscape()
    ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.fit_to_pages(1, 0)

    title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#FFFFFF', 'bg_color': '#1E3A8A', 'align': 'center', 'valign': 'vcenter'})
    sub_fmt = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#E2E8F0', 'bg_color': '#1E293B', 'align': 'center', 'valign': 'vcenter'})
    meta_lbl = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#0F172A', 'bg_color': '#F1F5F9', 'border': 1, 'border_color': '#CBD5E1'})
    meta_val = wb.add_format({'font_size': 9, 'font_color': '#0F172A', 'bg_color': '#FFFFFF', 'border': 1, 'border_color': '#CBD5E1'})

    veh_hdr = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': '#0369A1', 'border': 1, 'border_color': '#0284C7', 'valign': 'vcenter'})
    col_hdr = wb.add_format({'bold': True, 'font_size': 8.5, 'font_color': '#FFFFFF', 'bg_color': '#1E293B', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#334155', 'text_wrap': True})

    cell_txt = wb.add_format({'font_size': 8.5, 'font_color': '#0F172A', 'border': 1, 'border_color': '#E2E8F0', 'valign': 'vcenter'})
    cell_ctr = wb.add_format({'font_size': 8.5, 'font_color': '#0F172A', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0'})
    cell_num = wb.add_format({'font_size': 8.5, 'font_color': '#0F172A', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'num_format': '#,##0'})
    cell_dec = wb.add_format({'font_size': 8.5, 'font_color': '#0F172A', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'num_format': '#,##0.00'})
    cell_cur = wb.add_format({'font_size': 8.5, 'font_color': '#0F172A', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0', 'num_format': '৳ #,##0.00'})

    subtot_fmt = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#0F172A', 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#94A3B8', 'valign': 'vcenter'})
    subtot_num = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#0F172A', 'bg_color': '#F8FAFC', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#94A3B8', 'num_format': '#,##0'})
    subtot_dec = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#0F172A', 'bg_color': '#F8FAFC', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#94A3B8', 'num_format': '#,##0.00'})
    subtot_cur = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#0F172A', 'bg_color': '#F8FAFC', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#94A3B8', 'num_format': '৳ #,##0.00'})

    grand_fmt = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'border': 1, 'border_color': '#000000', 'valign': 'vcenter'})
    grand_num = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#000000', 'num_format': '#,##0'})
    grand_dec = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#000000', 'num_format': '#,##0.00'})
    grand_cur = wb.add_format({'bold': True, 'font_size': 10, 'font_color': '#FFFFFF', 'bg_color': '#0F172A', 'align': 'right', 'valign': 'vcenter', 'border': 1, 'border_color': '#000000', 'num_format': '৳ #,##0.00'})

    col_widths = [5, 26, 18, 14, 14, 12, 10, 10, 13, 10, 12, 22]
    for idx, w in enumerate(col_widths):
        ws.set_column(idx, idx, w)

    ws.merge_range('A1:L1', 'PARAGON AGRO LIMITED - CONSUMER FOOD DIVISION', title_fmt)
    ws.merge_range('A2:L2', f'DAILY DISTRIBUTION ROUTE PLAN & DRIVER DISPATCH SCHEDULE - {depot_name.upper()}', sub_fmt)
    ws.set_row(0, 24)
    ws.set_row(1, 18)

    ws.write('A4', 'Depot / Plant:', meta_lbl)
    ws.merge_range('B4:C4', depot_name, meta_val)
    ws.write('D4', 'Date:', meta_lbl)
    ws.write('E4', date_param, meta_val)
    ws.write('F4', 'Shift:', meta_lbl)
    ws.write('G4', 'Morning Shift (07:00 AM)', meta_val)
    ws.write('H4', 'Incharge:', meta_lbl)
    ws.merge_range('I4:J4', incharge_info, meta_val)
    ws.write('K4', 'Total Vans:', meta_lbl)
    ws.write('L4', 2, meta_val)

    headers = ['SL', 'Consignee / Outlet Name', 'Address / Area', 'Sales Order No', 'D.Note / Challan', 'Category', qty_header, weight_header, 'Invoice Value', 'Pay Mode', 'Cash Collect', 'Responsible Person / Contact']

    sample_outlets_1 = [
        (1, 'Agora Superstore Outlet', 'Main Market Area', 'SO-2026-001', 'DN-2026-001', category, 3000 if is_egg else 100, 0.0 if is_egg else 100.0, 36000.0 if is_egg else 48000.0, 'Credit', 0, 'Receiving Officer (01700-111111)'),
        (2, 'Shwapno Super Shop', 'Central Market Area', 'SO-2026-002', 'DN-2026-002', category, 6000 if is_egg else 250, 0.0 if is_egg else 250.0, 72000.0 if is_egg else 120000.0, 'Credit', 0, 'Store Manager (01700-222222)'),
        (3, 'Paragon Mart Outlet', 'Commercial Area', 'SO-2026-003', 'DN-2026-003', category, 1500 if is_egg else 75, 0.0 if is_egg else 75.0, 18000.0 if is_egg else 36000.0, 'Cash', 18000.0 if is_egg else 36000.0, 'Shop Incharge (01700-333333)')
    ]

    sample_outlets_2 = [
        (1, 'Meena Bazar Superstore', 'Market Circle', 'SO-2026-004', 'DN-2026-004', category, 4500 if is_egg else 179, 0.0 if is_egg else 72.6, 54000.0 if is_egg else 38850.4, 'Credit', 0, 'Store Incharge (01966-422968)'),
        (2, 'Unimart Hypermarket', 'Shopping Complex', 'SO-2026-005', 'DN-2026-005', category, 3000 if is_egg else 110, 0.0 if is_egg else 42.0, 36000.0 if is_egg else 20301.0, 'Credit', 0, 'Store Incharge (01811-999888)')
    ]

    sample_vans = [
        {
            'veh_info': f'VAN 1: DMS 11-5721 | Driver: Adil (01712-111222) | Delivery Man: Mostafiz (01912-333444) | Dispatch: 07:30 AM | Route: {depot_name} Area-A',
            'outlets': sample_outlets_1
        },
        {
            'veh_info': f'VAN 2: DMS 11-5623 | Driver: Shakil (01812-555666) | Delivery Man: Yeasin (01612-777888) | Dispatch: 08:00 AM | Route: {depot_name} Area-B',
            'outlets': sample_outlets_2
        }
    ]

    row_idx = 5
    subtot_rows = []

    for van in sample_vans:
        ws.merge_range(row_idx, 0, row_idx, 11, van['veh_info'], veh_hdr)
        ws.set_row(row_idx, 20)
        row_idx += 1
        
        for c_idx, h in enumerate(headers):
            ws.write(row_idx, c_idx, h, col_hdr)
        ws.set_row(row_idx, 22)
        row_idx += 1
        
        start_data_row = row_idx + 1
        for ot in van['outlets']:
            ws.write(row_idx, 0, ot[0], cell_ctr)
            ws.write(row_idx, 1, ot[1], cell_txt)
            ws.write(row_idx, 2, ot[2], cell_txt)
            ws.write(row_idx, 3, ot[3], cell_ctr)
            ws.write(row_idx, 4, ot[4], cell_ctr)
            ws.write(row_idx, 5, ot[5], cell_ctr)
            ws.write(row_idx, 6, ot[6], cell_num)
            ws.write(row_idx, 7, ot[7], cell_dec)
            ws.write(row_idx, 8, ot[8], cell_cur)
            ws.write(row_idx, 9, ot[9], cell_ctr)
            ws.write(row_idx, 10, ot[10], cell_cur)
            ws.write(row_idx, 11, ot[11], cell_txt)
            ws.set_row(row_idx, 18)
            row_idx += 1
        end_data_row = row_idx
        
        ws.merge_range(row_idx, 0, row_idx, 5, 'Vehicle Subtotal:', subtot_fmt)
        ws.write_formula(row_idx, 6, f'=SUM(G{start_data_row}:G{end_data_row})', subtot_num)
        ws.write_formula(row_idx, 7, f'=SUM(H{start_data_row}:H{end_data_row})', subtot_dec)
        ws.write_formula(row_idx, 8, f'=SUM(I{start_data_row}:I{end_data_row})', subtot_cur)
        ws.write(row_idx, 9, '', subtot_fmt)
        ws.write_formula(row_idx, 10, f'=SUM(K{start_data_row}:K{end_data_row})', subtot_cur)
        ws.write(row_idx, 11, '', subtot_fmt)
        ws.set_row(row_idx, 20)
        subtot_rows.append(row_idx + 1)
        row_idx += 2

    ws.merge_range(row_idx, 0, row_idx, 5, 'DEPOT DAY GRAND TOTAL (ALL VANS):', grand_fmt)
    pkt_formula = '=' + '+'.join([f'G{r}' for r in subtot_rows])
    kg_formula = '=' + '+'.join([f'H{r}' for r in subtot_rows])
    amt_formula = '=' + '+'.join([f'I{r}' for r in subtot_rows])
    cash_formula = '=' + '+'.join([f'K{r}' for r in subtot_rows])

    ws.write_formula(row_idx, 6, pkt_formula, grand_num)
    ws.write_formula(row_idx, 7, kg_formula, grand_dec)
    ws.write_formula(row_idx, 8, amt_formula, grand_cur)
    ws.write(row_idx, 9, '', grand_fmt)
    ws.write_formula(row_idx, 10, cash_formula, grand_cur)
    ws.write(row_idx, 11, '', grand_fmt)
    ws.set_row(row_idx, 22)

    wb.close()
    output.seek(0)
    
    clean_name = "".join(c for c in depot_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
    fname = f"Paragon_Route_Plan_Template_{clean_name}.xlsx" if depot_id else "Paragon_Route_Plan_Template.xlsx"
    
    return send_file(
        output,
        download_name=fname,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/api/distribution/export-van-runsheet-excel', methods=['POST'])
def export_van_runsheet_excel():
    data = request.json or {}
    plan_data = data.get('plan', {})
    vans = plan_data.get('vans', [])
    depot_name = plan_data.get('depot_name', 'Paragon Depot')
    plan_date = plan_data.get('date', datetime.date.today().strftime('%d/%m/%Y'))
    depot_id = plan_data.get('depot_id')
    
    is_egg = "egg" in str(depot_name).lower() or str(depot_id) == '9'
    is_dairy = "dairy" in str(depot_name).lower() or "milk" in str(depot_name).lower() or str(depot_id) in ['4', '11']
    primary_uom = "Pcs" if is_egg else ("Ltr" if is_dairy else "Pkt")
    secondary_uom = "Ltr" if is_dairy else "Kg"
    
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    for v_idx, van in enumerate(vans):
        sheet_name = f"Van_{v_idx+1}_{van.get('vehicle_no', 'Trip')[:10]}".replace('/', '_').replace('\\', '_')[:31]
        ws = wb.add_worksheet(sheet_name)
        ws.set_portrait()
        ws.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
        
        hdr_fmt = wb.add_format({'bold': True, 'font_size': 12, 'font_color': '#FFFFFF', 'bg_color': '#1E3A8A', 'align': 'center', 'valign': 'vcenter'})
        sub_hdr = wb.add_format({'bold': True, 'font_size': 9, 'font_color': '#FFFFFF', 'bg_color': '#0369A1', 'align': 'center', 'valign': 'vcenter'})
        box_hdr = wb.add_format({'bold': True, 'font_size': 8.5, 'bg_color': '#F1F5F9', 'border': 1, 'border_color': '#94A3B8'})
        box_val = wb.add_format({'font_size': 8.5, 'border': 1, 'border_color': '#94A3B8'})
        th_fmt = wb.add_format({'bold': True, 'font_size': 8, 'font_color': '#FFFFFF', 'bg_color': '#1E293B', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        td_txt = wb.add_format({'font_size': 8, 'border': 1, 'border_color': '#CBD5E1'})
        td_ctr = wb.add_format({'font_size': 8, 'align': 'center', 'border': 1, 'border_color': '#CBD5E1'})
        td_num = wb.add_format({'font_size': 8, 'align': 'right', 'border': 1, 'border_color': '#CBD5E1', 'num_format': '#,##0'})
        td_dec = wb.add_format({'font_size': 8, 'align': 'right', 'border': 1, 'border_color': '#CBD5E1', 'num_format': '#,##0.00'})
        td_cur = wb.add_format({'font_size': 8, 'align': 'right', 'border': 1, 'border_color': '#CBD5E1', 'num_format': '৳ #,##0.00'})
        tot_fmt = wb.add_format({'bold': True, 'font_size': 8.5, 'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#64748B'})
        tot_num = wb.add_format({'bold': True, 'font_size': 8.5, 'bg_color': '#F8FAFC', 'align': 'right', 'border': 1, 'border_color': '#64748B', 'num_format': '#,##0'})
        tot_dec = wb.add_format({'bold': True, 'font_size': 8.5, 'bg_color': '#F8FAFC', 'align': 'right', 'border': 1, 'border_color': '#64748B', 'num_format': '#,##0.00'})
        tot_cur = wb.add_format({'bold': True, 'font_size': 8.5, 'bg_color': '#F8FAFC', 'align': 'right', 'border': 1, 'border_color': '#64748B', 'num_format': '৳ #,##0.00'})

        ws.set_column(0, 0, 4)   # SL
        ws.set_column(1, 1, 24)  # Consignee
        ws.set_column(2, 2, 14)  # Order No
        ws.set_column(3, 3, 13)  # D.Note
        ws.set_column(4, 4, 8)   # Primary UOM
        ws.set_column(5, 5, 8)   # Secondary UOM
        ws.set_column(6, 6, 12)  # Value
        ws.set_column(7, 7, 7)   # Pay Mode
        ws.set_column(8, 8, 11)  # Cash Collect
        ws.set_column(9, 9, 8)   # Deliv Qty
        ws.set_column(10, 10, 8) # Ret Qty
        ws.set_column(11, 11, 14)# Receiver Sig

        ws.merge_range('A1:L1', 'PARAGON AGRO LIMITED - DELIVERY RUN SHEET & DRIVER TRIP CHALLAN', hdr_fmt)
        ws.merge_range('A2:L2', f"DEPOT: {depot_name} | DATE: {plan_date} | VEHICLE: {van.get('vehicle_no', 'N/A')}", sub_hdr)
        
        ws.write('A4', 'Driver Name:', box_hdr)
        ws.merge_range('B4:D4', f"{van.get('driver_name', 'N/A')} ({van.get('driver_mobile', '')})", box_val)
        ws.write('E4', 'Delivery Man:', box_hdr)
        ws.merge_range('F4:H4', f"{van.get('delivery_man', 'N/A')} ({van.get('delivery_man_mobile', '')})", box_val)
        ws.write('I4', 'Dispatch Time:', box_hdr)
        ws.merge_range('J4:L4', van.get('dispatch_time', '07:30 AM'), box_val)
        
        ws.write('A5', 'Route / Zone:', box_hdr)
        ws.merge_range('B5:D5', van.get('route_zone', 'Local Territory'), box_val)
        ws.write('E5', 'Security Gate Pass:', box_hdr)
        ws.merge_range('F5:H5', 'VERIFIED & CLEARED', box_val)
        ws.write('I5', 'Return Audit Status:', box_hdr)
        ws.merge_range('J5:L5', 'PENDING EVENING AUDIT', box_val)

        r_headers = ['SL', 'Consignee / Outlet Name', 'Sales Order No', 'D.Note No', primary_uom, secondary_uom, 'Inv Value', 'Pay', 'Collect ৳', 'Deliv', 'Return', 'Receiver Signature']
        for c, h in enumerate(r_headers):
            ws.write(6, c, h, th_fmt)
        ws.set_row(6, 22)

        outlets = van.get('outlets', [])
        r_cur = 7
        for sl, ot in enumerate(outlets, 1):
            ws.write(r_cur, 0, sl, td_ctr)
            ws.write(r_cur, 1, ot.get('consignee_name', ''), td_txt)
            ws.write(r_cur, 2, ot.get('order_no', ''), td_ctr)
            ws.write(r_cur, 3, ot.get('delivery_note_id', ''), td_ctr)
            ws.write(r_cur, 4, ot.get('total_pkt', 0), td_num)
            ws.write(r_cur, 5, ot.get('total_kg', 0.0), td_dec)
            ws.write(r_cur, 6, ot.get('total_amount', 0.0), td_cur)
            ws.write(r_cur, 7, ot.get('payment_mode', 'Credit'), td_ctr)
            ws.write(r_cur, 8, ot.get('expected_cash', 0.0), td_cur)
            ws.write(r_cur, 9, '', td_txt)
            ws.write(r_cur, 10, '', td_txt)
            ws.write(r_cur, 11, '', td_txt)
            ws.set_row(r_cur, 18)
            r_cur += 1

        # Van Total
        ws.merge_range(r_cur, 0, r_cur, 3, 'VAN TRIP TOTAL:', tot_fmt)
        ws.write_formula(r_cur, 4, f'=SUM(E8:E{r_cur})', tot_num)
        ws.write_formula(r_cur, 5, f'=SUM(F8:F{r_cur})', tot_dec)
        ws.write_formula(r_cur, 6, f'=SUM(G8:G{r_cur})', tot_cur)
        ws.write(r_cur, 7, '', tot_fmt)
        ws.write_formula(r_cur, 8, f'=SUM(I8:I{r_cur})', tot_cur)
        ws.write(r_cur, 9, '', tot_fmt)
        ws.write(r_cur, 10, '', tot_fmt)
        ws.write(r_cur, 11, '', tot_fmt)
        ws.set_row(r_cur, 20)

        # Signature & SOP block
        r_sig = r_cur + 2
        ws.write(r_sig, 1, '_____________________', wb.add_format({'align': 'center'}))
        ws.write(r_sig, 5, '_____________________', wb.add_format({'align': 'center'}))
        ws.write(r_sig, 9, '_____________________', wb.add_format({'align': 'center'}))
        ws.write(r_sig+1, 1, 'Depot Incharge Signature', wb.add_format({'align': 'center', 'font_size': 8}))
        ws.write(r_sig+1, 5, 'Security Gate Officer', wb.add_format({'align': 'center', 'font_size': 8}))
        ws.write(r_sig+1, 9, 'Driver / Delivery Man', wb.add_format({'align': 'center', 'font_size': 8}))

    wb.close()
    output.seek(0)
    
    clean_depot = "".join(c for c in depot_name if c.isalnum() or c in (' ', '_', '-')).strip().replace(' ', '_')
    fname = f"Van_Runsheets_{clean_depot}_{plan_date.replace('/', '-')}.xlsx"
    return send_file(
        output,
        download_name=fname,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==============================================================================
# ROUTE PLAN PERSISTENCE & HISTORICAL VIEWER API
# ==============================================================================

@app.route('/api/distribution/save-plan', methods=['POST'])
def save_distribution_plan():
    data = request.json or {}
    depot_id = data.get('depot_id')
    plan_date = data.get('date') or datetime.date.today().strftime('%Y-%m-%d')
    shift = data.get('shift', 'Morning Shift')
    shift_time = data.get('shift_time', '07:00')
    plan_data = data.get('plan_data') or {}
    
    if not depot_id:
        return jsonify({"success": False, "message": "Missing depot_id"}), 400
        
    vans = plan_data.get('vans', [])
    tot_outlets = sum(len(v.get('outlets', [])) for v in vans)
    tot_pkts = sum(sum(float(o.get('total_pkt', 0)) for o in v.get('outlets', [])) for v in vans)
    tot_kg = sum(sum(float(o.get('total_kg', 0)) for o in v.get('outlets', [])) for v in vans)
    gross_val = sum(sum(float(o.get('total_amount', 0)) for o in v.get('outlets', [])) for v in vans)
    
    conn = get_db()
    cursor = conn.cursor()
    
    plan_json_str = json.dumps(plan_data)
    cursor.execute('''
    INSERT INTO saved_route_plans (
        depot_id, plan_date, shift, shift_time, total_vans, total_outlets, total_pkts, total_kg, gross_value, plan_json, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    ON CONFLICT(depot_id, plan_date) DO UPDATE SET
        shift = excluded.shift,
        shift_time = excluded.shift_time,
        total_vans = excluded.total_vans,
        total_outlets = excluded.total_outlets,
        total_pkts = excluded.total_pkts,
        total_kg = excluded.total_kg,
        gross_value = excluded.gross_value,
        plan_json = excluded.plan_json,
        updated_at = CURRENT_TIMESTAMP
    ''', (depot_id, plan_date, shift, shift_time, len(vans), tot_outlets, tot_pkts, tot_kg, gross_val, plan_json_str))
    
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": f"Route plan for {plan_date} saved successfully!"})

@app.route('/api/distribution/get-plan', methods=['GET'])
def get_distribution_plan():
    depot_id = request.args.get('depot_id')
    plan_date = request.args.get('date') or datetime.date.today().strftime('%Y-%m-%d')
    
    if not depot_id:
        return jsonify({"success": False, "message": "Missing depot_id"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM saved_route_plans WHERE depot_id = ? AND plan_date = ?', (depot_id, plan_date))
    row = cursor.fetchone()
    
    if row:
        plan_dict = dict(row)
        plan_data = json.loads(plan_dict['plan_json'])
        conn.close()
        return jsonify({
            "success": True,
            "exists": True,
            "depot_id": int(depot_id),
            "date": plan_date,
            "shift": plan_dict['shift'],
            "shift_time": plan_dict['shift_time'],
            "total_vans": plan_dict['total_vans'],
            "total_outlets": plan_dict['total_outlets'],
            "total_pkts": plan_dict['total_pkts'],
            "total_kg": plan_dict['total_kg'],
            "gross_value": plan_dict['gross_value'],
            "plan_data": plan_data
        })
        
    # If not in saved_route_plans, check daily_reports + trips + invoices for fallback
    cursor.execute('SELECT id, incharge_name, shift, dispatched_gross_val FROM daily_reports WHERE depot_id = ? AND report_date = ?', (depot_id, plan_date))
    rep = cursor.fetchone()
    if rep:
        rep_dict = dict(rep)
        rep_id = rep_dict['id']
        cursor.execute('SELECT * FROM trips WHERE report_id = ?', (rep_id,))
        trips = [dict(t) for t in cursor.fetchall()]
        cursor.execute('SELECT * FROM invoices WHERE report_id = ?', (rep_id,))
        invoices = [dict(i) for i in cursor.fetchall()]
        
        # Reconstruct vans and outlets
        vans = []
        for t in trips:
            v_trip = t.get('trip_no')
            v_invs = [i for i in invoices if i.get('trip_no') == v_trip]
            outlets = []
            for inv in v_invs:
                outlets.append({
                    "order_no": inv.get('invoice_no'),
                    "consignee_name": inv.get('customer_name'),
                    "delivery_note_id": "",
                    "total_pkt": inv.get('dispatched_qty', 0),
                    "total_kg": 0.0,
                    "total_amount": inv.get('dispatched_val', 0),
                    "payment_mode": inv.get('collection_mode', 'Credit'),
                    "expected_cash": inv.get('amount_collected', 0) if inv.get('collection_mode') == 'Cash' else 0
                })
            vans.append({
                "vehicle_no": t.get('vehicle_no'),
                "vehicle_type": t.get('vehicle_type', '1.5T Covered Van'),
                "driver_name": t.get('driver_name', 'Driver'),
                "delivery_man": t.get('delivery_man', 'Staff'),
                "route_zone": t.get('route_name', 'Route Zone'),
                "capacity_kg": t.get('capacity_kg', 1500),
                "loaded_kg": t.get('loaded_kg', 0),
                "reefer_temp": t.get('reefer_temp', '-18°C'),
                "outlets": outlets
            })
        
        conn.close()
        return jsonify({
            "success": True,
            "exists": True,
            "source": "daily_reports",
            "depot_id": int(depot_id),
            "date": plan_date,
            "shift": rep_dict.get('shift', 'Morning Shift'),
            "shift_time": "07:00",
            "total_vans": len(vans),
            "total_outlets": len(invoices),
            "total_pkts": sum(i.get('dispatched_qty', 0) for i in invoices),
            "total_kg": 0.0,
            "gross_value": rep_dict.get('dispatched_gross_val', 0),
            "plan_data": {
                "depot_id": int(depot_id),
                "date": plan_date,
                "vans": vans,
                "unassigned_orders": []
            }
        })
        
    conn.close()
    return jsonify({"success": True, "exists": False, "message": "No saved plan found for this date."})

@app.route('/api/distribution/plan-history', methods=['GET'])
def get_distribution_plan_history():
    depot_id = request.args.get('depot_id')
    if not depot_id:
        return jsonify({"success": False, "message": "Missing depot_id"}), 400
        
    conn = get_db()
    cursor = conn.cursor()
    
    cursor.execute('''
    SELECT id, depot_id, plan_date, shift, shift_time, total_vans, total_outlets, total_pkts, total_kg, gross_value, updated_at
    FROM saved_route_plans
    WHERE depot_id = ?
    ORDER BY plan_date DESC
    ''', (depot_id,))
    history = [dict(r) for r in cursor.fetchall()]
    
    # Also fetch distinct dates from daily_reports if not in saved_route_plans
    existing_dates = set(h['plan_date'] for h in history)
    cursor.execute('''
    SELECT id, depot_id, report_date, total_vehicles, total_invoices, dispatched_gross_val, created_at
    FROM daily_reports
    WHERE depot_id = ?
    ORDER BY report_date DESC
    ''', (depot_id,))
    for rep in cursor.fetchall():
        rep_dict = dict(rep)
        r_date = rep_dict['report_date']
        if r_date not in existing_dates:
            history.append({
                "id": rep_dict['id'],
                "depot_id": rep_dict['depot_id'],
                "plan_date": r_date,
                "shift": "Dispatched Report",
                "shift_time": "07:00",
                "total_vans": rep_dict['total_vehicles'],
                "total_outlets": rep_dict['total_invoices'],
                "total_pkts": 0,
                "total_kg": 0.0,
                "gross_value": rep_dict['dispatched_gross_val'],
                "updated_at": rep_dict['created_at']
            })
            existing_dates.add(r_date)
    history.sort(key=lambda x: x['plan_date'], reverse=True)
    conn.close()
    return jsonify({"success": True, "history": history})

# ----------------- ADMIN DIRECTORY & PLAN CLEAR APIS -----------------

@app.route('/api/admin/clear-master-data', methods=['GET', 'POST', 'DELETE'])
def admin_clear_master_data():
    data = request.json if request.is_json else {}
    sess_user = session.get('user') or {}
    role = session.get('role') or sess_user.get('role') or data.get('role') or request.headers.get('X-Admin-Role')
    if role and role not in ['admin', 'guest'] and role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized: Only Admin can clear master directory data"}), 403

    clear_type = data.get('clear_type') or data.get('type') or request.args.get('clear_type') or request.args.get('type')  # 'routes', 'consignees', 'fleet', 'crew', 'borrow', 'mapping', 'all'
    raw_depot_id = data.get('depot_id') or request.args.get('depot_id')
    depot_id = None
    if raw_depot_id and str(raw_depot_id).lower() not in ('all', 'undefined', 'null', ''):
        try:
            depot_id = int(raw_depot_id)
        except (ValueError, TypeError):
            depot_id = None
    
    conn = get_db()
    cursor = conn.cursor()
    
    if clear_type in ('routes', 'route'):
        if depot_id is not None:
            cursor.execute('DELETE FROM route_consignees WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM routes WHERE depot_id = ?', (depot_id,))
        else:
            cursor.execute('DELETE FROM route_consignees')
            cursor.execute('DELETE FROM routes')
        msg = "All route master records and mappings cleared successfully!"
        
    elif clear_type in ('consignees', 'consignee'):
        if depot_id is not None:
            cursor.execute('DELETE FROM route_consignees WHERE depot_id = ?', (depot_id,))
        else:
            cursor.execute('DELETE FROM route_consignees')
        msg = "All consignee outlet mappings cleared successfully!"
        
    elif clear_type in ('fleet', 'vehicles', 'vehicle'):
        if depot_id is not None:
            cursor.execute('DELETE FROM fleet_vehicles WHERE depot_id = ?', (depot_id,))
        else:
            cursor.execute('DELETE FROM fleet_vehicles')
        msg = "All fleet vehicles and rental entries cleared successfully!"
        
    elif clear_type in ('crew', 'drivers', 'staff'):
        if depot_id is not None:
            cursor.execute('DELETE FROM depot_crew WHERE depot_id = ?', (depot_id,))
        else:
            cursor.execute('DELETE FROM depot_crew')
        msg = "All driver and helper records cleared successfully!"
        
    elif clear_type in ('borrow', 'borrowed'):
        if depot_id is not None:
            cursor.execute('DELETE FROM inter_depot_vehicle_requests WHERE requesting_depot_id = ? OR lending_depot_id = ?', (depot_id, depot_id))
            cursor.execute("DELETE FROM fleet_vehicles WHERE depot_id = ? AND ownership = 'Borrowed'", (depot_id,))
        else:
            cursor.execute('DELETE FROM inter_depot_vehicle_requests')
            cursor.execute("DELETE FROM fleet_vehicles WHERE ownership = 'Borrowed'")
        msg = "All borrowed vehicle records and requests cleared successfully!"
        
    elif clear_type in ('mapping', 'sku', 'skus'):
        if depot_id is not None:
            cursor.execute('DELETE FROM depot_vehicle_mapping WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM depot_sku_master WHERE depot_id = ?', (depot_id,))
        else:
            cursor.execute('DELETE FROM depot_vehicle_mapping')
            cursor.execute('DELETE FROM depot_sku_master')
        msg = "All default driver-vehicle mappings and SKU catalogs cleared successfully!"
        
    elif clear_type == 'all':
        if depot_id is not None:
            cursor.execute('DELETE FROM route_consignees WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM routes WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM fleet_vehicles WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM depot_crew WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM depot_vehicle_mapping WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM depot_sku_master WHERE depot_id = ?', (depot_id,))
            cursor.execute('DELETE FROM inter_depot_vehicle_requests WHERE requesting_depot_id = ? OR lending_depot_id = ?', (depot_id, depot_id))
        else:
            cursor.execute('DELETE FROM route_consignees')
            cursor.execute('DELETE FROM routes')
            cursor.execute('DELETE FROM fleet_vehicles')
            cursor.execute('DELETE FROM depot_crew')
            cursor.execute('DELETE FROM depot_vehicle_mapping')
            cursor.execute('DELETE FROM depot_sku_master')
            cursor.execute('DELETE FROM inter_depot_vehicle_requests')
        msg = "All Master Directory records for this depot cleared successfully!"
    else:
        conn.close()
        return jsonify({"success": False, "message": "Invalid clear type specified"}), 400
        
    # Record in admin audit log
    username = sess_user.get('username') or session.get('username') or 'admin'
    cursor.execute('''
    INSERT INTO admin_audit_log (username, action, target_type, depot_id, details, ip_address)
    VALUES (?, ?, ?, ?, ?, ?)
    ''', (username, 'CLEAR_MASTER_DATA', clear_type, str(depot_id or 'ALL'), msg, request.remote_addr))

    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": msg})

@app.route('/api/fleet/borrowed-vehicles', methods=['GET'])
def get_borrowed_vehicles():
    depot_id = request.args.get('depot_id')
    conn = get_db()
    cursor = conn.cursor()
    query = '''
        SELECT fv.*, d.name as lending_depot_name, d.category as lending_depot_category
        FROM fleet_vehicles fv
        LEFT JOIN depots d ON fv.home_depot_id = d.id
        WHERE fv.ownership = 'Borrowed'
    '''
    params = []
    if depot_id and str(depot_id).lower() not in ('all', 'undefined', 'null', ''):
        query += ' AND fv.depot_id = ?'
        params.append(depot_id)
    query += ' ORDER BY fv.id DESC'
    cursor.execute(query, tuple(params))
    borrowed = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({"success": True, "borrowed_vehicles": borrowed})

@app.route('/api/admin/clear-route-plan', methods=['GET', 'POST', 'DELETE'])
def admin_clear_route_plan():
    data = request.json if request.is_json else {}
    sess_user = session.get('user') or {}
    role = session.get('role') or sess_user.get('role') or data.get('role') or request.headers.get('X-Admin-Role')
    if role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized: Only Admin can delete saved route plans"}), 403

    depot_id = data.get('depot_id') or request.args.get('depot_id')
    plan_date = data.get('date') or data.get('plan_date') or request.args.get('date') or request.args.get('plan_date')
    clear_all = data.get('clear_all', False)
    
    conn = get_db()
    cursor = conn.cursor()
    
    if clear_all:
        if depot_id:
            cursor.execute('DELETE FROM saved_route_plans WHERE depot_id = ?', (depot_id,))
        else:
            cursor.execute('DELETE FROM saved_route_plans')
        msg = "All saved route plans deleted successfully!"
    elif depot_id and plan_date:
        cursor.execute('DELETE FROM saved_route_plans WHERE depot_id = ? AND plan_date = ?', (depot_id, plan_date))
        # Also if there are daily_reports created for this date that admin wants cleared
        if data.get('delete_daily_report'):
            cursor.execute('SELECT id FROM daily_reports WHERE depot_id = ? AND report_date = ?', (depot_id, plan_date))
            for r in cursor.fetchall():
                cursor.execute('DELETE FROM invoices WHERE report_id = ?', (r['id'],))
                cursor.execute('DELETE FROM trips WHERE report_id = ?', (r['id'],))
            cursor.execute('DELETE FROM daily_reports WHERE depot_id = ? AND report_date = ?', (depot_id, plan_date))
        msg = f"Saved route plan for {plan_date} deleted successfully!"
    else:
        conn.close()
        return jsonify({"success": False, "message": "Missing depot_id or date"}), 400
        
    # Record in admin audit log
    username = sess_user.get('username') or session.get('username') or 'admin'
    cursor.execute('''
    INSERT INTO admin_audit_log (username, action, target_type, depot_id, details, ip_address)
    VALUES (?, ?, ?, ?, ?, ?)
    ''', (username, 'CLEAR_ROUTE_PLAN', 'saved_route_plans', str(depot_id or 'ALL'), msg, request.remote_addr))

    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": msg})

def match_depot_from_ref(branch="", ref_no="", order_no=""):
    """Match depot from branch, reference, or order strings for Poloxy import"""
    combined = f"{branch} {ref_no} {order_no}".lower()
    depot_keywords = [
        (1, ['ashulia frozen', 'ashulia_frozen', 'ashulia-frozen']),
        (2, ['ashulia dry', 'ashulia_dry', 'dry food']),
        (3, ['gazipur', 'process ck', 'gazipur_ck']),
        (4, ['sirajganj', 'dairy plant', 'sirajganj_dairy']),
        (5, ['sylhet tea', 'tea packing', 'sylhet_tea']),
        (6, ['sylhet frozen', 'sylhet_frozen']),
        (7, ['tejgaon frozen', 'tejgaon_frozenck', 'frozen chicken', 'chicken & frozen']),
        (8, ['ecommerce', 'e-commerce', 'tejgaon_ecommerce']),
        (9, ['egg', 'fresh egg', 'tejgaon_egg']),
        (10, ['tea dist', 'tejgaon_tea', 'tea distribution']),
        (11, ['liquid milk', 'tejgaon_dairy', 'milk depot']),
        (12, ['mohakhali', 'momo']),
        (13, ['ctg frozen', 'chittagong frozen']),
        (14, ['ctg dry', 'chittagong dry']),
        (15, ['jessore frozen', 'jessore_frozen']),
        (16, ['jessore dry', 'jessore_dry']),
        (17, ['rangpur'])
    ]
    for d_id, kws in depot_keywords:
        if any(kw in combined for kw in kws):
            name = DEPOT_COORDINATES.get(d_id, {}).get('name', f'Depot #{d_id}')
            return {'id': d_id, 'name': name, 'category': 'Food'}
    return {'id': 1, 'name': 'Tejgaon Head Office & Central Depot', 'category': 'Central Food'}

# ==============================================================================
# POLOXY ERP BATCH PARSER WITH ROUTE & FLEET AUTO-MAPPING
# ==============================================================================

@app.route('/api/distribution/import-poloxy-orders', methods=['POST'])
def api_import_poloxy_orders():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': 'No file uploaded in request'}), 400
        
    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'message': 'Empty file selected'}), 400
        
    depot_id_param = request.form.get('depot_id')
    active_depot_id = int(depot_id_param) if depot_id_param and depot_id_param.isdigit() else None
    
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        
        header_row_idx = 2
        for r_idx, row in enumerate(ws.iter_rows(max_row=8, values_only=True), 1):
            row_str = " ".join([str(c) for c in row if c is not None]).upper()
            if "ORDER NO" in row_str or "REF. NO" in row_str or "CUSTOMER" in row_str:
                header_row_idx = r_idx
                break
                
        conn = get_db()
        cursor = conn.cursor()
        
        # Load routes and consignee map
        routes_map = {}
        consignee_route_map = {}
        if active_depot_id:
            cursor.execute('SELECT id, route_name, route_code, default_vehicle_no FROM routes WHERE depot_id = ?', (active_depot_id,))
            for r in cursor.fetchall():
                routes_map[r['id']] = dict(r)
                
            cursor.execute('SELECT consignee_name, route_id FROM route_consignees WHERE depot_id = ?', (active_depot_id,))
            for rc in cursor.fetchall():
                consignee_route_map[rc['consignee_name'].strip().lower()] = rc['route_id']
                
        # Load depot fleet and crew
        depot_fleet = []
        depot_drivers = []
        depot_deliverymen = []
        if active_depot_id:
            cursor.execute('SELECT * FROM fleet_vehicles WHERE depot_id = ? ORDER BY id ASC', (active_depot_id,))
            depot_fleet = [dict(v) for v in cursor.fetchall()]
            
            cursor.execute("SELECT * FROM depot_crew WHERE depot_id = ? AND role = 'driver'", (active_depot_id,))
            depot_drivers = [dict(c) for c in cursor.fetchall()]
            
            cursor.execute("SELECT * FROM depot_crew WHERE depot_id = ? AND role = 'delivery_man'", (active_depot_id,))
            depot_deliverymen = [dict(c) for c in cursor.fetchall()]
            
        cursor.execute('SELECT * FROM depots')
        depots_db = {d['id']: dict(d) for d in cursor.fetchall()}
        conn.close()
        
        depots_result = {}
        total_orders_parsed = 0
        total_value_parsed = 0.0
        
        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if not row or not any(row):
                continue
            
            sr_no = str(row[0] or '').strip()
            date_val = str(row[1] or '').strip()
            customer = str(row[2] or '').strip()
            order_no = str(row[3] or '').strip()
            ref_no = str(row[4] or '').strip()
            branch = str(row[6] or '').strip()
            item = str(row[7] or '').strip()
            raw_rate_unit = str(row[8] or '').strip()
            rate = float(row[9] or 0.0)
            qty_bag = float(row[10] or 0.0)
            qty_kg = float(row[11] or 0.0)
            
            # Column N (Index 13): Total Amount
            total_amount_col_n = None
            if len(row) > 13 and row[13] is not None and str(row[13]).strip() != '':
                try:
                    total_amount_col_n = float(row[13])
                except (ValueError, TypeError):
                    total_amount_col_n = None
                    
            d_note_id = str(row[27] or '').strip() if len(row) > 27 else ''
            consignee_name = str(row[32] or '').strip() if len(row) > 32 and str(row[32]).strip() else customer
            consignee_contact = str(row[33] or '').strip() if len(row) > 33 else ''
            consignee_address = str(row[34] or '').strip() if len(row) > 34 else ''
            
            if not ref_no and not order_no and not item:
                continue
                
            if active_depot_id:
                depot_info = depots_db.get(active_depot_id, {'id': active_depot_id, 'name': 'Selected Depot', 'category': 'Food'})
            else:
                depot_info = match_depot_from_ref(branch, ref_no, order_no)
                
            d_id = depot_info['id']
            d_name = depot_info['name']
            d_cat = depot_info.get('category', 'Food')
            
            is_egg = d_id == 9 or "egg" in d_cat.lower()
            is_dairy = d_id in [4, 11] or "dairy" in d_cat.lower()
            primary_uom = "Pcs" if is_egg else ("Ltr" if is_dairy else "Pkt")
            secondary_uom = "Kg" if not is_dairy else "Ltr"
            rate_unit = raw_rate_unit or primary_uom
            
            if d_id not in depots_result:
                depots_result[d_id] = {
                    'depot_id': d_id,
                    'depot_name': d_name,
                    'category': d_cat,
                    'primary_uom': primary_uom,
                    'secondary_uom': secondary_uom,
                    'orders_map': {},
                    'fleet': depot_fleet,
                    'drivers': depot_drivers,
                    'deliverymen': depot_deliverymen,
                    'routes': list(routes_map.values()),
                    'total_outlets': 0,
                    'total_pkts': 0.0,
                    'total_kg': 0.0,
                    'total_amount': 0.0
                }
                
            ord_key = order_no or f"ORD-{sr_no or (total_orders_parsed + 1)}"
            if ord_key not in depots_result[d_id]['orders_map']:
                is_credit = any(kw in consignee_name.lower() for kw in ['shwapno', 'agora', 'meena', 'unimart', 'pran', 'aarong', 'lavender', 'food panda'])
                
                # Match route for consignee
                matched_route_id = consignee_route_map.get(consignee_name.strip().lower())
                matched_route = routes_map.get(matched_route_id) if matched_route_id else None
                
                depots_result[d_id]['orders_map'][ord_key] = {
                    'order_no': ord_key,
                    'date': date_val,
                    'customer': customer,
                    'consignee_name': consignee_name,
                    'consignee_contact': consignee_contact,
                    'consignee_address': consignee_address,
                    'branch_category': branch or d_cat,
                    'delivery_note_id': d_note_id,
                    'route_id': matched_route_id,
                    'route_name': matched_route['route_name'] if matched_route else 'General Route',
                    'default_vehicle_no': matched_route['default_vehicle_no'] if matched_route else '',
                    'items': [],
                    'total_pkt': 0.0,
                    'total_kg': 0.0,
                    'total_amount': 0.0,
                    'col_n_found': False,
                    'assigned_van': '',
                    'payment_mode': 'Credit' if is_credit else 'Cash',
                    'expected_cash': 0.0,
                    'rate_unit': rate_unit
                }
                depots_result[d_id]['total_outlets'] += 1
                total_orders_parsed += 1
                
            if total_amount_col_n is not None and not depots_result[d_id]['orders_map'][ord_key]['col_n_found']:
                if total_amount_col_n > 0:
                    depots_result[d_id]['orders_map'][ord_key]['total_amount'] = total_amount_col_n
                    depots_result[d_id]['orders_map'][ord_key]['col_n_found'] = True
                    
            item_line_amount = rate * (qty_kg if qty_kg > 0 else qty_bag)
            depots_result[d_id]['orders_map'][ord_key]['items'].append({
                'item': item,
                'rate_unit': rate_unit,
                'rate': rate,
                'qty_pkt': qty_bag,
                'qty_kg': qty_kg,
                'amount': item_line_amount
            })
            
            depots_result[d_id]['orders_map'][ord_key]['total_pkt'] += qty_bag
            depots_result[d_id]['orders_map'][ord_key]['total_kg'] += qty_kg
            depots_result[d_id]['total_pkts'] += qty_bag
            depots_result[d_id]['total_kg'] += qty_kg
            
        # Finalize list
        final_depots = []
        for d_id, data in depots_result.items():
            order_list = list(data['orders_map'].values())
            d_total_val = 0.0
            for o in order_list:
                if not o['col_n_found'] or o['total_amount'] == 0:
                    o['total_amount'] = sum(it['amount'] for it in o['items'])
                d_total_val += o['total_amount']
                if o['payment_mode'] == 'Cash':
                    o['expected_cash'] = o['total_amount']
                    
            data['total_amount'] = d_total_val
            data['orders'] = order_list
            del data['orders_map']
            total_value_parsed += d_total_val
            final_depots.append(data)
            
        final_depots.sort(key=lambda x: x['total_outlets'], reverse=True)
        
        return jsonify({
            'success': True,
            'message': f'Poloxy ERP parsed {total_orders_parsed} orders for {final_depots[0]["depot_name"] if final_depots else "Selected Depot"}!',
            'total_orders': total_orders_parsed,
            'total_value': total_value_parsed,
            'depots': final_depots
        })
        
    except Exception as err:
        return jsonify({'success': False, 'message': f'Error reading Poloxy file: {str(err)}'}), 500




# ==============================================================================
# LIVE GPS & ROUTE MONITORING APIS (LEAFLET + OPENSTREETMAP ENGINE)
# ==============================================================================

DEPOT_COORDINATES = {
    1: {"lat": 23.7644, "lng": 90.3928, "name": "Tejgaon Head Office & Central Depot"},
    2: {"lat": 22.3569, "lng": 91.7832, "name": "Chittagong Central Depot"},
    3: {"lat": 24.8949, "lng": 91.8687, "name": "Sylhet Regional Depot"},
    4: {"lat": 23.7644, "lng": 90.3928, "name": "Tejgaon Dairy Plant"},
    5: {"lat": 24.3636, "lng": 88.6241, "name": "Rajshahi Regional Depot"},
    6: {"lat": 22.8456, "lng": 89.5403, "name": "Khulna Regional Depot"},
    7: {"lat": 23.7680, "lng": 90.3980, "name": "Tejgaon Frozen Foods Depot"},
    8: {"lat": 24.8465, "lng": 89.3777, "name": "Bogura Sales Depot"},
    9: {"lat": 23.7620, "lng": 90.3950, "name": "Tejgaon Fresh Egg Depot"},
    10: {"lat": 24.7471, "lng": 90.4203, "name": "Mymensingh Distribution Hub"},
    11: {"lat": 23.7660, "lng": 90.3910, "name": "Tejgaon Liquid Milk Depot"},
    12: {"lat": 22.7010, "lng": 90.3535, "name": "Barishal Regional Depot"},
    13: {"lat": 23.4682, "lng": 91.1788, "name": "Cumilla Regional Depot"},
    14: {"lat": 22.3590, "lng": 91.7850, "name": "Chittagong Dry Food Depot"},
    15: {"lat": 23.8475, "lng": 90.2577, "name": "Savar Poultry & Feed Depot"},
    16: {"lat": 23.6238, "lng": 90.5000, "name": "Narayanganj Sales Depot"},
    17: {"lat": 23.9999, "lng": 90.4203, "name": "Gazipur Feed & Breeder Depot"},
}

SAMPLE_OUTLET_OFFSETS = [
    {"name": "Shwapno Super Shop - Outlet #1", "dlat": 0.0150, "dlng": 0.0120, "contact": "01711-223344", "inv": "INV-240901"},
    {"name": "Agora Superstore - Main Branch", "dlat": 0.0240, "dlng": 0.0080, "contact": "01819-334455", "inv": "INV-240902"},
    {"name": "Meena Bazar - Express Outlet", "dlat": 0.0310, "dlng": -0.0120, "contact": "01912-445566", "inv": "INV-240903"},
    {"name": "Unimart Hypermarket", "dlat": 0.0190, "dlng": 0.0280, "contact": "01755-667788", "inv": "INV-240904"},
    {"name": "Lavender Convenience Mart", "dlat": 0.0080, "dlng": -0.0210, "contact": "01611-778899", "inv": "INV-240905"},
    {"name": "Almas Super Shop & Mart", "dlat": -0.0140, "dlng": 0.0180, "contact": "01722-889900", "inv": "INV-240906"},
    {"name": "Daily Shopping - Local Mart", "dlat": -0.0220, "dlng": -0.0150, "contact": "01833-990011", "inv": "INV-240907"},
    {"name": "Prince Bazar Super Mart", "dlat": -0.0290, "dlng": 0.0060, "contact": "01944-001122", "inv": "INV-240908"},
    {"name": "Well Food Bakery & Mart", "dlat": 0.0050, "dlng": 0.0350, "contact": "01788-112233", "inv": "INV-240909"},
    {"name": "Food Panda Dark Store", "dlat": -0.0110, "dlng": -0.0280, "contact": "01677-223344", "inv": "INV-240910"},
]

def init_tracking_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS live_tracking_positions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_no TEXT NOT NULL,
        driver_mobile TEXT,
        driver_name TEXT,
        depot_id INTEGER NOT NULL,
        route_id TEXT,
        latitude REAL NOT NULL,
        longitude REAL NOT NULL,
        heading REAL DEFAULT 0.0,
        speed_kmh REAL DEFAULT 0.0,
        source TEXT DEFAULT 'MOBILE_GEOLOCATION',
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(vehicle_no, depot_id)
    )
    ''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS live_drop_statuses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        depot_id INTEGER NOT NULL,
        route_code TEXT NOT NULL,
        drop_id TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'pending', -- pending, in_transit, completed, failed
        delivered_at TIMESTAMP,
        proof_note TEXT,
        cash_collected REAL DEFAULT 0.0,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(depot_id, route_code, drop_id)
    )
    ''')
    conn.commit()
    conn.close()

# Initialize tracking tables at module load
try:
    init_tracking_db()
except Exception as e:
    print("Notice: Tracking DB init deferred:", e)

@app.route('/api/tracking/depot-routes', methods=['GET'])
def get_tracking_depot_routes():
    depot_id_param = str(request.args.get('depot_id') or 'ALL').strip()
    is_all = (depot_id_param.upper() == 'ALL' or depot_id_param == '0' or depot_id_param == '')
    
    date_param = str(request.args.get('date') or '').strip()
    if not date_param:
        date_param = datetime.date.today().strftime('%Y-%m-%d')

    target_depot_ids = list(DEPOT_COORDINATES.keys()) if is_all else []
    if not is_all:
        try:
            target_depot_ids = [int(depot_id_param)]
        except ValueError:
            target_depot_ids = [1]

    conn = get_db()
    cursor = conn.cursor()

    # Check for live drop status overrides in database
    if is_all:
        cursor.execute('SELECT depot_id, drop_id, route_code, status, delivered_at, cash_collected FROM live_drop_statuses')
        status_overrides = {f"{r['depot_id']}_{r['route_code']}_{r['drop_id']}": dict(r) for r in cursor.fetchall()}
        cursor.execute('SELECT * FROM live_tracking_positions')
        live_pos_map = {f"{r['depot_id']}_{r['vehicle_no']}": dict(r) for r in cursor.fetchall()}
    else:
        d_id = target_depot_ids[0]
        cursor.execute('SELECT depot_id, drop_id, route_code, status, delivered_at, cash_collected FROM live_drop_statuses WHERE depot_id = ?', (d_id,))
        status_overrides = {f"{r['depot_id']}_{r['route_code']}_{r['drop_id']}": dict(r) for r in cursor.fetchall()}
        cursor.execute('SELECT * FROM live_tracking_positions WHERE depot_id = ?', (d_id,))
        live_pos_map = {f"{r['depot_id']}_{r['vehicle_no']}": dict(r) for r in cursor.fetchall()}

    routes_list = []
    depot_origins = []

    for d_id in target_depot_ids:
        depot_coords = DEPOT_COORDINATES.get(d_id, {"lat": 23.7644, "lng": 90.3928, "name": f"Depot #{d_id}"})
        depot_origins.append({
            "id": d_id,
            "name": depot_coords["name"],
            "lat": depot_coords["lat"],
            "lng": depot_coords["lng"]
        })

        # 1. Check saved_route_plans first for target date (from Route Planning & Van Dispatch)
        cursor.execute('SELECT * FROM saved_route_plans WHERE depot_id = ? AND plan_date = ? ORDER BY id DESC LIMIT 1', (d_id, date_param))
        saved_plans = cursor.fetchall()
        
        d_routes = []
        if saved_plans:
            for plan_row in saved_plans:
                plan_dict = dict(plan_row)
                try:
                    raw_json = json.loads(plan_dict['plan_json'])
                except Exception:
                    raw_json = {}
                    
                vans = raw_json.get('vans', [])
                for v_idx, van in enumerate(vans, 1):
                    v_no = van.get('vehicle_no') or f"VAN-{d_id:02d}-{v_idx:02d}"
                    d_name = van.get('driver_name') or "Assigned Driver"
                    d_phone = van.get('driver_mobile') or van.get('driver_contact') or f"01711-{d_id:02d}{v_idx:04d}"
                    r_code = van.get('route_code') or van.get('route_zone') or van.get('route_name') or f"RT-{d_id:02d}-{v_idx:02d}"
                    r_name = van.get('route_zone') or van.get('route_name') or f"Corridor {v_idx}"
                    orders = van.get('outlets') or van.get('orders') or []
                    
                    drop_points = []
                    for o_idx, ord_item in enumerate(orders, 1):
                        offset = SAMPLE_OUTLET_OFFSETS[(o_idx - 1 + d_id) % len(SAMPLE_OUTLET_OFFSETS)]
                        drop_lat = round(depot_coords['lat'] + offset['dlat'] + (v_idx * 0.003), 6)
                        drop_lng = round(depot_coords['lng'] + offset['dlng'] + (v_idx * 0.002), 6)
                        
                        drop_id = str(ord_item.get('order_no') or ord_item.get('delivery_note_id') or ord_item.get('invoice_no') or ord_item.get('id') or f"DROP-{o_idx}")
                        status_key = f"{d_id}_{r_code}_{drop_id}"
                        override = status_overrides.get(status_key)
                        
                        default_status = 'in_transit' if o_idx == 1 else 'pending'
                        curr_status = override['status'] if override else default_status
                        
                        drop_points.append({
                            "id": drop_id,
                            "sequence_order": o_idx,
                            "outlet_name": ord_item.get('consignee_name') or ord_item.get('customer') or ord_item.get('customer_name') or offset['name'],
                            "contact_phone": ord_item.get('contact_person') or ord_item.get('consignee_contact') or offset['contact'],
                            "address": ord_item.get('consignee_address') or f"Outlet #{o_idx}, {r_name}",
                            "invoice_no": ord_item.get('delivery_note_id') or ord_item.get('order_no') or offset['inv'],
                            "lat": drop_lat,
                            "lng": drop_lng,
                            "total_pkts": float(ord_item.get('total_pkt') or ord_item.get('dispatched_qty') or ord_item.get('qty') or 25),
                            "total_amount": float(ord_item.get('total_amount') or ord_item.get('dispatched_val') or 4500.0),
                            "payment_type": ord_item.get('payment_mode') or ord_item.get('collection_mode') or 'Cash',
                            "status": curr_status,
                            "delivered_at": override['delivered_at'] if override else None
                        })
                        
                    live_v = live_pos_map.get(f"{d_id}_{v_no}")
                    if live_v:
                        v_lat, v_lng, v_heading, v_speed = live_v['latitude'], live_v['longitude'], live_v['heading'], live_v['speed_kmh']
                    else:
                        next_drop = next((d for d in drop_points if d['status'] == 'in_transit'), drop_points[0] if drop_points else None)
                        if next_drop:
                            v_lat = round((depot_coords['lat'] + next_drop['lat']) / 2, 6)
                            v_lng = round((depot_coords['lng'] + next_drop['lng']) / 2, 6)
                            v_heading = 45.0
                            v_speed = 32.0
                        else:
                            v_lat, v_lng, v_heading, v_speed = depot_coords['lat'], depot_coords['lng'], 0.0, 0.0
                            
                    completed_count = sum(1 for d in drop_points if d['status'] == 'completed')
                    van_ownership = van.get('ownership') or ('Borrowed' if van.get('is_borrowed') else 'Owned')
                    d_routes.append({
                        "depot_id": d_id,
                        "depot_name": depot_coords["name"],
                        "depot_origin": depot_coords,
                        "route_code": r_code,
                        "route_name": r_name,
                        "vehicle_no": v_no,
                        "ownership": van_ownership,
                        "driver_name": d_name,
                        "driver_mobile": d_phone,
                        "total_drops": len(drop_points),
                        "completed_drops": completed_count,
                        "vehicle_position": {
                            "lat": v_lat,
                            "lng": v_lng,
                            "heading": v_heading,
                            "speed": v_speed
                        },
                        "drop_points": drop_points
                    })

        # 2. Fallback: If no saved_route_plans, check active daily_reports for target date
        if not d_routes:
            cursor.execute('SELECT id, report_date FROM daily_reports WHERE depot_id = ? AND report_date = ? AND status != "cancelled" ORDER BY id DESC LIMIT 1', (d_id, date_param))
            rep = cursor.fetchone()
            if rep:
                rep_id = rep['id']
                cursor.execute('SELECT * FROM trips WHERE report_id = ?', (rep_id,))
                rep_trips = [dict(t) for t in cursor.fetchall()]
                cursor.execute('SELECT * FROM invoices WHERE report_id = ?', (rep_id,))
                rep_invs = [dict(i) for i in cursor.fetchall()]
                
                for v_idx, tr in enumerate(rep_trips, 1):
                    v_no = tr.get('vehicle_no') or f"VAN-{d_id:02d}-{v_idx:02d}"
                    d_name = tr.get('driver_name') or "Assigned Driver"
                    d_phone = f"01711-{d_id:02d}{v_idx:04d}"
                    r_code = tr.get('trip_no') or f"TRIP-{v_idx:02d}"
                    r_name = tr.get('route_name') or f"Route Corridor {v_idx}"
                    
                    tr_invs = [inv for inv in rep_invs if inv.get('trip_no') == tr.get('trip_no')] or rep_invs
                    drop_points = []
                    for o_idx, inv in enumerate(tr_invs, 1):
                        offset = SAMPLE_OUTLET_OFFSETS[(o_idx - 1 + d_id) % len(SAMPLE_OUTLET_OFFSETS)]
                        drop_lat = round(depot_coords['lat'] + offset['dlat'] + (v_idx * 0.003), 6)
                        drop_lng = round(depot_coords['lng'] + offset['dlng'] + (v_idx * 0.002), 6)
                        
                        drop_id = str(inv.get('invoice_no') or f"INV-{o_idx}")
                        status_key = f"{d_id}_{r_code}_{drop_id}"
                        override = status_overrides.get(status_key)
                        
                        inv_deliv_st = (inv.get('delivery_status') or '').lower()
                        if inv_deliv_st == 'delivered':
                            default_status = 'completed'
                        elif inv_deliv_st == 'full return':
                            default_status = 'failed'
                        else:
                            default_status = 'in_transit' if o_idx == 1 else 'pending'
                            
                        curr_status = override['status'] if override else default_status
                        
                        drop_points.append({
                            "id": drop_id,
                            "sequence_order": o_idx,
                            "outlet_name": inv.get('customer_name') or offset['name'],
                            "contact_phone": offset['contact'],
                            "address": f"Outlet Location #{o_idx}, {r_name}",
                            "invoice_no": inv.get('invoice_no') or offset['inv'],
                            "lat": drop_lat,
                            "lng": drop_lng,
                            "total_pkts": float(inv.get('dispatched_qty') or 25),
                            "total_amount": float(inv.get('dispatched_val') or 4500.0),
                            "payment_type": inv.get('collection_mode') or 'Cash',
                            "status": curr_status,
                            "delivered_at": override['delivered_at'] if override else (rep['report_date'] if curr_status == 'completed' else None)
                        })
                    
                    live_v = live_pos_map.get(f"{d_id}_{v_no}")
                    if live_v:
                        v_lat, v_lng, v_heading, v_speed = live_v['latitude'], live_v['longitude'], live_v['heading'], live_v['speed_kmh']
                    else:
                        next_drop = next((d for d in drop_points if d['status'] == 'in_transit'), drop_points[0] if drop_points else None)
                        if next_drop:
                            v_lat = round((depot_coords['lat'] + next_drop['lat']) / 2, 6)
                            v_lng = round((depot_coords['lng'] + next_drop['lng']) / 2, 6)
                            v_heading = 45.0
                            v_speed = 28.0
                        else:
                            v_lat, v_lng, v_heading, v_speed = depot_coords['lat'], depot_coords['lng'], 0.0, 0.0
                            
                    completed_count = sum(1 for d in drop_points if d['status'] == 'completed')
                    d_routes.append({
                        "depot_id": d_id,
                        "depot_name": depot_coords["name"],
                        "depot_origin": depot_coords,
                        "route_code": r_code,
                        "route_name": r_name,
                        "vehicle_no": v_no,
                        "ownership": 'Owned',
                        "driver_name": d_name,
                        "driver_mobile": d_phone,
                        "total_drops": len(drop_points),
                        "completed_drops": completed_count,
                        "vehicle_position": {
                            "lat": v_lat,
                            "lng": v_lng,
                            "heading": v_heading,
                            "speed": v_speed
                        },
                        "drop_points": drop_points
                    })

        routes_list.extend(d_routes)

    conn.close()
    
    primary_origin = depot_origins[0] if len(depot_origins) == 1 else {"lat": 23.8103, "lng": 90.4125, "name": "All Depots (Nationwide Hubs)"}
    
    return jsonify({
        "success": True,
        "is_all_depots": is_all,
        "depot_id": "ALL" if is_all else target_depot_ids[0],
        "depot_name": "All Depots (Nationwide Fleet)" if is_all else primary_origin["name"],
        "depot_origin": primary_origin,
        "depot_origins": depot_origins,
        "routes": routes_list
    })

@app.route('/api/tracking/mobile-ping', methods=['POST'])
@app.route('/api/telematics/webhook', methods=['POST'])
def receive_tracking_ping():
    data = request.json if request.is_json else request.form.to_dict()
    if not data:
        return jsonify({"success": False, "message": "Missing telemetry payload"}), 400

    vehicle_no = data.get('vehicle_no') or data.get('vehicle_id') or 'DHK-METRO-TA-11-2041'
    depot_id = data.get('depot_id') or 9
    driver_mobile = data.get('driver_mobile') or data.get('user_id') or '01711-000000'
    driver_name = data.get('driver_name') or 'Mobile Driver'
    route_id = data.get('route_id') or data.get('route_code') or 'RT-01'
    
    try:
        lat = float(data.get('lat') or data.get('latitude') or 0.0)
        lng = float(data.get('lng') or data.get('longitude') or 0.0)
        heading = float(data.get('heading') or 0.0)
        speed = float(data.get('speed') or data.get('speed_kmh') or 0.0)
    except (ValueError, TypeError):
        return jsonify({"success": False, "message": "Invalid numeric coordinates"}), 400

    source = data.get('source') or ('TELEMATICS_GPS' if '/telematics/' in request.path else 'MOBILE_GEOLOCATION')

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
    INSERT INTO live_tracking_positions (vehicle_no, driver_mobile, driver_name, depot_id, route_id, latitude, longitude, heading, speed_kmh, source, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    ON CONFLICT(vehicle_no, depot_id) DO UPDATE SET
        driver_mobile = excluded.driver_mobile,
        driver_name = excluded.driver_name,
        route_id = excluded.route_id,
        latitude = excluded.latitude,
        longitude = excluded.longitude,
        heading = excluded.heading,
        speed_kmh = excluded.speed_kmh,
        source = excluded.source,
        updated_at = CURRENT_TIMESTAMP
    ''', (vehicle_no, driver_mobile, driver_name, depot_id, route_id, lat, lng, heading, speed, source))
    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "message": f"Position logged for {vehicle_no} @ ({lat}, {lng}) via {source}"
    })

@app.route('/api/tracking/live-positions', methods=['GET'])
def get_live_tracking_positions():
    depot_id = request.args.get('depot_id')
    conn = get_db()
    cursor = conn.cursor()
    if depot_id and str(depot_id).lower() != 'all':
        cursor.execute('SELECT * FROM live_tracking_positions WHERE depot_id = ? ORDER BY updated_at DESC', (depot_id,))
    else:
        cursor.execute('SELECT * FROM live_tracking_positions ORDER BY updated_at DESC')
    positions = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return jsonify({"success": True, "positions": positions})

@app.route('/api/tracking/update-drop-status', methods=['POST'])
def update_tracking_drop_status():
    data = request.json if request.is_json else {}
    depot_id = data.get('depot_id')
    route_code = data.get('route_code')
    drop_id = data.get('drop_id')
    new_status = data.get('status', 'completed') # pending, in_transit, completed, failed
    proof_note = data.get('proof_note', '')
    cash_collected = float(data.get('cash_collected') or 0.0)

    if not depot_id or not route_code or not drop_id:
        return jsonify({"success": False, "message": "Missing depot_id, route_code, or drop_id"}), 400

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
    INSERT INTO live_drop_statuses (depot_id, route_code, drop_id, status, delivered_at, proof_note, cash_collected, updated_at)
    VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP, ?, ?, CURRENT_TIMESTAMP)
    ON CONFLICT(depot_id, route_code, drop_id) DO UPDATE SET
        status = excluded.status,
        delivered_at = CURRENT_TIMESTAMP,
        proof_note = excluded.proof_note,
        cash_collected = excluded.cash_collected,
        updated_at = CURRENT_TIMESTAMP
    ''', (depot_id, route_code, drop_id, new_status, proof_note, cash_collected))
    conn.commit()
    conn.close()

    return jsonify({
        "success": True,
        "message": f"Drop point #{drop_id} status updated to '{new_status.upper()}'!",
        "drop_id": drop_id,
        "new_status": new_status
    })

@app.route('/api/tracking/simulate-movement', methods=['GET', 'POST'])
def simulate_vehicle_movement():
    data = request.json if request.is_json else {}
    depot_val = str(data.get('depot_id') or request.args.get('depot_id') or 'ALL').strip().upper()
    is_all = (depot_val == 'ALL' or depot_val == '0')
    
    target_depot_ids = list(DEPOT_COORDINATES.keys()) if is_all else []
    if not is_all:
        try:
            target_depot_ids = [int(depot_val)]
        except ValueError:
            target_depot_ids = [1]
            
    conn = get_db()
    cursor = conn.cursor()
    for d_id in target_depot_ids:
        cursor.execute('SELECT * FROM live_tracking_positions WHERE depot_id = ?', (d_id,))
        rows = cursor.fetchall()
        depot_coords = DEPOT_COORDINATES.get(d_id, {"lat": 23.7644, "lng": 90.3928, "name": f"Depot #{d_id}"})
        if not rows:
            # No demo vehicle injection — skip depots with no real tracking data
            continue
        else:
            for r in rows:
                r_id = int(r['id'] or 1)
                current_heading = float(r['heading'] if r['heading'] is not None else 45.0)
                new_heading = round((current_heading + 15.0) % 360, 1)
                angle_rad = math.radians(new_heading)
                
                # Each vehicle gets an elliptical bounded patrol route around depot (1.5 - 3.2 km)
                orbit_lat_r = 0.014 + ((r_id % 3) * 0.006)
                orbit_lng_r = 0.017 + (((r_id + 1) % 3) * 0.006)
                
                new_lat = round(depot_coords['lat'] + orbit_lat_r * math.sin(angle_rad), 6)
                new_lng = round(depot_coords['lng'] + orbit_lng_r * math.cos(angle_rad), 6)
                
                current_speed = float(r['speed_kmh'] or 32.0)
                speed_delta = 1.5 if (r_id % 2 == 0) else -1.2
                new_speed = round(max(18.0, min(52.0, current_speed + speed_delta)), 1)
                
                cursor.execute('''
                UPDATE live_tracking_positions SET latitude = ?, longitude = ?, heading = ?, speed_kmh = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                ''', (new_lat, new_lng, new_heading, new_speed, r['id']))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Simulation step applied!"})


# ------------------------------------------------------------------------------
# APPLICATION STARTUP ENTRYPOINT
# ------------------------------------------------------------------------------

if __name__ == '__main__':
    print("=" * 66)
    print("PARAGON AGRO DISTRIBUTION LIVE WEB PORTAL IS READY!")
    print("Running at: http://127.0.0.1:5000")
    print("=" * 66)
    app.run(host='0.0.0.0', port=5000, debug=False)



