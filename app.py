from datetime import datetime, timedelta, timezone
from collections import defaultdict
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, jsonify
from database import get_db_connection, init_db
from io import BytesIO
import openpyxl
from fpdf import FPDF
import json
import random
import re

# Granular Action-Level Permission Catalog for Staff Access Control
PERMISSION_CATALOG = [
    {
        'category': 'Inventory Management',
        'icon': 'fa-boxes-stacked',
        'permissions': [
            {'key': 'inventory_view', 'action': 'VIEW', 'label': 'View Inventory', 'desc': 'Browse medicine catalog, current stock levels, and item details'},
            {'key': 'inventory_add', 'action': 'ADD', 'label': 'Add New Product', 'desc': 'Create new medicine / product records in the catalog'},
            {'key': 'inventory_edit', 'action': 'EDIT', 'label': 'Edit Product', 'desc': 'Modify product names, categories, pricing, and reorder points'},
            {'key': 'inventory_archive', 'action': 'ARCHIVE', 'label': 'Archive & Restore', 'desc': 'Archive inactive medicines and restore archived items'}
        ]
    },
    {
        'category': 'Batch Stocks & Expiry',
        'icon': 'fa-boxes-packing',
        'permissions': [
            {'key': 'batches_view', 'action': 'VIEW', 'label': 'View Batches', 'desc': 'Browse batch inventory, current quantities, and batch statuses'},
            {'key': 'batches_set_expiry', 'action': 'EDIT', 'label': 'Set Expiry Date', 'desc': 'Set or update expiration dates on batch stock'},
            {'key': 'expiry_view', 'action': 'VIEW', 'label': 'Expiry Monitoring', 'desc': 'Check near-expiry batches and expiration tracking dashboard'},
            {'key': 'expiry_dispose', 'action': 'DISPOSE', 'label': 'Record Disposal', 'desc': 'Permanently dispose expired batches and adjust stock levels'}
        ]
    },
    {
        'category': 'Suppliers Management',
        'icon': 'fa-truck-field',
        'permissions': [
            {'key': 'suppliers_view', 'action': 'VIEW', 'label': 'View Suppliers', 'desc': 'Browse supplier directory and contact information'},
            {'key': 'suppliers_add', 'action': 'ADD', 'label': 'Add Supplier', 'desc': 'Register new pharmaceutical distributors and suppliers'},
            {'key': 'suppliers_edit', 'action': 'EDIT', 'label': 'Edit Supplier', 'desc': 'Update supplier addresses, contact persons, and phone numbers'},
            {'key': 'suppliers_archive', 'action': 'ARCHIVE', 'label': 'Archive Supplier', 'desc': 'Archive and restore inactive suppliers'}
        ]
    },
    {
        'category': 'Purchase Orders (Procurement)',
        'icon': 'fa-cart-flatbed',
        'permissions': [
            {'key': 'po_view', 'action': 'VIEW', 'label': 'View Purchase Orders', 'desc': 'View purchase order records, statuses, and ordered items'},
            {'key': 'po_create', 'action': 'CREATE', 'label': 'Create Purchase Order', 'desc': 'Create and submit new purchase orders to suppliers'},
            {'key': 'po_edit', 'action': 'EDIT', 'label': 'Edit Purchase Order', 'desc': 'Modify quantities, items, and notes of pending orders'},
            {'key': 'po_receive', 'action': 'RECEIVE', 'label': 'Receive Deliveries', 'desc': 'Accept shipments and automatically stock-in into batch inventory'},
            {'key': 'po_delete', 'action': 'DELETE', 'label': 'Cancel / Delete PO', 'desc': 'Cancel and delete pending purchase orders'}
        ]
    },
    {
        'category': 'Point of Sale & Transactions',
        'icon': 'fa-cash-register',
        'permissions': [
            {'key': 'sales_view', 'action': 'VIEW', 'label': 'View Transactions', 'desc': 'Browse sales history, itemized receipts, and cash logs'},
            {'key': 'sales_create', 'action': 'POS', 'label': 'Process Sale / POS', 'desc': 'Execute customer checkouts and record cash sales transactions'},
            {'key': 'sales_edit', 'action': 'EDIT', 'label': 'Edit Sale Record', 'desc': 'Modify sold quantities and adjust recorded sale transactions'},
            {'key': 'movements_view', 'action': 'VIEW', 'label': 'Stock Movements', 'desc': 'View audit trail of stock adjustments, stock-in, and sales'}
        ]
    },
    {
        'category': 'Analytics, Reports & Settings',
        'icon': 'fa-chart-pie',
        'permissions': [
            {'key': 'reports_view', 'action': 'VIEW', 'label': 'View DSS Analytics', 'desc': 'View DSS charts, safety buffer, FEFO timeline, and velocity'},
            {'key': 'reports_download', 'action': 'DOWNLOAD', 'label': 'Download Reports', 'desc': 'Export sales and stock data to formatted PDF and Excel'},
            {'key': 'settings_view', 'action': 'VIEW', 'label': 'View Preferences', 'desc': 'View personal theme, display settings, and system parameters'},
            {'key': 'settings_edit', 'action': 'EDIT', 'label': 'Save Preferences', 'desc': 'Save changes to theme, accessibility options, and alerts'}
        ]
    }
]

# Standard default permissions for staff (safe operational viewing + POS checkout)
DEFAULT_STAFF_PERMISSIONS = [
    'inventory_view', 'batches_view', 'sales_view', 'sales_create', 'expiry_view', 'movements_view', 'settings_view'
]

def has_permission(perm_key):
    role = session.get('role', '')
    # Superadmin and Owner have full access across all system modules
    if role in ['Superadmin', 'Owner / Pharmacist'] or 'Super' in role or 'Owner' in role:
        return True
    user_perms = session.get('permissions')
    if user_perms is None or user_perms == '*':
        # Admin defaults to full access; Staff defaults to standard operational perms
        if role in ['Admin', 'Administrator']:
            return True
        return perm_key in DEFAULT_STAFF_PERMISSIONS
    if isinstance(user_perms, str):
        try:
            user_perms = json.loads(user_perms)
        except Exception:
            return False
    if isinstance(user_perms, list):
        return perm_key in user_perms
    return False

def is_superadmin():
    role = session.get('role', '')
    return role in ['Superadmin', 'Owner / Pharmacist'] or 'Super' in role or 'Owner' in role

def get_current_ph_time():
    try:
        import zoneinfo
        tz = zoneinfo.ZoneInfo('Asia/Manila')
        return datetime.now(tz)
    except Exception:
        tz = timezone(timedelta(hours=8))
        return datetime.now(tz)

def parse_dt_safe(dt_str):
    if not dt_str:
        return None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %I:%M:%S %p', '%Y-%m-%d %I:%M %p', '%Y-%m-%d %H:%M', '%Y-%m-%dT%H:%M:%S'):
        try:
            return datetime.strptime(str(dt_str).strip(), fmt)
        except (ValueError, TypeError):
            continue
    return None

def compute_user_online_status(u, now_ph=None):
    if now_ph is None:
        now_ph = get_current_ph_time().replace(tzinfo=None)
    is_active = u.get('is_active', 1)
    if not is_active:
        return False
    if not u.get('is_online'):
        return False
    last_seen = u.get('last_seen')
    if not last_seen:
        return False
    dt = parse_dt_safe(last_seen)
    if not dt:
        return False
    return (now_ph - dt).total_seconds() <= 45

reset_otps = {}

# System technical integer handling capacity limit (prevents crashes/overflow)
SYSTEM_MAX_PARAMETER_LIMIT = 1000000

# Maximum catalog limit of products allowed in inventory
MAX_PRODUCT_CATALOG_LIMIT = 2500

# Per-product maximum stock capacity limit
MAX_STOCK_PER_PRODUCT = 2500
MAX_STOCK_LIMIT = MAX_STOCK_PER_PRODUCT

import os
import sys

if getattr(sys, 'frozen', False):
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=template_folder, static_folder=static_folder)
else:
    app = Flask(__name__)

app.secret_key = 'super_secret_key_farmacia'
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Ensure database is initialized
init_db()

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

@app.context_processor
def inject_system_info():
    return {'local_ip': get_local_ip(), 'is_superadmin': is_superadmin()}

def parse_date_range(frequency, period_week=None, period_month=None, period_year=None, period_date=None):
    if frequency in ('all', '', None):
        return None, None
    if frequency == 'daily' and period_date:
        try:
            d = datetime.strptime(str(period_date).strip(), '%Y-%m-%d')
            return d, d.replace(hour=23, minute=59, second=59)
        except Exception:
            pass
    elif frequency == 'weekly' and period_week:
        try:
            year, week = period_week.split('-W')
            # fromisocalendar is Monday start of week
            start_date = datetime.fromisocalendar(int(year), int(week), 1)
            end_date = start_date + timedelta(days=6)
            return start_date, end_date
        except Exception:
            pass
    elif frequency == 'monthly' and period_month:
        try:
            year, month = period_month.split('-')
            start_date = datetime(int(year), int(month), 1)
            if int(month) == 12:
                end_date = datetime(int(year) + 1, 1, 1) - timedelta(days=1)
            else:
                end_date = datetime(int(year), int(month) + 1, 1) - timedelta(days=1)
            return start_date, end_date
        except Exception:
            pass
    elif frequency == 'annual' and period_year:
        try:
            start_date = datetime(int(period_year), 1, 1)
            end_date = datetime(int(period_year), 12, 31)
            return start_date, end_date
        except Exception:
            pass
    return None, None


def filter_by_date(row_date_str, start_dt, end_dt):
    try:
        dt = datetime.strptime(row_date_str, '%m/%d/%Y')
        return start_dt <= dt <= end_dt
    except Exception:
        return False

def get_desktop_folder():
    """
    Returns the absolute path to the active Windows Desktop folder,
    correctly resolving OneDrive redirected Desktop if active.
    """
    try:
        import ctypes
        from ctypes import wintypes
        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf)
        if buf.value and os.path.exists(buf.value):
            return buf.value
    except Exception:
        pass

    userprofile = os.environ.get('USERPROFILE', '')
    candidates = [
        os.path.join(userprofile, 'OneDrive', 'Desktop'),
        os.path.join(userprofile, 'Desktop'),
        os.path.join(os.environ.get('HOMEDRIVE', 'C:'), os.environ.get('HOMEPATH', ''), 'Desktop')
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return userprofile or os.getcwd()


def add_notification_with_conn(conn, type, title, subtitle, color):
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    existing = conn.execute('SELECT id FROM notifications WHERE title = ?', (title,)).fetchone()
    if not existing:
        conn.execute('''
            INSERT INTO notifications (type, title, subtitle, color, created_at, seen)
            VALUES (?, ?, ?, ?, ?, 0)
        ''', (type, title, subtitle, color, now_str))
    else:
        conn.execute('''
            UPDATE notifications
            SET subtitle = ?, color = ?, created_at = ?, seen = 0
            WHERE id = ?
        ''', (subtitle, color, now_str, existing['id']))

def add_notification(type, title, subtitle, color):
    conn = get_db_connection()
    add_notification_with_conn(conn, type, title, subtitle, color)
    conn.commit()
    conn.close()

def get_low_stock_threshold(conn):
    try:
        row = conn.execute("SELECT value FROM settings WHERE key = 'low_stock_threshold'").fetchone()
        if row and row['value'].isdigit():
            return int(row['value'])
    except Exception:
        pass
    return 40

def get_near_expiry_warning_days(conn):
    try:
        row = conn.execute("SELECT value FROM settings WHERE key = 'near_expiry_warning'").fetchone()
        if row and row['value'].isdigit():
            return int(row['value'])
    except Exception:
        pass
    return 90

def compute_batch_expiry_status(expiry_date, near_expiry_days=90, expiry_pending=False):
    if expiry_pending or not expiry_date or not str(expiry_date).strip():
        return 'Pending'
    try:
        exp_dt = datetime.strptime(str(expiry_date).strip(), '%m/%d/%Y')
        days_left = (exp_dt - datetime.now()).days
        if days_left <= 0:
            return 'Expired'
        elif days_left <= near_expiry_days:
            return 'Near Expiry'
        else:
            return 'Good'
    except Exception:
        return 'Pending'

def compute_stock_status(stock, low_threshold=40, reorder_point=None):
    if stock == 0:
        return 'No Stock'
    if stock <= low_threshold:
        return 'Low Stock'
    if reorder_point is not None and stock > reorder_point:
        return 'Overstocked'
    return 'Good'

def get_sales_trend_data(conn):
    """
    Computes real-time sales trend datasets for the dashboard graph:
    - this_week: Monday to Sunday of the current week (based on real-time now)
    - last_week: Monday to Sunday of the prior week
    - monthly: last 12 continuous calendar months ending at current real-time month
    - annual: continuous sequence of recent years ending at current year
    All dates are arranged in strict chronological order with clear, professional labels.
    """
    from datetime import datetime, timedelta
    from collections import defaultdict
    
    now = datetime.now()
    today = now.date()
    
    all_sales = conn.execute('SELECT sale_date, qty FROM sales').fetchall()
    sales_by_date = defaultdict(int)
    for s in all_sales:
        raw = str(s['sale_date']).strip()
        parsed = None
        for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%d/%m/%Y'):
            try:
                parsed = datetime.strptime(raw, fmt).date()
                break
            except ValueError:
                pass
        if parsed:
            sales_by_date[parsed] += int(s['qty'] or 0)
            
    # 1. This Week: Monday to Sunday of current real-time week
    mon_tw = today - timedelta(days=today.weekday())
    this_week_labels = []
    this_week_values = []
    this_week_dates  = []
    for i in range(7):
        d = mon_tw + timedelta(days=i)
        label = f"{d.strftime('%a')}, {d.strftime('%b')} {d.day:02d}"
        this_week_labels.append(label)
        this_week_values.append(sales_by_date.get(d, 0))
        this_week_dates.append(d.strftime('%m/%d/%Y'))
        
    # 2. Last Week: Monday to Sunday of previous week
    mon_lw = mon_tw - timedelta(days=7)
    last_week_labels = []
    last_week_values = []
    last_week_dates  = []
    for i in range(7):
        d = mon_lw + timedelta(days=i)
        label = f"{d.strftime('%a')}, {d.strftime('%b')} {d.day:02d}"
        last_week_labels.append(label)
        last_week_values.append(sales_by_date.get(d, 0))
        last_week_dates.append(d.strftime('%m/%d/%Y'))
        
    # 3. Monthly: Continuous last 12 months up to current real-time month
    monthly_labels = []
    monthly_values = []
    for i in range(11, -1, -1):
        m = now.month - i
        y = now.year
        while m <= 0:
            m += 12
            y -= 1
        dt_month = datetime(y, m, 1)
        monthly_labels.append(dt_month.strftime('%b %Y'))
        val = sum(qty for d, qty in sales_by_date.items() if d.year == y and d.month == m)
        monthly_values.append(val)
        
    # 4. Annual: Continuous range of years up to current real-time year
    years_with_sales = [d.year for d in sales_by_date.keys()]
    if years_with_sales:
        start_year = min(min(years_with_sales), now.year - 4)
    else:
        start_year = now.year - 4
    annual_labels = [str(y) for y in range(start_year, now.year + 1)]
    annual_values = [
        sum(qty for d, qty in sales_by_date.items() if d.year == y)
        for y in range(start_year, now.year + 1)
    ]
    
    return {
        'this_week': {
            'labels': this_week_labels,
            'values': this_week_values,
            'dates': this_week_dates,
            'range': f"{this_week_labels[0]} - {this_week_labels[-1]}, {now.year}"
        },
        'last_week': {
            'labels': last_week_labels,
            'values': last_week_values,
            'dates': last_week_dates,
            'range': f"{last_week_labels[0]} - {last_week_labels[-1]}, {mon_lw.year}"
        },
        'monthly': {
            'labels': monthly_labels,
            'values': monthly_values,
            'range': f"{monthly_labels[0]} - {monthly_labels[-1]}"
        },
        'annual': {
            'labels': annual_labels,
            'values': annual_values,
            'range': f"{annual_labels[0]} - {annual_labels[-1]}"
        }
    }

@app.route('/api/sales_trend')
def api_sales_trend():
    """Real-time API endpoint returning live sales trend data for the requested period."""
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    period = request.args.get('period', 'this_week')
    conn = get_db_connection()
    data = get_sales_trend_data(conn)
    conn.close()
    
    if period in data:
        return jsonify({
            'status': 'success',
            'period': period,
            'labels': data[period]['labels'],
            'values': data[period]['values'],
            'range': data[period].get('range', ''),
            'total_qty': sum(data[period]['values'])
        })
    return jsonify({'status': 'error', 'message': f'Invalid period: {period}'}), 400

def log_activity(conn, action, target_type, target_id, target_name, performed_by=None, details=''):
    """Log an action to the activity_log table with rich audit details and PH time."""
    try:
        by = performed_by
        if not by:
            try:
                from flask import has_request_context
                if has_request_context():
                    by = session.get('name', 'System')
                else:
                    by = 'System'
            except Exception:
                by = 'System'
        now_ph = get_current_ph_time()
        now_str = now_ph.strftime('%Y-%m-%d %I:%M %p')
        conn.execute(
            'INSERT INTO activity_log (action, target_type, target_id, target_name, performed_by, performed_at, details) VALUES (?, ?, ?, ?, ?, ?, ?)',
            (action, str(target_type), str(target_id), str(target_name), str(by), now_str, str(details or ''))
        )
    except Exception as e:
        print('log_activity error:', e)

@app.route('/api/user_ping', methods=['POST'])
def user_ping():
    """Background heartbeat to track live online/offline presence for active users."""
    if 'user_id' in session:
        uid = session['user_id']
        now_str = get_current_ph_time().strftime('%Y-%m-%d %H:%M:%S')
        try:
            conn = get_db_connection()
            conn.execute("UPDATE users SET is_online = 1, last_seen = ? WHERE id = ?", (now_str, uid))
            conn.commit()
            conn.close()
        except Exception:
            pass
        return jsonify({'status': 'ok'})
    return jsonify({'status': 'unauthenticated'}), 401

@app.route('/api/user_offline', methods=['POST'])
def user_offline():
    """Sets user offline when browser tab closes or unloads."""
    if 'user_id' in session:
        uid = session['user_id']
        now_str = get_current_ph_time().strftime('%Y-%m-%d %H:%M:%S')
        try:
            conn = get_db_connection()
            conn.execute("UPDATE users SET is_online = 0, last_seen = ? WHERE id = ?", (now_str, uid))
            conn.commit()
            conn.close()
        except Exception:
            pass
    return jsonify({'status': 'ok'})

@app.route('/api/users/live_status')
def api_users_live_status():
    """Real-time live presence status for users list page and header presence indicator."""
    if 'user_id' not in session:
        return jsonify({'status': 'unauthenticated'}), 401
    conn = get_db_connection()
    raw_users = conn.execute('SELECT id, name, username, role, is_active, is_online, last_seen, deactivated_at FROM users').fetchall()
    conn.close()
    
    now_ph = get_current_ph_time().replace(tzinfo=None)
    users_status = []
    online_users = []
    online_count = 0
    deactivated_count = 0
    current_uid = session.get('user_id')
    
    for u in raw_users:
        u_dict = dict(u)
        is_act = u_dict.get('is_active', 1)
        days_left = 30
        if is_act == 0:
            deactivated_count += 1
            deact_str = u_dict.get('deactivated_at')
            if deact_str:
                deact_dt = parse_dt_safe(deact_str)
                if deact_dt:
                    days_left = max(0, 30 - (now_ph - deact_dt).days)
            is_on = False
        else:
            is_on = compute_user_online_status(u_dict, now_ph)
            if is_on:
                online_count += 1
                online_users.append({
                    'id': u_dict['id'],
                    'name': u_dict.get('name') or u_dict.get('username') or 'User',
                    'username': u_dict.get('username', ''),
                    'role': u_dict.get('role', 'Staff'),
                    'is_you': (u_dict['id'] == current_uid)
                })
        
        users_status.append({
            'id': u_dict['id'],
            'is_active': is_act,
            'is_online': is_on,
            'days_remaining': days_left
        })
        
    total_users = len(users_status)
    active_users = total_users - deactivated_count
    offline_count = max(0, active_users - online_count)
    
    return jsonify({
        'users': users_status,
        'online_users': online_users,
        'online_count': online_count,
        'offline_count': offline_count,
        'total_users': total_users
    })

@app.context_processor
def inject_settings():
    conn = get_db_connection()
    try:
        settings_raw = conn.execute('SELECT * FROM settings').fetchall()
        settings_dict = {r['key']: r['value'] for r in settings_raw}
    except Exception:
        settings_dict = {}
        
    # Check batches for near expiry / expired and auto-insert alerts
    try:
        expiry_alerts_enabled = settings_dict.get('expiry_alerts', 'true') == 'true'
        if expiry_alerts_enabled:
            try:
                near_expiry_days = int(settings_dict.get('near_expiry_warning', 90))
            except (ValueError, TypeError):
                near_expiry_days = 90
            now_dt = datetime.now()
            batches = conn.execute('''
                SELECT b.id as batch_id, i.brand as medicine, b.expiry_date, b.current_qty
                FROM batches b
                JOIN inventory i ON b.medicine_id = i.id
                WHERE b.current_qty > 0
            ''').fetchall()
            for b in batches:
                try:
                    expiry_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
                    days_left = (expiry_dt - now_dt).days
                    if days_left <= 0:
                        add_notification_with_conn(conn, 'expiry', f"{b['medicine']} — Expired", f"Batch {b['batch_id']} expired on {b['expiry_date']}", '#ef4444')
                    elif 0 < days_left <= near_expiry_days:
                        add_notification_with_conn(conn, 'expiry', f"{b['medicine']} — Near Expiry", f"Expires in {days_left} days · Batch {b['batch_id']}", '#ffb800')
                except Exception:
                    pass
            conn.commit()
    except Exception as e:
        print("Error auto-checking expiry batches:", e)
        
    # Fetch notifications from table
    notifications = []
    unseen_count = 0
    try:
        notifications_raw = conn.execute('SELECT * FROM notifications ORDER BY id DESC').fetchall()
        for r in notifications_raw:
            notifications.append({
                'id': r['id'],
                'type': r['type'],
                'title': r['title'],
                'subtitle': r['subtitle'],
                'color': r['color'],
                'created_at': r['created_at'],
                'seen': r['seen']
            })
        unseen_count = conn.execute('SELECT COUNT(*) FROM notifications WHERE seen = 0').fetchone()[0]
    except Exception as e:
        print("Error fetching notifications:", e)
        
    conn.close()
    return dict(app_settings=settings_dict, app_notifications=notifications, app_unseen_count=unseen_count, has_permission=has_permission, permission_catalog=PERMISSION_CATALOG)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        
        # Validation 1: Username max 15 chars, Password exactly 10 chars
        if len(username) > 15:
            flash('Username cannot exceed 15 characters.', 'danger')
            return render_template('login.html')
            
        if len(password) != 10:
            flash('Password must be exactly 10 characters.', 'danger')
            return render_template('login.html')
            
        # Validation 2: No special characters allowed (alphanumeric only)
        if not re.match(r'^[a-zA-Z0-9]{1,15}$', username) or not re.match(r'^[a-zA-Z0-9]{10}$', password):
            flash('Username and password must contain only letters and numbers (no special characters allowed). Password must be exactly 10 characters.', 'danger')
            return render_template('login.html')
            
        conn = get_db_connection()
        user = conn.execute('SELECT * FROM users WHERE LOWER(username) = LOWER(?) AND password = ?', (username, password)).fetchone()
        
        # Superadmin fallback check for superadmin / admin credentials
        if not user and username.lower() in ('superadmin', 'admin') and password in ('adminowner', 'owner12345'):
            user = conn.execute("SELECT * FROM users WHERE role = 'Superadmin' LIMIT 1").fetchone()
        
        if user:
            # Prevent deactivated accounts from logging in
            if user['is_active'] == 0:
                conn.close()
                flash('This account has been deactivated. Please contact the administrator.', 'danger')
                return render_template('login.html')

            session['user_id'] = user['id']
            session['name'] = user['name']
            session['role'] = user['role']
            session['permissions'] = user['permissions']
            
            now_str = get_current_ph_time().strftime('%Y-%m-%d %H:%M:%S')
            conn.execute("UPDATE users SET is_online = 1, last_login = ?, last_seen = ? WHERE id = ?", (now_str, now_str, user['id']))
            log_activity(conn, 'login', 'User', user['id'], user['name'], performed_by=user['name'], details=f"Signed in as {user['role']}")
            conn.commit()
            conn.close()
            return redirect(url_for('dashboard'))
        else:
            conn.close()
            flash('Invalid credentials', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    uid = session.get('user_id')
    uname = session.get('name', 'User')
    if uid:
        try:
            now_str = get_current_ph_time().strftime('%Y-%m-%d %H:%M:%S')
            conn = get_db_connection()
            conn.execute("UPDATE users SET is_online = 0, last_logout = ?, last_seen = ? WHERE id = ?", (now_str, now_str, uid))
            log_activity(conn, 'logout', 'User', uid, uname, performed_by=uname, details="Signed out")
            conn.commit()
            conn.close()
        except Exception:
            pass
    session.clear()
    return redirect(url_for('login'))

@app.route('/send-otp', methods=['POST'])
def send_otp():
    email = request.form.get('email', '').strip()
    if not email:
        return jsonify({'success': False, 'message': 'Please enter your registered email address.'})
        
    conn = get_db_connection()
    user = conn.execute('SELECT * FROM users WHERE LOWER(email) = LOWER(?) OR LOWER(username) = LOWER(?)', (email, email)).fetchone()
    conn.close()
    
    if not user:
        return jsonify({'success': False, 'message': 'No registered account found with that email address.'})
        
    otp = f"{random.randint(100000, 999999)}"
    reset_otps[user['email'].lower()] = {
        'otp': otp,
        'user_id': user['id'],
        'email': user['email']
    }
    
    print(f"\n==========================================")
    print(f"  [OTP DISPATCH] Account: {user['username']} ({user['email']})")
    print(f"  [OTP DISPATCH] Generated OTP Code: {otp}")
    print(f"==========================================\n")
    
    return jsonify({
        'success': True,
        'message': f"OTP code has been sent to {user['email']}",
        'email': user['email'],
        'otp_demo': otp
    })

@app.route('/verify-otp', methods=['POST'])
def verify_otp():
    email = request.form.get('email', '').strip().lower()
    otp_entered = request.form.get('otp', '').strip()
    
    record = reset_otps.get(email)
    if not record:
        return jsonify({'success': False, 'message': 'OTP expired or not requested. Please request a new OTP.'})
        
    if record['otp'] == otp_entered:
        session['otp_verified_user_id'] = record['user_id']
        return jsonify({'success': True, 'user_id': record['user_id']})
    else:
        return jsonify({'success': False, 'message': 'Invalid OTP code. Please check and try again.'})

@app.route('/forgot-password', methods=['POST'])
def forgot_password():
    user_id = request.form.get('user_id', '').strip()
    new_password = request.form.get('new_password', '').strip()
    confirm_password = request.form.get('confirm_password', '').strip()
    
    verified_session_user = session.get('otp_verified_user_id')
    if not user_id or str(user_id) != str(verified_session_user):
        flash('Security verification incomplete. Please verify your OTP code first.', 'danger')
        return redirect(url_for('login'))
        
    if confirm_password and new_password != confirm_password:
        flash('Passwords do not match. Please try again.', 'danger')
        return redirect(url_for('login'))
        
    if len(new_password) != 10 or not re.match(r'^[a-zA-Z0-9]{10}$', new_password):
        flash('Password must be exactly 10 alphanumeric characters (letters and numbers only, no special characters allowed).', 'danger')
        return redirect(url_for('login'))
        
    conn = get_db_connection()
    user = conn.execute('SELECT id FROM users WHERE id = ?', (user_id,)).fetchone()
    if user:
        conn.execute('UPDATE users SET password = ? WHERE id = ?', (new_password, user['id']))
        conn.commit()
        conn.close()
        session.pop('otp_verified_user_id', None)
        flash('Password reset successfully! You can now log in with your new password.', 'success')
    else:
        conn.close()
        flash('User account not found.', 'danger')
        
    return redirect(url_for('login'))

@app.before_request
def require_login():
    allowed_routes = ['login', 'static', 'forgot_password', 'send_otp', 'verify_otp']
    if request.endpoint not in allowed_routes and 'user_id' not in session:
        return redirect(url_for('login'))
    # Dynamically synchronize permissions from DB for Staff and Admin so permission adjustments take effect immediately
    if 'user_id' in session and session.get('role') not in ['Superadmin', 'Owner / Pharmacist'] and 'Super' not in session.get('role', '') and 'Owner' not in session.get('role', ''):
        try:
            conn = get_db_connection()
            user_row = conn.execute('SELECT permissions, is_active FROM users WHERE id = ?', (session['user_id'],)).fetchone()
            conn.close()
            if user_row:
                if user_row['is_active'] == 0:
                    session.clear()
                    flash('This account has been deactivated.', 'danger')
                    return redirect(url_for('login'))
                session['permissions'] = user_row['permissions']
        except Exception:
            pass

@app.route('/')
def dashboard():
    conn = get_db_connection()
    # Stats
    total_medicines = conn.execute("SELECT COUNT(*) FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL").fetchone()[0]
    low_stock = conn.execute("SELECT COUNT(*) FROM inventory WHERE status IN ('Low Stock', 'No Stock') AND archived_at IS NULL AND deleted_at IS NULL").fetchone()[0]
    # Near expiry threshold from settings
    near_expiry_setting = conn.execute("SELECT value FROM settings WHERE key = 'near_expiry_warning'").fetchone()
    try:
        near_expiry_days = int(near_expiry_setting['value']) if near_expiry_setting else 90
    except (ValueError, TypeError):
        near_expiry_days = 90

    from datetime import datetime
    all_batches_raw = conn.execute("SELECT expiry_date FROM batches WHERE current_qty > 0").fetchall()
    near_expiry = 0
    for b in all_batches_raw:
        try:
            expiry_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
            days_left = (expiry_dt - datetime.now()).days
            if 0 < days_left <= near_expiry_days:
                near_expiry += 1
        except Exception:
            pass

    sales_recorded = conn.execute('SELECT COUNT(*) FROM sales').fetchone()[0]
    po_receiving = conn.execute("SELECT COUNT(*) FROM purchase_orders WHERE status IN ('For Receiving', 'For Replacement')").fetchone()[0]
    try:
        pending_batches_raw = conn.execute('''
            SELECT b.id as batch_id, b.medicine_id, i.brand as medicine, i.generic, i.category,
                   b.current_qty, b.expiry_date, b.mfg_date, b.expiry_pending, b.status
            FROM batches b
            JOIN inventory i ON b.medicine_id = i.id
            WHERE b.current_qty > 0
              AND (
                  b.expiry_pending = 1 
                  OR b.expiry_date IS NULL 
                  OR TRIM(b.expiry_date) = '' 
                  OR b.expiry_date = '—'
                  OR b.expiry_date = 'Pending'
                  OR b.mfg_date IS NULL 
                  OR TRIM(b.mfg_date) = '' 
                  OR b.mfg_date = '—'
                  OR b.mfg_date = 'Pending'
              )
            ORDER BY b.id ASC
        ''').fetchall()
        pending_batches = []
        for r in pending_batches_raw:
            d = dict(r)
            exp_val = (d.get('expiry_date') or '').strip()
            mfg_val = (d.get('mfg_date') or '').strip()
            missing_exp = bool(d.get('expiry_pending') == 1 or not exp_val or exp_val in ('—', 'Pending'))
            missing_mfg = bool(not mfg_val or mfg_val in ('—', 'Pending'))
            d['missing_expiry'] = missing_exp
            d['missing_mfg'] = missing_mfg
            
            exp_iso = ''
            if not missing_exp and exp_val:
                for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
                    try:
                        exp_iso = datetime.strptime(exp_val, fmt).strftime('%Y-%m-%d')
                        break
                    except Exception:
                        pass
            d['expiry_date_iso'] = exp_iso
            
            mfg_iso = ''
            if not missing_mfg and mfg_val:
                for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
                    try:
                        mfg_iso = datetime.strptime(mfg_val, fmt).strftime('%Y-%m-%d')
                        break
                    except Exception:
                        pass
            d['mfg_date_iso'] = mfg_iso
            pending_batches.append(d)
        pending_expiry = len(pending_batches)
    except Exception:
        pending_batches = []
        pending_expiry = 0
    
    # Tables
    low_stock_meds = conn.execute("SELECT brand as medicine, stock, reorder_point, status FROM inventory WHERE status IN ('Low Stock', 'No Stock') AND archived_at IS NULL AND deleted_at IS NULL LIMIT 5").fetchall()
    
    # Overstocked medicines: where stock exceeds reorder point (storage capacity threshold)
    overstocked_raw = conn.execute('''
        SELECT id, brand as medicine, stock, reorder_point, category
        FROM inventory
        WHERE stock > reorder_point AND archived_at IS NULL AND deleted_at IS NULL
        ORDER BY (stock - reorder_point) DESC
    ''').fetchall()
    overstocked_count = len(overstocked_raw)
    overstocked_meds = [
        {
            'id': r['id'],
            'medicine': r['medicine'],
            'stock': r['stock'],
            'reorder_point': r['reorder_point'],
            'excess': r['stock'] - r['reorder_point'],
            'category': r['category'],
            'status': 'Overstocked'
        }
        for r in overstocked_raw
    ]
    
    # For near expiry medicines table, join with inventory to get the brand name 
    near_expiry_meds_raw = conn.execute('''
        SELECT i.brand as medicine, b.expiry_date, b.status 
        FROM batches b 
        JOIN inventory i ON b.medicine_id = i.id
        WHERE b.current_qty > 0 AND b.status NOT IN ('Disposed', 'Expired')
    ''').fetchall()
    
    near_expiry_meds = []
    for b in near_expiry_meds_raw:
        try:
            expiry_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
            days_left = (expiry_dt - datetime.now()).days
        except Exception:
            days_left = None
        
        if days_left is not None and 0 < days_left <= near_expiry_days:
            near_expiry_meds.append({
                'medicine': b['medicine'],
                'expiry_date': b['expiry_date'],
                'days_left': f"{days_left} days",
                'status': 'Near Expiry'
            })
    near_expiry_meds = near_expiry_meds[:5]

    # Real-time Sales Trend datasets (This Week, Last Week, Monthly, Annual)
    trend_data = get_sales_trend_data(conn)
    last_week_labels = trend_data['last_week']['labels']
    last_week_values = trend_data['last_week']['values']
    this_week_labels = trend_data['this_week']['labels']
    this_week_values = trend_data['this_week']['values']
    monthly_labels = trend_data['monthly']['labels']
    monthly_values = trend_data['monthly']['values']
    annual_labels = trend_data['annual']['labels']
    annual_values = trend_data['annual']['values']
    trend_ranges = {k: trend_data[k].get('range', '') for k in trend_data}

    # Top 10 Products Customers Always Buy (Most Purchased / Fast-Moving)
    top_products_query = '''
        SELECT i.brand, i.generic, SUM(s.qty) as total_sold
        FROM sales s
        JOIN inventory i ON s.medicine_id = i.id
        WHERE i.archived_at IS NULL AND i.deleted_at IS NULL
        GROUP BY s.medicine_id
        ORDER BY total_sold DESC
        LIMIT 10
    '''
    top_products_raw = conn.execute(top_products_query).fetchall()
    top_product_labels = [row['brand'] for row in top_products_raw]
    top_product_values = [int(row['total_sold']) for row in top_products_raw]
    top_product_generics = [row['generic'] if row['generic'] and row['generic'] != 'N/A' else '' for row in top_products_raw]

    # Inventory stock levels (retained for backward compatibility)
    inventory_stocks = conn.execute('SELECT brand, stock FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL').fetchall()
    stock_labels = [item['brand'] for item in inventory_stocks]
    stock_values = [item['stock'] for item in inventory_stocks]

    # Recent activities (Superadmin only)
    current_role = session.get('role', '')
    is_superadmin = (
        current_role in ['Superadmin', 'Owner / Pharmacist']
        or 'Super' in current_role
        or 'Owner' in current_role
    )
    if is_superadmin:
        try:
            recent_activities_raw = conn.execute(
                'SELECT * FROM activity_log ORDER BY id DESC LIMIT 8'
            ).fetchall()
            recent_activities = [dict(r) for r in recent_activities_raw]
        except Exception:
            recent_activities = []
    else:
        recent_activities = []

    try:
        pending_audits_count = conn.execute("SELECT COUNT(*) FROM batch_count_audits WHERE status = 'Pending'").fetchone()[0]
    except Exception:
        pending_audits_count = 0

    conn.close()

    
    return render_template('dashboard.html', 
        total_medicines=total_medicines,
        low_stock=low_stock,
        near_expiry=near_expiry,
        sales_recorded=sales_recorded,
        po_receiving=po_receiving,
        low_stock_meds=low_stock_meds,
        near_expiry_meds=near_expiry_meds,
        last_week_labels=last_week_labels,
        last_week_values=last_week_values,
        this_week_labels=this_week_labels,
        this_week_values=this_week_values,
        weekly_labels=None,
        weekly_values=None,
        monthly_labels=monthly_labels,
        monthly_values=monthly_values,
        annual_labels=annual_labels,
        annual_values=annual_values,
        top_product_labels=top_product_labels,
        top_product_values=top_product_values,
        top_product_generics=top_product_generics,
        stock_labels=stock_labels,
        stock_values=stock_values,
        recent_activities=recent_activities,
        is_superadmin=is_superadmin,
        pending_expiry=pending_expiry,
        pending_batches=pending_batches,
        pending_audits_count=pending_audits_count,
        overstocked_count=overstocked_count,
        overstocked_meds=overstocked_meds,
        trend_ranges=trend_ranges,
        active_page='dashboard'
    )

def extract_dosage_and_size(description, product_name):
    """
    Extracts normalized (dosage, size) from description and product_name.
    """
    dosage = ""
    size = ""
    desc = (description or '').strip()
    if desc:
        for part in desc.split(','):
            part = part.strip()
            if ':' in part:
                k, v = part.split(':', 1)
                k_clean = k.strip().lower()
                v_clean = v.strip().lower()
                if k_clean in ('dosage', 'strength'):
                    dosage = v_clean
                elif k_clean in ('size', 'volume/size', 'volume'):
                    size = v_clean

    pname = (product_name or '').strip().lower()
    if not dosage:
        m = re.search(r'\b(\d+(?:\.\d+)?\s*(?:mg/5ml|mg/ml|mcg|mg|g|ml|iu))\b', pname, re.I)
        if m:
            dosage = m.group(1).replace(' ', '').lower()
    if not size:
        m_paren = re.search(r'\(([^)]+)\)', pname)
        if m_paren:
            size = m_paren.group(1).replace(' ', '').lower()
        else:
            m = re.search(r'\b(?:size[:\s]*)?(nb|s|m|l|xl|xxl|small|medium|large|extra\s*large)\b', pname, re.I)
            if m:
                size = m.group(1).replace(' ', '').lower()
                if size == 'small': size = 's'
                elif size == 'medium': size = 'm'
                elif size == 'large': size = 'l'
                elif size in ('extra large', 'extralarge'): size = 'xl'

    return dosage, size


def get_product_spec_tag(brand, product_name, description):
    """
    Extracts a concise, standardized size and dosage/strength specification tag
    (e.g., '[Size: M]', '[Size: S (100 g)]', '[500 mg]') to clearly differentiate
    variants of the same brand or base product in dropdowns and records.
    """
    desc = (description or '').strip()
    pname = (product_name or '').strip()
    brand_s = (brand or '').strip()
    
    size_val = ''
    wv_val = ''
    dos_val = ''
    
    for part in desc.split(','):
        part = part.strip()
        if ':' in part:
            k, v = part.split(':', 1)
            k = k.strip().lower()
            v = v.strip()
            if not v or v.lower() in ('n/a', 'standard', 'general'):
                continue
            if k in ('size', 'volume/size') and not re.match(r'^(?:box|bottle|pack|roll)\s+of', v, re.I):
                size_val = v
            elif k in ('weight/volume', 'volume/weight', 'volume', 'weight') and not re.match(r'^(?:box|bottle|pack|roll)\s+of', v, re.I):
                wv_val = v
            elif k in ('strength', 'dosage'):
                dos_val = v

    if not size_val:
        m_s = re.search(r'\b(?:Size[:\s]*)?(NB|S|M|L|XL|XXL|Small|Medium|Large|Extra\s*Large)\b', desc, re.I)
        if m_s:
            size_val = m_s.group(1).strip()
            
    if size_val.lower() == 'small': size_val = 'S'
    elif size_val.lower() == 'medium': size_val = 'M'
    elif size_val.lower() == 'large': size_val = 'L'
    elif size_val.lower() in ('extra large', 'extralarge'): size_val = 'XL'
    elif size_val.lower() in ('s', 'm', 'l', 'xl', 'xxl', 'nb'): size_val = size_val.upper()

    if not dos_val:
        m_d = re.search(r'\b(\d+(?:\.\d+)?\s*(?:mg\/5ml|mg\/ml|mcg|mg|iu))\b', desc, re.I)
        if m_d:
            dos_val = m_d.group(1).strip()

    pname_clean = pname.lower().replace(' ', '')
    brand_clean = brand_s.lower().replace(' ', '')
    
    has_size = bool(size_val and not re.search(r'\b' + re.escape(size_val) + r'\b', pname, re.I))
    has_wv = bool(wv_val and (wv_val.lower().replace(' ', '') not in pname_clean))
    has_dos = bool(dos_val and (dos_val.lower().replace(' ', '') not in pname_clean and dos_val.lower().replace(' ', '') not in brand_clean))

    parts = []
    if has_size:
        if has_wv and wv_val.lower().replace(' ', '') not in size_val.lower().replace(' ', ''):
            parts.append(f'Size: {size_val} ({wv_val})')
            has_wv = False
        else:
            parts.append(f'Size: {size_val}')
            
    if has_wv:
        parts.append(wv_val)
        
    if has_dos and (not wv_val or dos_val.lower().replace(' ', '') != wv_val.lower().replace(' ', '')):
        parts.append(dos_val)
        
    if parts:
        return ' [' + ' • '.join(parts) + ']'
    return ''


def clean_product_name_only(brand, product_name):
    """
    Returns only the clean name of the product without size, volume, or dosage specifications.
    Example:
      'Biogesic 500mg' -> 'Biogesic'
      'Pampers', 'Baby Diapers' -> 'Pampers - Baby Diapers'
      'Ventolin', 'Salbutamol' -> 'Ventolin - Salbutamol'
    """
    b = (brand or '').strip()
    p = (product_name or '').strip()
    if not b and not p:
        return 'Unknown Product'
    if not p:
        raw = b
    elif not b:
        raw = p
    elif b.lower() == p.lower():
        raw = b
    elif p.lower().startswith(b.lower()):
        raw = p
    elif b.lower() in p.lower():
        raw = p
    else:
        raw = f"{b} - {p}"
        
    # Strip compound dosages e.g. 500mg/2mg/10mg, 15mg/100mg
    s = re.sub(r'\b\d+(?:\.\d+)?\s*(?:mg|mcg|g|iu|ml|mL|kg)\s*(?:\/\s*\d+(?:\.\d+)?\s*(?:mg|mcg|g|iu|ml|mL|kg)\s*)+', '', raw, flags=re.I)
    # Strip single dosages/volumes e.g. 500mg, 200mg, 100mL, 100 g, 50mL, 400g, 10mg
    s = re.sub(r'\b\d+(?:\.\d+)?\s*(?:mg|mcg|iu|ml|mL|g|kg)\b', '', s, flags=re.I)
    # Clean up whitespace and trailing dashes or separators
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'[\s\-]+$', '', s).strip()
    return s if s else (b or p)


def extract_spec_detail(brand, product_name, description):
    """
    Extracts the size, dosage, or strength specifications as detail/record subtext.
    """
    spec = get_product_spec_tag(brand, product_name, description)
    parts = []
    if spec:
        spec_inner = spec.strip(' []')
        if spec_inner:
            parts.extend([p.strip() for p in spec_inner.split('•') if p.strip()])
            
    # Extract dosages from product_name if any
    found_dosages = re.findall(r'\b\d+(?:\.\d+)?\s*(?:mg(?:\/\d+(?:\.\d+)?(?:mg|ml|mL|mcg))?|mcg|iu|g|ml|mL|kg)\b(?:\/\d+(?:\.\d+)?\s*(?:mg|mcg|iu|g|ml|mL))?', product_name or '', re.I)
    for fd in found_dosages:
        fd_clean = fd.strip()
        if not any(fd_clean.lower().replace(' ', '') in p.lower().replace(' ', '') for p in parts):
            parts.append(fd_clean)

    if not parts:
        desc = (description or '').strip()
        for part in desc.split(','):
            if ':' in part:
                k, v = part.split(':', 1)
                k, v = k.strip().lower(), v.strip()
                if not v or v.lower() in ('n/a', 'standard', 'general'):
                    continue
                if k in ('dosage', 'strength'):
                    if not any(v.lower().replace(' ', '') in p.lower().replace(' ', '') for p in parts):
                        parts.append(v)
                elif k in ('size', 'volume/size') and not re.match(r'^(?:box|bottle|pack|roll)\s+of', v, re.I):
                    sz_str = f"Size: {v}"
                    if not any(v.lower().replace(' ', '') in p.lower().replace(' ', '') for p in parts):
                        parts.append(sz_str)
                elif k in ('weight/volume', 'volume/weight', 'volume', 'weight') and not re.match(r'^(?:box|bottle|pack|roll)\s+of', v, re.I):
                    if not any(v.lower().replace(' ', '') in p.lower().replace(' ', '') for p in parts):
                        parts.append(v)

    seen = set()
    final = []
    for p in parts:
        norm = p.lower().replace(' ', '')
        if norm not in seen:
            seen.add(norm)
            final.append(p)
    return ' • '.join(final)



def find_duplicate_product(conn, brand, product_name, generic, description, exclude_id=None):
    """
    Finds if an identical product already exists in inventory.
    Products with the same name/brand but DIFFERENT size or dosage are NOT duplicates.
    Only if brand/base name match AND dosage/size match (or both have neither).
    """
    brand_clean = (brand or '').strip().lower()
    prod_clean = (product_name or brand or '').strip().lower()
    target_dosage, target_size = extract_dosage_and_size(description, product_name)

    def clean_base(s):
        s = re.sub(r'\b\d+(?:\.\d+)?\s*(?:mg/5ml|mg/ml|mcg|mg|g|ml|iu)\b', '', s, flags=re.I)
        s = re.sub(r'\(.*?\)', '', s)
        s = re.sub(r'\b(?:size[:\s]*)?(nb|s|m|l|xl|xxl|small|medium|large)\b', '', s, flags=re.I)
        return ' '.join(s.strip().split())

    target_base = clean_base(prod_clean) or brand_clean

    query = "SELECT id, brand, product_name, generic, description FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL"
    params = []
    if exclude_id:
        query += " AND id != ?"
        params.append(exclude_id)

    rows = conn.execute(query, params).fetchall()
    for r in rows:
        r_brand = (r['brand'] or '').strip().lower()
        r_prod = (r['product_name'] or r['brand'] or '').strip().lower()
        r_base = clean_base(r_prod) or r_brand

        is_name_match = (
            (r_brand == brand_clean and r_base == target_base) or
            (r_prod == prod_clean) or
            (r_brand == brand_clean and (r_prod == target_base or prod_clean == r_base))
        )
        if not is_name_match:
            continue

        r_dosage, r_size = extract_dosage_and_size(r['description'], r['product_name'])

        # Compare dosage
        if target_dosage != r_dosage:
            continue

        # Compare size
        if target_size != r_size:
            continue

        # Both match
        return r

    return None


def generate_next_med_ids(conn, count=1):
    """
    Generates the next batch of unique MED-XXX IDs based on current numeric maximum.
    """
    rows = conn.execute("SELECT id FROM inventory").fetchall()
    max_num = 0
    for r in rows:
        rid = str(r['id'] or '').strip()
        m = re.search(r'MED-(\d+)', rid, re.I)
        if m:
            max_num = max(max_num, int(m.group(1)))
    new_ids = []
    for i in range(1, count + 1):
        new_ids.append(f"MED-{max_num + i:03d}")
    return new_ids


@app.route('/inventory', methods=['GET', 'POST'])
def inventory():
    if not has_permission('inventory_view'):
        flash('Access denied. You do not have permission to access Inventory.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    if request.method == 'POST':
        if not has_permission('inventory_add'):
            conn.close()
            flash('Access denied. You do not have permission to add new products.', 'warning')
            return redirect(url_for('inventory'))
        prod_count = conn.execute("SELECT COUNT(*) FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL").fetchone()[0]
        if prod_count >= MAX_PRODUCT_CATALOG_LIMIT:
            conn.close()
            flash('Maximum product limit reached (2,500 products max). Cannot add additional products.', 'error')
            return redirect(url_for('inventory'))

        generic = request.form.get('generic', '').strip()
        brand = request.form['brand'].strip()
        product_name = request.form.get('product_name', brand).strip()
        supplier = request.form.get('supplier', '').strip()

        # Handle new category creation under supplier
        new_category = request.form.get('new_category_input', '').strip()
        category = new_category if new_category else request.form.get('category', '').strip()
        if not category or category == '__create_new__':
            conn.close()
            flash('Please select or enter a category for this product.', 'error')
            return redirect(url_for('inventory'))

        reorder_point = min(MAX_STOCK_LIMIT, max(0, int(request.form.get('reorder_point', 100))))
        low_thresh = get_low_stock_threshold(conn)
        if reorder_point < low_thresh:
            conn.close()
            flash(f"Cannot add product: Reorder Point ({reorder_point}) cannot be less than the Low Stock Threshold ({low_thresh} units).", "error")
            return redirect(url_for('inventory'))
        
        unit_of_measure = request.form.get('unit_of_measure', 'Piece')
        base_unit = request.form.get('base_unit', unit_of_measure).strip() or unit_of_measure
        packaging_unit = request.form.get('packaging_unit', 'Box').strip() or 'Box'
        try:
            conversion_qty = max(1, int(request.form.get('conversion_qty', 1)))
        except (ValueError, TypeError):
            conversion_qty = 1
        purchase_price = float(request.form.get('purchase_price', 0.0))

        # Ensure category is registered under this supplier in supplier_categories
        if supplier and category:
            conn.execute("""
                INSERT OR IGNORE INTO supplier_categories (supplier_name, category_name)
                VALUES (?, ?)
            """, (supplier, category))

        # Collect size / dosage variants (multiple or single)
        variant_type = request.form.get('variant_type', '').strip().lower()  # 'dosage' or 'size'
        raw_variants = request.form.getlist('product_variants[]')
        if not raw_variants:
            raw_variants = request.form.getlist('product_dosages[]') or request.form.getlist('product_sizes[]')
            
        if not raw_variants:
            # Fallback to single inputs
            if request.form.get('med_dosage'):
                raw_variants = [request.form.get('med_dosage').strip()]
                variant_type = 'dosage'
            elif request.form.get('baby_size'):
                raw_variants = [request.form.get('baby_size').strip()]
                variant_type = 'size'

        variants = []
        for v in raw_variants:
            v_str = str(v).strip()
            if v_str and v_str != '__custom__' and v_str not in variants:
                variants.append(v_str)

        def get_variant_details(v_val):
            if not v_val:
                desc = request.form.get('description', '').strip()
                return product_name, desc
            
            # Check if variant is a combined dosage|size
            if '|' in v_val:
                parts = v_val.split('|')
                v_dos = parts[0].strip()
                v_sz = parts[1].strip() if len(parts) > 1 else ''
                pname = product_name
                desc_parts = []
                if v_dos:
                    if v_dos.lower() not in pname.lower():
                        pname = f"{pname} {v_dos}"
                    desc_parts.append(f"Dosage: {v_dos}")
                if v_sz:
                    if v_sz.lower() not in pname.lower():
                        pname = f"{pname} ({v_sz})"
                    desc_parts.append(f"Size: {v_sz}")
                return pname, ', '.join(desc_parts)

            if variant_type == 'dosage':
                pname = product_name if v_val.lower() in product_name.lower() else f"{product_name} {v_val}"
                desc = f"Dosage: {v_val}"
            elif variant_type == 'size':
                pname = product_name if v_val.lower() in product_name.lower() else f"{product_name} ({v_val})"
                desc = f"Size: {v_val}"
            else:
                pname = product_name if v_val.lower() in product_name.lower() else f"{product_name} {v_val}"
                desc = f"Detail: {v_val}"
            return pname, desc

        # Single product without size/dosage variants
        if not variants:
            pname, desc = get_variant_details(None)
            dup = find_duplicate_product(conn, brand, pname, generic, desc)
            if dup:
                matched_name = dup['product_name'] or dup['brand']
                conn.close()
                flash(f"Duplicate product detected! Product '{matched_name}' already exists in the system (SKU: {dup['id']}). Cannot save duplicate product.", "danger")
                return redirect(url_for('inventory'))
            
            med_id = generate_next_med_ids(conn, count=1)[0]
            conn.execute("""
                INSERT INTO inventory (id, generic, brand, category, stock, reorder_point, status, product_name, description, unit_of_measure, purchase_price, supplier, base_unit, packaging_unit, conversion_qty) 
                VALUES (?, ?, ?, ?, 0, ?, 'No Stock', ?, ?, ?, ?, ?, ?, ?, ?)
            """, (med_id, generic, brand, category, reorder_point, pname, desc, base_unit, purchase_price, supplier, base_unit, packaging_unit, conversion_qty))
            log_activity(conn, 'add', 'product', med_id, brand, details=f"Added product {brand} ({generic}) in category '{category}' under {supplier}")
            conn.commit()
            session['newly_added_id'] = med_id
            flash(f"Product '{brand}' was successfully added!", "success")
            return redirect(url_for('inventory'))

        # Check all variants for duplicates first
        dup_conflicts = []
        for v in variants:
            pname, desc = get_variant_details(v)
            dup = find_duplicate_product(conn, brand, pname, generic, desc)
            if dup:
                dup_conflicts.append((v, dup))

        if dup_conflicts:
            conn.close()
            conflicts_str = ", ".join([f"{v} (SKU {d['id']})" for v, d in dup_conflicts])
            flash(f"Duplicate product detected! The following variant(s) of '{product_name}' already exist in the system: {conflicts_str}. Products with the same name must have distinct sizes/dosages.", "danger")
            return redirect(url_for('inventory'))

        if prod_count + len(variants) > MAX_PRODUCT_CATALOG_LIMIT:
            conn.close()
            flash(f"Cannot add {len(variants)} products. It would exceed the maximum catalog limit of {MAX_PRODUCT_CATALOG_LIMIT} products.", "error")
            return redirect(url_for('inventory'))

        next_ids = generate_next_med_ids(conn, count=len(variants))
        added_names = []
        for idx, v in enumerate(variants):
            med_id = next_ids[idx]
            pname, desc = get_variant_details(v)
            conn.execute("""
                INSERT INTO inventory (id, generic, brand, category, stock, reorder_point, status, product_name, description, unit_of_measure, purchase_price, supplier, base_unit, packaging_unit, conversion_qty) 
                VALUES (?, ?, ?, ?, 0, ?, 'No Stock', ?, ?, ?, ?, ?, ?, ?, ?)
            """, (med_id, generic, brand, category, reorder_point, pname, desc, base_unit, purchase_price, supplier, base_unit, packaging_unit, conversion_qty))
            log_activity(conn, 'add', 'product', med_id, brand, details=f"Added product variant {pname} ({desc}) in category '{category}' under {supplier}")
            added_names.append(pname)

        conn.commit()
        session['newly_added_id'] = next_ids[0]
        if len(variants) == 1:
            flash(f"Product '{added_names[0]}' was successfully added!", "success")
        else:
            flash(f"Successfully added {len(variants)} individual products to inventory: {', '.join(added_names)}!", "success")
        return redirect(url_for('inventory'))
        
    newly_added_id = session.pop('newly_added_id', None)
    items_raw = conn.execute('SELECT rowid, * FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL ORDER BY rowid ASC').fetchall()
    items = [dict(r) for r in items_raw]
    if newly_added_id:
        idx = next((i for i, r in enumerate(items) if r['id'] == newly_added_id), None)
        if idx is not None:
            new_item = items.pop(idx)
            items.insert(0, new_item)

    suppliers = conn.execute('SELECT name FROM suppliers ORDER BY name ASC').fetchall()
    
    # Query supplier categories mapping: supplier_name -> [categories]
    sup_cat_rows = conn.execute("""
        SELECT supplier_name, category_name 
        FROM supplier_categories 
        ORDER BY category_name ASC
    """).fetchall()
    supplier_categories_map = {}
    for r in sup_cat_rows:
        s_name = r['supplier_name']
        supplier_categories_map.setdefault(s_name, []).append(r['category_name'])

    conn.close()
    return render_template('inventory.html', 
                           items=items, 
                           newly_added_id=newly_added_id,
                           suppliers=suppliers, 
                           supplier_categories_map=supplier_categories_map,
                           active_page='inventory')

@app.route('/inventory/edit', methods=['POST'])
def edit_medicine():
    if not has_permission('inventory_edit'):
        flash('Access denied. You do not have permission to edit inventory.', 'warning')
        return redirect(url_for('inventory'))
    conn = get_db_connection()
    med_id = request.form['id']
    generic = request.form.get('generic', '').strip()
    brand = request.form['brand'].strip()
    product_name = request.form.get('product_name', brand).strip()
    
    # Assemble description from dosage or size if provided
    baby_size = request.form.get('baby_size', '').strip()
    med_dosage = request.form.get('med_dosage', '').strip()
    description_parts = []
    if request.form.get('category') == 'Baby Care' and baby_size:
        description_parts.append(f"Size: {baby_size}")
    elif request.form.get('category') == 'Medicines' and med_dosage:
        description_parts.append(f"Dosage: {med_dosage}")
    elif request.form.get('description', '').strip():
        description_parts.append(request.form.get('description', '').strip())
    description = ', '.join(description_parts)
    
    # Check if another product already exists with this exact brand, product_name, and description/size/dosage
    existing_prod = find_duplicate_product(conn, brand, product_name, generic, description, exclude_id=med_id)
    if existing_prod:
        matched_name = existing_prod['product_name'] or existing_prod['brand']
        conn.close()
        flash(f"Duplicate product detected! Product '{matched_name}' already exists in the system (SKU: {existing_prod['id']}). Cannot save duplicate product.", "danger")
        return redirect(url_for('inventory'))

    # Handle new category creation under supplier
    new_category = request.form.get('new_category_input', '').strip()
    category = new_category if new_category else request.form.get('category', '').strip()
    if not category or category == '__create_new__':
        conn.close()
        flash('Please select or enter a category for this product.', 'error')
        return redirect(url_for('inventory'))

    reorder_point = min(MAX_STOCK_LIMIT, max(0, int(request.form.get('reorder_point', 100))))
    low_thresh = get_low_stock_threshold(conn)
    if reorder_point < low_thresh:
        conn.close()
        flash(f"Cannot update product: Reorder Point ({reorder_point}) cannot be less than the Low Stock Threshold ({low_thresh} units).", "error")
        return redirect(url_for('inventory'))
    
    description = request.form.get('description', '')
    unit_of_measure = request.form.get('unit_of_measure', 'Piece')
    base_unit = request.form.get('base_unit', unit_of_measure).strip() or unit_of_measure
    old_med = conn.execute("SELECT * FROM inventory WHERE id = ?", (med_id,)).fetchone()
    packaging_unit = (old_med['packaging_unit'] if old_med and old_med['packaging_unit'] else 'Box')
    conversion_qty = (old_med['conversion_qty'] if old_med and old_med['conversion_qty'] else 1)
    purchase_price = float(request.form.get('purchase_price', 0.0))
    supplier = request.form.get('supplier', '').strip()
    
    # Fetch old medicine details to record exact changes in activity log
    changes = []
    if old_med:
        if old_med['brand'] != brand:
            changes.append(f"Brand: '{old_med['brand']}' -> '{brand}'")
        if (old_med['product_name'] or old_med['brand']) != product_name:
            changes.append(f"Product Name: '{old_med['product_name'] or old_med['brand']}' -> '{product_name}'")
        if (old_med['generic'] or '') != generic:
            changes.append(f"Generic: '{old_med['generic'] or 'None'}' -> '{generic or 'None'}'")
        if old_med['category'] != category:
            changes.append(f"Category: '{old_med['category']}' -> '{category}'")
        if old_med['supplier'] != supplier:
            changes.append(f"Supplier: '{old_med['supplier']}' -> '{supplier}'")
        if old_med['reorder_point'] != reorder_point:
            changes.append(f"Reorder Point: {old_med['reorder_point']} -> {reorder_point}")
        if (old_med['base_unit'] or old_med['unit_of_measure'] or '') != base_unit:
            changes.append(f"Base Unit: '{old_med['base_unit'] or old_med['unit_of_measure']}' -> '{base_unit}'")
        if (old_med['description'] or '') != description:
            changes.append(f"Details: '{description}'")

    # Ensure category is registered under this supplier in supplier_categories
    if supplier and category:
        conn.execute("""
            INSERT OR IGNORE INTO supplier_categories (supplier_name, category_name)
            VALUES (?, ?)
        """, (supplier, category))

    # Recalculate status using low_stock_threshold
    stock = old_med['stock'] if old_med else 0
    new_status = compute_stock_status(stock, low_thresh, reorder_point)
    
    conn.execute('''
        UPDATE inventory 
        SET generic = ?, brand = ?, category = ?, reorder_point = ?, status = ?,
            product_name = ?, description = ?, unit_of_measure = ?, purchase_price = ?, supplier = ?,
            base_unit = ?, packaging_unit = ?, conversion_qty = ?
        WHERE id = ?
    ''', (generic, brand, category, reorder_point, new_status, product_name, description, base_unit, purchase_price, supplier, base_unit, packaging_unit, conversion_qty, med_id))
    details_str = f"Updated: {', '.join(changes)}" if changes else f"Updated details for {brand}"
    log_activity(conn, 'edit', 'product', med_id, brand, details=details_str)
    conn.commit()
    conn.close()
    flash('Medicine updated successfully!')
    return redirect(url_for('inventory'))

@app.route('/inventory/api/create_category', methods=['POST'])
def api_create_supplier_category():
    if not has_permission('inventory_add') and not has_permission('inventory_edit'):
        return jsonify({'success': False, 'error': 'Permission denied'}), 403
    data = request.get_json() or {}
    supplier = data.get('supplier', '').strip()
    category = data.get('category', '').strip()
    if not supplier:
        return jsonify({'success': False, 'error': 'Please select a supplier first.'}), 400
    if not category:
        return jsonify({'success': False, 'error': 'Category name cannot be empty.'}), 400
    
    conn = get_db_connection()
    conn.execute("""
        INSERT OR IGNORE INTO supplier_categories (supplier_name, category_name)
        VALUES (?, ?)
    """, (supplier, category))
    log_activity(conn, 'add', 'category', supplier, category, details=f"Created new category '{category}' under supplier '{supplier}'")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'supplier': supplier, 'category': category})

# ─── ARCHIVE / TRASH / RESTORE ROUTES ───────────────────────────────────────

@app.route('/inventory/archive/<med_id>', methods=['POST'])
def archive_medicine(med_id):
    if not has_permission('inventory_archive'):
        flash('Access denied. You do not have permission to archive products.', 'warning')
        return redirect(url_for('inventory'))
    conn = get_db_connection()
    med = conn.execute('SELECT brand FROM inventory WHERE id = ?', (med_id,)).fetchone()
    if med:
        now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
        conn.execute('UPDATE inventory SET archived_at = ?, deleted_at = NULL WHERE id = ?', (now_str, med_id))
        log_activity(conn, 'archive', 'product', med_id, med['brand'], details=f"Archived product {med['brand']}")
        conn.commit()
    conn.close()
    flash(f'{med["brand"]} has been archived.')
    return redirect(url_for('inventory'))

@app.route('/inventory/restore/<med_id>', methods=['POST'])
def restore_medicine(med_id):
    if not has_permission('inventory_archive'):
        flash('Access denied. You do not have permission to restore products.', 'warning')
        return redirect(url_for('history'))
    conn = get_db_connection()
    med = conn.execute('SELECT brand FROM inventory WHERE id = ?', (med_id,)).fetchone()
    if med:
        conn.execute('UPDATE inventory SET archived_at = NULL, deleted_at = NULL WHERE id = ?', (med_id,))
        log_activity(conn, 'restore', 'product', med_id, med['brand'], details=f"Restored {med['brand']} to active inventory")
        conn.commit()
    conn.close()
    flash(f'{med["brand"]} has been restored to active inventory.')
    return redirect(url_for('history'))

@app.route('/inventory/trash/<med_id>', methods=['POST'])
def trash_medicine(med_id):
    flash('Product deletion has been disabled. Please archive the product instead.', 'warning')
    return redirect(url_for('inventory'))

@app.route('/inventory/delete-permanent/<med_id>', methods=['POST'])
def delete_medicine_permanent(med_id):
    if not has_permission('inventory_delete'):
        flash('Access denied. You do not have permission to permanently delete products.', 'warning')
        return redirect(url_for('history'))
    conn = get_db_connection()
    med = conn.execute('SELECT brand FROM inventory WHERE id = ?', (med_id,)).fetchone()
    if med:
        log_activity(conn, 'delete', 'product', med_id, med['brand'], details=f"Permanently purged {med['brand']} (ID: {med_id}) from database")
        conn.execute('DELETE FROM inventory WHERE id = ?', (med_id,))
        conn.commit()
    conn.close()
    flash(f'{med["brand"]} has been permanently deleted.')
    return redirect(url_for('history'))

@app.route('/history')
def history():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    conn = get_db_connection()
    now = datetime.now()

    # Auto-purge products whose deleted_at > 30 days
    try:
        all_trashed = conn.execute("SELECT id, brand, deleted_at FROM inventory WHERE deleted_at IS NOT NULL").fetchall()
        for t in all_trashed:
            try:
                deleted_dt = datetime.strptime(t['deleted_at'], '%Y-%m-%d %I:%M %p')
                if (now - deleted_dt).days >= 30:
                    log_activity(conn, 'delete', 'product', t['id'], t['brand'], performed_by='System (Auto-Purge)')
                    conn.execute('DELETE FROM inventory WHERE id = ?', (t['id'],))
            except Exception:
                pass
        conn.commit()
    except Exception as e:
        print('Auto-purge error:', e)

    # Archived products
    archived = conn.execute(
        "SELECT * FROM inventory WHERE archived_at IS NOT NULL AND deleted_at IS NULL ORDER BY archived_at DESC"
    ).fetchall()

    # Archived suppliers
    archived_suppliers = conn.execute(
        "SELECT * FROM suppliers WHERE archived_at IS NOT NULL ORDER BY archived_at DESC"
    ).fetchall()

    # Combine into unified archived list
    all_archived = []
    for p in archived:
        all_archived.append({
            'type': 'Product',
            'id': p['id'],
            'name': p['product_name'] or p['brand'],
            'brand': p['brand'] if p['product_name'] and p['product_name'] != p['brand'] else '',
            'category': p['category'],
            'stock': p['stock'],
            'archived_at': p['archived_at'] or '',
            'restore_url': url_for('restore_medicine', med_id=p['id']),
            'trash_url': None,
            'can_trash': False
        })

    for s in archived_suppliers:
        all_archived.append({
            'type': 'Supplier',
            'id': s['id'],
            'name': s['name'],
            'brand': s['contact'] or '',
            'category': s['address'] or s['contact'] or '',
            'stock': None,
            'archived_at': s['archived_at'] or '',
            'restore_url': url_for('restore_supplier', sup_id=s['id']),
            'trash_url': None,
            'can_trash': False
        })

    def parse_archive_date(item):
        d_str = item.get('archived_at') or ''
        for fmt in ('%Y-%m-%d %I:%M %p', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
            try:
                return datetime.strptime(d_str, fmt)
            except ValueError:
                pass
        return datetime.min

    all_archived.sort(key=parse_archive_date, reverse=True)

    # Trashed products with days_remaining
    trashed_raw = conn.execute(
        "SELECT * FROM inventory WHERE deleted_at IS NOT NULL ORDER BY deleted_at DESC"
    ).fetchall()
    trashed = []
    for t in trashed_raw:
        try:
            deleted_dt = datetime.strptime(t['deleted_at'], '%Y-%m-%d %I:%M %p')
            days_left = 30 - (now - deleted_dt).days
            days_left = max(0, days_left)
        except Exception:
            days_left = 30
        trashed.append({'item': dict(t), 'days_remaining': days_left})

    # Activity log (paginated — last 1000)
    activity_log_raw = conn.execute(
        'SELECT * FROM activity_log ORDER BY id DESC LIMIT 1000'
    ).fetchall()
    activity_log = [dict(r) for r in activity_log_raw]

    conn.close()
    return render_template('history.html',
        archived=archived,
        archived_suppliers=archived_suppliers,
        all_archived=all_archived,
        trashed=trashed,
        activity_log=activity_log,
        active_page='history'
    )


@app.route('/batches')
def batches():
    if not has_permission('batches_view'):
        flash('Access denied. You do not have permission to access Batch Stocks.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    low_thresh = get_low_stock_threshold(conn)
    items = conn.execute('''
        SELECT b.id as batch_id, i.brand as medicine, i.generic, b.expiry_date, b.mfg_date, b.current_qty, b.expiry_pending, b.status as batch_status,
               b.received_date, b.supplier_name, b.packaging_unit, b.conversion_qty,
               i.supplier as default_supplier, i.category,
               i.stock as total_stock, i.reorder_point, i.status as inventory_status, i.base_unit
        FROM batches b 
        JOIN inventory i ON b.medicine_id = i.id
        WHERE b.current_qty > 0
        ORDER BY b.id ASC
    ''').fetchall()
    # Fetch pending actual count audits
    pending_audits_raw = conn.execute("""
        SELECT a.*, b.id as batch_id, i.brand as medicine_name, i.generic, i.base_unit, i.packaging_unit, i.conversion_qty,
               b.current_qty as live_batch_qty
        FROM batch_count_audits a
        JOIN batches b ON a.batch_id = b.id
        JOIN inventory i ON a.medicine_id = i.id
        WHERE a.status = 'Pending'
        ORDER BY a.submitted_at DESC
    """).fetchall()
    pending_audits = [dict(r) for r in pending_audits_raw]
    pending_by_batch = {a['batch_id']: a for a in pending_audits}
    conn.close()
    
    list_items = []
    for b in items:
        try:
            expiry_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
            days_left = (expiry_dt - datetime.now()).days
            if days_left <= 0:
                expiry_status = 'Expired'
            elif days_left <= 50:
                expiry_status = 'Near Expiry'
            else:
                expiry_status = 'Good'
        except Exception:
            expiry_status = 'Pending' if b['expiry_pending'] == 1 else 'Good'
            
        qty = b['current_qty']
        rop = b['reorder_point'] if b['reorder_point'] is not None else 100
        total_stock = b['total_stock'] if b['total_stock'] is not None else 0
        
        # Batch stock status:
        if qty > rop:
            stock_status = 'Overstocked'
        elif qty == 0:
            stock_status = 'No Stock'
        elif qty <= low_thresh:
            stock_status = 'Low Stock'
        else:
            stock_status = 'Good Stock'

        # Format ISO dates for modal pre-filling
        exp_iso = ''
        if b['expiry_date']:
            for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
                try:
                    exp_iso = datetime.strptime(b['expiry_date'], fmt).strftime('%Y-%m-%d')
                    break
                except Exception:
                    pass

        mfg_iso = ''
        if b['mfg_date']:
            for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d'):
                try:
                    mfg_iso = datetime.strptime(b['mfg_date'], fmt).strftime('%Y-%m-%d')
                    break
                except Exception:
                    pass

        missing_dates = bool(
            not b['mfg_date'] or str(b['mfg_date']).strip() in ('', '—', 'Pending')
            or not b['expiry_date'] or str(b['expiry_date']).strip() in ('', '—', 'Pending')
            or b['expiry_pending'] == 1
        )

        rec_date = (b['received_date'] or '').strip()
        if not rec_date or rec_date in ('—', 'Pending'):
            rec_date = '03/20/2026'

        supp_name = (b['supplier_name'] or b['default_supplier'] or 'Direct Supplier').strip()

        list_items.append({
            'batch_id': b['batch_id'],
            'medicine': b['medicine'],
            'generic': b['generic'] or '',
            'mfg_date': b['mfg_date'] or '—',
            'mfg_date_iso': mfg_iso,
            'expiry_date': b['expiry_date'] or '—',
            'expiry_date_iso': exp_iso,
            'received_date': rec_date,
            'supplier': supp_name,
            'current_qty': b['current_qty'],
            'base_unit': b['base_unit'] or 'Piece',
            'packaging_unit': b['packaging_unit'] or 'Box',
            'conversion_qty': b['conversion_qty'] or 1,
            'total_stock': total_stock,
            'reorder_point': rop,
            'status': expiry_status,
            'stock_status': stock_status,
            'category': b['category'],
            'missing_dates': missing_dates,
            'dates_saved': not missing_dates,
            'pending_actual_count': pending_by_batch.get(b['batch_id'])
        })
    return render_template('batches.html', items=list_items, pending_audits=pending_audits, is_superadmin=is_superadmin(), active_page='batches')

@app.route('/suppliers', methods=['GET', 'POST'])
def suppliers():
    if not has_permission('suppliers_view'):
        flash('Access denied. You do not have permission to access Suppliers.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    if request.method == 'POST':
        if not has_permission('suppliers_add'):
            conn.close()
            flash('Access denied. You do not have permission to add new suppliers.', 'warning')
            return redirect(url_for('suppliers'))
        name = request.form['name'].strip()
        address = request.form['address'].strip()
        contact = request.form['contact'].strip()
        
        # Check duplicate supplier by name or contact
        existing_sup = conn.execute("""
            SELECT id, name, contact FROM suppliers 
            WHERE (LOWER(TRIM(name)) = LOWER(?) OR (contact != '' AND contact = ?))
              AND archived_at IS NULL
            LIMIT 1
        """, (name, contact)).fetchone()
        if existing_sup:
            conn.close()
            flash(f"Duplicate supplier detected! Supplier '{existing_sup['name']}' ({existing_sup['contact']}) already exists in the system (ID: {existing_sup['id']}). Cannot save duplicate supplier.", "danger")
            return redirect(url_for('suppliers'))
        
        # Generate next Supplier ID (e.g. SUP-003)
        last_sup = conn.execute("SELECT id FROM suppliers ORDER BY id DESC LIMIT 1").fetchone()
        if last_sup:
            try:
                last_num = int(last_sup['id'].split('-')[1])
                next_sup_id = f"SUP-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_sup_id = "SUP-001"
        else:
            next_sup_id = "SUP-001"
            
        conn.execute("INSERT INTO suppliers (id, name, address, contact) VALUES (?, ?, ?, ?)",
                     (next_sup_id, name, address, contact))
        log_activity(conn, 'add_supplier', 'Supplier', next_sup_id, name, details=f"Added supplier {name} ({contact})")
        conn.commit()
        conn.close()
        return redirect(url_for('suppliers'))
        
    suppliers_raw = conn.execute('SELECT * FROM suppliers WHERE archived_at IS NULL ORDER BY id ASC').fetchall()
    items = []
    for s in suppliers_raw:
        s_dict = dict(s)
        prods = conn.execute('''
            SELECT id, brand, generic, category, stock 
            FROM inventory 
            WHERE supplier = ? AND archived_at IS NULL AND deleted_at IS NULL
            ORDER BY brand ASC
        ''', (s['name'],)).fetchall()
        s_dict['product_count'] = len(prods)
        s_dict['products'] = [dict(p) for p in prods]
        items.append(s_dict)
    conn.close()
    return render_template('suppliers.html', items=items, active_page='suppliers')

@app.route('/suppliers/<sup_id>/archive', methods=['POST'])
def archive_supplier(sup_id):
    if not has_permission('suppliers_archive') and not has_permission('suppliers_edit'):
        flash('Access denied. You do not have permission to archive suppliers.', 'warning')
        return redirect(url_for('suppliers'))
    conn = get_db_connection()
    sup = conn.execute("SELECT name FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    if sup:
        now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
        conn.execute("UPDATE suppliers SET archived_at = ? WHERE id = ?", (now_str, sup_id))
        log_activity(conn, 'archive_supplier', 'Supplier', sup_id, sup['name'], details=f"Archived supplier {sup['name']} (ID: {sup_id})")
        conn.commit()
        flash(f"Supplier {sup['name']} has been archived.")
    conn.close()
    return redirect(url_for('suppliers'))

@app.route('/suppliers/<sup_id>/restore', methods=['POST'])
def restore_supplier(sup_id):
    if not has_permission('suppliers_archive') and not has_permission('suppliers_edit'):
        flash('Access denied. You do not have permission to restore suppliers.', 'warning')
        return redirect(url_for('history'))
    conn = get_db_connection()
    sup = conn.execute("SELECT name FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    if sup:
        conn.execute("UPDATE suppliers SET archived_at = NULL WHERE id = ?", (sup_id,))
        log_activity(conn, 'restore_supplier', 'Supplier', sup_id, sup['name'], details=f"Restored supplier {sup['name']} (ID: {sup_id})")
        conn.commit()
        flash(f"Supplier {sup['name']} has been restored.")
    conn.close()
    return redirect(url_for('history'))

@app.route('/suppliers/edit', methods=['POST'])
def edit_supplier():
    if not has_permission('suppliers_edit'):
        flash('Access denied. You do not have permission to edit suppliers.', 'warning')
        return redirect(url_for('suppliers'))
    conn = get_db_connection()
    sup_id = request.form['id'].strip()
    name = request.form['name'].strip()
    address = request.form['address'].strip()
    contact = request.form['contact'].strip()
    
    # Check duplicate supplier by name or contact on another record
    existing_sup = conn.execute("""
        SELECT id, name, contact FROM suppliers 
        WHERE (LOWER(TRIM(name)) = LOWER(?) OR (contact != '' AND contact = ?))
          AND id != ?
          AND archived_at IS NULL
        LIMIT 1
    """, (name, contact, sup_id)).fetchone()
    if existing_sup:
        conn.close()
        flash(f"Duplicate supplier detected! Supplier '{existing_sup['name']}' ({existing_sup['contact']}) already exists in the system (ID: {existing_sup['id']}). Cannot save duplicate supplier.", "danger")
        return redirect(url_for('suppliers'))
    
    old_sup = conn.execute("SELECT * FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    old_name = old_sup['name'] if old_sup else None

    changes = []
    if old_sup:
        if old_sup['name'] != name:
            changes.append(f"Name: '{old_sup['name']}' -> '{name}'")
        if old_sup['address'] != address:
            changes.append(f"Address: '{old_sup['address']}' -> '{address}'")
        if old_sup['contact'] != contact:
            changes.append(f"Contact: '{old_sup['contact']}' -> '{contact}'")

    conn.execute('''
        UPDATE suppliers 
        SET name = ?, address = ?, contact = ?
        WHERE id = ?
    ''', (name, address, contact, sup_id))
    
    # Keep inventory and supplier category associations synchronized if supplier name changes
    if old_name and old_name != name:
        conn.execute("UPDATE inventory SET supplier = ? WHERE supplier = ?", (name, old_name))
        conn.execute("UPDATE supplier_categories SET supplier_name = ? WHERE supplier_name = ?", (name, old_name))

    details_str = f"Updated: {', '.join(changes)}" if changes else f"Updated supplier {name}"
    log_activity(conn, 'edit_supplier', 'Supplier', sup_id, name, details=details_str)
    conn.commit()
    conn.close()
    flash('Supplier updated successfully!')
    return redirect(url_for('suppliers'))

@app.route('/suppliers/<sup_id>/delete', methods=['POST'])
def delete_supplier(sup_id):
    flash('Supplier deletion has been disabled. Please archive the supplier instead.', 'warning')
    return redirect(url_for('suppliers'))

@app.route('/purchase_orders', methods=['GET', 'POST'])
def purchase_orders():
    if not has_permission('po_view'):
        flash('Access denied. You do not have permission to view Purchase Orders.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    
    if request.method == 'POST':
        if not has_permission('po_create'):
            conn.close()
            flash('Access denied. You do not have permission to create purchase orders.', 'warning')
            return redirect(url_for('purchase_orders'))
        supplier_id = request.form['supplier_id']
        notes = request.form.get('notes', '')
        medicines = request.form.getlist('medicine[]')
        quantities = request.form.getlist('quantity[]')
        packaging_units = request.form.getlist('packaging_unit[]')
        conversion_qtys = request.form.getlist('conversion_qty[]')
        
        # 1. Sanitize and structure items per row
        parsed_items = []
        for i in range(len(medicines)):
            m_id = medicines[i] if i < len(medicines) else ''
            q_val = quantities[i] if i < len(quantities) else '0'
            pkg_val = packaging_units[i] if i < len(packaging_units) else 'Box'
            conv_val = conversion_qtys[i] if i < len(conversion_qtys) else '1'
            if m_id and q_val:
                try:
                    q_int = int(q_val)
                    conv_int = max(1, int(conv_val))
                    pkg_str = (pkg_val or 'Box').strip()
                    if q_int > 0:
                        parsed_items.append({
                            'medicine_id': m_id,
                            'quantity': q_int,
                            'packaging_unit': pkg_str,
                            'conversion_qty': conv_int,
                            'base_qty': q_int * conv_int
                        })
                except (ValueError, TypeError):
                    pass
                    
        if not parsed_items:
            conn.close()
            flash('Please specify at least one valid product and quantity.', 'warning')
            return redirect(url_for('purchase_orders'))

        # Aggregate total base units ordered per medicine
        med_base_map = defaultdict(int)
        for it in parsed_items:
            med_base_map[it['medicine_id']] += it['base_qty']

        # 2. Validate against storage threshold (Reorder Point in base units) to prevent overstocking
        for med_id, total_base_order in med_base_map.items():
            med = conn.execute("SELECT id, brand, stock, reorder_point, base_unit, packaging_unit, conversion_qty FROM inventory WHERE id = ?", (med_id,)).fetchone()
            if not med:
                conn.close()
                flash(f"Invalid medicine selected ({med_id}).", "danger")
                return redirect(url_for('purchase_orders'))
            
            cur_stock = med['stock'] or 0
            threshold = med['reorder_point'] if (med['reorder_point'] is not None and med['reorder_point'] > 0) else 100
            base_unit = med['base_unit'] or 'Piece'
            
            pending_row = conn.execute('''
                SELECT SUM(poi.quantity * COALESCE(poi.conversion_qty, i.conversion_qty, 1)) as pending_base_qty
                FROM purchase_order_items poi
                JOIN purchase_orders po ON poi.purchase_order_id = po.id
                JOIN inventory i ON poi.medicine_id = i.id
                WHERE po.status IN ('For Receiving', 'For Replacement') AND poi.medicine_id = ?
            ''', (med_id,)).fetchone()
            pending_base_qty = pending_row['pending_base_qty'] or 0 if pending_row else 0
            
            max_orderable_base = max(0, threshold - cur_stock - pending_base_qty)
            
            if total_base_order > max_orderable_base or (cur_stock + pending_base_qty + total_base_order) > threshold:
                conn.close()
                flash(
                    f"Order rejected to prevent overstocking: {med['brand']} ordered total ({total_base_order:,} {base_unit}s) "
                    f"exceeds the storage capacity threshold ({threshold:,} {base_unit}s). "
                    f"Current stock: {cur_stock:,} {base_unit}s, Pending incoming: {pending_base_qty:,} {base_unit}s. "
                    f"Maximum allowable incoming order is {max_orderable_base:,} {base_unit}s.",
                    "danger"
                )
                return redirect(url_for('purchase_orders'))

        # Generate next PO ID
        last_po = conn.execute("SELECT id FROM purchase_orders ORDER BY id DESC LIMIT 1").fetchone()
        if last_po:
            try:
                last_num = int(last_po['id'].split('-')[1])
                next_po_id = f"PO-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_po_id = "PO-001"
        else:
            next_po_id = "PO-001"
            
        order_date = datetime.now().strftime('%m/%d/%Y')
        prepared_by = session.get('name', 'Owner')
        
        # Insert PO
        conn.execute('''
            INSERT INTO purchase_orders (id, supplier_id, prepared_by, order_date, status, notes) 
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (next_po_id, supplier_id, prepared_by, order_date, 'For Receiving', notes))
        
        # Insert PO Items with packaging_unit and conversion_qty
        for it in parsed_items:
            sanitized_qty = min(MAX_STOCK_LIMIT, max(1, it['quantity']))
            sanitized_conv = max(1, it['conversion_qty'])
            sanitized_pkg = it['packaging_unit'][:50] if it['packaging_unit'] else 'Box'
            conn.execute('''
                INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity, packaging_unit, conversion_qty) 
                VALUES (?, ?, ?, ?, ?)
            ''', (next_po_id, it['medicine_id'], sanitized_qty, sanitized_pkg, sanitized_conv))
            
        # Add PO pending notification
        supplier_name = conn.execute("SELECT name FROM suppliers WHERE id = ?", (supplier_id,)).fetchone()['name']
        add_notification_with_conn(conn, 'po', f"PO {next_po_id} pending receipt", f"From {supplier_name} · {len(parsed_items)} items · For Receiving", '#3b82f6')
        log_activity(conn, 'create_po', 'purchase_order', next_po_id, f"PO {next_po_id}", details=f"Supplier: {supplier_name} ({len(parsed_items)} items ordered)")
        
        conn.commit()
        conn.close()
        return redirect(url_for('purchase_orders'))
        
    # GET
    suppliers = conn.execute('SELECT * FROM suppliers ORDER BY id ASC').fetchall()
    medicines_raw = conn.execute('''
        SELECT i.id, i.brand, i.generic, i.product_name, i.category, i.supplier, i.stock, i.reorder_point, s.id as supplier_id,
               i.description, i.base_unit, i.packaging_unit, i.conversion_qty
        FROM inventory i
        LEFT JOIN suppliers s ON i.supplier = s.name
        WHERE i.archived_at IS NULL AND i.deleted_at IS NULL
        ORDER BY i.brand ASC
    ''').fetchall()
    medicines = [dict(m) for m in medicines_raw]
    
    # Calculate pending incoming quantities across all 'For Receiving' POs (in base units and pkg units)
    pending_rows = conn.execute('''
        SELECT poi.medicine_id, 
               SUM(poi.quantity * COALESCE(poi.conversion_qty, i.conversion_qty, 1)) as pending_base_qty,
               SUM(poi.quantity) as pending_pkg_qty
        FROM purchase_order_items poi
        JOIN purchase_orders po ON poi.purchase_order_id = po.id
        JOIN inventory i ON poi.medicine_id = i.id
        WHERE po.status IN ('For Receiving', 'For Replacement')
        GROUP BY poi.medicine_id
    ''').fetchall()
    pending_base_map = {r['medicine_id']: (r['pending_base_qty'] or 0) for r in pending_rows}
    pending_pkg_map = {r['medicine_id']: (r['pending_pkg_qty'] or 0) for r in pending_rows}

    # Build medicines stock info map for quick frontend threshold lookup
    medicines_stock_info = {}
    for m in medicines:
        m_id = m['id']
        stk = m.get('stock') or 0
        rop = m.get('reorder_point') if (m.get('reorder_point') is not None and m.get('reorder_point') > 0) else 100
        conv = m.get('conversion_qty') or 1
        base_unit = m.get('base_unit') or m.get('unit_of_measure') or 'Piece'
        pkg_unit = m.get('packaging_unit') or 'Box'
        p_base_qty = pending_base_map.get(m_id, 0)
        p_pkg_qty = pending_pkg_map.get(m_id, 0)
        max_ord_base = max(0, rop - stk - p_base_qty)
        max_ord_pkg = max_ord_base // conv
        medicines_stock_info[m_id] = {
            'id': m_id,
            'brand': m['brand'],
            'stock': stk,
            'reorder_point': rop,
            'pending_qty': p_pkg_qty,
            'pending_base_qty': p_base_qty,
            'conversion_qty': conv,
            'base_unit': base_unit,
            'packaging_unit': pkg_unit,
            'max_orderable': max_ord_pkg,
            'max_orderable_base': max_ord_base
        }

    # Build structured supplier catalog for supplier-specific category and product filtering
    supplier_catalog = {}
    for s in suppliers:
        s_id = s['id']
        supplier_catalog[s_id] = {
            'id': s_id,
            'name': s['name'],
            'categories': {}
        }
        # Pre-populate registered categories from supplier_categories
        registered_cats = conn.execute("SELECT category_name FROM supplier_categories WHERE supplier_name = ? ORDER BY category_name ASC", (s['name'],)).fetchall()
        for rc in registered_cats:
            cname = rc['category_name']
            if cname not in supplier_catalog[s_id]['categories']:
                supplier_catalog[s_id]['categories'][cname] = []
        
    for m in medicines:
        sid = m.get('supplier_id')
        if not sid or sid not in supplier_catalog:
            continue
        cat = m.get('category') or 'General'
        if cat not in supplier_catalog[sid]['categories']:
            supplier_catalog[sid]['categories'][cat] = []
        info = medicines_stock_info.get(m['id'], {})
        dos, sz = extract_dosage_and_size(m.get('description') or '', m.get('product_name') or '')
        spec = get_product_spec_tag(m['brand'], m.get('product_name'), m.get('description'))
        supplier_catalog[sid]['categories'][cat].append({
            'id': m['id'],
            'brand': m['brand'],
            'generic': m.get('generic') or '',
            'product_name': m.get('product_name') or '',
            'description': m.get('description') or '',
            'spec_tag': spec,
            'size': sz.upper() if sz in ('nb', 's', 'm', 'l', 'xl', 'xxl') else sz.title() if sz else '',
            'dosage': dos or '',
            'stock': info.get('stock', 0),
            'reorder_point': info.get('reorder_point', 100),
            'pending_qty': info.get('pending_qty', 0),
            'pending_base_qty': info.get('pending_base_qty', 0),
            'conversion_qty': info.get('conversion_qty', 1),
            'base_unit': info.get('base_unit', 'Piece'),
            'packaging_unit': info.get('packaging_unit', 'Box'),
            'max_orderable': info.get('max_orderable', 0),
            'max_orderable_base': info.get('max_orderable_base', 0)
        })
    
    # Fetch POs with individual items
    orders_raw = conn.execute('''
        SELECT po.id, po.supplier_id, s.name as supplier, po.prepared_by, po.order_date, po.status, po.notes, po.received_by,
               i.brand as medicine_name, poi.quantity as item_qty, i.product_name, i.description,
               i.base_unit,
               COALESCE(poi.packaging_unit, i.packaging_unit, 'Box') as packaging_unit,
               COALESCE(poi.conversion_qty, i.conversion_qty, 1) as conversion_qty
        FROM purchase_order_items poi
        JOIN purchase_orders po ON poi.purchase_order_id = po.id
        JOIN suppliers s ON po.supplier_id = s.id
        JOIN inventory i ON poi.medicine_id = i.id
        ORDER BY po.id DESC, poi.id ASC
    ''').fetchall()
    
    orders = []
    for o in orders_raw:
        notes_escaped = (o['notes'] or '').replace("'", "\\'").replace("\n", " ")
        _dos, _sz = extract_dosage_and_size(o['description'] or '', o['product_name'] or '')
        p_display = o['medicine_name']
        if _sz:
            p_display += f" - Size: {_sz.upper() if _sz in ('nb', 's', 'm', 'l', 'xl', 'xxl') else _sz.title()}"
        elif _dos:
            p_display += f" - {_dos}"
        orders.append({
            'id': o['id'],
            'supplier_id': o['supplier_id'],
            'supplier': o['supplier'],
            'prepared_by': o['prepared_by'],
            'order_date': o['order_date'],
            'status': o['status'],
            'notes': notes_escaped,
            'products_list': p_display,
            'total_quantity': o['item_qty'],
            'base_unit': o['base_unit'] or 'Piece',
            'packaging_unit': o['packaging_unit'] or 'Box',
            'conversion_qty': o['conversion_qty'] or 1,
            'received_by': o['received_by'] or ''
        })
    
    # Fetch detailed PO items for the Details Modal
    po_items_raw = conn.execute('''
        SELECT poi.purchase_order_id, poi.medicine_id, i.brand as medicine_name, i.generic as generic_name, i.category, poi.quantity,
               i.product_name, i.description, i.base_unit,
               COALESCE(poi.packaging_unit, i.packaging_unit, 'Box') as packaging_unit,
               COALESCE(poi.conversion_qty, i.conversion_qty, 1) as conversion_qty
        FROM purchase_order_items poi
        JOIN inventory i ON poi.medicine_id = i.id
    ''').fetchall()
    
    po_items_map = {}
    for item in po_items_raw:
        po_id = item['purchase_order_id']
        if po_id not in po_items_map:
            po_items_map[po_id] = []
        _dos, _sz = extract_dosage_and_size(item['description'] or '', item['product_name'] or '')
        item_disp_name = item['medicine_name']
        if _sz:
            item_disp_name += f" - Size: {_sz.upper() if _sz in ('nb', 's', 'm', 'l', 'xl', 'xxl') else _sz.title()}"
        elif _dos:
            item_disp_name += f" - {_dos}"
        po_items_map[po_id].append({
            'medicine_id': item['medicine_id'],
            'medicine_name': item_disp_name,
            'generic_name': item['generic_name'],
            'category': item['category'],
            'quantity': item['quantity'],
            'base_unit': item['base_unit'] or 'Piece',
            'packaging_unit': item['packaging_unit'] or 'Box',
            'conversion_qty': item['conversion_qty'] or 1,
            'size': _sz.upper() if _sz in ('nb', 's', 'm', 'l', 'xl', 'xxl') else _sz.title() if _sz else '',
            'dosage': _dos or ''
        })
        
    conn.close()
    return render_template('purchase_orders.html', 
                           suppliers=suppliers, 
                           medicines=medicines, 
                           orders=orders, 
                           po_items_map=po_items_map,
                           supplier_catalog=supplier_catalog,
                           medicines_stock_info=medicines_stock_info,
                           active_page='purchase_orders')

@app.route('/purchase_orders/<po_id>/items-json')
def po_items_json(po_id):
    """Return PO items as JSON for the expiry date modal."""
    if 'user_id' not in session:
        from flask import jsonify
        return jsonify({'items': []})
    conn = get_db_connection()
    items_raw = conn.execute('''
        SELECT poi.medicine_id, i.brand as medicine_name, poi.quantity, i.product_name, i.description,
               i.base_unit,
               COALESCE(poi.packaging_unit, i.packaging_unit, 'Box') as packaging_unit,
               COALESCE(poi.conversion_qty, i.conversion_qty, 1) as conversion_qty
        FROM purchase_order_items poi
        JOIN inventory i ON poi.medicine_id = i.id
        WHERE poi.purchase_order_id = ?
    ''', (po_id,)).fetchall()
    conn.close()
    
    items = []
    for r in items_raw:
        _dos, _sz = extract_dosage_and_size(r['description'] or '', r['product_name'] or '')
        disp_name = r['medicine_name']
        sz_formatted = _sz.upper() if _sz in ('nb', 's', 'm', 'l', 'xl', 'xxl') else (_sz.title() if _sz else '')
        size_dosage_str = f"Size: {sz_formatted}" if _sz else (_dos or '')
        if _sz:
            disp_name += f" - Size: {sz_formatted}"
        elif _dos:
            disp_name += f" - {_dos}"
        items.append({
            'medicine_id': r['medicine_id'],
            'brand': r['medicine_name'],
            'size_dosage': size_dosage_str,
            'medicine_name': disp_name,
            'quantity': r['quantity'],
            'base_unit': r['base_unit'] or 'Piece',
            'packaging_unit': r['packaging_unit'] or 'Box',
            'conversion_qty': r['conversion_qty'] or 1,
            'size': sz_formatted,
            'dosage': _dos or ''
        })
    from flask import jsonify
    return jsonify({'items': items})

@app.route('/purchase_orders/<po_id>/receive', methods=['POST'])
def receive_po(po_id):
    if not has_permission('po_receive'):
        flash('Access denied. You do not have permission to receive purchase order deliveries.', 'warning')
        return redirect(url_for('purchase_orders'))
    conn = get_db_connection()
    # Ensure expiry_pending column exists
    try:
        conn.execute("ALTER TABLE batches ADD COLUMN expiry_pending INTEGER DEFAULT 0")
        conn.commit()
    except Exception:
        pass  # already exists

    po = conn.execute("SELECT * FROM purchase_orders WHERE id = ?", (po_id,)).fetchone()
    if not po or po['status'] not in ('For Receiving', 'For Replacement'):
        conn.close()
        return redirect(url_for('purchase_orders'))

    po_supplier = conn.execute("SELECT name FROM suppliers WHERE id = ?", (po['supplier_id'],)).fetchone()
    po_supplier_name = po_supplier['name'] if po_supplier else ''
    rec_date = datetime.now().strftime('%m/%d/%Y')

    receiver = session.get('name', 'Unknown')
    skipped  = request.form.get('skip_expiry') == '1'
    delivery_status = request.form.get('delivery_status', 'complete')

    items = conn.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = ?", (po_id,)).fetchall()

    # Determine highest current batch number to sequentially assign next BAT-xxx IDs
    all_batches = conn.execute("SELECT id FROM batches").fetchall()
    max_batch_num = 0
    for b in all_batches:
        try:
            num = int(str(b['id']).split('-')[1])
            if num > max_batch_num:
                max_batch_num = num
        except Exception:
            pass

    # Determine highest current stock movement number to sequentially assign next TRN-xxx IDs
    all_movs = conn.execute("SELECT id FROM stock_movements").fetchall()
    max_mov_num = 0
    for m in all_movs:
        try:
            num = int(str(m['id']).split('-')[1])
            if num > max_mov_num:
                max_mov_num = num
        except Exception:
            pass

    remainder_items = []
    return_items = []
    received_items = []

    for item in items:
        med_id = item['medicine_id']
        ordered_qty = int(item['quantity'])

        if delivery_status == 'incomplete':
            try:
                raw_rem = request.form.get(f'remaining_{med_id}', 0)
                rem_qty = max(0, min(ordered_qty, int(raw_rem)))
            except (ValueError, TypeError):
                rem_qty = 0
            rec_qty = max(0, ordered_qty - rem_qty)
            conn.execute("UPDATE purchase_order_items SET quantity = ? WHERE id = ?", (rec_qty, item['id']))
            if rem_qty > 0:
                remainder_items.append({
                    'medicine_id': med_id,
                    'quantity': rem_qty,
                    'packaging_unit': item['packaging_unit'] if ('packaging_unit' in item.keys() and item['packaging_unit']) else 'Box',
                    'conversion_qty': item['conversion_qty'] if ('conversion_qty' in item.keys() and item['conversion_qty']) else 1
                })
            if rec_qty > 0:
                received_items.append({'medicine_id': med_id, 'quantity': rec_qty})
        elif delivery_status == 'return_replacement':
            try:
                raw_rec = request.form.get(f'received_{med_id}', ordered_qty)
                rec_qty = max(0, min(ordered_qty, int(raw_rec)))
            except (ValueError, TypeError):
                rec_qty = ordered_qty
            ret_qty = max(0, ordered_qty - rec_qty)
            ret_reason = (request.form.get(f'return_reason_{med_id}') or request.form.get('overall_return_reason') or 'Returned for replacement due to incident/defect').strip()
            conn.execute("UPDATE purchase_order_items SET quantity = ? WHERE id = ?", (rec_qty, item['id']))
            if ret_qty > 0:
                return_items.append({
                    'medicine_id': med_id,
                    'quantity': ret_qty,
                    'packaging_unit': item['packaging_unit'] if ('packaging_unit' in item.keys() and item['packaging_unit']) else 'Box',
                    'conversion_qty': item['conversion_qty'] if ('conversion_qty' in item.keys() and item['conversion_qty']) else 1,
                    'reason': ret_reason
                })
            if rec_qty > 0:
                received_items.append({'medicine_id': med_id, 'quantity': rec_qty})
        else:
            rem_qty = 0
            rec_qty = min(MAX_STOCK_LIMIT, max(0, ordered_qty))
            received_items.append({'medicine_id': med_id, 'quantity': rec_qty})

        if rec_qty > 0:
            # 1. Update inventory stock (converted to base units using item-specific conversion_qty)
            med = conn.execute("SELECT stock, reorder_point, brand, base_unit, packaging_unit, conversion_qty FROM inventory WHERE id = ?", (med_id,)).fetchone()
            current_stock = med['stock'] if med else 0
            conv = item['conversion_qty'] if ('conversion_qty' in item.keys() and item['conversion_qty']) else (med['conversion_qty'] if (med and med['conversion_qty']) else 1)
            base_unit = med['base_unit'] if (med and med['base_unit']) else 'Piece'
            pkg_unit = item['packaging_unit'] if ('packaging_unit' in item.keys() and item['packaging_unit']) else (med['packaging_unit'] if (med and med['packaging_unit']) else 'Box')
            base_rec_qty = rec_qty * conv
            new_stock = min(SYSTEM_MAX_PARAMETER_LIMIT, current_stock + base_rec_qty)
            conn.execute("UPDATE inventory SET stock = ? WHERE id = ?", (new_stock, med_id))
            if med:
                rop = med['reorder_point'] if med['reorder_point'] is not None else 100
                new_status = compute_stock_status(new_stock, get_low_stock_threshold(conn), rop)
                conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, med_id))
                if new_stock > rop:
                    add_notification_with_conn(conn, 'overstock', f"Overstock Warning: {med['brand']}", f"Stock level reached {new_stock} units. Storage capacity threshold (ROP: {rop} units) exceeded!", '#dc2626')
                    flash(f"Storage Capacity Warning: {med['brand']} stock is now {new_stock} units, exceeding storage capacity threshold (ROP: {rop} units)!", "warning")

            # 2. Parse Expiry Slices
            exp_mode = request.form.get(f'expiry_mode_{med_id}', 'same')
            exp_slices = []
            if exp_mode == 'split':
                sq = request.form.getlist(f'split_expiry_qty_{med_id}') or request.form.getlist(f'split_qty_{med_id}')
                sd = request.form.getlist(f'split_expiry_date_{med_id}') or request.form.getlist(f'split_date_{med_id}')
                for q, d in zip(sq, sd):
                    try:
                        qi = int(q)
                        if qi > 0:
                            exp_slices.append((qi, (d or '').strip()))
                    except (ValueError, TypeError):
                        continue
            if not exp_slices:
                raw_exp = request.form.get(f'expiry_{med_id}', '').strip()
                exp_slices = [(rec_qty, raw_exp)]

            # Align expiry slices strictly to rec_qty
            total_exp = sum(s[0] for s in exp_slices)
            if total_exp < rec_qty:
                exp_slices.append((rec_qty - total_exp, ''))
            elif total_exp > rec_qty:
                clamped = []
                acc = 0
                for q, d in exp_slices:
                    if acc + q <= rec_qty:
                        clamped.append((q, d))
                        acc += q
                    else:
                        rem = rec_qty - acc
                        if rem > 0:
                            clamped.append((rem, d))
                        break
                exp_slices = clamped

            # 3. Parse Manufactured Slices
            mfg_mode = request.form.get(f'mfg_mode_{med_id}', 'same')
            mfg_slices = []
            if mfg_mode == 'split':
                mq = request.form.getlist(f'split_mfg_qty_{med_id}')
                md = request.form.getlist(f'split_mfg_date_{med_id}')
                for q, d in zip(mq, md):
                    try:
                        qi = int(q)
                        if qi > 0:
                            mfg_slices.append((qi, (d or '').strip()))
                    except (ValueError, TypeError):
                        continue
            if not mfg_slices:
                raw_mfg = request.form.get(f'mfg_{med_id}', '').strip()
                mfg_slices = [(rec_qty, raw_mfg)]

            # Align manufactured slices strictly to rec_qty
            total_mfg = sum(s[0] for s in mfg_slices)
            if total_mfg < rec_qty:
                mfg_slices.append((rec_qty - total_mfg, ''))
            elif total_mfg > rec_qty:
                clamped = []
                acc = 0
                for q, d in mfg_slices:
                    if acc + q <= rec_qty:
                        clamped.append((q, d))
                        acc += q
                    else:
                        rem = rec_qty - acc
                        if rem > 0:
                            clamped.append((rem, d))
                        break
                mfg_slices = clamped

            # 4. Merge Expiry & Manufactured Slices into Batches
            item_batches = []
            e_idx, m_idx = 0, 0
            e_rem = exp_slices[0][0] if exp_slices else 0
            m_rem = mfg_slices[0][0] if mfg_slices else 0

            while e_idx < len(exp_slices) and m_idx < len(mfg_slices):
                chunk = min(e_rem, m_rem)
                if chunk > 0:
                    item_batches.append((chunk, exp_slices[e_idx][1], mfg_slices[m_idx][1]))
                    e_rem -= chunk
                    m_rem -= chunk
                if e_rem == 0:
                    e_idx += 1
                    if e_idx < len(exp_slices):
                        e_rem = exp_slices[e_idx][0]
                if m_rem == 0:
                    m_idx += 1
                    if m_idx < len(mfg_slices):
                        m_rem = mfg_slices[m_idx][0]

            if not item_batches:
                item_batches = [(rec_qty, '', '')]

            # 5. Create individual batch records & stock movements
            trn_ref = f'PO {po_id} (Partial)' if (delivery_status == 'incomplete' and rem_qty > 0) else f'PO {po_id}'
            trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')

            for b_qty, b_exp_raw, b_mfg_raw in item_batches:
                # b_qty is in packaging units; convert to base units for the batch current_qty
                b_base_qty = b_qty * conv
                max_batch_num += 1
                next_batch_id = f"BAT-{max_batch_num:03d}"

                if skipped:
                    expiry_date = ''
                    expiry_pending = 1
                else:
                    if b_exp_raw:
                        try:
                            dt = datetime.strptime(b_exp_raw, '%Y-%m-%d')
                            expiry_date = dt.strftime('%m/%d/%Y')
                        except Exception:
                            expiry_date = b_exp_raw
                        expiry_pending = 0
                    else:
                        expiry_date = ''
                        expiry_pending = 1

                # Format Manufactured Date
                mfg_date = ''
                if b_mfg_raw:
                    try:
                        dt = datetime.strptime(b_mfg_raw, '%Y-%m-%d')
                        mfg_date = dt.strftime('%m/%d/%Y')
                    except Exception:
                        mfg_date = b_mfg_raw

                batch_status = compute_batch_expiry_status(expiry_date, get_near_expiry_warning_days(conn), expiry_pending)
                conn.execute(
                    "INSERT INTO batches (id, medicine_id, expiry_date, current_qty, status, expiry_pending, mfg_date, received_date, supplier_name, packaging_unit, conversion_qty) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (next_batch_id, med_id, expiry_date, b_base_qty, batch_status, expiry_pending, mfg_date, rec_date, po_supplier_name, pkg_unit, conv)
                )

                max_mov_num += 1
                next_trn_id = f"TRN-{max_mov_num:03d}"
                trn_ref_desc = f"{trn_ref} ({b_qty} {pkg_unit}{'s' if b_qty > 1 else ''})" if conv > 1 else trn_ref
                conn.execute(
                    "INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (next_trn_id, 'Stock-In', med_id, next_batch_id, b_base_qty, trn_date, trn_ref_desc)
                )

    if delivery_status == 'incomplete' and remainder_items:
        all_pos = conn.execute("SELECT id FROM purchase_orders").fetchall()
        max_po_num = 0
        for r in all_pos:
            try:
                num = int(r['id'].split('-')[1])
                if num > max_po_num:
                    max_po_num = num
            except Exception:
                pass
        next_po_id = f"PO-{max_po_num + 1:03d}"

        total_rem_units = sum(it['quantity'] for it in remainder_items)
        total_rec_units = sum(it['quantity'] for it in received_items)

        rem_order_date = datetime.now().strftime('%m/%d/%Y')
        rem_notes = f"Remainder delivery from PO {po_id} ({total_rem_units} unit(s) pending delivery)"
        conn.execute('''
            INSERT INTO purchase_orders (id, supplier_id, prepared_by, order_date, status, notes)
            VALUES (?, ?, ?, ?, 'For Receiving', ?)
        ''', (next_po_id, po['supplier_id'], po['prepared_by'], rem_order_date, rem_notes))

        for rem_it in remainder_items:
            conn.execute('''
                INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity, packaging_unit, conversion_qty)
                VALUES (?, ?, ?, ?, ?)
            ''', (next_po_id, rem_it['medicine_id'], rem_it['quantity'], rem_it.get('packaging_unit', 'Box'), rem_it.get('conversion_qty', 1)))

        orig_notes = (po['notes'] or '').strip()
        updated_orig_notes = f"{orig_notes} [Partially received ({total_rec_units} units). Remainder ({total_rem_units} units) moved to {next_po_id}]".strip()
        conn.execute("UPDATE purchase_orders SET status = 'Received', received_by = ?, notes = ? WHERE id = ?", (receiver, updated_orig_notes, po_id))

        supplier_row = conn.execute("SELECT name FROM suppliers WHERE id = ?", (po['supplier_id'],)).fetchone()
        supplier_name = supplier_row['name'] if supplier_row else 'Supplier'

        add_notification_with_conn(conn, 'po', f"PO {next_po_id} created for remainder", f"Pending {total_rem_units} items from {supplier_name} · For Receiving", '#3b82f6')
        add_notification_with_conn(conn, 'po', f"PO {po_id} partially received", f"{total_rec_units} units stocked in. Remainder moved to {next_po_id}", '#10b981')
        log_activity(conn, 'po_partial_received', 'purchase_order', po_id, f"PO {po_id}", details=f"Partially received {total_rec_units} units. Remainder ({total_rem_units} units) moved to new PO {next_po_id}")
        flash(f"Incomplete delivery processed: {total_rec_units} unit(s) received into stock. A new purchase order record ({next_po_id}) has been created to receive the remaining {total_rem_units} unit(s).", "success")
    elif delivery_status == 'return_replacement' and return_items:
        all_pos = conn.execute("SELECT id FROM purchase_orders").fetchall()
        max_po_num = 0
        for r in all_pos:
            try:
                num = int(r['id'].split('-')[1])
                if num > max_po_num:
                    max_po_num = num
            except Exception:
                pass
        next_po_id = f"PO-{max_po_num + 1:03d}"

        total_ret_units = sum(it['quantity'] for it in return_items)
        total_rec_units = sum(it['quantity'] for it in received_items)

        ret_order_date = datetime.now().strftime('%m/%d/%Y')
        reason_list = list(dict.fromkeys(it.get('reason') for it in return_items if it.get('reason')))
        reason_summary = ('; '.join(reason_list))[:120] if reason_list else 'Incident/Defect'
        ret_notes = f"Replacement from PO {po_id}: {total_ret_units} unit(s) returned. Reason: {reason_summary}"
        conn.execute('''
            INSERT INTO purchase_orders (id, supplier_id, prepared_by, order_date, status, notes)
            VALUES (?, ?, ?, ?, 'For Replacement', ?)
        ''', (next_po_id, po['supplier_id'], po['prepared_by'], ret_order_date, ret_notes))

        for ret_it in return_items:
            conn.execute('''
                INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity, packaging_unit, conversion_qty)
                VALUES (?, ?, ?, ?, ?)
            ''', (next_po_id, ret_it['medicine_id'], ret_it['quantity'], ret_it.get('packaging_unit', 'Box'), ret_it.get('conversion_qty', 1)))

        orig_notes = (po['notes'] or '').strip()
        updated_orig_notes = f"{orig_notes} [Accepted {total_rec_units} unit(s). {total_ret_units} unit(s) returned for replacement in {next_po_id}]".strip()
        conn.execute("UPDATE purchase_orders SET status = 'Received', received_by = ?, notes = ? WHERE id = ?", (receiver, updated_orig_notes, po_id))

        supplier_row = conn.execute("SELECT name FROM suppliers WHERE id = ?", (po['supplier_id'],)).fetchone()
        supplier_name = supplier_row['name'] if supplier_row else 'Supplier'

        add_notification_with_conn(conn, 'po', f"PO {next_po_id} created for supplier replacement", f"{total_ret_units} item(s) to be replaced by {supplier_name} · For Replacement", '#d97706')
        add_notification_with_conn(conn, 'po', f"PO {po_id} received with returns", f"{total_rec_units} units stocked in. Returns sent in {next_po_id}", '#10b981')
        log_activity(conn, 'po_return_replacement', 'purchase_order', po_id, f"PO {po_id}", details=f"Received {total_rec_units} units. Returned {total_ret_units} units for replacement to {supplier_name} ({next_po_id})")
        flash(f"Delivery processed: {total_rec_units} acceptable unit(s) received into stock. {total_ret_units} unit(s) returned to supplier — new Purchase Order record {next_po_id} created under 'For Replacement'.", "success")
    elif po['status'] == 'For Replacement':
        conn.execute("UPDATE purchase_orders SET status = 'Received', received_by = ? WHERE id = ?", (receiver, po_id))
        add_notification_with_conn(conn, 'po', f"Replacement PO {po_id} received", "Replacement stock has been successfully received and added to inventory", '#10b981')
        log_activity(conn, 'po_replacement_received', 'purchase_order', po_id, f"PO {po_id}", details=f"Received replacement PO items into stock. Batch inventory updated")
        flash(f"Replacement PO {po_id} received completely. Stock levels updated.", "success")
    else:
        conn.execute("UPDATE purchase_orders SET status = 'Received', received_by = ? WHERE id = ?", (receiver, po_id))
        add_notification_with_conn(conn, 'po', f"PO {po_id} received", "Stock levels have been successfully updated", '#10b981')
        log_activity(conn, 'po_received', 'purchase_order', po_id, f"PO {po_id}", details=f"Received PO items into stock. Batch inventory updated")
        flash(f"PO {po_id} received completely. Stock levels updated.", "success")

    if skipped:
        add_notification_with_conn(conn, 'warning', f"Expiry dates pending for PO {po_id}", "Please set expiry dates in Expiry Monitoring", '#f59e0b')
    conn.commit()
    conn.close()
    return redirect(url_for('purchase_orders'))

@app.route('/batches/<batch_id>/set-expiry', methods=['POST'])
def set_batch_expiry(batch_id):
    """Update expiry date on a batch from Expiry Monitoring or Batches page."""
    if not has_permission('batches_set_expiry'):
        flash('Access denied. You do not have permission to set or update batch expiry dates.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    expiry_raw = request.form.get('expiry_date', '').strip()
    if expiry_raw:
        expiry_date = expiry_raw
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d'):
            try:
                dt = datetime.strptime(expiry_raw, fmt)
                expiry_date = dt.strftime('%m/%d/%Y')
                break
            except Exception:
                pass
        batch = conn.execute("SELECT b.id, b.medicine_id, b.expiry_date as old_expiry, i.brand FROM batches b JOIN inventory i ON b.medicine_id = i.id WHERE b.id = ?", (batch_id,)).fetchone()
        brand_name = batch['brand'] if batch and 'brand' in batch.keys() else f"Batch {batch_id}"
        old_expiry = batch['old_expiry'] if batch and batch['old_expiry'] else 'Pending'
        batch_status = compute_batch_expiry_status(expiry_date, get_near_expiry_warning_days(conn), expiry_pending=False)
        conn.execute("UPDATE batches SET expiry_date = ?, expiry_pending = 0, status = ? WHERE id = ?", (expiry_date, batch_status, batch_id))
        log_activity(conn, 'update_expiry', 'Batch', batch_id, brand_name, details=f"Changed expiry date from '{old_expiry}' to '{expiry_date}' for {brand_name}")
        conn.commit()
        flash(f'Expiry date updated successfully for batch {batch_id} ({brand_name}).', 'success')
    conn.close()
    ref = request.referrer
    if ref and ('/batches' in ref or '/expiry' in ref or '/dashboard' in ref or ref.endswith('/') or ref.endswith(':5000')):
        return redirect(ref)
    return redirect(url_for('expiry_monitoring'))


@app.route('/batches/<batch_id>/edit-stock', methods=['POST'])
def edit_batch_stock(batch_id):
    """Update manufactured date, expiry date, and current quantity for a batch (one-time completion for Admin/Superadmin)."""
    user_role = session.get('role', '')
    if user_role not in ('Admin', 'Superadmin', 'Super Admin') and not (has_permission('batches_set_expiry') or has_permission('inventory_edit')):
        flash('Access denied. Only Admin and Super Admin can fill in missing batch stock dates.', 'warning')
        return redirect(url_for('dashboard'))

    conn = get_db_connection()
    batch = conn.execute("""
        SELECT b.id, b.medicine_id, b.expiry_date, b.mfg_date, b.current_qty, b.expiry_pending,
               b.packaging_unit, b.conversion_qty,
               i.brand, i.generic, i.stock, i.reorder_point, i.base_unit
        FROM batches b
        JOIN inventory i ON b.medicine_id = i.id
        WHERE b.id = ?
    """, (batch_id,)).fetchone()

    if not batch:
        conn.close()
        flash(f'Batch {batch_id} was not found.', 'danger')
        return redirect(url_for('dashboard'))

    # Check if dates have already been saved and finalized
    already_saved_exp = bool(batch['expiry_pending'] == 0 and batch['expiry_date'] and str(batch['expiry_date']).strip() not in ('', '—', 'Pending'))
    already_saved_mfg = bool(batch['mfg_date'] and str(batch['mfg_date']).strip() not in ('', '—', 'Pending'))

    if already_saved_exp and already_saved_mfg:
        conn.close()
        flash(f'Expiration and manufacturing dates for batch {batch_id} are already saved and locked from editing.', 'warning')
        ref = request.referrer
        if ref and ('/batches' in ref or '/dashboard' in ref or '/expiry' in ref or ref.endswith('/') or ref.endswith(':5000')):
            return redirect(ref)
        return redirect(url_for('batches'))

    brand_name = batch['brand'] if batch and 'brand' in batch.keys() else f"Batch {batch_id}"
    old_expiry = batch['expiry_date'] or ''
    old_mfg = batch['mfg_date'] or ''
    old_qty = batch['current_qty']

    mfg_raw = request.form.get('mfg_date', '').strip()
    expiry_raw = request.form.get('expiry_date', '').strip()
    qty_raw = request.form.get('current_qty', '').strip()

    # Format mfg_date
    formatted_mfg = old_mfg
    if mfg_raw:
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d'):
            try:
                dt = datetime.strptime(mfg_raw, fmt)
                formatted_mfg = dt.strftime('%m/%d/%Y')
                break
            except Exception:
                pass
    elif 'mfg_date' in request.form:
        formatted_mfg = None

    # Format expiry_date
    formatted_expiry = old_expiry
    expiry_pending = batch['expiry_pending']
    if expiry_raw:
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d'):
            try:
                dt = datetime.strptime(expiry_raw, fmt)
                formatted_expiry = dt.strftime('%m/%d/%Y')
                expiry_pending = 0
                break
            except Exception:
                pass
    elif 'expiry_date' in request.form and not expiry_raw:
        # User explicitly cleared expiry date
        formatted_expiry = ''
        expiry_pending = 1

    # Optional current_qty adjustment
    new_qty = old_qty
    if qty_raw != '':
        try:
            q_val = int(qty_raw)
            if q_val >= 0:
                new_qty = q_val
        except (ValueError, TypeError):
            pass

    # Optional packaging_unit & conversion_qty adjustment
    old_pkg_unit = batch['packaging_unit'] or 'Box'
    old_conv = batch['conversion_qty'] or 1
    pkg_unit_raw = request.form.get('packaging_unit', '').strip()
    new_pkg_unit = pkg_unit_raw if pkg_unit_raw else old_pkg_unit
    conv_raw = request.form.get('conversion_qty', '').strip()
    try:
        new_conv = max(1, int(conv_raw)) if conv_raw else old_conv
    except (ValueError, TypeError):
        new_conv = old_conv

    # Recalculate expiry status
    batch_status = compute_batch_expiry_status(formatted_expiry, get_near_expiry_warning_days(conn), expiry_pending=bool(expiry_pending))

    # Update batch
    conn.execute("""
        UPDATE batches 
        SET expiry_date = ?, mfg_date = ?, current_qty = ?, expiry_pending = ?, status = ?,
            packaging_unit = ?, conversion_qty = ?
        WHERE id = ?
    """, (formatted_expiry or '', formatted_mfg, new_qty, expiry_pending, batch_status, new_pkg_unit, new_conv, batch_id))

    # If quantity changed, adjust parent inventory stock & status
    qty_delta = new_qty - old_qty
    if qty_delta != 0:
        new_inv_stock = max(0, (batch['stock'] or 0) + qty_delta)
        low_thresh = get_low_stock_threshold(conn)
        new_inv_status = compute_stock_status(new_inv_stock, low_thresh, batch['reorder_point'])
        conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_inv_stock, new_inv_status, batch['medicine_id']))

        # Log movement
        last_movement = conn.execute("SELECT id FROM stock_movements ORDER BY id DESC LIMIT 1").fetchone()
        next_trn_id = "TRN-001"
        if last_movement:
            try:
                last_num = int(last_movement['id'].split('-')[1])
                next_trn_id = f"TRN-{last_num + 1:03d}"
            except Exception:
                pass
        trn_type = 'Stock-In' if qty_delta > 0 else 'Stock-Out'
        trn_date = datetime.now().strftime('%m/%d/%Y')
        conn.execute(
            "INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (next_trn_id, trn_type, batch['medicine_id'], batch_id, abs(qty_delta), trn_date, f"Manual Batch Stock Adjustment ({batch_id})")
        )

    # Activity log
    log_details = f"Updated batch {batch_id} ({brand_name}): Mfg: {formatted_mfg or 'None'}, Expiry: {formatted_expiry or 'None'}"
    if qty_delta != 0:
        log_details += f", Qty: {old_qty} -> {new_qty}"
    if new_pkg_unit != old_pkg_unit or new_conv != old_conv:
        log_details += f", Packaging: 1 {new_pkg_unit} = {new_conv} {batch['base_unit'] or 'Piece'}"
    log_activity(conn, 'edit_batch_stock', 'Batch', batch_id, brand_name, details=log_details)

    conn.commit()
    conn.close()

    flash(f"Batch {batch_id} ({brand_name}) stock details updated successfully.", "success")
    ref = request.referrer
    if ref and ('/batches' in ref or '/dashboard' in ref or '/expiry' in ref or ref.endswith('/') or ref.endswith(':5000')):
        return redirect(ref)
    return redirect(url_for('dashboard'))


@app.route('/batches/<batch_id>/actual-count', methods=['POST'])
def submit_batch_actual_count(batch_id):
    """Staff or Admin submits physically counted shelf stock for a batch."""
    if not has_permission('batches_view'):
        flash('Access denied. You do not have permission to perform physical shelf counts.', 'warning')
        return redirect(url_for('batches'))

    conn = get_db_connection()
    batch = conn.execute("""
        SELECT b.id, b.medicine_id, b.current_qty, b.packaging_unit, b.conversion_qty,
               i.brand, i.generic, i.stock, i.reorder_point, i.base_unit
        FROM batches b
        JOIN inventory i ON b.medicine_id = i.id
        WHERE b.id = ?
    """, (batch_id,)).fetchone()

    if not batch:
        conn.close()
        flash(f'Batch {batch_id} was not found.', 'danger')
        return redirect(url_for('batches'))

    try:
        actual_qty = int(request.form.get('actual_qty', 0))
        if actual_qty < 0:
            actual_qty = 0
    except (ValueError, TypeError):
        conn.close()
        flash('Invalid actual count quantity entered.', 'danger')
        return redirect(url_for('batches'))

    reason = request.form.get('reason', '').strip()
    system_qty = batch['current_qty']
    discrepancy = actual_qty - system_qty
    base_unit = batch['base_unit'] or 'Piece'
    user_name = session.get('name', 'Staff Member')
    user_role = session.get('role', 'Staff')
    now_str = get_current_ph_time().strftime('%m/%d/%Y %I:%M %p')

    # If count exactly matches system quantity
    if discrepancy == 0:
        conn.close()
        flash(f"Actual shelf count for Batch {batch_id} ({batch['brand']}) matches the current system quantity of {system_qty:,} {base_unit}s. No adjustment required.", "info")
        return redirect(url_for('batches'))

    # If Superadmin directly submits, apply immediately and audit under their account
    if is_superadmin():
        delta = discrepancy
        new_inv_stock = max(0, (batch['stock'] or 0) + delta)
        low_thresh = get_low_stock_threshold(conn)
        new_inv_status = compute_stock_status(new_inv_stock, low_thresh, batch['reorder_point'])

        conn.execute("UPDATE batches SET current_qty = ? WHERE id = ?", (actual_qty, batch_id))
        conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_inv_stock, new_inv_status, batch['medicine_id']))

        max_trn_num = 0
        for r in conn.execute("SELECT id FROM stock_movements").fetchall():
            try:
                parts = r['id'].split('-')
                if len(parts) == 2 and parts[1].isdigit():
                    val = int(parts[1])
                    if val > max_trn_num:
                        max_trn_num = val
            except Exception:
                pass
        next_trn_id = f"TRN-{max_trn_num + 1:03d}"
        trn_type = 'Stock-In' if delta > 0 else 'Stock-Out'
        trn_date = datetime.now().strftime('%m/%d/%Y')
        conn.execute(
            "INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (next_trn_id, trn_type, batch['medicine_id'], batch_id, abs(delta), trn_date, f"Physical Count Audit ({batch_id}: {system_qty} -> {actual_qty}) by Superadmin {user_name}")
        )

        conn.execute("""
            INSERT INTO batch_count_audits
            (batch_id, medicine_id, system_qty, actual_qty, discrepancy, reason, submitted_by, submitted_by_role, submitted_at, status, reviewed_by, reviewed_at, review_notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Approved', ?, ?, ?)
        """, (batch_id, batch['medicine_id'], system_qty, actual_qty, discrepancy, reason, user_name, user_role, now_str, user_name, now_str, 'Directly verified & applied by Superadmin'))

        diff_str = f"+{delta:,}" if delta > 0 else f"{delta:,}"
        log_activity(conn, 'actual_count_applied', 'Batch', batch_id, batch['brand'],
                     details=f"Directly applied shelf count: {system_qty} -> {actual_qty} {base_unit}s ({diff_str}). Reason: {reason}")
        conn.commit()
        conn.close()
        flash(f"Actual shelf count for Batch {batch_id} ({batch['brand']}) applied successfully. Quantity updated from {system_qty:,} to {actual_qty:,} {base_unit}s ({diff_str}).", "success")
        return redirect(url_for('batches'))

    # Record or update pending verification request
    existing_pending = conn.execute("SELECT id FROM batch_count_audits WHERE batch_id = ? AND status = 'Pending'", (batch_id,)).fetchone()
    if existing_pending:
        conn.execute("""
            UPDATE batch_count_audits
            SET actual_qty = ?, discrepancy = ?, reason = ?, submitted_by = ?, submitted_by_role = ?, submitted_at = ?
            WHERE id = ?
        """, (actual_qty, discrepancy, reason, user_name, user_role, now_str, existing_pending['id']))
    else:
        conn.execute("""
            INSERT INTO batch_count_audits
            (batch_id, medicine_id, system_qty, actual_qty, discrepancy, reason, submitted_by, submitted_by_role, submitted_at, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pending')
        """, (batch_id, batch['medicine_id'], system_qty, actual_qty, discrepancy, reason, user_name, user_role, now_str))

    diff_str = f"+{discrepancy:,}" if discrepancy > 0 else f"{discrepancy:,}"
    # Send high-priority alert notification to Superadmin
    add_notification_with_conn(
        conn,
        'audit_pending',
        f"Actual Count Verification: {batch_id}",
        f"{user_name} ({user_role}) reported actual count of {actual_qty:,} {base_unit}s for {batch['brand']} (System: {system_qty:,}, Diff: {diff_str}). Superadmin verification required.",
        '#f59e0b'
    )

    log_activity(conn, 'actual_count_submitted', 'Batch', batch_id, batch['brand'],
                 details=f"Submitted physical count: {actual_qty} vs system {system_qty} {base_unit}s (Diff: {diff_str}). Reason: {reason}")

    conn.commit()
    conn.close()

    flash(f"Physical count for Batch {batch_id} ({batch['brand']}) submitted: {actual_qty:,} {base_unit}s (Discrepancy: {diff_str}). Notification sent to Superadmin for verification and approval.", "warning")
    return redirect(url_for('batches'))


@app.route('/batches/actual-count/<int:audit_id>/approve', methods=['POST'])
def approve_batch_actual_count(audit_id):
    """Superadmin approves actual count adjustment, updating batch and inventory numbers."""
    if not is_superadmin():
        flash('Access denied. Only Superadmin can verify and approve actual count adjustments.', 'warning')
        return redirect(url_for('batches'))

    conn = get_db_connection()
    audit = conn.execute("""
        SELECT a.*, b.current_qty as live_batch_qty, i.stock as inv_stock, i.reorder_point, i.brand, i.base_unit
        FROM batch_count_audits a
        JOIN batches b ON a.batch_id = b.id
        JOIN inventory i ON a.medicine_id = i.id
        WHERE a.id = ? AND a.status = 'Pending'
    """, (audit_id,)).fetchone()

    if not audit:
        conn.close()
        flash('Pending count verification request was not found or has already been reviewed.', 'warning')
        return redirect(url_for('batches'))

    batch_id = audit['batch_id']
    med_id = audit['medicine_id']
    brand_name = audit['brand']
    base_unit = audit['base_unit'] or 'Piece'
    actual_qty = audit['actual_qty']
    old_batch_qty = audit['live_batch_qty']
    delta = actual_qty - old_batch_qty

    superadmin_name = session.get('name', 'Superadmin')
    now_str = get_current_ph_time().strftime('%m/%d/%Y %I:%M %p')
    review_notes = request.form.get('review_notes', 'Verified and approved by Superadmin.').strip()

    # Update batch quantity
    conn.execute("UPDATE batches SET current_qty = ? WHERE id = ?", (actual_qty, batch_id))

    # Update inventory stock and status
    new_inv_stock = max(0, (audit['inv_stock'] or 0) + delta)
    low_thresh = get_low_stock_threshold(conn)
    new_inv_status = compute_stock_status(new_inv_stock, low_thresh, audit['reorder_point'])
    conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_inv_stock, new_inv_status, med_id))

    # Record stock movement
    if delta != 0:
        max_trn_num = 0
        for r in conn.execute("SELECT id FROM stock_movements").fetchall():
            try:
                parts = r['id'].split('-')
                if len(parts) == 2 and parts[1].isdigit():
                    val = int(parts[1])
                    if val > max_trn_num:
                        max_trn_num = val
            except Exception:
                pass
        next_trn_id = f"TRN-{max_trn_num + 1:03d}"
        trn_type = 'Stock-In' if delta > 0 else 'Stock-Out'
        trn_date = datetime.now().strftime('%m/%d/%Y')
        conn.execute(
            "INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (next_trn_id, trn_type, med_id, batch_id, abs(delta), trn_date, f"Physical Count Audit ({batch_id}: {old_batch_qty} -> {actual_qty}) approved by {superadmin_name}")
        )

    # Mark audit approved
    conn.execute("""
        UPDATE batch_count_audits
        SET status = 'Approved', reviewed_by = ?, reviewed_at = ?, review_notes = ?
        WHERE id = ?
    """, (superadmin_name, now_str, review_notes, audit_id))

    diff_str = f"+{delta:,}" if delta > 0 else f"{delta:,}"
    log_activity(conn, 'audit_approved', 'Batch', batch_id, brand_name,
                 details=f"Superadmin approved actual count: {old_batch_qty} -> {actual_qty} {base_unit}s ({diff_str}). Submitter: {audit['submitted_by']}")

    add_notification_with_conn(
        conn,
        'audit_approved',
        f"Actual Count Approved: {batch_id}",
        f"Superadmin {superadmin_name} approved shelf count for {brand_name}. Quantity updated to {actual_qty:,} {base_unit}s ({diff_str}).",
        '#10b981'
    )

    conn.commit()
    conn.close()

    flash(f"Actual count for Batch {batch_id} ({brand_name}) approved! Batch quantity updated to {actual_qty:,} {base_unit}s ({diff_str}).", "success")
    return redirect(url_for('batches'))


@app.route('/batches/actual-count/<int:audit_id>/deny', methods=['POST'])
def deny_batch_actual_count(audit_id):
    """Superadmin denies actual count adjustment; system quantity is maintained as already correct."""
    if not is_superadmin():
        flash('Access denied. Only Superadmin can verify and deny actual count adjustments.', 'warning')
        return redirect(url_for('batches'))

    conn = get_db_connection()
    audit = conn.execute("""
        SELECT a.*, b.current_qty as live_batch_qty, i.brand, i.base_unit
        FROM batch_count_audits a
        JOIN batches b ON a.batch_id = b.id
        JOIN inventory i ON a.medicine_id = i.id
        WHERE a.id = ? AND a.status = 'Pending'
    """, (audit_id,)).fetchone()

    if not audit:
        conn.close()
        flash('Pending count verification request was not found or has already been reviewed.', 'warning')
        return redirect(url_for('batches'))

    batch_id = audit['batch_id']
    brand_name = audit['brand']
    base_unit = audit['base_unit'] or 'Piece'
    system_qty = audit['live_batch_qty']

    superadmin_name = session.get('name', 'Superadmin')
    now_str = get_current_ph_time().strftime('%m/%d/%Y %I:%M %p')
    review_notes = request.form.get('review_notes', 'Denied: current system quantity confirmed correct.').strip()

    # System stock remains untouched!
    conn.execute("""
        UPDATE batch_count_audits
        SET status = 'Denied', reviewed_by = ?, reviewed_at = ?, review_notes = ?
        WHERE id = ?
    """, (superadmin_name, now_str, review_notes, audit_id))

    log_activity(conn, 'audit_denied', 'Batch', batch_id, brand_name,
                 details=f"Superadmin denied actual count adjustment ({audit['actual_qty']} reported by {audit['submitted_by']}). System quantity of {system_qty:,} {base_unit}s maintained. Reason: {review_notes}")

    add_notification_with_conn(
        conn,
        'audit_denied',
        f"Actual Count Denied: {batch_id}",
        f"Superadmin {superadmin_name} denied shelf count adjustment for {brand_name}. Current system quantity of {system_qty:,} {base_unit}s confirmed correct.",
        '#6b7280'
    )

    conn.commit()
    conn.close()

    flash(f"Actual count adjustment for Batch {batch_id} ({brand_name}) was denied. Current system quantity ({system_qty:,} {base_unit}s) confirmed as correct.", "info")
    return redirect(url_for('batches'))


@app.route('/purchase_orders/edit', methods=['POST'])
def edit_po():
    if not has_permission('po_edit'):
        flash('Access denied. You do not have permission to edit purchase orders.', 'warning')
        return redirect(url_for('purchase_orders'))
    conn = get_db_connection()
    po_id = request.form['id']
    supplier_id = request.form['supplier_id']
    notes = request.form['notes']
    medicines = request.form.getlist('medicine[]')
    quantities = request.form.getlist('quantity[]')
    packaging_units = request.form.getlist('packaging_unit[]')
    conversion_qtys = request.form.getlist('conversion_qty[]')
    
    # 1. Sanitize and structure items per row
    parsed_items = []
    for i in range(len(medicines)):
        m_id = medicines[i] if i < len(medicines) else ''
        q_val = quantities[i] if i < len(quantities) else '0'
        pkg_val = packaging_units[i] if i < len(packaging_units) else 'Box'
        conv_val = conversion_qtys[i] if i < len(conversion_qtys) else '1'
        if m_id and q_val:
            try:
                q_int = int(q_val)
                conv_int = max(1, int(conv_val))
                pkg_str = (pkg_val or 'Box').strip()
                if q_int > 0:
                    parsed_items.append({
                        'medicine_id': m_id,
                        'quantity': q_int,
                        'packaging_unit': pkg_str,
                        'conversion_qty': conv_int,
                        'base_qty': q_int * conv_int
                    })
            except (ValueError, TypeError):
                pass

    if not parsed_items:
        conn.close()
        flash('Please specify at least one valid product and quantity.', 'warning')
        return redirect(url_for('purchase_orders'))

    # Aggregate total base units ordered per medicine
    med_base_map = defaultdict(int)
    for it in parsed_items:
        med_base_map[it['medicine_id']] += it['base_qty']

    # 2. Validate against storage threshold (Reorder Point in base units) to prevent overstocking (excluding this PO's own items)
    for med_id, total_base_order in med_base_map.items():
        med = conn.execute("SELECT id, brand, stock, reorder_point, base_unit, packaging_unit, conversion_qty FROM inventory WHERE id = ?", (med_id,)).fetchone()
        if not med:
            conn.close()
            flash(f"Invalid medicine selected ({med_id}).", "danger")
            return redirect(url_for('purchase_orders'))
        
        cur_stock = med['stock'] or 0
        threshold = med['reorder_point'] if (med['reorder_point'] is not None and med['reorder_point'] > 0) else 100
        base_unit = med['base_unit'] or 'Piece'
        
        pending_row = conn.execute('''
            SELECT SUM(poi.quantity * COALESCE(poi.conversion_qty, i.conversion_qty, 1)) as pending_base_qty
            FROM purchase_order_items poi
            JOIN purchase_orders po ON poi.purchase_order_id = po.id
            JOIN inventory i ON poi.medicine_id = i.id
            WHERE po.status IN ('For Receiving', 'For Replacement') AND po.id != ? AND poi.medicine_id = ?
        ''', (po_id, med_id)).fetchone()
        pending_base_qty = pending_row['pending_base_qty'] or 0 if pending_row else 0
        
        max_orderable_base = max(0, threshold - cur_stock - pending_base_qty)
        
        if total_base_order > max_orderable_base or (cur_stock + pending_base_qty + total_base_order) > threshold:
            conn.close()
            flash(
                f"Order update rejected to prevent overstocking: {med['brand']} ordered total ({total_base_order:,} {base_unit}s) "
                f"exceeds the storage capacity threshold ({threshold:,} {base_unit}s). "
                f"Current stock: {cur_stock:,} {base_unit}s, Pending other orders: {pending_base_qty:,} {base_unit}s. "
                f"Maximum allowable incoming order is {max_orderable_base:,} {base_unit}s.",
                "danger"
            )
            return redirect(url_for('purchase_orders'))

    old_po = conn.execute("SELECT po.*, s.name as supplier_name FROM purchase_orders po LEFT JOIN suppliers s ON po.supplier_id = s.id WHERE po.id = ?", (po_id,)).fetchone()

    # Delete old items
    conn.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = ?", (po_id,))
    
    # Update PO
    conn.execute('''
        UPDATE purchase_orders 
        SET supplier_id = ?, notes = ?
        WHERE id = ?
    ''', (supplier_id, notes, po_id))
    
    # Insert new items with packaging_unit and conversion_qty
    item_summaries = []
    for it in parsed_items:
        sanitized_qty = min(MAX_STOCK_LIMIT, max(1, it['quantity']))
        sanitized_conv = max(1, it['conversion_qty'])
        sanitized_pkg = it['packaging_unit'][:50] if it['packaging_unit'] else 'Box'
        conn.execute('''
            INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity, packaging_unit, conversion_qty)
            VALUES (?, ?, ?, ?, ?)
        ''', (po_id, it['medicine_id'], sanitized_qty, sanitized_pkg, sanitized_conv))
        med_row = conn.execute("SELECT brand, base_unit FROM inventory WHERE id = ?", (it['medicine_id'],)).fetchone()
        med_name = med_row['brand'] if med_row else it['medicine_id']
        b_unit = (med_row['base_unit'] if med_row and med_row['base_unit'] else 'Piece')
        item_summaries.append(f"{med_name} ({sanitized_qty} {sanitized_pkg}{'s' if sanitized_qty > 1 else ''} · {sanitized_conv} {b_unit}s/pkg)")
            
    sup_name_row = conn.execute("SELECT name FROM suppliers WHERE id = ?", (supplier_id,)).fetchone()
    new_sup_name = sup_name_row['name'] if sup_name_row else supplier_id
    
    changes = []
    if old_po and old_po['supplier_id'] != supplier_id:
        changes.append(f"Supplier: '{old_po['supplier_name'] or old_po['supplier_id']}' -> '{new_sup_name}'")
    if old_po and (old_po['notes'] or '') != notes:
        changes.append(f"Notes: '{notes}'")
    changes.append(f"Items: [{', '.join(item_summaries)}]")

    details_str = f"Updated PO: {'; '.join(changes)}"
    log_activity(conn, 'edit_po', 'purchase_order', po_id, f"PO {po_id}", details=details_str)
    conn.commit()
    conn.close()
    flash('Purchase Order updated successfully!')
    return redirect(url_for('purchase_orders'))

@app.route('/purchase_orders/<po_id>/delete', methods=['POST'])
def delete_po(po_id):
    if not has_permission('po_delete'):
        flash('Access denied. You do not have permission to cancel or delete purchase orders.', 'warning')
        return redirect(url_for('purchase_orders'))
    conn = get_db_connection()
    po = conn.execute("SELECT * FROM purchase_orders WHERE id = ?", (po_id,)).fetchone()
    if po and po['status'] in ('For Receiving', 'For Replacement'):
        conn.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = ?", (po_id,))
        conn.execute("DELETE FROM purchase_orders WHERE id = ?", (po_id,))
        log_activity(conn, 'delete_po', 'purchase_order', po_id, f"PO {po_id}", details=f"Cancelled and deleted pending PO {po_id}")
        conn.commit()
        flash('Purchase Order deleted successfully!')
    conn.close()
    return redirect(url_for('purchase_orders'))

@app.route('/purchase_orders/<po_id>/return', methods=['POST'])
def return_po_to_supplier(po_id):
    if not has_permission('po_receive'):
        flash('Access denied. You do not have permission to return purchase order items to supplier.', 'warning')
        return redirect(url_for('purchase_orders'))
    conn = get_db_connection()
    po = conn.execute("SELECT * FROM purchase_orders WHERE id = ?", (po_id,)).fetchone()
    if not po or po['status'] != 'Received':
        conn.close()
        flash('Invalid purchase order or PO is not in Received status.', 'danger')
        return redirect(url_for('purchase_orders'))

    med_id = request.form.get('medicine_id')
    try:
        ret_pkg_qty = int(request.form.get('return_quantity', 0))
    except (ValueError, TypeError):
        ret_pkg_qty = 0
    reason = request.form.get('return_reason', '').strip() or 'Defective quality / Incident'
    notes_extra = request.form.get('notes', '').strip()

    if not med_id or ret_pkg_qty <= 0:
        conn.close()
        flash('Please select a product and enter a valid quantity to return.', 'danger')
        return redirect(url_for('purchase_orders'))

    po_item = conn.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = ? AND medicine_id = ?", (po_id, med_id)).fetchone()
    if not po_item:
        conn.close()
        flash('Selected item is not part of this purchase order.', 'danger')
        return redirect(url_for('purchase_orders'))

    med = conn.execute("SELECT * FROM inventory WHERE id = ?", (med_id,)).fetchone()
    if not med:
        conn.close()
        flash('Product not found in inventory.', 'danger')
        return redirect(url_for('purchase_orders'))

    conv = po_item['conversion_qty'] if ('conversion_qty' in po_item.keys() and po_item['conversion_qty']) else (med['conversion_qty'] or 1)
    pkg_unit = po_item['packaging_unit'] if ('packaging_unit' in po_item.keys() and po_item['packaging_unit']) else (med['packaging_unit'] or 'Box')
    base_unit = med['base_unit'] or 'Piece'
    ret_base_qty = ret_pkg_qty * conv

    current_stock = med['stock'] or 0
    if current_stock < ret_base_qty:
        conn.close()
        flash(f"Cannot return {ret_pkg_qty} {pkg_unit}(s) ({ret_base_qty} {base_unit}s): Current inventory stock is only {current_stock} {base_unit}s.", 'danger')
        return redirect(url_for('purchase_orders'))

    # Deduct from batches
    batches = conn.execute("SELECT * FROM batches WHERE medicine_id = ? AND current_qty > 0 ORDER BY id DESC", (med_id,)).fetchall()
    remaining_to_deduct = ret_base_qty
    for b in batches:
        b_qty = b['current_qty'] or 0
        deduct_from_b = min(b_qty, remaining_to_deduct)
        new_b_qty = b_qty - deduct_from_b
        conn.execute("UPDATE batches SET current_qty = ? WHERE id = ?", (new_b_qty, b['id']))
        remaining_to_deduct -= deduct_from_b
        if remaining_to_deduct <= 0:
            break

    # Deduct from inventory stock
    new_inv_stock = max(0, current_stock - ret_base_qty)
    conn.execute("UPDATE inventory SET stock = ? WHERE id = ?", (new_inv_stock, med_id))
    rop = med['reorder_point'] if med['reorder_point'] is not None else 100
    new_status = compute_stock_status(new_inv_stock, get_low_stock_threshold(conn), rop)
    conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, med_id))

    # Log stock movement
    all_movs = conn.execute("SELECT id FROM stock_movements").fetchall()
    max_mov_num = 0
    for m in all_movs:
        try:
            num = int(str(m['id']).split('-')[1])
            if num > max_mov_num:
                max_mov_num = num
        except Exception:
            pass
    next_trn_id = f"TRN-{max_mov_num + 1:03d}"
    trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    conn.execute(
        "INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (next_trn_id, 'Stock-Out', med_id, '', ret_base_qty, trn_date, f"Return to Supplier: PO {po_id} incident ({reason})")
    )

    # Create new PO with status 'For Replacement'
    all_pos = conn.execute("SELECT id FROM purchase_orders").fetchall()
    max_po_num = 0
    for r in all_pos:
        try:
            num = int(r['id'].split('-')[1])
            if num > max_po_num:
                max_po_num = num
        except Exception:
            pass
    next_po_id = f"PO-{max_po_num + 1:03d}"
    ret_order_date = datetime.now().strftime('%m/%d/%Y')
    full_reason = f"{reason} - {notes_extra}" if notes_extra else reason
    po_notes = f"Incident return from PO {po_id}: {ret_pkg_qty} {pkg_unit}(s) returned. Reason: {full_reason}"

    conn.execute('''
        INSERT INTO purchase_orders (id, supplier_id, prepared_by, order_date, status, notes)
        VALUES (?, ?, ?, ?, 'For Replacement', ?)
    ''', (next_po_id, po['supplier_id'], session.get('name', 'Staff'), ret_order_date, po_notes))

    conn.execute('''
        INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity, packaging_unit, conversion_qty)
        VALUES (?, ?, ?, ?, ?)
    ''', (next_po_id, med_id, ret_pkg_qty, pkg_unit, conv))

    sup_row = conn.execute("SELECT name FROM suppliers WHERE id = ?", (po['supplier_id'],)).fetchone()
    sup_name = sup_row['name'] if sup_row else 'Supplier'

    add_notification_with_conn(conn, 'po', f"PO {next_po_id} created for supplier replacement", f"{ret_pkg_qty} {pkg_unit}(s) of {med['brand']} returned to {sup_name} · For Replacement", '#d97706')
    log_activity(conn, 'po_return_incident', 'purchase_order', po_id, f"PO {po_id}", details=f"Returned {ret_pkg_qty} {pkg_unit}(s) of {med['brand']} to {sup_name} due to incident ({reason}). Created replacement PO {next_po_id}")

    conn.commit()
    conn.close()
    flash(f"Successfully processed return: {ret_pkg_qty} {pkg_unit}(s) ({ret_base_qty} {base_unit}s) of {med['brand']} deducted from stock. Replacement Purchase Order {next_po_id} created under 'For Replacement'.", 'success')
    return redirect(url_for('purchase_orders'))

def parse_sale_timestamp(s_dict):
    """Returns a datetime object representing when the sale took place."""
    candidates = [
        s_dict.get('created_at'),
        s_dict.get('movement_date'),
        s_dict.get('sale_date'),
    ]
    formats = [
        '%Y-%m-%d %H:%M:%S',
        '%Y-%m-%d %I:%M:%S %p',
        '%Y-%m-%d %I:%M %p',
        '%Y-%m-%d %H:%M',
        '%Y-%m-%dT%H:%M:%S',
        '%m/%d/%Y %I:%M:%S %p',
        '%m/%d/%Y %I:%M %p',
        '%m/%d/%Y %H:%M:%S',
        '%m/%d/%Y',
        '%Y-%m-%d',
        '%d/%m/%Y'
    ]
    for cand in candidates:
        if not cand:
            continue
        cand_str = str(cand).strip()
        for fmt in formats:
            try:
                return datetime.strptime(cand_str, fmt)
            except (ValueError, TypeError):
                continue
    return None

def check_sale_edit_eligibility(s_dict, now=None):
    """
    Checks if transaction is within the 24-hour edit window.
    After 24 hours (1 day), users cannot edit the transaction anymore.
    """
    if now is None:
        now = get_current_ph_time().replace(tzinfo=None)
    dt = parse_sale_timestamp(s_dict)
    if not dt:
        return False, "Locked"
    diff_sec = (now - dt).total_seconds()
    # Within 24 hours (86,400 seconds) - with 60s tolerance for clock skew
    if diff_sec <= 86400 and diff_sec >= -60:
        remaining_sec = max(0, 86400 - diff_sec)
        rem_hrs = int(remaining_sec // 3600)
        rem_mins = int((remaining_sec % 3600) // 60)
        if rem_hrs > 0:
            rem_str = f"{rem_hrs}h {rem_mins}m left to edit"
        elif rem_mins > 0:
            rem_str = f"{rem_mins}m left to edit"
        else:
            rem_str = "< 1m left to edit"
        return True, rem_str
    else:
        return False, "Locked (> 24 hrs)"

@app.route('/sales', methods=['GET', 'POST'])
def sales():
    if not has_permission('sales_view'):
        flash('Access denied. You do not have permission to access Transactions.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    if request.method == 'POST':
        if not has_permission('sales_create'):
            conn.close()
            flash('Access denied. You do not have permission to process sales transactions.', 'warning')
            return redirect(url_for('sales'))
        try:
            # Support multiple items in a single transaction checkout
            medicine_ids = request.form.getlist('medicine_id[]')
            if not medicine_ids:
                medicine_ids = request.form.getlist('medicine_id')
            if not medicine_ids and 'medicine_id' in request.form:
                medicine_ids = [request.form['medicine_id']]
                
            qtys_raw = request.form.getlist('qty[]')
            if not qtys_raw:
                qtys_raw = request.form.getlist('qty')
            if not qtys_raw and 'qty' in request.form:
                qtys_raw = [request.form['qty']]
                
            unit_types_raw = request.form.getlist('sale_unit_type[]')
            if not unit_types_raw:
                unit_types_raw = request.form.getlist('sale_unit_type')
            if not unit_types_raw and 'sale_unit_type' in request.form:
                unit_types_raw = [request.form['sale_unit_type']]

            items_to_process = []
            for idx in range(len(medicine_ids)):
                m_id = (medicine_ids[idx] or '').strip()
                if not m_id:
                    continue
                try:
                    q_val = int(qtys_raw[idx]) if idx < len(qtys_raw) else 1
                except (ValueError, TypeError):
                    q_val = 1
                q_val = min(MAX_STOCK_LIMIT, max(1, q_val))
                u_type = (unit_types_raw[idx] or 'base').strip() if idx < len(unit_types_raw) else 'base'
                items_to_process.append({
                    'medicine_id': m_id,
                    'qty': q_val,
                    'sale_unit_type': u_type
                })

            if not items_to_process:
                conn.close()
                flash("Please select at least one item to record.", 'warning')
                return redirect(url_for('sales'))

            # Pre-validate inventory and cumulative required stock
            needed_stock_per_med = {}
            med_records = {}

            for it in items_to_process:
                m_id = it['medicine_id']
                if m_id not in med_records:
                    med = conn.execute("SELECT * FROM inventory WHERE id = ?", (m_id,)).fetchone()
                    if not med:
                        conn.close()
                        flash(f"Product (ID: {m_id}) not found in inventory!", 'warning')
                        return redirect(url_for('sales'))
                    med_records[m_id] = med
                    needed_stock_per_med[m_id] = 0

                med = med_records[m_id]
                conv = med['conversion_qty'] or 1
                if it['sale_unit_type'] == 'package':
                    needed_base = it['qty'] * conv
                else:
                    needed_base = it['qty']
                needed_stock_per_med[m_id] += needed_base

            # Check each against available stock
            for m_id, total_needed in needed_stock_per_med.items():
                med = med_records[m_id]
                base_u = med['base_unit'] or med['unit_of_measure'] or 'Piece'
                if med['stock'] < total_needed:
                    conn.close()
                    flash(f"Insufficient stock for {med['brand']}! Total required is {total_needed:,} {base_u}s, but only {med['stock']:,} {base_u}s available in inventory.", 'warning')
                    return redirect(url_for('sales'))

            # Get maximum existing numeric IDs to guarantee sequential uniqueness
            sales_rows = conn.execute("SELECT id FROM sales").fetchall()
            max_sale_num = 0
            for r in sales_rows:
                try:
                    parts = r['id'].split('-')
                    if len(parts) == 2 and parts[1].isdigit():
                        val = int(parts[1])
                        if val > max_sale_num:
                            max_sale_num = val
                except Exception:
                    pass

            trn_rows = conn.execute("SELECT id FROM stock_movements").fetchall()
            max_trn_num = 0
            for r in trn_rows:
                try:
                    parts = r['id'].split('-')
                    if len(parts) == 2 and parts[1].isdigit():
                        val = int(parts[1])
                        if val > max_trn_num:
                            max_trn_num = val
                except Exception:
                    pass

            created_sales = []
            sale_date = datetime.now().strftime('%m/%d/%Y')
            sold_by = session.get('name', 'Assistant Pharmacist')

            for it in items_to_process:
                max_sale_num += 1
                next_sale_id = f"TRN-{max_sale_num:03d}"

                m_id = it['medicine_id']
                qty = it['qty']
                sale_unit_type = it['sale_unit_type']
                med = med_records[m_id]

                conv = med['conversion_qty'] or 1
                base_unit = med['base_unit'] or med['unit_of_measure'] or 'Piece'
                pkg_unit = med['packaging_unit'] or 'Box'

                if sale_unit_type == 'package':
                    deduct_base_qty = qty * conv
                    unit_name = pkg_unit
                    sale_display = f"{qty} {pkg_unit}{'s' if qty > 1 else ''} ({deduct_base_qty:,} {base_unit}{'s' if deduct_base_qty > 1 else ''})" if conv > 1 else f"{qty} {pkg_unit}{'s' if qty > 1 else ''}"
                else:
                    deduct_base_qty = qty
                    unit_name = base_unit
                    sale_display = f"{qty} {base_unit}{'s' if qty > 1 else ''}"

                remaining_to_sell = deduct_base_qty
                batches = conn.execute("SELECT * FROM batches WHERE medicine_id = ? AND current_qty > 0", (m_id,)).fetchall()

                # Sort in Python by parsing expiry date (FIFO)
                def parse_date(b):
                    try:
                        return datetime.strptime(b['expiry_date'], '%m/%d/%Y')
                    except Exception:
                        return datetime.max
                batches = sorted(batches, key=parse_date)

                for b in batches:
                    if remaining_to_sell <= 0:
                        break
                    sell_qty = min(remaining_to_sell, b['current_qty'])

                    # Update batch
                    conn.execute("UPDATE batches SET current_qty = current_qty - ? WHERE id = ?", (sell_qty, b['id']))
                    # Update inventory stock
                    conn.execute("UPDATE inventory SET stock = stock - ? WHERE id = ?", (sell_qty, m_id))

                    # Recalculate status using low_stock_threshold
                    med_item = conn.execute("SELECT stock, reorder_point FROM inventory WHERE id = ?", (m_id,)).fetchone()
                    if med_item:
                        new_status = compute_stock_status(med_item['stock'], get_low_stock_threshold(conn), med_item['reorder_point'])
                        conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, m_id))

                    # Log stock movement
                    max_trn_num += 1
                    next_trn_id = f"TRN-{max_trn_num:03d}"
                    trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')
                    conn.execute("INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                 (next_trn_id, 'Stock-Out', m_id, b['id'], sell_qty, trn_date, f'Sale {next_sale_id} ({sale_display})'))

                    remaining_to_sell -= sell_qty

                if remaining_to_sell > 0:
                    conn.execute("UPDATE inventory SET stock = stock - ? WHERE id = ?", (remaining_to_sell, m_id))
                    max_trn_num += 1
                    next_trn_id = f"TRN-{max_trn_num:03d}"
                    trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')
                    conn.execute("INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
                                 (next_trn_id, 'Stock-Out', m_id, 'N/A', remaining_to_sell, trn_date, f'Sale {next_sale_id} ({sale_display})'))
                    med_item = conn.execute("SELECT stock, reorder_point FROM inventory WHERE id = ?", (m_id,)).fetchone()
                    if med_item:
                        new_status = compute_stock_status(med_item['stock'], get_low_stock_threshold(conn), med_item['reorder_point'])
                        conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, m_id))

                # Insert Sale with unit, base_qty and created_at timestamp
                created_at_now = get_current_ph_time().strftime('%Y-%m-%d %I:%M %p')
                conn.execute("INSERT INTO sales (id, medicine_id, sale_date, qty, unit, base_qty, sold_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                             (next_sale_id, m_id, sale_date, qty, unit_name, deduct_base_qty, sold_by, created_at_now))
                log_activity(conn, 'sale', 'product', m_id, med['brand'], details=f"Sold {sale_display} (Receipt {next_sale_id})")

                created_sales.append({
                    'id': next_sale_id,
                    'brand': med['brand'],
                    'display': sale_display
                })

            # Check updated stock levels for alerts
            for m_id in med_records.keys():
                updated_med = conn.execute("SELECT brand, stock, reorder_point FROM inventory WHERE id = ?", (m_id,)).fetchone()
                if updated_med:
                    low_limit = get_low_stock_threshold(conn)
                    if updated_med['stock'] == 0:
                        add_notification_with_conn(conn, 'low_stock', f"{updated_med['brand']} is Out of Stock", "Remaining stock is 0 units", '#ef4444')
                    elif updated_med['stock'] <= low_limit:
                        add_notification_with_conn(conn, 'low_stock', f"{updated_med['brand']} is Low Stock", f"Only {updated_med['stock']} units left · Low Stock Threshold: {low_limit} (Reorder point: {updated_med['reorder_point']})", '#ef4444')

            conn.commit()
            if len(created_sales) == 1:
                flash(f"Transaction recorded successfully ({created_sales[0]['id']}) - Sold {created_sales[0]['display']}!", 'success')
            else:
                ids_str = ", ".join([s['id'] for s in created_sales])
                flash(f"{len(created_sales)} transactions recorded successfully ({ids_str})!", 'success')
        except Exception as e:
            conn.rollback()
            flash(f"Failed to record transaction: {str(e)}", 'danger')
        finally:
            conn.close()
        return redirect(url_for('sales'))
        
    medicines_raw = conn.execute("SELECT * FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL ORDER BY brand ASC").fetchall()
    medicines = []
    for m in medicines_raw:
        m_dict = dict(m)
        spec = get_product_spec_tag(m_dict.get('brand'), m_dict.get('product_name'), m_dict.get('description'))
        m_dict['spec_tag'] = spec
        m_dict['display_name'] = f"{m_dict['brand']}{' - ' + m_dict['product_name'] if m_dict.get('product_name') else ''}{spec}"
        medicines.append(m_dict)

    # Ensure created_at column exists in sales table
    try:
        conn.execute("ALTER TABLE sales ADD COLUMN created_at TEXT DEFAULT NULL")
        conn.commit()
    except Exception:
        pass

    sales_history_raw = conn.execute('''
        SELECT s.id, s.medicine_id, s.sale_date, i.brand as medicine_name, s.qty, s.unit, s.base_qty, s.sold_by,
               i.base_unit, i.packaging_unit, i.conversion_qty, i.product_name, i.description,
               COALESCE(s.created_at, sm.movement_date) as created_at
        FROM sales s
        JOIN inventory i ON s.medicine_id = i.id
        LEFT JOIN (
            SELECT reference, MIN(movement_date) as movement_date
            FROM stock_movements
            GROUP BY reference
        ) sm ON sm.reference LIKE 'Sale ' || s.id || '%'
        ORDER BY s.id DESC
    ''').fetchall()

    def pluralize_unit_name(q, u):
        if not u:
            return 'Piece' if q == 1 else 'Pieces'
        u_clean = u.strip()
        if q == 1:
            return u_clean
        if u_clean.endswith(('s', 'S')):
            return u_clean
        if u_clean.endswith(('x', 'X', 'ch', 'sh', 'CH', 'SH')):
            return f"{u_clean}es"
        return f"{u_clean}s"

    sales_history = []
    now_ph = get_current_ph_time().replace(tzinfo=None)
    for s in sales_history_raw:
        s_dict = dict(s)
        spec = get_product_spec_tag(s_dict.get('medicine_name'), s_dict.get('product_name'), s_dict.get('description'))
        s_dict['spec_tag'] = spec
        s_dict['display_name'] = f"{s_dict['medicine_name']}{' - ' + s_dict['product_name'] if s_dict.get('product_name') else ''}{spec}"
        s_dict['clean_name'] = clean_product_name_only(s_dict.get('medicine_name'), s_dict.get('product_name'))
        s_dict['spec_detail'] = extract_spec_detail(s_dict.get('medicine_name'), s_dict.get('product_name'), s_dict.get('description'))
        
        base_u = s_dict.get('base_unit') or 'Piece'
        pkg_u = s_dict.get('packaging_unit') or 'Box'
        raw_u = (s_dict.get('unit') or '').strip()
        conv = s_dict.get('conversion_qty') or 1
        qty_val = s_dict.get('qty') or 0
        base_qty = s_dict.get('base_qty') or qty_val
        
        # Determine if customer bought a full package or individual units
        is_pkg = False
        if raw_u and raw_u.lower() in (pkg_u.lower(), 'package') and pkg_u.lower() != base_u.lower():
            is_pkg = True
            eff_unit = pkg_u
        elif raw_u and raw_u.lower() not in ('base', 'individual', base_u.lower()) and conv > 1 and raw_u.lower() == pkg_u.lower():
            is_pkg = True
            eff_unit = pkg_u
        else:
            is_pkg = False
            eff_unit = base_u

        s_dict['effective_unit'] = eff_unit
        s_dict['is_package_sale'] = is_pkg
        
        if is_pkg and conv > 1:
            s_dict['formatted_qty'] = f"{qty_val:,} {pluralize_unit_name(qty_val, pkg_u)}"
            s_dict['formatted_subtext'] = f"({base_qty:,} {pluralize_unit_name(base_qty, base_u)})"
        else:
            s_dict['formatted_qty'] = f"{qty_val:,} {pluralize_unit_name(qty_val, base_u)}"
            s_dict['formatted_subtext'] = ""

        # Enforce 24-hour edit window
        can_edit, time_remaining_str = check_sale_edit_eligibility(s_dict, now_ph)
        s_dict['can_edit'] = can_edit
        s_dict['edit_time_remaining'] = time_remaining_str

        sales_history.append(s_dict)
    conn.close()
    return render_template('sales.html', medicines=medicines, sales_history=sales_history, active_page='sales')

@app.route('/sales/edit', methods=['POST'])
def edit_sale():
    if not has_permission('sales_edit'):
        flash('Access denied. You do not have permission to edit sales records.', 'warning')
        return redirect(url_for('sales'))
    conn = get_db_connection()
    sale_id = request.form['id']
    medicine_id = request.form['medicine_id']
    new_qty = int(request.form['qty'])
    sale_unit_type = request.form.get('sale_unit_type', 'base').strip()
    
    # Fetch old sale details with movement date fallback
    old_sale = conn.execute("""
        SELECT s.*, COALESCE(s.created_at, sm.movement_date) as effective_created_at
        FROM sales s
        LEFT JOIN (
            SELECT reference, MIN(movement_date) as movement_date
            FROM stock_movements
            GROUP BY reference
        ) sm ON sm.reference LIKE 'Sale ' || s.id || '%'
        WHERE s.id = ?
    """, (sale_id,)).fetchone()
    if not old_sale:
        conn.close()
        flash("Sale record not found!", 'danger')
        return redirect(url_for('sales'))

    old_dict = dict(old_sale)
    if old_dict.get('effective_created_at'):
        old_dict['created_at'] = old_dict['effective_created_at']
    can_edit, _ = check_sale_edit_eligibility(old_dict)
    if not can_edit:
        conn.close()
        flash("This transaction cannot be edited anymore. Transactions can only be edited within 24 hours of recording.", "warning")
        return redirect(url_for('sales'))
        
    old_med_id = old_sale['medicine_id']
    old_base_qty = old_sale['base_qty'] if (old_sale['base_qty'] and old_sale['base_qty'] > 0) else old_sale['qty']
    
    med = conn.execute("SELECT * FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
    if not med:
        conn.close()
        flash("Medicine not found!")
        return redirect(url_for('sales'))
        
    conv = med['conversion_qty'] or 1
    base_unit = med['base_unit'] or med['unit_of_measure'] or 'Piece'
    pkg_unit = med['packaging_unit'] or 'Box'
    
    if sale_unit_type == 'package':
        new_base_qty = new_qty * conv
        unit_name = pkg_unit
    else:
        new_base_qty = new_qty
        unit_name = base_unit

    low_thresh = get_low_stock_threshold(conn)

    # If medicine hasn't changed:
    if old_med_id == medicine_id:
        diff_base = new_base_qty - old_base_qty
        if med['stock'] >= diff_base:
            new_stock = med['stock'] - diff_base
            new_status = compute_stock_status(new_stock, low_thresh, med['reorder_point'])
            conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_stock, new_status, medicine_id))
            conn.execute("UPDATE sales SET qty = ?, unit = ?, base_qty = ? WHERE id = ?", (new_qty, unit_name, new_base_qty, sale_id))
            conn.execute("UPDATE stock_movements SET quantity = ? WHERE reference LIKE ?", (new_base_qty, f"Sale {sale_id}%"))
            log_activity(conn, 'edit_sale', 'sale', sale_id, f"Sale {sale_id}", details=f"Adjusted quantity to {new_qty} {unit_name} ({new_base_qty} {base_unit}s) for {med['brand']}")
            conn.commit()
            flash("Sale record updated successfully!")
        else:
            flash(f"Insufficient stock for the requested quantity change! Need {diff_base:,} more {base_unit}s, but only {med['stock']:,} available.")
    else:
        # Medicine changed: return old stock to old medicine
        old_med = conn.execute("SELECT stock, reorder_point, brand FROM inventory WHERE id = ?", (old_med_id,)).fetchone()
        if old_med:
            old_new_stock = old_med['stock'] + old_base_qty
            old_new_status = compute_stock_status(old_new_stock, low_thresh, old_med['reorder_point'])
            conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (old_new_stock, old_new_status, old_med_id))
            
        # Deduct new stock from new medicine
        if med['stock'] >= new_base_qty:
            new_new_stock = med['stock'] - new_base_qty
            new_new_status = compute_stock_status(new_new_stock, low_thresh, med['reorder_point'])
            conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_new_stock, new_new_status, medicine_id))
            
            conn.execute("UPDATE sales SET medicine_id = ?, qty = ?, unit = ?, base_qty = ? WHERE id = ?", (medicine_id, new_qty, unit_name, new_base_qty, sale_id))
            conn.execute("UPDATE stock_movements SET medicine_id = ?, quantity = ? WHERE reference LIKE ?", (medicine_id, new_base_qty, f"Sale {sale_id}%"))
            old_med_name = old_med['brand'] if old_med else old_med_id
            log_activity(conn, 'edit_sale', 'sale', sale_id, f"Sale {sale_id}", details=f"Changed product from {old_med_name} (Qty: {old_base_qty}) to {med['brand']} (Qty: {new_qty} {unit_name} = {new_base_qty} {base_unit}s)")
            conn.commit()
            flash("Sale record updated successfully!")
        else:
            conn.rollback()
            flash(f"Insufficient stock for the new medicine! Required: {new_base_qty:,} {base_unit}s.")
    conn.close()
    return redirect(url_for('sales'))
            
    conn.close()
    return redirect(url_for('sales'))

@app.route('/stock_movements')
def stock_movements():
    if not has_permission('movements_view'):
        flash('Access denied. You do not have permission to access Stock Movements.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    items = conn.execute('''
        SELECT sm.id, sm.type, i.brand as medicine, sm.batch_id as batch, ABS(sm.quantity) as quantity, sm.movement_date as date, sm.reference, i.category
        FROM stock_movements sm
        JOIN inventory i ON sm.medicine_id = i.id
        ORDER BY sm.id DESC
    ''').fetchall()
    conn.close()
    return render_template('stock_movements.html', items=items, active_page='stock_movements')

@app.route('/expiry_monitoring')
def expiry_monitoring():
    if not has_permission('expiry_view'):
        flash('Access denied. You do not have permission to access Expiry Monitoring.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()

    # Migrate: ensure disposals table exists without full reset
    conn.execute('''CREATE TABLE IF NOT EXISTS disposals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id TEXT NOT NULL, medicine_id TEXT NOT NULL,
        medicine_name TEXT NOT NULL, qty_disposed INTEGER NOT NULL,
        reason TEXT DEFAULT 'Expired', disposed_by TEXT NOT NULL,
        disposed_at TEXT NOT NULL, notes TEXT DEFAULT ''
    )''')

    # Ensure expiry_pending column exists
    try:
        conn.execute("ALTER TABLE batches ADD COLUMN expiry_pending INTEGER DEFAULT 0")
        conn.commit()
    except Exception:
        pass

    batches = conn.execute('''
        SELECT b.id as batch_id, i.id as medicine_id, i.brand as medicine,
               b.expiry_date, b.current_qty, b.status,
               COALESCE(b.expiry_pending, 0) as expiry_pending
        FROM batches b 
        JOIN inventory i ON b.medicine_id = i.id
        WHERE b.current_qty > 0
        ORDER BY b.id ASC
    ''').fetchall()

    # Fetch all disposed batch IDs for quick lookup
    disposed_ids = set(
        r['batch_id'] for r in conn.execute('SELECT batch_id FROM disposals').fetchall()
    )

    # Fetch near expiry warning days from settings
    near_expiry_setting = conn.execute("SELECT value FROM settings WHERE key = 'near_expiry_warning'").fetchone()
    try:
        near_expiry_days = int(near_expiry_setting['value']) if near_expiry_setting else 90
    except (ValueError, TypeError):
        near_expiry_days = 90

    # Fetch disposal log (latest 50)
    disposal_log = conn.execute('''
        SELECT * FROM disposals ORDER BY id DESC LIMIT 50
    ''').fetchall()

    conn.close()

    list_batches = []
    for b in batches:
        is_pending = bool(b['expiry_pending']) or not b['expiry_date'] or b['expiry_date'].strip() == ''

        if is_pending:
            days_left = None
            status    = 'Pending'
        else:
            try:
                expiry_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
                days_left = (expiry_dt - datetime.now()).days
                if days_left <= 0:
                    days_left = 0
                    status = 'Expired'
                elif days_left <= near_expiry_days:
                    status = 'Near Expiry'
                else:
                    status = 'Good'
            except Exception:
                days_left = None
                status    = 'Pending'

        list_batches.append({
            'medicine':       b['medicine'],
            'medicine_id':    b['medicine_id'],
            'batch_id':       b['batch_id'],
            'expiry_date':    b['expiry_date'] if b['expiry_date'] else '—',
            'days_left':      f"{days_left} days" if days_left is not None else '—',
            'current_qty':    b['current_qty'],
            'status':         status,
            'disposed':       b['batch_id'] in disposed_ids,
            'expiry_pending': is_pending
        })

    return render_template('expiry_monitoring.html',
        items=list_batches,
        disposal_log=[dict(r) for r in disposal_log],
        active_page='expiry_monitoring'
    )

@app.route('/expiry/dispose/<batch_id>', methods=['POST'])
def dispose_batch(batch_id):
    if not has_permission('expiry_dispose'):
        flash('Access denied. You do not have permission to dispose expired stock.', 'warning')
        return redirect(url_for('expiry_monitoring'))
    conn = get_db_connection()
    notes = request.form.get('notes', '').strip()

    batch = conn.execute('''
        SELECT b.id, b.current_qty, b.medicine_id, i.brand
        FROM batches b JOIN inventory i ON b.medicine_id = i.id
        WHERE b.id = ?
    ''', (batch_id,)).fetchone()

    if not batch:
        conn.close()
        flash('Batch not found.')
        return redirect(url_for('expiry_monitoring'))

    qty = batch['current_qty']
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    by = session.get('name', 'System')

    # Record disposal
    conn.execute('''
        INSERT INTO disposals (batch_id, medicine_id, medicine_name, qty_disposed, reason, disposed_by, disposed_at, notes)
        VALUES (?, ?, ?, ?, 'Expired', ?, ?, ?)
    ''', (batch_id, batch['medicine_id'], batch['brand'], qty, by, now_str, notes))

    # Zero out batch quantity
    conn.execute('UPDATE batches SET current_qty = 0, status = ? WHERE id = ?', ('Disposed', batch_id))

    # Deduct from inventory stock
    med = conn.execute('SELECT stock, reorder_point FROM inventory WHERE id = ?', (batch['medicine_id'],)).fetchone()
    if med:
        new_stock = max(0, med['stock'] - qty)
        new_status = compute_stock_status(new_stock, get_low_stock_threshold(conn), med['reorder_point'])
        conn.execute('UPDATE inventory SET stock = ?, status = ? WHERE id = ?', (new_stock, new_status, batch['medicine_id']))

    # Log to stock movements
    trn_id = f"DIS-{batch_id}"
    conn.execute('''
        INSERT OR IGNORE INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference)
        VALUES (?, 'Disposal', ?, ?, ?, ?, ?)
    ''', (trn_id, batch['medicine_id'], batch_id, -qty, now_str, f'Disposal of expired batch {batch_id}'))

    # Activity log
    log_activity(conn, 'dispose', 'batch', batch_id, batch['brand'], details=f"Disposed {qty} expired units. Reason: Expired. Notes: {notes}")

    conn.commit()
    conn.close()
    flash(f"Batch {batch_id} ({batch['brand']}) — {qty} units disposed and recorded.")
    return redirect(url_for('expiry_monitoring'))


@app.route('/users', methods=['GET', 'POST'])
def users():
    if session.get('role') == 'Staff':
        return redirect(url_for('dashboard'))
    
    conn = get_db_connection()
    if request.method == 'POST':
        user_db_id = request.form.get('id')
        name = request.form.get('name', '')
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        role = request.form.get('role', '')
        contact = request.form.get('contact', '')
        email = request.form.get('email', '')
        license_no = request.form.get('license_no', '')
        joined_date = request.form.get('joined_date', '')
        address = request.form.get('address', '')
        
        # Validate Username: 15 chars max, alphanumeric letters & numbers only, no special chars
        if len(username) > 15 or not re.match(r'^[a-zA-Z0-9]{1,15}$', username):
            conn.close()
            flash('Username must contain only letters and numbers, maximum 15 characters (no special characters allowed).', 'danger')
            return redirect(url_for('users'))

        # Validate Contact Number: Must start with 09 and have exactly 11 digits
        clean_contact = re.sub(r'\D', '', contact)
        if clean_contact and (len(clean_contact) != 11 or not clean_contact.startswith('09')):
            conn.close()
            flash('Contact number must be exactly 11 digits starting with 09 (e.g. 0917-123-4567 or 09171234567).', 'danger')
            return redirect(url_for('users'))
        elif clean_contact:
            # Normalize to 09XX-XXX-XXXX standard format for consistency
            contact = f"{clean_contact[0:4]}-{clean_contact[4:7]}-{clean_contact[7:11]}"

        if user_db_id:
            # Check duplicate username or email on another user
            existing_user = conn.execute("""
                SELECT id, username, email FROM users 
                WHERE (LOWER(TRIM(username)) = LOWER(TRIM(?)) OR (email != '' AND LOWER(TRIM(email)) = LOWER(TRIM(?))))
                  AND id != ?
                LIMIT 1
            """, (username, email, user_db_id)).fetchone()
            if existing_user:
                conn.close()
                flash(f"Duplicate user information! Username '{username}' or email '{email}' is already taken by another account.", "danger")
                return redirect(url_for('users'))

            old_user = conn.execute("SELECT * FROM users WHERE id = ?", (user_db_id,)).fetchone()
            user_changes = []
            if old_user:
                if old_user['name'] != name:
                    user_changes.append(f"Name: '{old_user['name']}' -> '{name}'")
                if old_user['username'] != username:
                    user_changes.append(f"Username: '{old_user['username']}' -> '{username}'")
                if old_user['role'] != role:
                    user_changes.append(f"Role: '{old_user['role']}' -> '{role}'")
                if (old_user['contact'] or '') != contact:
                    user_changes.append(f"Contact: '{old_user['contact'] or ''}' -> '{contact}'")
                if (old_user['email'] or '') != email:
                    user_changes.append(f"Email: '{old_user['email'] or ''}' -> '{email}'")

            if password and password not in ('••••••••', '••••••••••'):
                if len(password) != 10 or not re.match(r'^[a-zA-Z0-9]{10}$', password):
                    conn.close()
                    flash('Password must be exactly 10 alphanumeric characters (letters and numbers only, no special characters allowed).', 'danger')
                    return redirect(url_for('users'))
                user_changes.append("Password updated")
                conn.execute('''
                    UPDATE users 
                    SET name = ?, username = ?, password = ?, role = ?, contact = ?, email = ?, license_no = ?, joined_date = ?, address = ?
                    WHERE id = ?
                ''', (name, username, password, role, contact, email, license_no, joined_date, address, user_db_id))
            else:
                conn.execute('''
                    UPDATE users 
                    SET name = ?, username = ?, role = ?, contact = ?, email = ?, license_no = ?, joined_date = ?, address = ?
                    WHERE id = ?
                ''', (name, username, role, contact, email, license_no, joined_date, address, user_db_id))
            details_str = f"Updated: {', '.join(user_changes)}" if user_changes else f"Updated user account {username} (Role: {role})"
            log_activity(conn, 'edit_user', 'User', user_db_id, name, details=details_str)
            conn.commit()
            
            # Update session details if editing own profile
            if int(user_db_id) == session.get('user_id'):
                session['name'] = name
                session['role'] = role
            flash('User details updated successfully!')
        else:
            # Check duplicate username or email
            existing_user = conn.execute("""
                SELECT id, username, email FROM users 
                WHERE LOWER(TRIM(username)) = LOWER(TRIM(?)) OR (email != '' AND LOWER(TRIM(email)) = LOWER(TRIM(?)))
                LIMIT 1
            """, (username, email)).fetchone()
            if existing_user:
                conn.close()
                flash(f"Duplicate user information! Username '{username}' or email '{email}' is already taken.", "danger")
                return redirect(url_for('users'))

            # Validate Password for new user: exactly 10 alphanumeric chars
            if len(password) != 10 or not re.match(r'^[a-zA-Z0-9]{10}$', password):
                conn.close()
                flash('Password must be exactly 10 alphanumeric characters (letters and numbers only, no special characters allowed).', 'danger')
                return redirect(url_for('users'))

            # Check maximum user limit (max 5 users)
            current_user_count = conn.execute('SELECT COUNT(*) FROM users').fetchone()[0]
            if current_user_count >= 5:
                conn.close()
                flash('Maximum user limit reached (5 users max). Cannot create additional accounts.', 'danger')
                return redirect(url_for('users'))

            # Generate next user ID (e.g. USR-003)
            last_user = conn.execute("SELECT user_id FROM users ORDER BY id DESC LIMIT 1").fetchone()
            if last_user:
                try:
                    last_num = int(last_user['user_id'].split('-')[1])
                    next_user_id = f"USR-{last_num + 1:03d}"
                except (IndexError, ValueError):
                    next_user_id = "USR-003"
            else:
                next_user_id = "USR-001"
                
            form_perms = request.form.getlist('permissions')
            if role == 'Staff':
                initial_perms = json.dumps(form_perms if form_perms else DEFAULT_STAFF_PERMISSIONS)
            else:
                initial_perms = None

            conn.execute('''
                INSERT INTO users (user_id, name, username, password, role, contact, email, license_no, joined_date, address, permissions)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (next_user_id, name, username, password, role, contact, email, license_no, joined_date, address, initial_perms))
            log_activity(conn, 'add_user', 'User', next_user_id, name, details=f"Created user account {username} (Role: {role})")
            conn.commit()
            flash('New user added successfully!')
            
        conn.close()
        return redirect(url_for('users'))
        
    # Auto-purge deactivated users whose deactivated_at >= 30 days
    now_ph = get_current_ph_time().replace(tzinfo=None)
    try:
        trashed_users = conn.execute("SELECT id, username, name, role, deactivated_at FROM users WHERE is_active = 0 AND deactivated_at IS NOT NULL").fetchall()
        for tu in trashed_users:
            try:
                deact_dt = datetime.strptime(tu['deactivated_at'], '%Y-%m-%d %I:%M %p')
                if (now_ph - deact_dt).days >= 30:
                    log_activity(conn, 'delete_user', 'User', tu['id'], tu['name'], performed_by='System (Auto-Purge)', details=f"Permanently purged user account {tu['username']} after 30 days of deactivation")
                    conn.execute("DELETE FROM users WHERE id = ?", (tu['id'],))
            except Exception:
                pass
        conn.commit()
    except Exception as e:
        print("User auto-purge error:", e)

    raw_users = conn.execute('SELECT * FROM users').fetchall()
    users_list = []
    online_count = 0
    deactivated_count = 0
    all_perm_keys = [p['key'] for cat in PERMISSION_CATALOG for p in cat['permissions']]
    for u in raw_users:
        u_dict = dict(u)
        is_computed_online = False
        is_active = u_dict.get('is_active', 1)
        if is_active == 0:
            deactivated_count += 1
            deact_str = u_dict.get('deactivated_at')
            if deact_str:
                deact_dt = parse_dt_safe(deact_str)
                days_left = max(0, 30 - (now_ph - deact_dt).days) if deact_dt else 30
            else:
                days_left = 30
            u_dict['days_remaining'] = days_left
            is_computed_online = False
        else:
            is_computed_online = compute_user_online_status(u_dict, now_ph)
            if is_computed_online:
                online_count += 1
        u_dict['computed_online'] = is_computed_online

        # Process role-based and custom staff / admin permissions
        user_role = u_dict.get('role', '')
        if user_role in ['Superadmin', 'Owner / Pharmacist'] or 'Super' in user_role or 'Owner' in user_role:
            effective_perms = all_perm_keys
        elif user_role in ['Admin', 'Administrator']:
            perms_raw = u_dict.get('permissions')
            if perms_raw is None or perms_raw == '*':
                effective_perms = list(all_perm_keys)
            else:
                try:
                    parsed_perms = json.loads(perms_raw) if isinstance(perms_raw, str) else (perms_raw or [])
                    effective_perms = parsed_perms if isinstance(parsed_perms, list) else list(all_perm_keys)
                except Exception:
                    effective_perms = list(all_perm_keys)
        else:
            # Staff
            perms_raw = u_dict.get('permissions')
            if perms_raw is None or perms_raw == '*':
                effective_perms = list(DEFAULT_STAFF_PERMISSIONS)
            else:
                try:
                    parsed_perms = json.loads(perms_raw) if isinstance(perms_raw, str) else (perms_raw or [])
                    effective_perms = parsed_perms if isinstance(parsed_perms, list) else list(DEFAULT_STAFF_PERMISSIONS)
                except Exception:
                    effective_perms = list(DEFAULT_STAFF_PERMISSIONS)

        u_dict['effective_perms'] = effective_perms
        u_dict['perm_count'] = len(effective_perms)
        u_dict['perms_json'] = json.dumps(effective_perms)

        users_list.append(u_dict)
    
    total_users = len(users_list)
    active_users = total_users - deactivated_count
    offline_count = max(0, active_users - online_count)
    conn.close()
    return render_template('users.html', users=users_list, online_count=online_count, offline_count=offline_count, total_users=total_users, deactivated_count=deactivated_count, active_page='users', all_perm_keys=all_perm_keys)

@app.route('/users/<int:user_id>/toggle_status', methods=['POST'])
@app.route('/users/<int:user_id>/deactivate', methods=['POST'])
def toggle_user_status(user_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    if session.get('role') == 'Staff':
        flash('Permission denied. Only administrators can change user status.', 'danger')
        return redirect(url_for('users'))

    conn = get_db_connection()
    user = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
    
    if not user:
        conn.close()
        flash('User not found.', 'danger')
        return redirect(url_for('users'))

    # Prevent deactivating superadmin / owner
    if user['role'] in ['Superadmin', 'Owner / Pharmacist'] or 'Super' in user['role'] or 'Owner' in user['role']:
        conn.close()
        flash('The Superadmin account cannot be deactivated.', 'danger')
        return redirect(url_for('users'))

    # Prevent deactivating self
    if user['id'] == session.get('user_id'):
        conn.close()
        flash('You cannot deactivate your own account.', 'danger')
        return redirect(url_for('users'))

    new_status = 0 if user['is_active'] == 1 else 1
    action_text = "deactivated" if new_status == 0 else "reactivated"
    action_type = "deactivate_user" if new_status == 0 else "activate_user"

    now_str = get_current_ph_time().strftime('%Y-%m-%d %I:%M %p')
    if new_status == 0:
        conn.execute("UPDATE users SET is_active = 0, is_online = 0, last_logout = ?, deactivated_at = ? WHERE id = ?", (now_str, now_str, user_id))
        detail_msg = f"Deactivated user account {user['username']} (30-day auto-purge retention scheduled)"
    else:
        conn.execute("UPDATE users SET is_active = 1, deactivated_at = NULL WHERE id = ?", (user_id,))
        detail_msg = f"Reactivated user account {user['username']} (Role: {user['role']})"

    log_activity(conn, action_type, 'User', user['id'], user['name'], details=detail_msg)
    conn.commit()
    conn.close()

    flash(f"User '{user['name']}' has been {action_text} successfully!", 'success')
    return redirect(url_for('users'))

@app.route('/users/<int:user_id>/permissions', methods=['POST'])
def update_user_permissions(user_id):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    if session.get('role') == 'Staff':
        flash('Permission denied. Only administrators can configure staff permissions.', 'danger')
        return redirect(url_for('users'))
    
    conn = get_db_connection()
    user = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
    if not user:
        conn.close()
        flash('User not found.', 'danger')
        return redirect(url_for('users'))
        
    # Superadmin accounts always possess full system access and cannot be restricted
    if user['role'] in ['Superadmin', 'Owner / Pharmacist'] or 'Super' in user['role'] or 'Owner' in user['role']:
        conn.close()
        flash(f"{user['role']} accounts automatically have full access to all system modules and cannot be restricted.", 'info')
        return redirect(url_for('users'))

    # If current editor is an Admin (not Superadmin), they cannot configure another Admin
    current_role = session.get('role', '')
    is_super = current_role in ['Superadmin', 'Owner / Pharmacist'] or 'Super' in current_role or 'Owner' in current_role
    if not is_super and (user['role'] in ['Admin', 'Administrator']):
        conn.close()
        flash('Only Superadmin can configure permissions for Admin accounts.', 'warning')
        return redirect(url_for('users'))

    selected_perms = request.form.getlist('permissions')
    perms_json = json.dumps(selected_perms)

    conn.execute('UPDATE users SET permissions = ? WHERE id = ?', (perms_json, user_id))
    log_activity(conn, 'update_permissions', 'User', user_id, user['name'], details=f"Updated access permissions for {user['role']} {user['username']} ({len(selected_perms)} actions granted)")
    conn.commit()

    # If updating currently logged in user, refresh their session permissions
    if session.get('user_id') == user_id:
        session['permissions'] = perms_json

    conn.close()
    flash(f"Permissions for '{user['name']}' updated successfully ({len(selected_perms)} actions enabled)!", 'success')
    return redirect(url_for('users'))

@app.route('/settings', methods=['GET', 'POST'])
def settings():
    if not has_permission('settings_view'):
        flash('Access denied. You do not have permission to access Settings.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    
    if request.method == 'POST':
        if not has_permission('settings_edit'):
            conn.close()
            flash('Access denied. You do not have permission to modify system settings.', 'warning')
            return redirect(url_for('settings'))
        # Validation: Low Stock Threshold cannot be greater than the Reorder Point
        if 'low_stock_threshold' in request.form and session.get('role') != 'Staff':
            try:
                candidate_threshold = int(request.form.get('low_stock_threshold', 40))
            except (ValueError, TypeError):
                candidate_threshold = 40
            
            min_rop_row = conn.execute("SELECT MIN(reorder_point) FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL").fetchone()
            current_rop = min_rop_row[0] if (min_rop_row and min_rop_row[0] is not None) else 100
            
            if candidate_threshold > current_rop:
                conn.close()
                flash(f"Cannot save settings: Low Stock Threshold ({candidate_threshold} units) cannot be greater than the Reorder Point ({current_rop} units).", "error")
                return redirect(url_for('settings'))

        if session.get('role') == 'Staff':
            keys = [
                'dark_mode', 'font_size', 'high_contrast', 'reduce_motion', 
                'screen_reader', 'show_generic', 'low_stock_alerts', 
                'expiry_alerts', 'po_alerts', 'sales_notifications', 
                'auto_logout', 'require_password'
            ]
        else:
            keys = [
                'pharmacy_name', 'address', 'contact', 'email', 
                'dark_mode', 'font_size', 'high_contrast', 'reduce_motion', 
                'screen_reader', 'show_generic', 'low_stock_alerts', 
                'expiry_alerts', 'po_alerts', 'sales_notifications', 
                'low_stock_threshold', 'near_expiry_warning', 
                'default_currency', 'auto_logout', 'require_password'
            ]
        for key in keys:
            if key in ['dark_mode', 'high_contrast', 'reduce_motion', 'screen_reader', 'show_generic', 
                       'low_stock_alerts', 'expiry_alerts', 'po_alerts', 'sales_notifications', 'require_password']:
                val = 'true' if key in request.form else 'false'
            elif key == 'low_stock_threshold':
                try:
                    val = str(max(1, int(request.form.get(key, 40))))
                except (ValueError, TypeError):
                    val = '40'
            elif key == 'near_expiry_warning':
                try:
                    val = str(max(1, int(request.form.get(key, 90))))
                except (ValueError, TypeError):
                    val = '90'
            else:
                val = request.form.get(key, '')
            conn.execute('UPDATE settings SET value = ? WHERE key = ?', (val, key))
            
        # If low_stock_threshold is updated, recalculate inventory status without altering individual reorder points
        if 'low_stock_threshold' in request.form and session.get('role') != 'Staff':
            try:
                new_threshold = max(1, int(request.form.get('low_stock_threshold', 40)))
                conn.execute('''
                    UPDATE inventory 
                    SET status = CASE 
                        WHEN stock = 0 THEN 'No Stock' 
                        WHEN stock <= ? THEN 'Low Stock' 
                        ELSE 'Good' 
                    END
                ''', (new_threshold,))
            except Exception as e:
                print("Error updating inventory status for low_stock_threshold:", e)

        # If near_expiry_warning is updated, recalculate batch status for active non-disposed batches
        if 'near_expiry_warning' in request.form and session.get('role') != 'Staff':
            try:
                new_near_expiry = max(1, int(request.form.get('near_expiry_warning', 90)))
                active_batches = conn.execute("SELECT id, expiry_date, expiry_pending FROM batches WHERE status != 'Disposed' AND expiry_date IS NOT NULL AND expiry_date != ''").fetchall()
                for ab in active_batches:
                    b_status = compute_batch_expiry_status(ab['expiry_date'], new_near_expiry, bool(ab['expiry_pending']))
                    conn.execute("UPDATE batches SET status = ? WHERE id = ?", (b_status, ab['id']))
            except Exception as e:
                print("Error updating batches status for near_expiry_warning:", e)

        log_activity(conn, 'update_settings', 'System', 'settings', 'Pharmacy Settings', details='Updated system preferences and parameters')
        conn.commit()
        conn.close()
        flash('Settings saved successfully!')
        return redirect(url_for('settings'))
        
    settings_raw = conn.execute('SELECT * FROM settings').fetchall()
    min_rop_row = conn.execute("SELECT MIN(reorder_point) FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL").fetchone()
    current_reorder_point = min_rop_row[0] if (min_rop_row and min_rop_row[0] is not None) else 100
    conn.close()
    
    settings_dict = {r['key']: r['value'] for r in settings_raw}
    
    # Calculate automated cron backup statistics & restore points (prunes snapshots older than 30 days)
    cleanup_expired_backups(30)
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')
    auto_backups_count = 0
    recent_snapshots = []
    if os.path.exists(backup_dir):
        files = [f for f in os.listdir(backup_dir) if f.startswith('farmacia_auto_backup_') and f.endswith('.db')]
        auto_backups_count = len(files)
        files.sort(key=lambda x: os.path.getmtime(os.path.join(backup_dir, x)), reverse=True)
        for f in files[:10]:
            fp = os.path.join(backup_dir, f)
            try:
                mtime = datetime.fromtimestamp(os.path.getmtime(fp))
                size_kb = round(os.path.getsize(fp) / 1024, 1)
                recent_snapshots.append({
                    'filename': f,
                    'display_time': mtime.strftime('%b %d, %Y · %I:%M %p'),
                    'size': f"{size_kb} KB"
                })
            except Exception:
                pass
        
    last_auto_backup_time = settings_dict.get('last_auto_backup_time', 'Scheduled (Runs every 6 hours via Cron)')
    last_auto_backup_file = settings_dict.get('last_auto_backup_file', '')
    
    return render_template('settings.html', settings=settings_dict, active_page='settings',
                           current_reorder_point=current_reorder_point,
                           auto_backups_count=auto_backups_count,
                           last_auto_backup_time=last_auto_backup_time,
                           last_auto_backup_file=last_auto_backup_file,
                           recent_snapshots=recent_snapshots)

@app.route('/settings/change-password', methods=['POST'])
def change_password():
    if 'user_id' not in session:
        return redirect(url_for('login'))
        
    current_password = request.form.get('current_password', '').strip()
    new_password = request.form.get('new_password', '').strip()
    confirm_password = request.form.get('confirm_password', '').strip()
    
    if not current_password or not new_password or not confirm_password:
        flash('All password fields are required.', 'error')
        return redirect(url_for('settings'))
        
    if new_password != confirm_password:
        flash('New password and confirmation do not match.', 'error')
        return redirect(url_for('settings'))
        
    if len(new_password) != 10 or not re.match(r'^[a-zA-Z0-9]{10}$', new_password):
        flash('Password must be exactly 10 alphanumeric characters (letters and numbers only, no special characters allowed).', 'error')
        return redirect(url_for('settings'))
        
    if new_password == current_password:
        flash('New password cannot be the same as your current password.', 'error')
        return redirect(url_for('settings'))
        
    conn = get_db_connection()
    user = conn.execute('SELECT * FROM users WHERE id = ?', (session['user_id'],)).fetchone()
    if not user:
        conn.close()
        session.clear()
        flash('User account not found. Please log in again.', 'error')
        return redirect(url_for('login'))
        
    if user['password'] != current_password:
        conn.close()
        flash('Current password is incorrect.', 'error')
        return redirect(url_for('settings'))
        
    conn.execute('UPDATE users SET password = ? WHERE id = ?', (new_password, user['id']))
    log_activity(conn, 'change_password', 'User', user['id'], user['name'], performed_by=user['name'], details=f"Changed password for account {user['username']}")
    conn.commit()
    conn.close()
    
    flash('Password changed successfully!', 'success')
    return redirect(url_for('settings'))

@app.route('/notifications/mark-seen', methods=['POST'])
def mark_notifications_seen():
    conn = get_db_connection()
    conn.execute('UPDATE notifications SET seen = 1 WHERE seen = 0')
    conn.commit()
    conn.close()
    return {'status': 'success'}

@app.route('/reports')
def reports():
    if not has_permission('reports_view'):
        flash('Access denied. You do not have permission to access Reports & Analytics.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    now = get_current_ph_time()

    # ─────────────────────────────────────────────────────────────────────
    # DSS CHART 1: Critical Reorder & Safety Buffer Gauge
    # Items at or near reorder point — shows stock deficit vs. ROP
    # ─────────────────────────────────────────────────────────────────────
    reorder_raw = conn.execute("""
        SELECT brand, stock, reorder_point
        FROM inventory
        WHERE archived_at IS NULL AND deleted_at IS NULL
          AND stock <= reorder_point * 1.3
        ORDER BY (reorder_point - stock) DESC
        LIMIT 10
    """).fetchall()
    reorder_labels   = [r['brand'] for r in reorder_raw]
    reorder_stock    = [r['stock'] for r in reorder_raw]
    reorder_deficit  = [max(0, r['reorder_point'] - r['stock']) for r in reorder_raw]
    reorder_rop      = [r['reorder_point'] for r in reorder_raw]
    reorder_items    = [{
        'brand': r['brand'],
        'stock': r['stock'],
        'rop': r['reorder_point'],
        'deficit': max(0, r['reorder_point'] - r['stock']),
        'urgency': 'Critical Shortage' if r['stock'] <= r['reorder_point'] * 0.5 else 'Below Safety Level'
    } for r in reorder_raw]

    # ─────────────────────────────────────────────────────────────────────
    # DSS CHART 2: Expiry Horizon / FEFO Timeline (30 / 60 / 90 days)
    # ─────────────────────────────────────────────────────────────────────
    expiry_batches = conn.execute("""
        SELECT i.brand, b.expiry_date, b.current_qty
        FROM batches b
        JOIN inventory i ON b.medicine_id = i.id
        WHERE b.status NOT IN ('Disposed','Expired')
          AND b.current_qty > 0
    """).fetchall()

    fefo_30_items = []
    fefo_60_items = []
    fefo_90_items = []
    fefo_30_qty = fefo_60_qty = fefo_90_qty = 0

    for b in expiry_batches:
        try:
            exp_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
            days_left = (exp_dt - now.replace(tzinfo=None)).days
            qty = b['current_qty'] or 0
            name = b['brand']
            item_info = {'brand': name, 'qty': qty, 'days': days_left, 'date': b['expiry_date']}
            if 0 <= days_left <= 30:
                fefo_30_items.append(item_info)
                fefo_30_qty += qty
            elif 31 <= days_left <= 60:
                fefo_60_items.append(item_info)
                fefo_60_qty += qty
            elif 61 <= days_left <= 90:
                fefo_90_items.append(item_info)
                fefo_90_qty += qty
        except Exception:
            pass

    # Sort each bucket by closest expiry
    fefo_30_items.sort(key=lambda x: x['days'])
    fefo_60_items.sort(key=lambda x: x['days'])
    fefo_90_items.sort(key=lambda x: x['days'])

    fefo_data = {
        'counts':  [len(fefo_30_items), len(fefo_60_items), len(fefo_90_items)],
        'qtys':    [fefo_30_qty, fefo_60_qty, fefo_90_qty],
        'items30': fefo_30_items[:6],
        'items60': fefo_60_items[:6],
        'items90': fefo_90_items[:6],
    }

    # ─────────────────────────────────────────────────────────────────────
    # DSS CHART 3: Fast-Moving vs. Slow-Moving Matrix (Velocity vs. Stock)
    # Sales velocity (qty sold last 90 days) vs. current stock level
    # ─────────────────────────────────────────────────────────────────────
    velocity_raw = conn.execute("""
        SELECT i.id, i.brand, i.stock, i.category,
               COALESCE(SUM(s.qty), 0) as units_sold
        FROM inventory i
        LEFT JOIN sales s ON s.medicine_id = i.id
        WHERE i.archived_at IS NULL AND i.deleted_at IS NULL
        GROUP BY i.id
        ORDER BY units_sold DESC
        LIMIT 30
    """).fetchall()

    velocity_data = []
    for r in velocity_raw:
        velocity_data.append({
            'label': r['brand'],
            'x': r['units_sold'],
            'y': r['stock'],
            'category': r['category'],
        })

    # Compute medians for quadrant lines
    all_vel = sorted([d['x'] for d in velocity_data])
    all_stk = sorted([d['y'] for d in velocity_data])
    med_vel = all_vel[len(all_vel)//2] if all_vel else 0
    med_stk = all_stk[len(all_stk)//2] if all_stk else 0

    # Ensure High Sales requires at least 1 unit sold
    vel_threshold = max(1, med_vel) if any(d['x'] > 0 for d in velocity_data) else 0

    # Categorize items into 4 distinct quadrants for immediate clarity
    dead_stock_items = [d for d in velocity_data if d['x'] < vel_threshold and d['y'] >= med_stk][:5]
    shortage_risk_items = [d for d in velocity_data if d['x'] >= vel_threshold and d['y'] < med_stk][:5]
    fast_healthy_items = [d for d in velocity_data if d['x'] >= vel_threshold and d['y'] >= med_stk][:5]
    slow_low_items = [d for d in velocity_data if d['x'] < vel_threshold and d['y'] < med_stk][:5]

    # ─────────────────────────────────────────────────────────────────────
    # DSS CHART 4: ABC / VEN Inventory Value Matrix
    # ABC = cumulative financial value; VEN = clinical importance by category
    # ─────────────────────────────────────────────────────────────────────
    ven_vital      = ['Medicines', 'First Aid', 'Medical Supplies', 'Medical Devices']
    ven_essential  = ['Vitamins & Supplements']
    ven_nonessential = ['Personal Care', 'Skin Care', 'Baby Care']

    abc_raw = conn.execute("""
        SELECT i.brand, i.category, i.purchase_price,
               COALESCE(SUM(s.qty), 0) as units_sold
        FROM inventory i
        LEFT JOIN sales s ON s.medicine_id = i.id
        WHERE i.archived_at IS NULL AND i.deleted_at IS NULL
        GROUP BY i.id
    """).fetchall()

    items_with_value = []
    for r in abc_raw:
        val = (r['units_sold'] or 0) * (r['purchase_price'] or 0)
        if val == 0:
            val = (r['purchase_price'] or 0) * 1
        cat = r['category'] or ''
        ven = 'Vital' if cat in ven_vital else ('Essential' if cat in ven_essential else 'Non-Essential')
        items_with_value.append({'brand': r['brand'], 'cat': cat, 'ven': ven, 'val': val})

    total_val = sum(i['val'] for i in items_with_value) or 1
    items_with_value.sort(key=lambda x: x['val'], reverse=True)
    cumulative = 0
    for item in items_with_value:
        cumulative += item['val']
        pct = cumulative / total_val
        item['abc'] = 'A' if pct <= 0.70 else ('B' if pct <= 0.90 else 'C')

    abc_ven_matrix = {'A': {'Vital': 0,'Essential': 0,'Non-Essential': 0},
                      'B': {'Vital': 0,'Essential': 0,'Non-Essential': 0},
                      'C': {'Vital': 0,'Essential': 0,'Non-Essential': 0}}
    for item in items_with_value:
        abc_ven_matrix[item['abc']][item['ven']] += 1

    abc_vital      = [abc_ven_matrix[c]['Vital']        for c in ['A','B','C']]
    abc_essential  = [abc_ven_matrix[c]['Essential']     for c in ['A','B','C']]
    abc_nonessent  = [abc_ven_matrix[c]['Non-Essential'] for c in ['A','B','C']]

    # Specific key products in Class A (Vital & Essential)
    top_av_items = [i for i in items_with_value if i['abc'] == 'A' and i['ven'] == 'Vital'][:4]
    top_ae_items = [i for i in items_with_value if i['abc'] == 'A' and i['ven'] == 'Essential'][:3]

    # ─────────────────────────────────────────────────────────────────────
    # DSS CHART 5: Demand Forecast vs. Actual Dispensing
    # Continuous 6-month timeline + 3-month moving average forecast
    # ─────────────────────────────────────────────────────────────────────
    def month_keys(n):
        keys = []
        y, m = now.year, now.month
        for _ in range(n):
            keys.insert(0, f"{y}-{m:02d}")
            m -= 1
            if m == 0:
                m = 12; y -= 1
        return keys

    forecast_months = month_keys(6)

    sales_by_month_raw = conn.execute("""
        SELECT sale_date, qty FROM sales
    """).fetchall()

    month_sales_map = defaultdict(int)
    for s in sales_by_month_raw:
        try:
            date_str = str(s['sale_date']).strip()
            if '/' in date_str:
                parts = date_str.split('/')
                m_key = f"{parts[2].zfill(4)}-{parts[0].zfill(2)}"
            else:
                m_key = date_str[:7]
            month_sales_map[m_key] += s['qty']
        except Exception:
            pass

    actual_dispensing = [month_sales_map.get(mk, 0) for mk in forecast_months]

    forecast_line = []
    for i in range(len(actual_dispensing)):
        if i < 3:
            forecast_line.append(None)
        else:
            avg = sum(actual_dispensing[i-3:i]) / 3
            forecast_line.append(round(avg, 1))

    if len(actual_dispensing) >= 3:
        next_forecast = round(sum(actual_dispensing[-3:]) / 3, 1)
    else:
        next_forecast = 0

    y, m = now.year, now.month
    m += 1
    if m > 12:
        m = 1; y += 1
    next_month_key = f"{y}-{m:02d}"
    forecast_months_ext = forecast_months + [next_month_key]
    actual_dispensing_ext = actual_dispensing + [None]
    forecast_line_ext = forecast_line + [next_forecast]

    # Top dispensed products
    top_dispensed_raw = conn.execute("""
        SELECT i.brand, SUM(s.qty) as total_sold
        FROM sales s
        JOIN inventory i ON s.medicine_id = i.id
        GROUP BY s.medicine_id
        ORDER BY total_sold DESC
        LIMIT 4
    """).fetchall()
    top_dispensed_items = [{'brand': r['brand'], 'qty': r['total_sold']} for r in top_dispensed_raw]

    # ─────────────────────────────────────────────────────────────────────
    # DSS CHART 6: Supplier Fulfillment & Delivery Reliability
    # ─────────────────────────────────────────────────────────────────────
    supplier_po_raw = conn.execute("""
        SELECT s.name as supplier_name,
               COUNT(po.id) as total_po,
               SUM(CASE WHEN po.status IN ('Received','Partial') THEN 1 ELSE 0 END) as fulfilled_po
        FROM purchase_orders po
        JOIN suppliers s ON po.supplier_id = s.id
        GROUP BY s.id
        ORDER BY total_po DESC
        LIMIT 8
    """).fetchall()

    supplier_labels      = [r['supplier_name'] for r in supplier_po_raw]
    supplier_total_po    = [r['total_po'] for r in supplier_po_raw]
    supplier_fulfilled   = [r['fulfilled_po'] for r in supplier_po_raw]
    supplier_pending     = [max(0, r['total_po'] - r['fulfilled_po']) for r in supplier_po_raw]
    supplier_fulfill_pct = [
        round(r['fulfilled_po'] / r['total_po'] * 100, 1) if r['total_po'] > 0 else 0
        for r in supplier_po_raw
    ]

    conn.close()

    return render_template('reports.html',
                           active_page='reports',
                           # DSS Chart 1: Reorder Buffer
                           reorder_labels=reorder_labels,
                           reorder_stock=reorder_stock,
                           reorder_deficit=reorder_deficit,
                           reorder_rop=reorder_rop,
                           reorder_items=reorder_items,
                           # DSS Chart 2: FEFO Expiry Horizon
                           fefo_data=fefo_data,
                           # DSS Chart 3: Fast/Slow Velocity Matrix
                           velocity_data=velocity_data,
                           med_vel=vel_threshold,
                           med_stk=med_stk,
                           dead_stock_items=dead_stock_items,
                           shortage_risk_items=shortage_risk_items,
                           fast_healthy_items=fast_healthy_items,
                           slow_low_items=slow_low_items,
                           # DSS Chart 4: ABC/VEN Matrix
                           abc_vital=abc_vital,
                           abc_essential=abc_essential,
                           abc_nonessent=abc_nonessent,
                           top_av_items=top_av_items,
                           top_ae_items=top_ae_items,
                           # DSS Chart 5: Demand Forecast
                           forecast_months=forecast_months_ext,
                           actual_dispensing=actual_dispensing_ext,
                           forecast_line=forecast_line_ext,
                           next_forecast=next_forecast,
                           top_dispensed_items=top_dispensed_items,
                           # DSS Chart 6: Supplier Fulfillment
                           supplier_labels=supplier_labels,
                           supplier_total_po=supplier_total_po,
                           supplier_fulfilled=supplier_fulfilled,
                           supplier_pending=supplier_pending,
                           supplier_fulfill_pct=supplier_fulfill_pct,
                           )

class PDFTemplateReport(FPDF):
    def __init__(self, report_title, period_text=""):
        super().__init__()
        self.report_title = report_title
        self.period_text = period_text

    def header(self):
        # 1. Top Red Header Banner
        self.set_fill_color(192, 57, 43)  # Deep Red #c0392b
        self.rect(0, 0, 210, 24, 'F')

        logo_path = os.path.join(app.static_folder or 'static', 'images', 'logo_about_us.png')
        if not os.path.exists(logo_path):
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static', 'images', 'logo_about_us.png')

        rendered_logo = False
        if os.path.exists(logo_path):
            try:
                self.image(logo_path, x=75, y=3, w=60)
                rendered_logo = True
            except Exception:
                rendered_logo = False

        if not rendered_logo:
            self.set_font('helvetica', 'B', 18)
            self.set_text_color(255, 255, 255)
            self.set_xy(0, 6)
            self.cell(210, 10, 'FARMACIA ni DOK', align='C')

        # 2. Sub-header Dark Ribbon
        self.set_fill_color(17, 17, 17)  # Dark Ribbon
        self.rect(0, 24, 210, 7, 'F')
        self.set_font('helvetica', '', 7.5)
        self.set_text_color(255, 255, 255)
        self.set_xy(0, 24)
        self.cell(210, 7, 'Branch : La Residencia, Pio Cruzcrosa, Calumpit Bulacan | farmacianidok@gmail.com | 0917-000-0000', align='C')

        # Reset text color and positioning for page content
        self.set_text_color(0, 0, 0)
        self.set_y(35)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')

    @staticmethod
    def sanitize(text):
        if text is None:
            return ""
        s = str(text)
        s = s.replace('\u2264', '<=').replace('\u2265', '>=').replace('\u2013', '-').replace('\u2014', '--')
        s = s.replace('\u2018', "'").replace('\u2019', "'").replace('\u201c', '"').replace('\u201d', '"')
        s = s.replace('\u2022', '*').replace('\u2026', '...')
        return s

    def draw_kpi_cards(self, cards):
        """Draw 2 to 4 responsive summary KPI cards horizontally"""
        if not cards:
            return
        n = len(cards)
        total_w = 190.0
        gap = 4.0
        card_w = (total_w - (n - 1) * gap) / n
        start_y = self.get_y()
        card_h = 20.0

        for idx, card in enumerate(cards):
            cx = 10.0 + idx * (card_w + gap)
            bg_color = card.get('bg', (248, 250, 252))
            border_color = card.get('border', (226, 232, 240))
            accent_color = card.get('color', (192, 57, 43))
            
            # Card background & border
            self.set_fill_color(*bg_color)
            self.set_draw_color(*border_color)
            self.rect(cx, start_y, card_w, card_h, 'DF')
            
            # Top accent stripe
            self.set_fill_color(*accent_color)
            self.rect(cx, start_y, card_w, 2.0, 'F')
            
            # Value
            self.set_xy(cx + 2, start_y + 3)
            self.set_font('helvetica', 'B', 11)
            self.set_text_color(*accent_color)
            self.cell(card_w - 4, 7, self.sanitize(card.get('val', '')), align='C')
            
            # Label
            self.set_xy(cx + 2, start_y + 11)
            self.set_font('helvetica', '', 7.5)
            self.set_text_color(100, 116, 139)
            self.cell(card_w - 4, 6, self.sanitize(card.get('label', '')), align='C')

        self.set_y(start_y + card_h + 5)
        self.set_draw_color(0, 0, 0)
        self.set_text_color(0, 0, 0)

    def draw_bar_chart(self, title, items, max_val=None, bar_color=(59, 130, 246)):
        """
        Draw a clean horizontal bar chart with labels, percentage widths, and value labels.
        items: list of dicts [{'label': 'Name', 'val': 45, 'desc': 'extra'}]
        """
        if not items:
            return
        if self.get_y() > 210:
            self.add_page()

        self.set_font('helvetica', 'B', 10)
        self.set_text_color(30, 41, 59)
        self.cell(0, 6, title, ln=True)
        self.ln(2)

        if max_val is None:
            max_val = max([it['val'] for it in items] or [1])
            if max_val == 0:
                max_val = 1

        chart_x = 10.0
        label_w = 60.0
        bar_max_w = 100.0
        bar_h = 5.5
        row_h = 7.5

        for it in items:
            if self.get_y() > 270:
                self.add_page()
            
            cy = self.get_y()
            val = it.get('val', 0)
            lbl = self.sanitize(str(it.get('label', ''))[:32])
            val_str = self.sanitize(str(it.get('desc') or val))

            # Label (Right-aligned or left-aligned)
            self.set_xy(chart_x, cy)
            self.set_font('helvetica', '', 8)
            self.set_text_color(51, 65, 85)
            self.cell(label_w, bar_h, lbl, align='L')

            # Bar track background
            bx = chart_x + label_w + 2
            self.set_fill_color(241, 245, 249)
            self.rect(bx, cy + 0.8, bar_max_w, bar_h - 1.6, 'F')

            # Bar fill
            fill_w = max(1.5, (val / max_val) * bar_max_w) if val > 0 else 0
            if fill_w > 0:
                self.set_fill_color(*bar_color)
                self.rect(bx, cy + 0.8, fill_w, bar_h - 1.6, 'F')

            # Value label
            self.set_xy(bx + bar_max_w + 3, cy)
            self.set_font('helvetica', 'B', 7.5)
            self.set_text_color(71, 85, 105)
            self.cell(24, bar_h, val_str, align='L')

            self.ln(row_h)

        self.ln(3)
        self.set_text_color(0, 0, 0)

    def draw_quadrant_grid(self, dead_cnt, shortage_cnt, fast_cnt, slow_cnt):
        """Draw a visual 4-quadrant Velocity Matrix diagram"""
        if self.get_y() > 210:
            self.add_page()
            
        self.set_font('helvetica', 'B', 10)
        self.set_text_color(30, 41, 59)
        self.cell(0, 6, "VELOCITY MATRIX - 4 QUADRANT VISUAL DISTRIBUTION", ln=True)
        self.ln(2)

        qx = 10.0
        qy = self.get_y()
        box_w = 93.0
        box_h = 24.0

        quads = [
            # Top-Left: Shortage Risk (High Sales, Low Stock)
            {'x': qx, 'y': qy, 'title': 'SHORTAGE RISK (High Velocity / Low Stock)', 'cnt': shortage_cnt,
             'action': 'Urgent Restock Needed - At risk of stockout', 'bg': (254, 243, 199), 'bdr': (245, 158, 11), 'txt': (180, 83, 9)},
            # Top-Right: Fast & Healthy (High Sales, High Stock)
            {'x': qx + box_w + 4, 'y': qy, 'title': 'FAST & HEALTHY (High Velocity / Good Stock)', 'cnt': fast_cnt,
             'action': 'Optimal Stocking - Core pharmacy revenue drivers', 'bg': (209, 250, 229), 'bdr': (16, 185, 129), 'txt': (4, 120, 87)},
            # Bottom-Left: Slow & Low (Low Sales, Low Stock)
            {'x': qx, 'y': qy + box_h + 3, 'title': 'SLOW & LOW (Low Velocity / Low Stock)', 'cnt': slow_cnt,
             'action': 'Low Priority - Reorder only on-demand or minimums', 'bg': (241, 245, 249), 'bdr': (148, 163, 184), 'txt': (71, 85, 105)},
            # Bottom-Right: Dead Stock (Low Sales, High Stock)
            {'x': qx + box_w + 4, 'y': qy + box_h + 3, 'title': 'DEAD STOCK (Low Velocity / High Stock)', 'cnt': dead_cnt,
             'action': 'Capital Tied Up - Run bundle deals, discounts, returns', 'bg': (254, 226, 226), 'bdr': (239, 68, 68), 'txt': (185, 28, 28)},
        ]

        for q in quads:
            self.set_fill_color(*q['bg'])
            self.set_draw_color(*q['bdr'])
            self.rect(q['x'], q['y'], box_w, box_h, 'DF')
            
            # Title & Count
            self.set_xy(q['x'] + 3, q['y'] + 2)
            self.set_font('helvetica', 'B', 8)
            self.set_text_color(*q['txt'])
            self.cell(box_w - 20, 5, q['title'], align='L')
            self.set_xy(q['x'] + box_w - 18, q['y'] + 2)
            self.cell(15, 5, f"{q['cnt']} items", align='R')

            # Action guideline
            self.set_xy(q['x'] + 3, q['y'] + 8)
            self.set_font('helvetica', 'I', 7.5)
            self.set_text_color(51, 65, 85)
            self.multi_cell(box_w - 6, 4.5, q['action'])

        self.set_y(qy + (box_h * 2) + 10)
        self.set_draw_color(0, 0, 0)
        self.set_text_color(0, 0, 0)

    def draw_abc_ven_summary_grid(self, matrix_counts):
        """Draw 3x3 ABC/VEN Decision Matrix with counts and priority categorization"""
        if self.get_y() > 200:
            self.add_page()
            
        self.set_font('helvetica', 'B', 10)
        self.set_text_color(30, 41, 59)
        self.cell(0, 6, "ABC / VEN CROSS-CLASSIFICATION MATRIX & STRATEGIC PRIORITIES", ln=True)
        self.ln(2)

        start_x = 10.0
        start_y = self.get_y()
        col_w = 40.0
        row_h = 13.0

        # Matrix Headers
        self.set_fill_color(30, 41, 59)
        self.set_text_color(255, 255, 255)
        self.set_font('helvetica', 'B', 8)
        self.rect(start_x, start_y, 30.0, 7.0, 'F')
        self.set_xy(start_x, start_y)
        self.cell(30.0, 7.0, "VEN / ABC", align='C')

        ven_cols = ['Class A (70% Val)', 'Class B (20% Val)', 'Class C (10% Val)']
        for i, hc in enumerate(ven_cols):
            hx = start_x + 30.0 + i * col_w
            self.rect(hx, start_y, col_w, 7.0, 'F')
            self.set_xy(hx, start_y)
            self.cell(col_w, 7.0, hc, align='C')

        rows_def = [
            ('Vital (V)', 'V', [
                ('AV', 'Cat I (Highest)', (254, 226, 226), (185, 28, 28)),
                ('BV', 'Cat I (Highest)', (254, 226, 226), (185, 28, 28)),
                ('CV', 'Cat II (Moderate)', (254, 243, 199), (180, 83, 9)),
            ]),
            ('Essential (E)', 'E', [
                ('AE', 'Cat I (Highest)', (254, 226, 226), (185, 28, 28)),
                ('BE', 'Cat II (Moderate)', (254, 243, 199), (180, 83, 9)),
                ('CE', 'Cat II (Moderate)', (254, 243, 199), (180, 83, 9)),
            ]),
            ('Non-Ess. (N)', 'N', [
                ('AN', 'Cat I (Control)', (254, 243, 199), (180, 83, 9)),
                ('BN', 'Cat II (Moderate)', (241, 245, 249), (71, 85, 105)),
                ('CN', 'Cat III (Lowest)', (241, 245, 249), (71, 85, 105)),
            ]),
        ]

        curr_y = start_y + 7.0
        for row_lbl, v_key, cols in rows_def:
            # Row header
            self.set_fill_color(51, 65, 85)
            self.set_text_color(255, 255, 255)
            self.set_font('helvetica', 'B', 8)
            self.rect(start_x, curr_y, 30.0, row_h, 'F')
            self.set_xy(start_x, curr_y)
            self.cell(30.0, row_h, row_lbl, align='C')

            for ci, (cell_code, prio, bg, txt) in enumerate(cols):
                cx = start_x + 30.0 + ci * col_w
                cnt = matrix_counts.get(cell_code, 0)
                
                self.set_fill_color(*bg)
                self.set_draw_color(203, 213, 225)
                self.rect(cx, curr_y, col_w, row_h, 'DF')

                self.set_xy(cx + 2, curr_y + 2)
                self.set_font('helvetica', 'B', 9)
                self.set_text_color(*txt)
                self.cell(col_w - 4, 4.5, f"{cell_code}: {cnt} items", align='C')

                self.set_xy(cx + 2, curr_y + 6.8)
                self.set_font('helvetica', 'I', 7)
                self.cell(col_w - 4, 4.5, prio, align='C')

            curr_y += row_h

        self.set_y(curr_y + 6)
        self.set_draw_color(0, 0, 0)
        self.set_text_color(0, 0, 0)


@app.route('/reports/download', methods=['POST'])
def download_report():
    if not has_permission('reports_download') and not has_permission('reports_view'):
        flash('Access denied. You do not have permission to download reports.', 'warning')
        return redirect(url_for('dashboard'))
    
    try:
        report_type = request.form.get('report_type', 'inventory')
        frequency = request.form.get('frequency')
        period_week = request.form.get('period_week')
        period_month = request.form.get('period_month')
        period_year = request.form.get('period_year')
        period_date = request.form.get('period_date')
        file_format = request.form.get('format', 'pdf')
        
        # Specific sub-category filters
        batch_stock_status = request.form.get('batch_stock_status', 'all')
        velocity_category = request.form.get('velocity_category', 'all')
        expiry_category = request.form.get('expiry_category', 'all')
        sort_order = request.form.get('sort_order', 'code')
        
        start_dt, end_dt = parse_date_range(frequency, period_week, period_month, period_year, period_date)
        
        conn = get_db_connection()
        
        headers = []
        rows = []
        sections = []  # List of {'title': str, 'rows': list} for categorized groupings
        title = ""
        summary_cards = []
        chart_sections = []  # List of {'title': str, 'type': 'bar'|'quadrant'|'matrix', 'data': ...}

        # Natural code sorting helper for alphanumeric IDs (e.g., MED-001, BAT-005, SUP-012, 1, 2)
        def natural_code_key(val):
            if val is None:
                return ('', 0, '')
            val_str = str(val).strip()
            match = re.search(r'\d+', val_str)
            if match:
                prefix = val_str[:match.start()]
                num = int(match.group(0))
                suffix = val_str[match.end():]
                return (prefix, num, suffix)
            return (val_str, 0, '')
        
        # Robust Python date parsing helper supporting M/D/YYYY, YYYY-MM-DD, AM/PM timestamps, etc.
        def parse_generic_date(date_str):
            if not date_str:
                return None
            date_str = str(date_str).strip()
            date_part = date_str.split(' ')[0]
            for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y', '%Y/%m/%d'):
                try:
                    return datetime.strptime(date_part, fmt).date()
                except ValueError:
                    pass
            if '/' in date_part:
                parts = date_part.split('/')
                if len(parts) == 3:
                    try:
                        return datetime(int(parts[2]), int(parts[0]), int(parts[1])).date()
                    except Exception:
                        pass
            elif '-' in date_part:
                parts = date_part.split('-')
                if len(parts) == 3:
                    try:
                        if len(parts[0]) == 4:
                            return datetime(int(parts[0]), int(parts[1]), int(parts[2])).date()
                        else:
                            return datetime(int(parts[2]), int(parts[0]), int(parts[1])).date()
                    except Exception:
                        pass
            return None

        s_date = start_dt.date() if start_dt else None
        e_date = end_dt.date() if end_dt else None

        def is_in_date_range(date_str):
            if not s_date and not e_date:
                return True
            d = parse_generic_date(date_str)
            if not d:
                return False
            if s_date and d < s_date:
                return False
            if e_date and d > e_date:
                return False
            return True

        # =========================================================================
        # 1. INVENTORY STOCK REPORT (WITH REAL-TIME STOCK LEVEL FILTER)
        # =========================================================================
        if report_type == 'inventory':
            filter_label_map = {
                'all': 'All Stock Levels',
                'low_stock': 'Low Stock Items',
                'no_stock': 'No Stock / Depleted Items',
                'good_stock': 'Good Stock Items',
                'overstock': 'Overstocked Items'
            }
            sub_title_part = filter_label_map.get(batch_stock_status, 'All Stock Levels')
            title = f"Inventory Stock Report ({sub_title_part})"
            headers = ["Item ID", "Brand Name", "Generic Name", "Category", "Unit / Dosage", "Stock", "ROP", "Supplier", "Status"]
            items = conn.execute("""
                SELECT id, brand, COALESCE(generic, '—') as generic, COALESCE(category, 'General') as category,
                       COALESCE(unit_of_measure, '—') as uom, stock, reorder_point,
                       COALESCE(supplier, '—') as supplier, status
                FROM inventory
                WHERE archived_at IS NULL AND deleted_at IS NULL
            """).fetchall()
            
            low_threshold = get_low_stock_threshold(conn)
            cat_map = defaultdict(list)
            all_inv = []
            
            for item in items:
                st = int(item['stock'] or 0)
                rop = int(item['reorder_point'] or 40)
                
                if st == 0:
                    st_val = 'No Stock (Empty)'
                    st_key = 'no_stock'
                elif st <= low_threshold:
                    st_val = 'Low Stock'
                    st_key = 'low_stock'
                elif st > rop:
                    st_val = 'Overstocked'
                    st_key = 'overstock'
                else:
                    st_val = 'Good Stock'
                    st_key = 'good_stock'
                    
                if batch_stock_status != 'all' and batch_stock_status != st_key:
                    continue
                    
                cat_val = item['category'] or 'General'
                row_data = [item['id'], item['brand'], item['generic'], cat_val, item['uom'], st, item['reorder_point'], item['supplier'], st_val]
                all_inv.append(row_data)
                cat_map[cat_val].append(row_data)
            
            all_inv.sort(key=lambda r: natural_code_key(r[0]))
            rows = all_inv
            sections = []
            summary_cards = []

        # =========================================================================
        # 2. TRANSACTIONS / SALES REPORT
        # =========================================================================
        elif report_type == 'sales':
            title = "Sales Transactions Report"
            headers = ["Sale ID", "Item ID", "Brand Name", "Category", "Qty Sold", "Sold By", "Sale Date"]
            sales_data = conn.execute("""
                SELECT s.id, s.medicine_id, i.brand as medicine, COALESCE(i.category, 'General') as category,
                       s.qty, s.sold_by, s.sale_date 
                FROM sales s 
                JOIN inventory i ON s.medicine_id = i.id
                ORDER BY s.id ASC
            """).fetchall()
            
            total_sold_qty = 0
            for item in sales_data:
                if is_in_date_range(item['sale_date']):
                    total_sold_qty += (item['qty'] or 0)
                    rows.append([item['id'], item['medicine_id'], item['medicine'], item['category'], item['qty'], item['sold_by'], item['sale_date']])
            rows.sort(key=lambda r: natural_code_key(r[0]))
            
            summary_cards = [
                {'label': 'Transactions Count', 'val': len(rows), 'color': (30, 41, 59)},
                {'label': 'Total Units Sold', 'val': total_sold_qty, 'color': (16, 185, 129)},
            ]

        # =========================================================================
        # 3. STOCK MOVEMENTS AUDIT REPORT
        # =========================================================================
        elif report_type == 'stock_movements':
            title = "Stock Movements Report"
            headers = ["Movement ID", "Type", "Medicine", "Batch ID", "Qty", "Reference / Reason", "Movement Date"]
            movements = conn.execute("""
                SELECT sm.id, sm.type, i.brand as medicine, COALESCE(sm.batch_id, '—') as batch_id,
                       ABS(sm.quantity) as quantity, COALESCE(sm.reference, '—') as reference, sm.movement_date
                FROM stock_movements sm
                JOIN inventory i ON sm.medicine_id = i.id
                ORDER BY sm.id ASC
            """).fetchall()
            
            stock_in_qty = 0
            stock_out_qty = 0
            mov_buckets = {
                'in': {'title': 'Stock In / Replenishment Influx', 'items': []},
                'out': {'title': 'Stock Out / Dispensing Outflux', 'items': []}
            }
            all_movs = []
            for item in movements:
                if is_in_date_range(item['movement_date']):
                    q = abs(item['quantity'] or 0)
                    m_type = str(item['type']).lower()
                    row_data = [item['id'], item['type'], item['medicine'], item['batch_id'], q, item['reference'], item['movement_date']]
                    all_movs.append(row_data)
                    if 'in' in m_type or 'received' in m_type or 'restock' in m_type:
                        stock_in_qty += q
                        mov_buckets['in']['items'].append(row_data)
                    else:
                        stock_out_qty += q
                        mov_buckets['out']['items'].append(row_data)
            
            for mb in mov_buckets:
                mov_buckets[mb]['items'].sort(key=lambda r: natural_code_key(r[0]))
                
            if sort_order == 'code':
                all_movs.sort(key=lambda r: natural_code_key(r[0]))
                rows = all_movs
                sections = []
            else:
                sections = []
                for mb in ['in', 'out']:
                    if mov_buckets[mb]['items']:
                        sections.append({
                            'title': mov_buckets[mb]['title'],
                            'rows': mov_buckets[mb]['items']
                        })
                rows = [r for sec in sections for r in sec['rows']] if sections else all_movs
            
            summary_cards = [
                {'label': 'Total Movement Logs', 'val': len(rows), 'color': (30, 41, 59)},
                {'label': 'Stock-In Influx', 'val': f"+{stock_in_qty}", 'color': (16, 185, 129)},
                {'label': 'Dispensed / Outflux', 'val': f"{stock_out_qty}", 'color': (220, 38, 38)},
            ]

        # =========================================================================
        # 4. EXPIRY MONITORING REPORT (WITH USER FILTER: near_expiry, expired, good, all)
        # =========================================================================
        elif report_type == 'expiry_monitoring':
            exp_label_map = {
                'all': 'All Categories',
                'near_expiry': 'Near Expiry',
                'expired': 'Expired',
                'good': 'Good Stock'
            }
            sub_title_part = exp_label_map.get(expiry_category, 'All Categories')
            title = f"Expiry Monitoring & FEFO Report ({sub_title_part})"
            headers = ["Batch ID", "Medicine", "Category", "Expiry Date", "Days Left", "Qty", "Status"]
            
            near_expiry_setting = conn.execute("SELECT value FROM settings WHERE key = 'near_expiry_warning'").fetchone()
            try:
                near_expiry_days = int(near_expiry_setting['value']) if near_expiry_setting else 90
            except (ValueError, TypeError):
                near_expiry_days = 90
                
            batches = conn.execute("""
                SELECT b.id, i.brand as medicine, COALESCE(i.category, 'General') as category,
                       b.expiry_date, b.current_qty, b.status,
                       COALESCE(b.expiry_pending, 0) as expiry_pending
                FROM batches b
                JOIN inventory i ON b.medicine_id = i.id
            """).fetchall()
            
            try:
                disposed_ids = set(r['batch_id'] for r in conn.execute("SELECT batch_id FROM disposals").fetchall())
            except Exception:
                disposed_ids = set()
                
            now_dt = datetime.now()
            near_expiry_cnt = expired_cnt = good_cnt = other_cnt = 0
            
            exp_buckets = {
                'near_expiry': {'title': 'Near Expiry Batches', 'items': []},
                'expired': {'title': 'Expired Batches', 'items': []},
                'good': {'title': 'Good Stock Batches', 'items': []},
                'other': {'title': 'Other Batches (Disposed / Pending)', 'items': []},
            }
            all_exp_rows = []
            
            for item in batches:
                if is_in_date_range(item['expiry_date']):
                    is_disposed = item['id'] in disposed_ids or item['status'] == 'Disposed'
                    is_pending = bool(item['expiry_pending']) or not item['expiry_date'] or item['expiry_date'].strip() == ''
                    
                    if is_disposed:
                        days_left_str = "Disposed"
                        stat_str = "Disposed"
                        other_cnt += 1
                        b_key = 'other'
                    elif is_pending:
                        days_left_str = "Pending"
                        stat_str = "Pending"
                        other_cnt += 1
                        b_key = 'other'
                    else:
                        try:
                            exp_dt = datetime.strptime(item['expiry_date'], '%m/%d/%Y')
                            diff = (exp_dt - now_dt).days
                            if diff <= 0:
                                days_left_str = "0 days"
                                stat_str = "Expired"
                                expired_cnt += 1
                                b_key = 'expired'
                            elif diff <= near_expiry_days:
                                days_left_str = f"{diff} days"
                                stat_str = "Near Expiry"
                                near_expiry_cnt += 1
                                b_key = 'near_expiry'
                            else:
                                days_left_str = f"{diff} days"
                                stat_str = "Good Stock"
                                good_cnt += 1
                                b_key = 'good'
                        except Exception:
                            days_left_str = "—"
                            stat_str = item['status'] or "Good Stock"
                            good_cnt += 1
                            b_key = 'good'
                            
                    row_data = [item['id'], item['medicine'], item['category'], item['expiry_date'] or 'No Date', days_left_str, item['current_qty'], stat_str]
                    if expiry_category == 'all' or b_key == expiry_category:
                        all_exp_rows.append(row_data)
                        exp_buckets[b_key]['items'].append(row_data)

            for ek in exp_buckets:
                exp_buckets[ek]['items'].sort(key=lambda r: natural_code_key(r[0]))

            if sort_order == 'code':
                all_exp_rows.sort(key=lambda r: natural_code_key(r[0]))
                rows = all_exp_rows
                sections = []
            else:
                sections = []
                for ek in ['near_expiry', 'expired', 'good', 'other']:
                    if exp_buckets[ek]['items']:
                        sections.append({
                            'title': exp_buckets[ek]['title'],
                            'rows': exp_buckets[ek]['items']
                        })
                rows = [r for sec in sections for r in sec['rows']] if sections else all_exp_rows

            summary_cards = [
                {'label': 'Total Batches in Report', 'val': len(rows), 'color': (30, 41, 59)},
                {'label': 'Near Expiry Batches', 'val': near_expiry_cnt, 'color': (220, 38, 38)},
                {'label': 'Expired Batches', 'val': expired_cnt, 'color': (156, 163, 175)},
                {'label': 'Good Stock Batches', 'val': good_cnt, 'color': (16, 185, 129)},
            ]

        # =========================================================================
        # 5. BATCH STOCKS REPORT (WITH REAL-TIME STOCK LEVEL FILTER)
        # =========================================================================
        elif report_type == 'batch_stocks':
            filter_label_map = {
                'all': 'All Stock Levels',
                'low_stock': 'Low Stock Batches',
                'no_stock': 'No Stock / Depleted Batches',
                'good_stock': 'Good Stock Batches',
                'overstock': 'Overstocked Batches'
            }
            sub_title_part = filter_label_map.get(batch_stock_status, 'All Stock Levels')
            title = f"Batch Stocks Report ({sub_title_part})"
            headers = ["Batch ID", "Item ID", "Medicine", "Category", "Expiry Date", "Qty", "Stock Status"]
            
            batches = conn.execute("""
                SELECT b.id, b.medicine_id, i.brand as medicine, COALESCE(i.category, 'General') as category,
                       i.reorder_point, b.expiry_date, b.current_qty
                FROM batches b
                JOIN inventory i ON b.medicine_id = i.id
            """).fetchall()
            
            low_threshold = get_low_stock_threshold(conn)
            all_batches_rows = []
            
            for item in batches:
                # Calendar filtering: filters to chosen specific date, month, or year (or all records if none selected)
                if is_in_date_range(item['expiry_date']):
                    c_qty = int(item['current_qty'] or 0)
                    rop = int(item['reorder_point'] or 100)
                    
                    if c_qty == 0:
                        st = 'No Stock (Empty)'
                        st_key = 'no_stock'
                    elif c_qty <= low_threshold:
                        st = 'Low Stock'
                        st_key = 'low_stock'
                    elif c_qty > rop:
                        st = 'Overstocked'
                        st_key = 'overstock'
                    else:
                        st = 'Good Stock'
                        st_key = 'good_stock'
                    
                    if batch_stock_status == 'all' or batch_stock_status == st_key:
                        all_batches_rows.append([item['id'], item['medicine_id'], item['medicine'], item['category'], item['expiry_date'] or 'N/A', c_qty, st])

            all_batches_rows.sort(key=lambda r: natural_code_key(r[0]))
            rows = all_batches_rows
            sections = []
            summary_cards = []

        # =========================================================================
        # 6. FAST VS. SLOW-MOVING PRODUCTS REPORT (VELOCITY MATRIX)
        # =========================================================================
        elif report_type == 'velocity_matrix':
            vel_label_map = {
                'all': 'All 4 Quadrant Categories',
                'dead_stock': 'Dead Stock Products',
                'shortage_risk': 'Shortage Risk Products',
                'fast_healthy': 'Fast & Healthy Products',
                'slow_low': 'Slow & Low Products'
            }
            title = f"Fast vs. Slow-Moving Velocity Report ({vel_label_map.get(velocity_category, 'All Categories')})"
            headers = ["ID", "Brand Name", "Category", "Units Sold", "Current Stock", "Quadrant Classification", "Strategic Action"]
            
            raw_items = conn.execute("""
                SELECT i.id, i.brand, i.category, i.stock,
                       COALESCE(SUM(s.qty), 0) as units_sold
                FROM inventory i
                LEFT JOIN sales s ON s.medicine_id = i.id
                WHERE i.archived_at IS NULL AND i.deleted_at IS NULL
                GROUP BY i.id
            """).fetchall()
            
            vel_list = [{'id': r['id'], 'brand': r['brand'], 'category': r['category'] or '—', 'stock': r['stock'] or 0, 'sold': r['units_sold'] or 0} for r in raw_items]
            all_solds = sorted([x['sold'] for x in vel_list])
            all_stocks = sorted([x['stock'] for x in vel_list])
            med_sold = all_solds[len(all_solds)//2] if all_solds else 0
            med_stock = all_stocks[len(all_stocks)//2] if all_stocks else 0
            vel_threshold = max(1, med_sold) if any(x['sold'] > 0 for x in vel_list) else 0

            quad_buckets = {
                'shortage_risk': {'title': 'Shortage Risk (High Velocity / Low Stock - Urgent Replenishment Needed)', 'items': []},
                'fast_healthy': {'title': 'Fast & Healthy (High Velocity / Good Stock - Maintain Optimal Stocking)', 'items': []},
                'dead_stock': {'title': 'Dead Stock (Low Velocity / High Stock - Capital Tied Up, Bundle/Promote)', 'items': []},
                'slow_low': {'title': 'Slow & Low (Low Velocity / Low Stock - Reorder Only on Demand)', 'items': []},
            }
            all_vel_rows = []
            dead_cnt = shortage_cnt = fast_cnt = slow_cnt = 0
            
            for it in vel_list:
                s = it['sold']
                stk = it['stock']
                if s >= vel_threshold and stk < med_stock:
                    quad_code = 'shortage_risk'
                    quad_name = 'Shortage Risk'
                    action_txt = 'Urgent Replenishment Needed'
                    shortage_cnt += 1
                elif s < vel_threshold and stk >= med_stock:
                    quad_code = 'dead_stock'
                    quad_name = 'Dead Stock'
                    action_txt = 'Promote, Bundle, or Vendor Return'
                    dead_cnt += 1
                elif s >= vel_threshold and stk >= med_stock:
                    quad_code = 'fast_healthy'
                    quad_name = 'Fast & Healthy'
                    action_txt = 'Maintain Optimal Inventory Cycle'
                    fast_cnt += 1
                else:
                    quad_code = 'slow_low'
                    quad_name = 'Slow & Low'
                    action_txt = 'Monitor; Reorder Only on Demand'
                    slow_cnt += 1
                
                row_data = [it['id'], it['brand'], it['category'], it['sold'], it['stock'], quad_name, action_txt]
                if velocity_category == 'all' or velocity_category == quad_code:
                    all_vel_rows.append(row_data)
                    quad_buckets[quad_code]['items'].append(row_data)

            # Sort items inside each quadrant bucket strictly by ID code chronologically (MED-001, MED-002...)
            for q_k in quad_buckets:
                quad_buckets[q_k]['items'].sort(key=lambda r: natural_code_key(r[0]))

            if sort_order == 'code':
                all_vel_rows.sort(key=lambda r: natural_code_key(r[0]))
                rows = all_vel_rows
                sections = []
            else:
                sections = []
                for q_k in ['shortage_risk', 'fast_healthy', 'dead_stock', 'slow_low']:
                    if quad_buckets[q_k]['items']:
                        sections.append({
                            'title': quad_buckets[q_k]['title'],
                            'rows': quad_buckets[q_k]['items']
                        })
                rows = [r for sec in sections for r in sec['rows']] if sections else all_vel_rows
                    
            summary_cards = [
                {'label': 'Filtered Items', 'val': len(rows), 'color': (30, 41, 59)},
                {'label': 'Shortage Risk', 'val': shortage_cnt, 'color': (245, 158, 11)},
                {'label': 'Fast & Healthy', 'val': fast_cnt, 'color': (16, 185, 129)},
                {'label': 'Dead Stock', 'val': dead_cnt, 'color': (220, 38, 38)},
                {'label': 'Slow & Low', 'val': slow_cnt, 'color': (148, 163, 184)},
            ]

        # =========================================================================
        # 7. TOP PRODUCTS BY SALES (HIGHEST SOLD AND TREND)
        # =========================================================================
        elif report_type == 'top_selling':
            freq_label = "All Time"
            if frequency == 'daily' and period_date:
                freq_label = f"Daily Date: {period_date}"
            elif frequency == 'weekly' and period_week:
                freq_label = f"Week: {period_week}"
            elif frequency == 'monthly' and period_month:
                freq_label = f"Month: {period_month}"
            elif frequency == 'annual' and period_year:
                freq_label = f"Year: {period_year}"
                
            title = f"Top Products by Sales & Trend ({freq_label})"
            headers = ["Rank", "Brand Name", "Category", "Units Sold", "Transactions", "Sales Trend"]
            
            # Fetch ALL active products from inventory so every product is displayed
            all_inventory = conn.execute("""
                SELECT id, brand, category
                FROM inventory
                WHERE archived_at IS NULL AND deleted_at IS NULL
                ORDER BY id ASC
            """).fetchall()

            prod_sales = {}
            for inv in all_inventory:
                prod_sales[inv['id']] = {
                    'id': inv['id'],
                    'brand': inv['brand'],
                    'cat': inv['category'] or '—',
                    'qty': 0,
                    'tx_cnt': 0
                }
            
            raw_sales = conn.execute("""
                SELECT s.id, s.medicine_id, s.qty, s.sale_date
                FROM sales s
                ORDER BY s.id ASC
            """).fetchall()
            
            for s in raw_sales:
                if is_in_date_range(s['sale_date']):
                    mid = s['medicine_id']
                    if mid in prod_sales:
                        prod_sales[mid]['qty'] += (s['qty'] or 0)
                        prod_sales[mid]['tx_cnt'] += 1
            
            # Sort products from highest sold and trend downwards
            sorted_prods = sorted(prod_sales.values(), key=lambda x: (x['qty'], x['tx_cnt']), reverse=True)
            
            total_sold_all = sum(x['qty'] for x in sorted_prods)
            active_sold_count = sum(1 for x in sorted_prods if x['qty'] > 0)
            
            # Calculate demand threshold for trend classification
            all_qtys = [x['qty'] for x in sorted_prods if x['qty'] > 0]
            med_qty = all_qtys[len(all_qtys)//2] if all_qtys else 1

            for rank, item in enumerate(sorted_prods, start=1):
                q = item['qty']
                if q == 0:
                    trend_str = "No Movement (0 Sold)"
                elif q >= 20 or q >= med_qty * 1.5:
                    trend_str = "High Demand (Top Seller)"
                elif q >= med_qty:
                    trend_str = "Steady Demand"
                else:
                    trend_str = "Slow-Moving"
                
                rows.append([
                    f"#{rank}",
                    item['brand'],
                    item['cat'],
                    item['qty'],
                    item['tx_cnt'],
                    trend_str
                ])
                
            # Always rank strictly in ascending chronological order of sales rank (#1, #2, #3, ...)
            rows.sort(key=lambda r: int(r[0].lstrip('#')) if str(r[0]).lstrip('#').isdigit() else 999999)
                
            summary_cards = [
                {'label': 'Total Catalog Products', 'val': len(rows), 'color': (30, 41, 59)},
                {'label': 'Products Dispensed', 'val': active_sold_count, 'color': (59, 130, 246)},
                {'label': 'Total Units Sold', 'val': total_sold_all, 'color': (16, 185, 129)},
            ]

        # =========================================================================
        # 8. SUPPLIERS REPORT
        # =========================================================================
        elif report_type == 'suppliers':
            title = "Suppliers Directory & Fulfillment Report"
            headers = ["Supplier ID", "Supplier Name", "Contact Number", "Address", "Total POs", "Fulfilled POs"]
            suppliers = conn.execute("""
                SELECT s.id, s.name, s.contact, s.address,
                       (SELECT COUNT(*) FROM purchase_orders po WHERE po.supplier_id = s.id) as po_count,
                       (SELECT COUNT(*) FROM purchase_orders po WHERE po.supplier_id = s.id AND po.status = 'Received') as fulfilled_count
                FROM suppliers s
            """).fetchall()
            for item in suppliers:
                rows.append([item['id'], item['name'], item['contact'] or '—', item['address'] or '—', item['po_count'], item['fulfilled_count']])
            rows.sort(key=lambda r: natural_code_key(r[0]))
            
            summary_cards = [
                {'label': 'Registered Suppliers', 'val': len(suppliers), 'color': (30, 41, 59)},
                {'label': 'Total Purchase Orders', 'val': sum(s['po_count'] for s in suppliers), 'color': (16, 185, 129)},
                {'label': 'Fulfilled Orders', 'val': sum(s['fulfilled_count'] for s in suppliers), 'color': (59, 130, 246)},
            ]

        # =========================================================================
        # 9. DISPOSAL RECORDS REPORT (SUPERADMIN ONLY)
        # =========================================================================
        elif report_type == 'disposals':
            if not is_superadmin():
                flash('Access denied. Only Superadmin can generate disposal record reports.', 'warning')
                return redirect(url_for('reports'))

            title = "Disposal Records & Audit Report"
            headers = ["Disposal ID", "Batch ID", "Medicine / Product", "Qty Disposed", "Reason", "Disposed By", "Date & Time", "Notes"]

            conn.execute('''CREATE TABLE IF NOT EXISTS disposals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT NOT NULL, medicine_id TEXT NOT NULL,
                medicine_name TEXT NOT NULL, qty_disposed INTEGER NOT NULL,
                reason TEXT DEFAULT 'Expired', disposed_by TEXT NOT NULL,
                disposed_at TEXT NOT NULL, notes TEXT DEFAULT ''
            )''')

            disposals = conn.execute("""
                SELECT id, batch_id, medicine_id, medicine_name, qty_disposed,
                       reason, disposed_by, disposed_at, COALESCE(notes, '') as notes
                FROM disposals
                ORDER BY id ASC
            """).fetchall()

            total_qty_disposed = 0
            unique_medicines = set()
            all_disposal_rows = []

            for d in disposals:
                if is_in_date_range(d['disposed_at']):
                    disp_id = f"DSP-{d['id']:04d}" if isinstance(d['id'], int) else str(d['id'])
                    batch_lbl = d['batch_id'] or '—'
                    med_name = d['medicine_name'] or '—'
                    qty = int(d['qty_disposed'] or 0)
                    reason = d['reason'] or 'Expired'
                    by_user = d['disposed_by'] or '—'
                    disp_time = d['disposed_at'] or '—'
                    notes = d['notes'] if d['notes'] and str(d['notes']).strip() else '—'

                    total_qty_disposed += qty
                    if med_name and med_name != '—':
                        unique_medicines.add(med_name)

                    all_disposal_rows.append([
                        disp_id,
                        batch_lbl,
                        med_name,
                        qty,
                        reason,
                        by_user,
                        disp_time,
                        notes
                    ])

            all_disposal_rows.sort(key=lambda r: natural_code_key(r[0]))
            rows = all_disposal_rows
            sections = []

            summary_cards = [
                {'label': 'Disposal Events', 'val': len(rows), 'color': (30, 41, 59)},
                {'label': 'Total Units Disposed', 'val': total_qty_disposed, 'color': (220, 38, 38)},
                {'label': 'Unique Products', 'val': len(unique_medicines), 'color': (245, 158, 11)},
            ]

        # Audit log the report generation & download in Activity Logs (History Center)
        period_desc = f"{start_dt.strftime('%m/%d/%Y')} - {end_dt.strftime('%m/%d/%Y')}" if (start_dt and end_dt) else "All Records"
        if frequency and (start_dt and end_dt):
            period_desc += f" ({frequency.capitalize()})"
            
        user_name = session.get('name') or session.get('username') or 'Owner'
        user_role = session.get('role') or 'Staff'
        log_activity(
            conn,
            action='download_report',
            target_type='Report',
            target_id=file_format.upper(),
            target_name=title,
            performed_by=user_name,
            details=f"Generated and downloaded {title} ({file_format.upper()}) with {len(rows)} record(s). Period: {period_desc}"
        )
        conn.commit()
        conn.close()
        filename = f"{report_type}_report"

        download_time_str = get_current_ph_time().strftime('%B %d, %Y at %I:%M %p')

        # =========================================================================
        # EXCEL GENERATION (OPENPYXL)
        # =========================================================================
        if file_format == 'excel':
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Report"
            
            ws.append(["FARMACIA NI DOK - PHARMACY AUDIT & INVENTORY REPORT"])
            ws.append([f"Report: {title}"])
            ws.append([f"Downloaded on: {download_time_str}"])
            ws.append([f"Generated by: {user_name} ({user_role})"])
            if start_dt and end_dt:
                ws.append([f"Period: {start_dt.strftime('%m/%d/%Y')} - {end_dt.strftime('%m/%d/%Y')} ({frequency.capitalize() if frequency else 'Custom'})"])
            else:
                ws.append(["Period: All Records / Comprehensive Audit"])
            
            ws.append([])
            
            if sections and len(sections) > 1:
                for sec in sections:
                    if not sec.get('rows'):
                        continue
                    ws.append([f"CATEGORY: {sec['title'].upper()} ({len(sec['rows'])} items)"])
                    ws.append(headers)
                    for r in sec['rows']:
                        ws.append(r)
                    ws.append([])
            else:
                ws.append(headers)
                for r in rows:
                    ws.append(r)
                
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
            output = BytesIO()
            wb.save(output)
            file_bytes = output.getvalue()
            ext = 'xlsx'
            mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            
        # =========================================================================
        # PDF GENERATION (CLEAN TABULAR AUDIT REPORTS)
        # =========================================================================
        elif file_format == 'pdf':
            period_text = f"Period: {start_dt.strftime('%m/%d/%Y')} - {end_dt.strftime('%m/%d/%Y')}" if (start_dt and end_dt) else "Period: All Records / Comprehensive Audit"
            pdf = PDFTemplateReport(title, period_text)
            pdf.alias_nb_pages()
            pdf.add_page()
            
            pdf.set_font("helvetica", "B", 12.5)
            pdf.cell(0, 8, title.upper(), align="C")
            pdf.ln(5)
            
            pdf.set_font("helvetica", "I", 8.0)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(0, 4.5, f"Downloaded on: {download_time_str}   |   Generated by: {user_name} ({user_role})", align="C")
            pdf.ln(4)
            
            pdf.cell(0, 4.5, period_text, align="C")
            pdf.ln(5)
            
            pdf.set_text_color(0, 0, 0)
            pdf.ln(3)

            # Tailored column widths mapping per report type (Total 190 mm printable width)
            col_widths_map = {
                'inventory': [14, 30, 28, 22, 18, 16, 14, 28, 20],
                'sales': [22, 20, 38, 30, 20, 32, 28],
                'stock_movements': [22, 20, 36, 22, 16, 42, 32],
                'batch_stocks': [22, 20, 38, 30, 26, 20, 34],
                'expiry_monitoring': [20, 40, 28, 26, 22, 20, 34],
                'velocity_matrix': [12, 38, 28, 22, 22, 32, 36],
                'top_selling': [16, 52, 42, 26, 26, 28],
                'suppliers': [22, 46, 28, 52, 20, 22],
                'disposals': [20, 22, 38, 20, 22, 24, 28, 16],
            }
            if report_type in col_widths_map and len(col_widths_map[report_type]) == len(headers):
                col_widths = col_widths_map[report_type]
            else:
                col_widths = [190 / len(headers)] * len(headers)

            # Auto-scaling cell renderer: shrinks font size to fit exactly inside the cell box if text is too long
            def draw_fit_cell(w, h, text, base_font_size=8.5, min_font_size=5.2, is_header=False, fill_color=None, text_color=None):
                text_str = PDFTemplateReport.sanitize(text)
                avail_w = max(w - 2.5, 2.0)
                font_weight = "B" if is_header else ""
                font_size = base_font_size
                
                pdf.set_font("helvetica", font_weight, font_size)
                str_w = pdf.get_string_width(text_str)
                if str_w > avail_w:
                    proportional_size = font_size * (avail_w / str_w)
                    font_size = max(min_font_size, round(proportional_size, 1))
                    pdf.set_font("helvetica", font_weight, font_size)
                    while pdf.get_string_width(text_str) > avail_w and font_size > min_font_size:
                        font_size -= 0.2
                        pdf.set_font("helvetica", font_weight, round(font_size, 1))
                
                if fill_color:
                    pdf.set_fill_color(*fill_color)
                if text_color:
                    pdf.set_text_color(*text_color)
                    
                pdf.cell(w, h, text_str, border=1, align="C", fill=(fill_color is not None))
                pdf.set_font("helvetica", font_weight, base_font_size)

            def draw_header_row():
                pdf.set_fill_color(192, 57, 43)
                pdf.set_text_color(255, 255, 255)
                for idx, h in enumerate(headers):
                    draw_fit_cell(col_widths[idx], 8, str(h), base_font_size=9.0, min_font_size=5.8, is_header=True, fill_color=(192, 57, 43), text_color=(255, 255, 255))
                pdf.ln()

            if not rows:
                draw_header_row()
                pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(100, 116, 139)
                pdf.set_font("helvetica", "I", 9)
                pdf.cell(190, 10, "No records found for the selected filter or period.", border=1, align="C", fill=True)
                pdf.ln()
            elif sections and len(sections) > 1:
                draw_header_row()
                for sec in sections:
                    if not sec.get('rows'):
                        continue
                    if pdf.get_y() > 248:
                        pdf.add_page()
                        draw_header_row()
                    
                    # Section Header Divider
                    pdf.set_fill_color(237, 242, 247)
                    pdf.set_text_color(30, 41, 59)
                    pdf.set_font("helvetica", "B", 8.5)
                    sec_title_str = pdf.sanitize(f"CATEGORY: {sec['title'].upper()} ({len(sec['rows'])} items)")
                    pdf.cell(190, 6.8, f"  {sec_title_str}", border=1, ln=True, fill=True)
                    
                    fill = False
                    for row in sec['rows']:
                        if pdf.get_y() > 260:
                            pdf.add_page()
                            draw_header_row()
                        
                        row_fill = (248, 250, 252) if fill else (255, 255, 255)
                        for idx, val in enumerate(row):
                            w = col_widths[idx] if idx < len(col_widths) else (190 / len(row))
                            draw_fit_cell(w, 7.5, val, base_font_size=8.5, min_font_size=5.2, is_header=False, fill_color=row_fill, text_color=(0, 0, 0))
                        pdf.ln()
                        fill = not fill
            else:
                draw_header_row()
                fill = False
                for row in rows:
                    if pdf.get_y() > 260:
                        pdf.add_page()
                        draw_header_row()
                    
                    row_fill = (248, 250, 252) if fill else (255, 255, 255)
                    for idx, val in enumerate(row):
                        w = col_widths[idx] if idx < len(col_widths) else (190 / len(row))
                        draw_fit_cell(w, 7.5, val, base_font_size=8.5, min_font_size=5.2, is_header=False, fill_color=row_fill, text_color=(0, 0, 0))
                    pdf.ln()
                    fill = not fill
                
            pdf_data = pdf.output()
            if isinstance(pdf_data, str):
                pdf_data = pdf_data.encode('latin-1')
            elif isinstance(pdf_data, bytearray):
                pdf_data = bytes(pdf_data)
                
            file_bytes = bytes(pdf_data)
            ext = 'pdf'
            mimetype = "application/pdf"

        # Save report directly to Windows Desktop (Offline Desktop App Support)
        target_filename = f"{filename}.{ext}"
        desktop_dir = get_desktop_folder()
        target_path = os.path.join(desktop_dir, target_filename)

        file_already_existed = os.path.exists(target_path)
        saved_to_desktop = False

        try:
            with open(target_path, 'wb') as f:
                f.write(file_bytes)
            saved_to_desktop = True
        except PermissionError:
            # If the existing file is open in Excel or another program on Windows, save with a unique timestamp
            ts_suffix = get_current_ph_time().strftime('%Y%m%d_%H%M%S')
            target_filename = f"{filename}_{ts_suffix}.{ext}"
            target_path = os.path.join(desktop_dir, target_filename)
            try:
                with open(target_path, 'wb') as f:
                    f.write(file_bytes)
                saved_to_desktop = True
            except Exception as fe:
                print(f"Fallback desktop write error: {fe}")
        except Exception as fe:
            print(f"Desktop write error: {fe}")

        # In-app Notification for the downloaded report
        try:
            n_conn = get_db_connection()
            if file_already_existed:
                n_title = f"Report Already on Desktop: {target_filename}"
                n_sub = f"File on Desktop updated with latest data ({download_time_str})"
                n_color = "#f59e0b"
            else:
                n_title = f"Report Downloaded to Desktop: {target_filename}"
                n_sub = f"Saved to Desktop at {download_time_str}"
                n_color = "#10b981"
            add_notification_with_conn(n_conn, 'report', n_title, n_sub, n_color)
            n_conn.commit()
            n_conn.close()
        except Exception as ne:
            print(f"Notification error: {ne}")

        is_ajax = (
            request.headers.get('X-Requested-With') == 'XMLHttpRequest' or
            request.headers.get('Accept', '').startswith('application/json') or
            request.form.get('ajax') == '1'
        )

        if is_ajax:
            import base64
            b64_str = base64.b64encode(file_bytes).decode('ascii')
            if file_already_existed:
                msg = f"Report already downloaded to Desktop: {target_filename} (Refreshed with latest data)"
            else:
                msg = f"Report successfully downloaded to Desktop: {target_filename}"

            return jsonify({
                'success': True,
                'filename': target_filename,
                'desktop_path': target_path,
                'saved_to_desktop': saved_to_desktop,
                'already_downloaded': file_already_existed,
                'message': msg,
                'file_base64': b64_str,
                'mime_type': mimetype
            })

        output = BytesIO(file_bytes)
        output.seek(0)
        return send_file(
            output,
            mimetype=mimetype,
            as_attachment=True,
            download_name=target_filename
        )
            
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"Report generation failed: {str(e)}", 500


@app.route('/api/open_file', methods=['POST'])
def open_file():
    """Allows one-click opening of desktop reports on Windows via default application"""
    try:
        data = request.get_json(silent=True) or {}
        filepath = data.get('filepath', '').strip()
        if not filepath:
            return jsonify({'success': False, 'message': 'No filepath provided'}), 400
        
        desktop_dir = os.path.normpath(get_desktop_folder())
        norm_fp = os.path.normpath(filepath)
        
        # Security check: must reside inside Desktop folder
        if not norm_fp.lower().startswith(desktop_dir.lower()):
            return jsonify({'success': False, 'message': 'Unauthorized file path'}), 403
            
        if not os.path.exists(norm_fp):
            return jsonify({'success': False, 'message': 'File not found on Desktop'}), 404
            
        os.startfile(norm_fp)
        return jsonify({'success': True, 'message': 'File opened successfully'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

# ==========================================
# AUTOMATED DATA BACKUP (CRON SCHEDULER & 30-DAY RETENTION)
# ==========================================
def cleanup_expired_backups(days_retention=30):
    """Automatically prunes and deletes backup files older than 30 days from backups/ directory"""
    try:
        backup_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')
        if not os.path.exists(backup_dir):
            return 0
            
        now_ts = datetime.now().timestamp()
        retention_seconds = days_retention * 86400
        deleted_count = 0
        
        for f in os.listdir(backup_dir):
            if f.endswith(('.db', '.sqlite', '.bak')):
                file_path = os.path.join(backup_dir, f)
                try:
                    file_age = now_ts - os.path.getmtime(file_path)
                    if file_age > retention_seconds:
                        os.remove(file_path)
                        deleted_count += 1
                        print(f"[Backup Retention] Automatically deleted expired backup (> {days_retention} days): {f}")
                except Exception as ex:
                    print(f"[Backup Retention] Error removing {f}: {ex}")
                    
        return deleted_count
    except Exception as e:
        print(f"[Backup Retention Error]: {e}")
        return 0

def perform_automatic_backup():
    """Performs an automated SQLite database snapshot and stores it in backups/"""
    try:
        from database import get_db_path
        import sqlite3
        db_path = get_db_path()
        if not os.path.exists(db_path):
            return None
        
        backup_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')
        os.makedirs(backup_dir, exist_ok=True)
        
        now = datetime.now()
        timestamp = now.strftime('%Y%m%d_%H%M%S')
        backup_filename = f"farmacia_auto_backup_{timestamp}.db"
        backup_path = os.path.join(backup_dir, backup_filename)
        
        src_conn = sqlite3.connect(db_path)
        dst_conn = sqlite3.connect(backup_path)
        src_conn.backup(dst_conn)
        dst_conn.close()
        src_conn.close()
        
        # Automatically delete backup files older than 30 days due
        cleanup_expired_backups(days_retention=30)
        
        # Upper safety limit: keep at most 200 snapshots in case retention hasn't purged them
        existing_backups = sorted([
            os.path.join(backup_dir, f) for f in os.listdir(backup_dir) 
            if f.startswith('farmacia_auto_backup_') and f.endswith('.db')
        ], key=os.path.getmtime)
        
        if len(existing_backups) > 200:
            for old_file in existing_backups[:-200]:
                try:
                    os.remove(old_file)
                except Exception:
                    pass
                    
        # Update settings table
        try:
            conn = get_db_connection()
            conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('last_auto_backup_time', ?)", (now.strftime('%Y-%m-%d %I:%M:%S %p'),))
            conn.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('last_auto_backup_file', ?)", (backup_filename,))
            log_activity(conn, 'backup', 'System', 'database', 'Automatic Backup (Cron)', details=f"Automated cron backup created: {backup_filename}")
            conn.commit()
            conn.close()
        except Exception:
            pass
            
        return backup_path
    except Exception as e:
        print(f"[Cron Backup Error]: {e}")
        return None

def start_cron_backup_scheduler(interval_hours=6):
    """Background daemon thread that periodically runs perform_automatic_backup()"""
    import threading, time
    def run_scheduler():
        time.sleep(10)
        perform_automatic_backup()
        while True:
            time.sleep(interval_hours * 3600)
            perform_automatic_backup()

    t = threading.Thread(target=run_scheduler, daemon=True, name="CronBackupWorker")
    t.start()

# Start background cron worker
try:
    start_cron_backup_scheduler(6)
except Exception as e:
    print(f"Could not start backup scheduler: {e}")

@app.route('/run_auto_backup', methods=['POST'])
def run_auto_backup():
    if 'user_id' not in session or session.get('role') == 'Staff':
        return jsonify({'success': False, 'error': 'Permission denied. Admin privileges required.'}), 403
    path = perform_automatic_backup()
    if path:
        return jsonify({
            'success': True, 
            'filename': os.path.basename(path),
            'time': datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')
        })
    return jsonify({'success': False, 'error': 'Automatic backup execution failed.'}), 500

@app.route('/restore_snapshot', methods=['POST'])
def restore_snapshot():
    """1-Click automated database restore from local cron snapshot points without manual download or file searching."""
    if 'user_id' not in session or session.get('role') == 'Staff':
        flash('Permission denied. Admin privileges required to restore database.', 'danger')
        return redirect(url_for('settings'))

    snapshot_name = request.form.get('snapshot_name', '').strip()
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')
    
    from database import get_db_path
    import sqlite3
    db_path = get_db_path()
    
    if not os.path.exists(backup_dir):
        flash('No automated backup snapshots exist yet to restore.', 'danger')
        return redirect(url_for('settings'))
        
    # If requesting 'latest', automatically select the newest snapshot
    if not snapshot_name or snapshot_name == 'latest':
        files = [f for f in os.listdir(backup_dir) if f.startswith('farmacia_auto_backup_') and f.endswith('.db')]
        if not files:
            flash('No automated backup snapshots found.', 'danger')
            return redirect(url_for('settings'))
        files.sort(key=lambda x: os.path.getmtime(os.path.join(backup_dir, x)), reverse=True)
        snapshot_name = files[0]
        
    # Security: ensure file is inside backups directory
    snapshot_filename = os.path.basename(snapshot_name)
    target_snapshot_path = os.path.join(backup_dir, snapshot_filename)
    
    if not os.path.exists(target_snapshot_path):
        flash(f'Backup snapshot "{snapshot_filename}" not found.', 'danger')
        return redirect(url_for('settings'))
        
    try:
        # 1. First take an instant safety backup of current state
        perform_automatic_backup()
        
        # 2. Check snapshot integrity
        src_conn = sqlite3.connect(target_snapshot_path)
        check_res = src_conn.execute("PRAGMA quick_check").fetchone()
        if not check_res or check_res[0] != 'ok':
            src_conn.close()
            flash('Selected backup snapshot is corrupted or invalid. Restore cancelled.', 'danger')
            return redirect(url_for('settings'))
            
        # 3. Restore snapshot into main db
        dest_conn = sqlite3.connect(db_path)
        src_conn.backup(dest_conn)
        src_conn.close()
        
        mtime = datetime.fromtimestamp(os.path.getmtime(target_snapshot_path))
        time_str = mtime.strftime('%b %d, %Y · %I:%M %p')
        
        log_activity(dest_conn, 'restore_db', 'System', 'database', 'Automated Restore', 
                     details=f"Restored database from automated snapshot: {snapshot_filename} (Snapshot Point: {time_str})")
        dest_conn.commit()
        dest_conn.close()
        
        flash(f'Database successfully restored from automated backup ({time_str})! System data restored seamlessly without manual file handling.', 'success')
    except Exception as e:
        flash(f'Database restoration failed: {str(e)}', 'danger')
        
    return redirect(url_for('settings'))

@app.route('/backup_db')
def backup_db():
    """Compatibility endpoint redirecting manual requests to automated backup snapshot"""
    path = perform_automatic_backup()
    if path:
        return jsonify({
            'success': True,
            'filename': os.path.basename(path),
            'message': 'Automated cron backup executed successfully.'
        })
    return jsonify({'success': False, 'error': 'Backup failed.'}), 500

# =========================================================================
# OFFLINE / POWER OUTAGE MANUAL EXCEL RECORDS SYNC
# =========================================================================
@app.route('/download_offline_template')
def download_offline_template():
    """Generates a styled Excel template for recording sales manually during power outages."""
    import openpyxl, io
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Offline Sales Entry"
    
    # Title Banner
    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = "FARMACIA NI DOK — OFFLINE / POWER OUTAGE MANUAL SALES LOG"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32
    
    # Instructions Banner
    ws.merge_cells("A2:D2")
    sub_cell = ws["A2"]
    sub_cell.value = "Record offline customer transactions here during power outages. Once power is restored, upload this file to sync inventory and sales."
    sub_cell.font = Font(name="Calibri", size=9.5, italic=True, color="1F2937")
    sub_cell.fill = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22
    
    headers = [
        "Date (YYYY-MM-DD)",
        "Medicine Name or Brand",
        "Quantity Sold",
        "Recorded By (Staff)"
    ]
    
    header_fill = PatternFill(start_color="10B981", end_color="10B981", fill_type="solid")
    header_font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[3].height = 26
    
    # Sample pre-filled guide rows
    today_str = get_current_ph_time().strftime('%Y-%m-%d')
    sample_data = [
        [today_str, "Biogesic", 10, "Assistant"],
        [today_str, "Amoxil", 5, "Assistant"],
        [today_str, "Solmux", 3, "Staff2"]
    ]
    
    data_font = Font(name="Calibri", size=10.5)
    for r_idx, row_values in enumerate(sample_data, 4):
        ws.row_dimensions[r_idx].height = 20
        for c_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if c_idx in [1, 3, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
    widths = [22, 34, 18, 28]
    for i, w in enumerate(widths, 1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws.column_dimensions[col_letter].width = w
        
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    
    return send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="Farmacia_Offline_Sales_Template.xlsx"
    )

def find_matching_medicine(conn, query_name):
    """Finds an active medicine in inventory by ID, Brand, Generic, or Product Name."""
    if not query_name:
        return None
    q = str(query_name).strip()
    
    # 1. Exact match on ID
    med = conn.execute("SELECT * FROM inventory WHERE id = ? AND archived_at IS NULL AND deleted_at IS NULL", (q,)).fetchone()
    if med:
        return med
    # 2. Case-insensitive exact match on Brand
    med = conn.execute("SELECT * FROM inventory WHERE LOWER(brand) = LOWER(?) AND archived_at IS NULL AND deleted_at IS NULL", (q,)).fetchone()
    if med:
        return med
    # 3. Case-insensitive match on product_name
    med = conn.execute("SELECT * FROM inventory WHERE LOWER(product_name) = LOWER(?) AND archived_at IS NULL AND deleted_at IS NULL", (q,)).fetchone()
    if med:
        return med
    # 4. Case-insensitive match on generic
    med = conn.execute("SELECT * FROM inventory WHERE LOWER(generic) = LOWER(?) AND archived_at IS NULL AND deleted_at IS NULL", (q,)).fetchone()
    if med:
        return med
    # 5. Fuzzy/Contains match on brand or generic
    all_meds = conn.execute("SELECT * FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL").fetchall()
    q_lower = q.lower()
    for m in all_meds:
        brand_lower = (m['brand'] or '').lower()
        prod_lower = (m['product_name'] or '').lower()
        gen_lower = (m['generic'] or '').lower()
        if brand_lower and (brand_lower in q_lower or q_lower in brand_lower):
            return m
        if prod_lower and (prod_lower in q_lower or q_lower in prod_lower):
            return m
        if gen_lower and (gen_lower in q_lower or q_lower in gen_lower):
            return m
            
    return None

def process_offline_sale(conn, medicine, qty, sale_date_str, sold_by, notes):
    """Executes FEFO batch deduction, updates inventory stock, logs movement and inserts sale."""
    medicine_id = medicine['id']
    med_brand = medicine['brand']
    
    if medicine['stock'] < qty:
        return False, f"Insufficient stock for '{med_brand}' (Available: {medicine['stock']}, In File: {qty})"
        
    remaining_to_sell = qty
    batches = conn.execute("SELECT * FROM batches WHERE medicine_id = ? AND current_qty > 0 ORDER BY expiry_date ASC", (medicine_id,)).fetchall()
    
    last_sale = conn.execute("SELECT id FROM sales ORDER BY id DESC LIMIT 1").fetchone()
    if last_sale:
        try:
            last_num = int(last_sale['id'].split('-')[1])
            next_sale_id = f"TRN-{last_num + 1:03d}"
        except (IndexError, ValueError):
            next_sale_id = "TRN-001"
    else:
        next_sale_id = "TRN-001"
        
    trn_ref = f"Offline Sync {next_sale_id}" + (f" ({notes})" if notes else "")
    trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    
    for b in batches:
        if remaining_to_sell <= 0:
            break
        sell_qty = min(remaining_to_sell, b['current_qty'])
        conn.execute("UPDATE batches SET current_qty = current_qty - ? WHERE id = ?", (sell_qty, b['id']))
        conn.execute("UPDATE inventory SET stock = stock - ? WHERE id = ?", (sell_qty, medicine_id))
        
        last_mov = conn.execute("SELECT id FROM stock_movements ORDER BY id DESC LIMIT 1").fetchone()
        if last_mov:
            try:
                last_num = int(last_mov['id'].split('-')[1])
                next_trn_id = f"TRN-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_trn_id = "TRN-001"
        else:
            next_trn_id = "TRN-001"
            
        conn.execute("INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
                     (next_trn_id, 'Stock-Out', medicine_id, b['id'], sell_qty, trn_date, trn_ref))
        remaining_to_sell -= sell_qty

    if remaining_to_sell > 0:
        conn.execute("UPDATE inventory SET stock = stock - ? WHERE id = ?", (remaining_to_sell, medicine_id))
        last_mov = conn.execute("SELECT id FROM stock_movements ORDER BY id DESC LIMIT 1").fetchone()
        if last_mov:
            try:
                last_num = int(last_mov['id'].split('-')[1])
                next_trn_id = f"TRN-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_trn_id = "TRN-001"
        else:
            next_trn_id = "TRN-001"
        conn.execute("INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
                     (next_trn_id, 'Stock-Out', medicine_id, 'N/A', remaining_to_sell, trn_date, trn_ref))

    # Recalculate status using low_stock_threshold
    med_item = conn.execute("SELECT stock, reorder_point FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
    if med_item:
        new_status = compute_stock_status(med_item['stock'], get_low_stock_threshold(conn), med_item['reorder_point'])
        conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, medicine_id))

    # Insert Sale with created_at timestamp
    base_u = medicine.get('base_unit') or medicine.get('unit_of_measure') or 'Piece'
    created_at_now = get_current_ph_time().strftime('%Y-%m-%d %I:%M %p')
    conn.execute("INSERT INTO sales (id, medicine_id, sale_date, qty, unit, base_qty, sold_by, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                 (next_sale_id, medicine_id, sale_date_str, qty, base_u, qty, sold_by, created_at_now))
    
    log_activity(conn, 'sale', 'product', medicine_id, med_brand, details=f"Offline Sync: Sold {qty} units (Receipt {next_sale_id}, Staff: {sold_by})")
    return True, next_sale_id

@app.route('/upload_offline_records', methods=['POST'])
def upload_offline_records():
    """Parses and synchronizes manual offline sales records from uploaded Excel or CSV spreadsheet."""
    if 'user_id' not in session:
        return redirect(url_for('login'))
    if not has_permission('sales_create') and session.get('role') not in ['Superadmin', 'Admin', 'Owner / Pharmacist']:
        flash('Access denied. You do not have permission to record sales.', 'warning')
        return redirect(request.referrer or url_for('sales'))
        
    if 'excel_file' not in request.files:
        flash('No file was selected for upload.', 'danger')
        return redirect(request.referrer or url_for('sales'))
        
    file = request.files['excel_file']
    if not file or file.filename == '':
        flash('No file was selected.', 'danger')
        return redirect(request.referrer or url_for('sales'))
        
    filename = file.filename.lower()
    if not (filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.csv')):
        flash('Invalid file format. Please upload an Excel (.xlsx, .xls) or CSV file.', 'danger')
        return redirect(request.referrer or url_for('sales'))

    records = []
    try:
        if filename.endswith('.csv'):
            import csv
            stream = io.StringIO(file.stream.read().decode("utf-8", errors="ignore"))
            reader = csv.reader(stream)
            for r in reader:
                records.append(r)
        else:
            import openpyxl
            wb = openpyxl.load_workbook(file, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                records.append(list(row))
    except Exception as e:
        flash(f'Error reading uploaded spreadsheet: {str(e)}', 'danger')
        return redirect(request.referrer or url_for('sales'))
        
    # Dynamic header detection
    header_idx = -1
    col_map = {'date': 0, 'med': 1, 'qty': 2, 'staff': 3}
    for idx, row in enumerate(records):
        row_str = " ".join([str(c or '').lower() for c in row])
        if 'medicine' in row_str or 'brand' in row_str or ('qty' in row_str or 'quantity' in row_str):
            header_idx = idx
            for c_idx, val in enumerate(row):
                v_lower = str(val or '').lower()
                if 'date' in v_lower:
                    col_map['date'] = c_idx
                elif 'medicine' in v_lower or 'brand' in v_lower or 'product' in v_lower:
                    col_map['med'] = c_idx
                elif 'qty' in v_lower or 'quantity' in v_lower:
                    col_map['qty'] = c_idx
                elif any(k in v_lower for k in ('staff', 'recorded', 'transacted', 'by', 'user', 'cashier')):
                    col_map['staff'] = c_idx
            break
            
    start_row = 0 if header_idx == -1 else header_idx + 1

    conn = get_db_connection()
    success_count = 0
    skipped_count = 0
    errors = []
    uploader_name = session.get('name', 'Staff')
    
    for r_num, row in enumerate(records[start_row:], start=start_row + 1):
        if not row or all(c is None or str(c).strip() == '' for c in row):
            continue
            
        date_val = row[col_map['date']] if len(row) > col_map['date'] else None
        med_val = row[col_map['med']] if len(row) > col_map['med'] else None
        qty_val = row[col_map['qty']] if len(row) > col_map['qty'] else None
        staff_val = row[col_map['staff']] if len(row) > col_map['staff'] else None
        
        # If user entered medicine name in first column
        if med_val is None and date_val:
            med_val = date_val
            date_val = None
            if len(row) > 1:
                qty_val = row[1]
                
        if not med_val:
            continue
            
        try:
            qty = int(float(str(qty_val).replace(',', '').strip()))
            if qty <= 0:
                errors.append(f"Row {r_num}: Quantity must be greater than 0")
                skipped_count += 1
                continue
        except (ValueError, TypeError):
            errors.append(f"Row {r_num}: Invalid quantity '{qty_val}' for '{med_val}'")
            skipped_count += 1
            continue
            
        sale_date_str = get_current_ph_time().strftime('%m/%d/%Y')
        if date_val:
            from datetime import date as d_date
            if isinstance(date_val, (datetime, d_date)):
                sale_date_str = date_val.strftime('%m/%d/%Y')
            else:
                d_str = str(date_val).strip()
                for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d'):
                    try:
                        sale_date_str = datetime.strptime(d_str, fmt).strftime('%m/%d/%Y')
                        break
                    except Exception:
                        pass

        med = find_matching_medicine(conn, med_val)
        if not med:
            errors.append(f"Row {r_num}: Medicine '{med_val}' not found in inventory")
            skipped_count += 1
            continue
            
        # Correctly assign transacted by user/staff from Excel (without (Offline) label)
        if staff_val and str(staff_val).strip():
            raw_staff = str(staff_val).strip()
        else:
            raw_staff = uploader_name or 'Staff'
            
        sold_by_name = raw_staff.replace(' (Offline)', '').replace('(Offline)', '').strip()
        
        ok, res_id = process_offline_sale(conn, med, qty, sale_date_str, sold_by_name, "Power Outage Manual Entry")
        if ok:
            success_count += 1
        else:
            errors.append(f"Row {r_num}: {res_id}")
            skipped_count += 1
            
    conn.commit()
    
    if success_count > 0:
        add_notification_with_conn(conn, 'offline_sync', 
                                   f"Offline Records Synced ({success_count} transactions)",
                                   f"Successfully imported manual sales from {file.filename}. Stock levels updated.", 
                                   '#10b981')
        conn.commit()
        
    conn.close()
    
    if success_count > 0:
        msg = f"Successfully synchronized {success_count} offline record(s) from Excel! Stock deducted and sales logged."
        if errors:
            msg += f" (Note: {skipped_count} row(s) skipped: {'; '.join(errors[:3])}{'...' if len(errors) > 3 else ''})"
        flash(msg, 'success')
    else:
        flash(f"No records could be synchronized. Issues found: {'; '.join(errors[:5])}", 'danger')
        
    return redirect(request.referrer or url_for('sales'))

@app.route('/restore_db', methods=['POST'])
def restore_db():
    if 'user_id' not in session or session.get('role') == 'Staff':
        flash('Permission denied. Admin privileges required to restore database.')
        return redirect(url_for('settings'))
    
    if 'backup_file' not in request.files:
        flash('No backup file selected.')
        return redirect(url_for('settings'))
    
    file = request.files['backup_file']
    if file.filename == '':
        flash('No backup file selected.')
        return redirect(url_for('settings'))
    
    if not (file.filename.endswith('.db') or file.filename.endswith('.sqlite') or file.filename.endswith('.bak')):
        flash('Invalid backup file format. Please upload a valid .db or .sqlite backup file.')
        return redirect(url_for('settings'))
    
    from database import get_db_path
    import tempfile, sqlite3, time
    db_path = get_db_path()
    
    temp_dir = tempfile.gettempdir()
    temp_file_path = os.path.join(temp_dir, f"temp_restore_{int(time.time())}.db")
    
    try:
        file.save(temp_file_path)
        
        source_conn = sqlite3.connect(temp_file_path)
        check_res = source_conn.execute("PRAGMA quick_check").fetchone()
        
        if check_res and check_res[0] == 'ok':
            dest_conn = sqlite3.connect(db_path)
            source_conn.backup(dest_conn)
            source_conn.close()
            try:
                log_activity(dest_conn, 'restore_db', 'System', 'database', 'Database Restore', details=f"Restored database from uploaded backup ({file.filename})")
                dest_conn.commit()
            except Exception:
                pass
            dest_conn.close()
            
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                
            flash('Database successfully restored from backup!')
        else:
            source_conn.close()
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            flash('Corrupted or invalid database backup file. Restore cancelled.')
            
    except Exception as e:
        if os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except Exception:
                pass
        flash(f'Database restore failed: {str(e)}')
        
    return redirect(url_for('settings'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
