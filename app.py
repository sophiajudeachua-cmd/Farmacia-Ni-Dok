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
            {'key': 'inventory_archive', 'action': 'ARCHIVE', 'label': 'Archive & Restore', 'desc': 'Archive inactive medicines and restore archived items'},
            {'key': 'inventory_delete', 'action': 'DELETE', 'label': 'Trash & Delete', 'desc': 'Move medicine items to trash or permanently purge them'}
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
            {'key': 'suppliers_delete', 'action': 'DELETE', 'label': 'Delete Supplier', 'desc': 'Remove suppliers that have no linked purchase orders'}
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
    return {'local_ip': get_local_ip()}

def parse_date_range(frequency, period_week, period_month, period_year):
    if frequency == 'weekly' and period_week:
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

def add_notification_with_conn(conn, type, title, subtitle, color):
    now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
    existing = conn.execute('SELECT id FROM notifications WHERE title = ?', (title,)).fetchone()
    if not existing:
        conn.execute('''
            INSERT INTO notifications (type, title, subtitle, color, created_at, seen)
            VALUES (?, ?, ?, ?, ?, 0)
        ''', (type, title, subtitle, color, now_str))

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

def compute_stock_status(stock, low_threshold=40):
    if stock == 0:
        return 'No Stock'
    if stock <= low_threshold:
        return 'Low Stock'
    return 'Good'

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
        
        # Validation 1: Maximum length of 15 characters
        if len(username) > 15 or len(password) > 15:
            flash('Username and password cannot exceed 15 characters.', 'danger')
            return render_template('login.html')
            
        # Validation 2: No special characters allowed (alphanumeric only)
        if not re.match(r'^[a-zA-Z0-9]{1,15}$', username) or not re.match(r'^[a-zA-Z0-9]{1,15}$', password):
            flash('Username and password must contain only letters and numbers, maximum 15 characters (no special characters allowed).', 'danger')
            return render_template('login.html')
            
        conn = get_db_connection()
        user = conn.execute('SELECT * FROM users WHERE username = ? AND password = ?', (username, password)).fetchone()
        
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
        
    if len(new_password) < 10 or len(new_password) > 15 or not re.match(r'^[a-zA-Z0-9]{10,15}$', new_password):
        flash('Password must be between 10 and 15 alphanumeric characters (no special characters allowed).', 'danger')
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
    po_receiving = conn.execute("SELECT COUNT(*) FROM purchase_orders WHERE status = 'For Receiving'").fetchone()[0]
    try:
        pending_expiry = conn.execute("SELECT COUNT(*) FROM batches WHERE expiry_pending = 1").fetchone()[0]
    except Exception:
        pending_expiry = 0
    
    # Tables
    low_stock_meds = conn.execute("SELECT brand as medicine, stock, reorder_point, status FROM inventory WHERE status IN ('Low Stock', 'No Stock') AND archived_at IS NULL AND deleted_at IS NULL LIMIT 5").fetchall()
    
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

    # Sales by date for Line Graph (with Last Week, Monthly, and Annual summarization)
    all_sales = conn.execute('SELECT sale_date, qty FROM sales').fetchall()
    from collections import defaultdict
    from datetime import datetime, timedelta

    sales_dates = []
    for s in all_sales:
        try:
            dt = datetime.strptime(s['sale_date'], '%m/%d/%Y')
            sales_dates.append((dt, s['qty']))
        except Exception:
            pass

    if sales_dates:
        anchor_date = max(d[0] for d in sales_dates)
    else:
        anchor_date = datetime.now()

    # 1. Last Week (last 7 days ending at anchor_date)
    last_week_data = {}
    for i in range(7):
        d = anchor_date - timedelta(days=6-i)
        last_week_data[d.strftime('%m/%d/%Y')] = 0
    
    for dt, qty in sales_dates:
        d_str = dt.strftime('%m/%d/%Y')
        if d_str in last_week_data:
            last_week_data[d_str] += qty
            
    last_week_labels = list(last_week_data.keys())
    last_week_values = list(last_week_data.values())

    # 2. This Week (daily sales from the Monday of anchor_date's week to Sunday/anchor_date)
    # anchor_date.weekday() returns 0 for Monday, 6 for Sunday
    monday_of_week = anchor_date - timedelta(days=anchor_date.weekday())
    this_week_data = {}
    for i in range(7):
        d = monday_of_week + timedelta(days=i)
        this_week_data[d.strftime('%m/%d/%Y')] = 0
        
    for dt, qty in sales_dates:
        d_str = dt.strftime('%m/%d/%Y')
        if d_str in this_week_data:
            this_week_data[d_str] += qty
            
    this_week_labels = list(this_week_data.keys())
    this_week_values = list(this_week_data.values())



    # 4. Monthly (group by Month/Year)
    monthly_map = defaultdict(int)
    for dt, qty in sales_dates:
        monthly_map[dt.strftime('%b %Y')] += qty
        
    sorted_months = sorted(monthly_map.keys(), key=lambda x: datetime.strptime(x, '%b %Y'))
    monthly_labels = sorted_months
    monthly_values = [monthly_map[m] for m in sorted_months]

    # 5. Annual (group by Year)
    annual_map = defaultdict(int)
    for dt, qty in sales_dates:
        annual_map[dt.strftime('%Y')] += qty
        
    sorted_years = sorted(annual_map.keys())
    annual_labels = sorted_years
    annual_values = [annual_map[y] for y in sorted_years]

    # Stocks for Pie Chart
    inventory_stocks = conn.execute('SELECT brand, stock FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL').fetchall()
    stock_labels = [item['brand'] for item in inventory_stocks]
    stock_values = [item['stock'] for item in inventory_stocks]

    # Recent activities
    try:
        recent_activities_raw = conn.execute(
            'SELECT * FROM activity_log ORDER BY id DESC LIMIT 8'
        ).fetchall()
        recent_activities = [dict(r) for r in recent_activities_raw]
    except Exception:
        recent_activities = []

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
        stock_labels=stock_labels,
        stock_values=stock_values,
        recent_activities=recent_activities,
        pending_expiry=pending_expiry,
        active_page='dashboard'
    )

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

        generic = request.form.get('generic', '')
        brand = request.form['brand']
        category = request.form['category']
        reorder_point = min(MAX_STOCK_LIMIT, max(0, int(request.form.get('reorder_point', 100))))
        low_thresh = get_low_stock_threshold(conn)
        if reorder_point < low_thresh:
            conn.close()
            flash(f"Cannot add product: Reorder Point ({reorder_point}) cannot be less than the Low Stock Threshold ({low_thresh} units).", "error")
            return redirect(url_for('inventory'))
        
        product_name = request.form.get('product_name', brand)
        description = request.form.get('description', '')
        unit_of_measure = request.form.get('unit_of_measure', 'Piece')
        purchase_price = float(request.form.get('purchase_price', 0.0))
        supplier = request.form.get('supplier', '')
        
        # Generate next medicine ID (e.g. MED-034)
        last_med = conn.execute("SELECT id FROM inventory ORDER BY id DESC LIMIT 1").fetchone()
        if last_med:
            try:
                last_num = int(last_med['id'].split('-')[1])
                next_med_id = f"MED-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_med_id = "MED-034"
        else:
            next_med_id = "MED-001"
            
        conn.execute("""
            INSERT INTO inventory (id, generic, brand, category, stock, reorder_point, status, product_name, description, unit_of_measure, purchase_price, supplier) 
            VALUES (?, ?, ?, ?, 0, ?, 'No Stock', ?, ?, ?, ?, ?)
        """, (next_med_id, generic, brand, category, reorder_point, product_name, description, unit_of_measure, purchase_price, supplier))
        log_activity(conn, 'add', 'product', next_med_id, brand, details=f"Added product {brand} ({generic}) in {category}")
        conn.commit()
        conn.close()
        return redirect(url_for('inventory'))
        
    items = conn.execute('SELECT * FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL').fetchall()
    suppliers = conn.execute('SELECT name FROM suppliers').fetchall()
    conn.close()
    return render_template('inventory.html', items=items, suppliers=suppliers, active_page='inventory')

@app.route('/inventory/edit', methods=['POST'])
def edit_medicine():
    if not has_permission('inventory_edit'):
        flash('Access denied. You do not have permission to edit inventory.', 'warning')
        return redirect(url_for('inventory'))
    conn = get_db_connection()
    med_id = request.form['id']
    generic = request.form.get('generic', '')
    brand = request.form['brand']
    category = request.form['category']
    reorder_point = min(MAX_STOCK_LIMIT, max(0, int(request.form.get('reorder_point', 100))))
    low_thresh = get_low_stock_threshold(conn)
    if reorder_point < low_thresh:
        conn.close()
        flash(f"Cannot update product: Reorder Point ({reorder_point}) cannot be less than the Low Stock Threshold ({low_thresh} units).", "error")
        return redirect(url_for('inventory'))
    
    product_name = request.form.get('product_name', brand)
    description = request.form.get('description', '')
    unit_of_measure = request.form.get('unit_of_measure', 'Piece')
    purchase_price = float(request.form.get('purchase_price', 0.0))
    supplier = request.form.get('supplier', '')
    
    # Recalculate status using low_stock_threshold
    med = conn.execute("SELECT stock FROM inventory WHERE id = ?", (med_id,)).fetchone()
    stock = med['stock'] if med else 0
    low_thresh = get_low_stock_threshold(conn)
    new_status = compute_stock_status(stock, low_thresh)
    
    conn.execute('''
        UPDATE inventory 
        SET generic = ?, brand = ?, category = ?, reorder_point = ?, status = ?,
            product_name = ?, description = ?, unit_of_measure = ?, purchase_price = ?, supplier = ?
        WHERE id = ?
    ''', (generic, brand, category, reorder_point, new_status, product_name, description, unit_of_measure, purchase_price, supplier, med_id))
    log_activity(conn, 'edit', 'product', med_id, brand, details=f"Updated details (Category: {category}, Reorder Point: {reorder_point})")
    conn.commit()
    conn.close()
    flash('Medicine updated successfully!')
    return redirect(url_for('inventory'))

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
    if not has_permission('inventory_delete'):
        flash('Access denied. You do not have permission to delete products.', 'warning')
        return redirect(url_for('history'))
    conn = get_db_connection()
    med = conn.execute('SELECT brand FROM inventory WHERE id = ?', (med_id,)).fetchone()
    if med:
        now_str = datetime.now().strftime('%Y-%m-%d %I:%M %p')
        conn.execute('UPDATE inventory SET deleted_at = ?, archived_at = NULL WHERE id = ?', (now_str, med_id))
        log_activity(conn, 'trash', 'product', med_id, med['brand'], details=f"Moved {med['brand']} to trash (30-day retention)")
        conn.commit()
    conn.close()
    flash(f'{med["brand"]} moved to trash. It will be permanently deleted after 30 days.')
    return redirect(url_for('history'))

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

    # Activity log (paginated — last 300)
    activity_log_raw = conn.execute(
        'SELECT * FROM activity_log ORDER BY id DESC LIMIT 300'
    ).fetchall()
    activity_log = [dict(r) for r in activity_log_raw]

    conn.close()
    return render_template('history.html',
        archived=archived,
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
    items = conn.execute('''
        SELECT b.id as batch_id, i.brand as medicine, b.expiry_date, b.current_qty, b.status, i.category
        FROM batches b 
        JOIN inventory i ON b.medicine_id = i.id
    ''').fetchall()
    conn.close()
    
    list_items = []
    for b in items:
        try:
            expiry_dt = datetime.strptime(b['expiry_date'], '%m/%d/%Y')
            days_left = (expiry_dt - datetime.now()).days
            if days_left <= 0:
                status = 'Expired'
            elif days_left <= 50:
                status = 'Near Expiry'
            else:
                status = 'Good'
        except Exception:
            status = 'Good'
            
        list_items.append({
            'batch_id': b['batch_id'],
            'medicine': b['medicine'],
            'expiry_date': b['expiry_date'],
            'current_qty': b['current_qty'],
            'status': status,
            'category': b['category']
        })
    return render_template('batches.html', items=list_items, active_page='batches')

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
        name = request.form['name']
        address = request.form['address']
        contact = request.form['contact']
        
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
        
    suppliers_raw = conn.execute('SELECT * FROM suppliers ORDER BY id ASC').fetchall()
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

@app.route('/suppliers/edit', methods=['POST'])
def edit_supplier():
    if not has_permission('suppliers_edit'):
        flash('Access denied. You do not have permission to edit suppliers.', 'warning')
        return redirect(url_for('suppliers'))
    conn = get_db_connection()
    sup_id = request.form['id']
    name = request.form['name']
    address = request.form['address']
    contact = request.form['contact']
    
    old_sup = conn.execute("SELECT name FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    old_name = old_sup['name'] if old_sup else None

    conn.execute('''
        UPDATE suppliers 
        SET name = ?, address = ?, contact = ?
        WHERE id = ?
    ''', (name, address, contact, sup_id))
    
    # Keep inventory product associations synchronized if supplier name changes
    if old_name and old_name != name:
        conn.execute("UPDATE inventory SET supplier = ? WHERE supplier = ?", (name, old_name))

    log_activity(conn, 'edit_supplier', 'Supplier', sup_id, name, details=f"Updated supplier {name} ({contact})")
    conn.commit()
    conn.close()
    flash('Supplier updated successfully!')
    return redirect(url_for('suppliers'))

@app.route('/suppliers/<sup_id>/delete', methods=['POST'])
def delete_supplier(sup_id):
    if not has_permission('suppliers_delete'):
        flash('Access denied. You do not have permission to delete suppliers.', 'warning')
        return redirect(url_for('suppliers'))
    conn = get_db_connection()
    sup = conn.execute("SELECT name FROM suppliers WHERE id = ?", (sup_id,)).fetchone()
    sup_name = sup['name'] if sup else sup_id
    po = conn.execute("SELECT * FROM purchase_orders WHERE supplier_id = ?", (sup_id,)).fetchone()
    inv = conn.execute("SELECT * FROM inventory WHERE supplier = ? AND archived_at IS NULL AND deleted_at IS NULL", (sup_name,)).fetchone()
    if po:
        flash("Cannot delete supplier: linked purchase orders exist!", "warning")
    elif inv:
        flash(f"Cannot delete supplier: products are currently assigned to {sup_name} in Inventory!", "warning")
    else:
        conn.execute("DELETE FROM suppliers WHERE id = ?", (sup_id,))
        log_activity(conn, 'delete_supplier', 'Supplier', sup_id, sup_name, details=f"Deleted supplier {sup_name}")
        conn.commit()
        flash("Supplier deleted successfully!")
    conn.close()
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
        notes = request.form['notes']
        medicines = request.form.getlist('medicine[]')
        quantities = request.form.getlist('quantity[]')
        
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
        
        # Insert PO Items
        for med_id, qty in zip(medicines, quantities):
            if med_id and qty:
                sanitized_qty = min(MAX_STOCK_LIMIT, max(1, int(qty)))
                conn.execute('''
                    INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity) 
                    VALUES (?, ?, ?)
                ''', (next_po_id, med_id, sanitized_qty))
                
        # Add PO pending notification
        supplier_name = conn.execute("SELECT name FROM suppliers WHERE id = ?", (supplier_id,)).fetchone()['name']
        valid_items = [m for m in medicines if m]
        add_notification_with_conn(conn, 'po', f"PO {next_po_id} pending receipt", f"From {supplier_name} · {len(valid_items)} items · For Receiving", '#3b82f6')
        log_activity(conn, 'create_po', 'purchase_order', next_po_id, f"PO {next_po_id}", details=f"Supplier: {supplier_name} ({len(valid_items)} items ordered)")
        
        conn.commit()
        conn.close()
        return redirect(url_for('purchase_orders'))
        
    # GET
    suppliers = conn.execute('SELECT * FROM suppliers ORDER BY id ASC').fetchall()
    medicines_raw = conn.execute('''
        SELECT i.id, i.brand, i.generic, i.product_name, i.category, i.supplier, s.id as supplier_id
        FROM inventory i
        LEFT JOIN suppliers s ON i.supplier = s.name
        WHERE i.archived_at IS NULL AND i.deleted_at IS NULL
        ORDER BY i.brand ASC
    ''').fetchall()
    medicines = [dict(m) for m in medicines_raw]
    
    # Build structured supplier catalog for supplier-specific category and product filtering
    supplier_catalog = {}
    for s in suppliers:
        s_id = s['id']
        supplier_catalog[s_id] = {
            'id': s_id,
            'name': s['name'],
            'categories': {}
        }
        
    for m in medicines:
        sid = m.get('supplier_id')
        if not sid or sid not in supplier_catalog:
            continue
        cat = m.get('category') or 'General'
        if cat not in supplier_catalog[sid]['categories']:
            supplier_catalog[sid]['categories'][cat] = []
        supplier_catalog[sid]['categories'][cat].append({
            'id': m['id'],
            'brand': m['brand'],
            'generic': m.get('generic') or '',
            'product_name': m.get('product_name') or ''
        })
    
    # Fetch POs with individual items
    orders_raw = conn.execute('''
        SELECT po.id, po.supplier_id, s.name as supplier, po.prepared_by, po.order_date, po.status, po.notes, po.received_by,
               i.brand as medicine_name, poi.quantity as item_qty
        FROM purchase_order_items poi
        JOIN purchase_orders po ON poi.purchase_order_id = po.id
        JOIN suppliers s ON po.supplier_id = s.id
        JOIN inventory i ON poi.medicine_id = i.id
        ORDER BY po.id DESC, poi.id ASC
    ''').fetchall()
    
    orders = []
    for o in orders_raw:
        notes_escaped = (o['notes'] or '').replace("'", "\\'").replace("\n", " ")
        orders.append({
            'id': o['id'],
            'supplier_id': o['supplier_id'],
            'supplier': o['supplier'],
            'prepared_by': o['prepared_by'],
            'order_date': o['order_date'],
            'status': o['status'],
            'notes': notes_escaped,
            'products_list': o['medicine_name'],
            'total_quantity': o['item_qty'],
            'received_by': o['received_by'] or ''
        })
    
    # Fetch detailed PO items for the Details Modal
    po_items_raw = conn.execute('''
        SELECT poi.purchase_order_id, poi.medicine_id, i.brand as medicine_name, i.generic as generic_name, i.category, poi.quantity
        FROM purchase_order_items poi
        JOIN inventory i ON poi.medicine_id = i.id
    ''').fetchall()
    
    po_items_map = {}
    for item in po_items_raw:
        po_id = item['purchase_order_id']
        if po_id not in po_items_map:
            po_items_map[po_id] = []
        po_items_map[po_id].append({
            'medicine_id': item['medicine_id'],
            'medicine_name': item['medicine_name'],
            'generic_name': item['generic_name'],
            'category': item['category'],
            'quantity': item['quantity']
        })
        
    conn.close()
    return render_template('purchase_orders.html', 
                           suppliers=suppliers, 
                           medicines=medicines, 
                           orders=orders, 
                           po_items_map=po_items_map,
                           supplier_catalog=supplier_catalog,
                           active_page='purchase_orders')

@app.route('/purchase_orders/<po_id>/items-json')
def po_items_json(po_id):
    """Return PO items as JSON for the expiry date modal."""
    if 'user_id' not in session:
        from flask import jsonify
        return jsonify({'items': []})
    conn = get_db_connection()
    items = conn.execute('''
        SELECT poi.medicine_id, i.brand as medicine_name, poi.quantity
        FROM purchase_order_items poi
        JOIN inventory i ON poi.medicine_id = i.id
        WHERE poi.purchase_order_id = ?
    ''', (po_id,)).fetchall()
    conn.close()
    from flask import jsonify
    return jsonify({'items': [dict(r) for r in items]})

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
    if not po or po['status'] != 'For Receiving':
        conn.close()
        return redirect(url_for('purchase_orders'))

    receiver = session.get('name', 'Unknown')
    skipped  = request.form.get('skip_expiry') == '1'

    conn.execute("UPDATE purchase_orders SET status = 'Received', received_by = ? WHERE id = ?", (receiver, po_id))
    items = conn.execute("SELECT * FROM purchase_order_items WHERE purchase_order_id = ?", (po_id,)).fetchall()

    for item in items:
        med_id = item['medicine_id']
        qty    = min(MAX_STOCK_LIMIT, max(0, int(item['quantity'])))

        med = conn.execute("SELECT stock, reorder_point, brand FROM inventory WHERE id = ?", (med_id,)).fetchone()
        current_stock = med['stock'] if med else 0
        new_stock = min(MAX_STOCK_LIMIT, current_stock + qty)
        conn.execute("UPDATE inventory SET stock = ? WHERE id = ?", (new_stock, med_id))
        if med:
            new_status = compute_stock_status(new_stock, get_low_stock_threshold(conn))
            conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, med_id))
            if new_stock >= 2500:
                add_notification_with_conn(conn, 'overstock', f"⚠️ Overstock Warning: {med['brand']}", f"Stock level reached {new_stock} units. Storage capacity threshold (2,500 units) exceeded!", '#f59e0b')
                flash(f"⚠️ Storage Capacity Warning: {med['brand']} stock is now {new_stock} units, exceeding storage capacity threshold (2,500 units)!", "warning")

        last_batch = conn.execute("SELECT id FROM batches ORDER BY id DESC LIMIT 1").fetchone()
        if last_batch:
            try:
                last_num = int(last_batch['id'].split('-')[1])
                next_batch_id = f"BAT-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_batch_id = "BAT-004"
        else:
            next_batch_id = "BAT-001"

        if skipped:
            expiry_date    = ''   # blank — user must fill in later
            expiry_pending = 1
        else:
            raw = request.form.get(f'expiry_{med_id}', '').strip()
            if raw:
                try:
                    dt = datetime.strptime(raw, '%Y-%m-%d')
                    expiry_date = dt.strftime('%m/%d/%Y')
                except Exception:
                    expiry_date = raw
                expiry_pending = 0
            else:
                expiry_date = ''
                expiry_pending = 1

        batch_status = compute_batch_expiry_status(expiry_date, get_near_expiry_warning_days(conn), expiry_pending)
        conn.execute(
            "INSERT INTO batches (id, medicine_id, expiry_date, current_qty, status, expiry_pending) VALUES (?, ?, ?, ?, ?, ?)",
            (next_batch_id, med_id, expiry_date, qty, batch_status, expiry_pending)
        )

        last_movement = conn.execute("SELECT id FROM stock_movements ORDER BY id DESC LIMIT 1").fetchone()
        if last_movement:
            try:
                last_num = int(last_movement['id'].split('-')[1])
                next_trn_id = f"TRN-{last_num + 1:03d}"
            except (IndexError, ValueError):
                next_trn_id = "TRN-002"
        else:
            next_trn_id = "TRN-001"

        trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')
        conn.execute(
            "INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (next_trn_id, 'Stock-In', med_id, next_batch_id, qty, trn_date, f'PO {po_id}')
        )

    add_notification_with_conn(conn, 'po', f"PO {po_id} received", "Stock levels have been successfully updated", '#10b981')
    log_activity(conn, 'po_received', 'purchase_order', po_id, f"PO {po_id}", details=f"Received PO items into stock. Batch inventory updated")
    if skipped:
        add_notification_with_conn(conn, 'warning', f"Expiry dates pending for PO {po_id}", "Please set expiry dates in Expiry Monitoring", '#f59e0b')
    conn.commit()
    conn.close()
    return redirect(url_for('purchase_orders'))

@app.route('/batches/<batch_id>/set-expiry', methods=['POST'])
def set_batch_expiry(batch_id):
    """Update expiry date on a pending batch from Expiry Monitoring page."""
    if not has_permission('batches_set_expiry'):
        flash('Access denied. You do not have permission to set or update batch expiry dates.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    expiry_raw = request.form.get('expiry_date', '').strip()
    if expiry_raw:
        try:
            dt = datetime.strptime(expiry_raw, '%Y-%m-%d')
            expiry_date = dt.strftime('%m/%d/%Y')
        except Exception:
            expiry_date = expiry_raw
        batch = conn.execute("SELECT b.id, b.medicine_id, i.brand FROM batches b JOIN inventory i ON b.medicine_id = i.id WHERE b.id = ?", (batch_id,)).fetchone()
        batch_status = compute_batch_expiry_status(expiry_date, get_near_expiry_warning_days(conn), expiry_pending=False)
        conn.execute("UPDATE batches SET expiry_date = ?, expiry_pending = 0, status = ? WHERE id = ?", (expiry_date, batch_status, batch_id))
        log_activity(conn, 'update_expiry', 'Batch', batch_id, brand_name, details=f"Set expiry date to {expiry_date}")
        conn.commit()
        flash(f'Expiry date updated for batch {batch_id}.')
    conn.close()
    return redirect(url_for('expiry_monitoring'))


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
    
    # Delete old items
    conn.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = ?", (po_id,))
    
    # Update PO
    conn.execute('''
        UPDATE purchase_orders 
        SET supplier_id = ?, notes = ?
        WHERE id = ?
    ''', (supplier_id, notes, po_id))
    
    # Insert new items
    for med_id, qty in zip(medicines, quantities):
        if med_id and qty:
            sanitized_qty = min(MAX_STOCK_LIMIT, max(1, int(qty)))
            conn.execute('''
                INSERT INTO purchase_order_items (purchase_order_id, medicine_id, quantity)
                VALUES (?, ?, ?)
            ''', (po_id, med_id, sanitized_qty))
            
    log_activity(conn, 'edit_po', 'purchase_order', po_id, f"PO {po_id}", details=f"Updated supplier/items for PO {po_id}")
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
    if po and po['status'] == 'For Receiving':
        conn.execute("DELETE FROM purchase_order_items WHERE purchase_order_id = ?", (po_id,))
        conn.execute("DELETE FROM purchase_orders WHERE id = ?", (po_id,))
        log_activity(conn, 'delete_po', 'purchase_order', po_id, f"PO {po_id}", details=f"Cancelled and deleted pending PO {po_id}")
        conn.commit()
        flash('Purchase Order deleted successfully!')
    conn.close()
    return redirect(url_for('purchase_orders'))

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
        medicine_id = request.form['medicine_id']
        qty = min(MAX_STOCK_LIMIT, max(1, int(request.form.get('qty', 1))))
        
        # Validate stock
        med = conn.execute("SELECT * FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
        if med and med['stock'] >= qty:
            remaining_to_sell = qty
            batches = conn.execute("SELECT * FROM batches WHERE medicine_id = ? AND current_qty > 0", (medicine_id,)).fetchall()
            
            # Sort in Python by parsing expiry date
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
                conn.execute("UPDATE inventory SET stock = stock - ? WHERE id = ?", (sell_qty, medicine_id))
                
                # Recalculate status using low_stock_threshold
                med_item = conn.execute("SELECT stock FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
                if med_item:
                    new_status = compute_stock_status(med_item['stock'], get_low_stock_threshold(conn))
                    conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, medicine_id))
                
                # Log stock movement
                last_movement = conn.execute("SELECT id FROM stock_movements ORDER BY id DESC LIMIT 1").fetchone()
                if last_movement:
                    try:
                        last_num = int(last_movement['id'].split('-')[1])
                        next_trn_id = f"TRN-{last_num + 1:03d}"
                    except (IndexError, ValueError):
                        next_trn_id = "TRN-002"
                else:
                    next_trn_id = "TRN-001"
                
                trn_date = datetime.now().strftime('%Y-%m-%d %I:%M %p')
                
                # Next SAL ID for reference
                last_sale = conn.execute("SELECT id FROM sales ORDER BY id DESC LIMIT 1").fetchone()
                if last_sale:
                    try:
                        last_num = int(last_sale['id'].split('-')[1])
                        next_sale_id = f"SAL-{last_num + 1:03d}"
                    except (IndexError, ValueError):
                        next_sale_id = "SAL-002"
                else:
                    next_sale_id = "SAL-001"
                
                conn.execute("INSERT INTO stock_movements (id, type, medicine_id, batch_id, quantity, movement_date, reference) VALUES (?, ?, ?, ?, ?, ?, ?)",
                             (next_trn_id, 'Stock-Out', medicine_id, b['id'], -sell_qty, trn_date, f'Sale {next_sale_id}'))
                
                remaining_to_sell -= sell_qty
            
            # Insert Sale
            last_sale = conn.execute("SELECT id FROM sales ORDER BY id DESC LIMIT 1").fetchone()
            if last_sale:
                try:
                    last_num = int(last_sale['id'].split('-')[1])
                    next_sale_id = f"SAL-{last_num + 1:03d}"
                except (IndexError, ValueError):
                    next_sale_id = "SAL-002"
            else:
                next_sale_id = "SAL-001"
            
            sale_date = datetime.now().strftime('%m/%d/%Y')
            sold_by = session.get('name', 'Assistant Pharmacist')
            conn.execute("INSERT INTO sales (id, medicine_id, sale_date, qty, sold_by) VALUES (?, ?, ?, ?, ?)",
                         (next_sale_id, medicine_id, sale_date, qty, sold_by))
            log_activity(conn, 'sale', 'product', medicine_id, med['brand'], details=f"Sold {qty} units (Receipt {next_sale_id})")
            
            # Check updated stock levels for alerts
            updated_med = conn.execute("SELECT brand, stock, reorder_point FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
            if updated_med:
                low_limit = get_low_stock_threshold(conn)
                if updated_med['stock'] == 0:
                    add_notification_with_conn(conn, 'low_stock', f"{updated_med['brand']} is Out of Stock", "Remaining stock is 0 units", '#ef4444')
                elif updated_med['stock'] <= low_limit:
                    add_notification_with_conn(conn, 'low_stock', f"{updated_med['brand']} is Low Stock", f"Only {updated_med['stock']} units left · Low Stock Threshold: {low_limit} (Reorder point: {updated_med['reorder_point']})", '#ef4444')
            
            conn.commit()
        else:
            flash("Insufficient stock or medicine not found!")
        
        conn.close()
        return redirect(url_for('sales'))
        
    medicines = conn.execute("SELECT * FROM inventory").fetchall()
    sales_history = conn.execute('''
        SELECT s.id, s.medicine_id, s.sale_date, i.brand as medicine_name, s.qty, s.sold_by
        FROM sales s
        JOIN inventory i ON s.medicine_id = i.id
        ORDER BY s.id DESC
    ''').fetchall()
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
    
    # Fetch old sale details
    old_sale = conn.execute("SELECT * FROM sales WHERE id = ?", (sale_id,)).fetchone()
    if not old_sale:
        conn.close()
        flash("Sale record not found!")
        return redirect(url_for('sales'))
        
    old_qty = old_sale['qty']
    old_med_id = old_sale['medicine_id']
    
    # If medicine hasn't changed:
    if old_med_id == medicine_id:
        # Check if we have enough stock to cover the difference
        med = conn.execute("SELECT stock, reorder_point, brand FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
        diff = new_qty - old_qty
        if med and med['stock'] >= diff:
            # Update stock
            new_stock = med['stock'] - diff
            low_thresh = get_low_stock_threshold(conn)
            new_status = compute_stock_status(new_stock, low_thresh)
            conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_stock, new_status, medicine_id))
            # Update sale
            conn.execute("UPDATE sales SET qty = ? WHERE id = ?", (new_qty, sale_id))
            log_activity(conn, 'edit_sale', 'sale', sale_id, f"Sale {sale_id}", details=f"Adjusted quantity from {old_qty} to {new_qty} for {med['brand']}")
            conn.commit()
            flash("Sale record updated successfully!")
        else:
            flash("Insufficient stock for the requested quantity change!")
    else:
        # Medicine changed!
        low_thresh = get_low_stock_threshold(conn)
        # Return old stock to old medicine
        old_med = conn.execute("SELECT stock, reorder_point FROM inventory WHERE id = ?", (old_med_id,)).fetchone()
        if old_med:
            old_new_stock = old_med['stock'] + old_qty
            old_new_status = compute_stock_status(old_new_stock, low_thresh)
            conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (old_new_stock, old_new_status, old_med_id))
            
        # Deduct new stock from new medicine
        new_med = conn.execute("SELECT stock, reorder_point, brand FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
        if new_med and new_med['stock'] >= new_qty:
            new_new_stock = new_med['stock'] - new_qty
            new_new_status = compute_stock_status(new_new_stock, low_thresh)
            conn.execute("UPDATE inventory SET stock = ?, status = ? WHERE id = ?", (new_new_stock, new_new_status, medicine_id))
            
            # Update sale
            conn.execute("UPDATE sales SET medicine_id = ?, qty = ? WHERE id = ?", (medicine_id, new_qty, sale_id))
            log_activity(conn, 'edit_sale', 'sale', sale_id, f"Sale {sale_id}", details=f"Changed product to {new_med['brand']} (Qty: {new_qty})")
            conn.commit()
            flash("Sale record updated successfully!")
        else:
            # Rollback old medicine stock return
            conn.rollback()
            flash("Insufficient stock for the new medicine!")
            
    conn.close()
    return redirect(url_for('sales'))

@app.route('/stock_movements')
def stock_movements():
    if not has_permission('movements_view'):
        flash('Access denied. You do not have permission to access Stock Movements.', 'warning')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    items = conn.execute('''
        SELECT sm.id, sm.type, i.brand as medicine, sm.batch_id as batch, sm.quantity, sm.movement_date as date, sm.reference, i.category
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
    med = conn.execute('SELECT stock FROM inventory WHERE id = ?', (batch['medicine_id'],)).fetchone()
    if med:
        new_stock = max(0, med['stock'] - qty)
        new_status = compute_stock_status(new_stock, get_low_stock_threshold(conn))
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
            if password and password != '••••••••':
                if len(password) < 10 or len(password) > 15 or not re.match(r'^[a-zA-Z0-9]{10,15}$', password):
                    conn.close()
                    flash('Password must be between 10 and 15 alphanumeric characters (no special characters allowed).', 'danger')
                    return redirect(url_for('users'))
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
            log_activity(conn, 'edit_user', 'User', user_db_id, name, details=f"Updated user account {username} (Role: {role})")
            conn.commit()
            
            # Update session details if editing own profile
            if int(user_db_id) == session.get('user_id'):
                session['name'] = name
                session['role'] = role
            flash('User details updated successfully!')
        else:
            # Validate Password for new user: 10 to 15 alphanumeric chars
            if len(password) < 10 or len(password) > 15 or not re.match(r'^[a-zA-Z0-9]{10,15}$', password):
                conn.close()
                flash('Password must be between 10 and 15 alphanumeric characters (no special characters allowed).', 'danger')
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


@app.route('/reports/download', methods=['POST'])
def download_report():
    if not has_permission('reports_view'):
        flash('Access denied. You do not have permission to download reports.', 'warning')
        return redirect(url_for('dashboard'))
    
    try:
        report_type = request.form.get('report_type', 'inventory')
        frequency = request.form.get('frequency')
        period_week = request.form.get('period_week')
        period_month = request.form.get('period_month')
        period_year = request.form.get('period_year')
        file_format = request.form.get('format', 'pdf')
        
        start_dt, end_dt = parse_date_range(frequency, period_week, period_month, period_year)
        
        conn = get_db_connection()
        
        headers = []
        rows = []
        title = ""
        
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

        if report_type == 'inventory':
            title = "Inventory Stock Report"
            headers = ["ID", "Generic Name", "Brand Name", "Category", "Stock", "Reorder Point", "Status"]
            items = conn.execute("SELECT * FROM inventory WHERE archived_at IS NULL AND deleted_at IS NULL ORDER BY id ASC").fetchall()
            for item in items:
                rows.append([item['id'], item['generic'], item['brand'], item['category'], item['stock'], item['reorder_point'], item['status']])
                
        elif report_type == 'sales':
            title = "Sales Report"
            headers = ["Sale ID", "Medicine", "Qty Sold", "Sold By", "Date"]
            sales_data = conn.execute("""
                SELECT s.id, i.brand as medicine, s.qty, s.sold_by, s.sale_date 
                FROM sales s 
                JOIN inventory i ON s.medicine_id = i.id
                ORDER BY s.id DESC
            """).fetchall()
            for item in sales_data:
                if is_in_date_range(item['sale_date']):
                    rows.append([item['id'], item['medicine'], item['qty'], item['sold_by'], item['sale_date']])
                
        elif report_type == 'stock_movements':
            title = "Stock Movements Report"
            headers = ["Movement ID", "Type", "Medicine", "Batch", "Qty", "Date", "Reference"]
            movements = conn.execute("""
                SELECT sm.id, sm.type, i.brand as medicine, sm.batch_id, sm.quantity, sm.movement_date, sm.reference
                FROM stock_movements sm
                JOIN inventory i ON sm.medicine_id = i.id
                ORDER BY sm.id DESC
            """).fetchall()
            for item in movements:
                if is_in_date_range(item['movement_date']):
                    rows.append([item['id'], item['type'], item['medicine'], item['batch_id'], item['quantity'], item['movement_date'], item['reference']])
                
        elif report_type == 'expiry_monitoring':
            title = "Expiry Monitoring Report"
            headers = ["Batch ID", "Medicine", "Expiry Date", "Current Qty", "Status"]
            batches = conn.execute("""
                SELECT b.id, i.brand as medicine, b.expiry_date, b.current_qty, b.status
                FROM batches b
                JOIN inventory i ON b.medicine_id = i.id
                ORDER BY b.id DESC
            """).fetchall()
            for item in batches:
                if is_in_date_range(item['expiry_date']):
                    rows.append([item['id'], item['medicine'], item['expiry_date'], item['current_qty'], item['status']])
                
        elif report_type == 'suppliers':
            title = "Suppliers Report"
            headers = ["ID", "Supplier Name", "Contact Person", "Phone", "Email", "Address"]
            suppliers = conn.execute("SELECT * FROM suppliers").fetchall()
            for item in suppliers:
                rows.append([item['id'], item['name'], item['contact_person'], item['phone'], item['email'], item['address']])
                
        conn.close()
        filename = f"{report_type}_report"

        download_time_str = get_current_ph_time().strftime('%B %d, %Y at %I:%M %p')

        if file_format == 'excel':
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Report"
            
            ws.append([title])
            ws.append([f"Downloaded on: {download_time_str}"])
            if start_dt and end_dt:
                ws.append([f"Period: {start_dt.strftime('%m/%d/%Y')} - {end_dt.strftime('%m/%d/%Y')}"])
            ws.append([])
            
            ws.append(headers)
            for r in rows:
                ws.append(r)
                
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            return send_file(
                output,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=f"{filename}.xlsx"
            )
            
        elif file_format == 'pdf':
            period_text = f"Period: {start_dt.strftime('%m/%d/%Y')} - {end_dt.strftime('%m/%d/%Y')}" if (start_dt and end_dt) else ""
            pdf = PDFTemplateReport(title, period_text)
            pdf.alias_nb_pages()
            pdf.add_page()
            
            pdf.set_font("helvetica", "B", 13)
            pdf.cell(0, 8, title.upper(), align="C")
            pdf.ln(6)
            
            pdf.set_font("helvetica", "I", 8.5)
            pdf.set_text_color(100, 116, 139)
            pdf.cell(0, 5, f"Downloaded on: {download_time_str}", align="C")
            pdf.ln(5)
            
            if period_text:
                pdf.cell(0, 5, period_text, align="C")
                pdf.ln(5)
                
            pdf.set_text_color(0, 0, 0)
            pdf.ln(3)

            # Tailored column widths mapping per report type (Total 190 mm printable width)
            col_widths_map = {
                'inventory': [14, 40, 38, 40, 18, 20, 20],
                'sales': [28, 56, 22, 48, 36],
                'stock_movements': [22, 22, 42, 24, 16, 36, 28],
                'expiry_monitoring': [30, 65, 35, 30, 30],
                'suppliers': [16, 40, 35, 28, 36, 35],
            }
            if report_type in col_widths_map and len(col_widths_map[report_type]) == len(headers):
                col_widths = col_widths_map[report_type]
            else:
                col_widths = [190 / len(headers)] * len(headers)

            # Auto-scaling cell renderer: shrinks font size to fit exactly inside the cell box if text is too long
            def draw_fit_cell(w, h, text, base_font_size=8.5, min_font_size=5.2, is_header=False, fill_color=None, text_color=None):
                text_str = str(text) if text is not None else ""
                avail_w = max(w - 2.5, 2.0)
                font_weight = "B" if is_header else ""
                font_size = base_font_size
                
                pdf.set_font("helvetica", font_weight, font_size)
                str_w = pdf.get_string_width(text_str)
                if str_w > avail_w:
                    # Proportionally scale font down to match box width
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
                    draw_fit_cell(col_widths[idx], 8, str(h), base_font_size=9.0, min_font_size=6.0, is_header=True, fill_color=(192, 57, 43), text_color=(255, 255, 255))
                pdf.ln()

            draw_header_row()
            
            fill = False
            if not rows:
                pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(100, 116, 139)
                pdf.set_font("helvetica", "I", 9)
                pdf.cell(190, 10, "No records found for the selected period.", border=1, align="C", fill=True)
                pdf.ln()
            else:
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
                
            output = BytesIO(pdf_data)
            output.seek(0)
            
            return send_file(
                output,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=f"{filename}.pdf"
            )
            
        return "Invalid format", 400
    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"Report generation failed: {str(e)}", 500

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
            next_sale_id = f"SAL-{last_num + 1:03d}"
        except (IndexError, ValueError):
            next_sale_id = "SAL-001"
    else:
        next_sale_id = "SAL-001"
        
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
                     (next_trn_id, 'Stock-Out', medicine_id, b['id'], -sell_qty, trn_date, trn_ref))
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
                     (next_trn_id, 'Stock-Out', medicine_id, 'N/A', -remaining_to_sell, trn_date, trn_ref))

    # Recalculate status using low_stock_threshold
    med_item = conn.execute("SELECT stock FROM inventory WHERE id = ?", (medicine_id,)).fetchone()
    if med_item:
        new_status = compute_stock_status(med_item['stock'], get_low_stock_threshold(conn))
        conn.execute("UPDATE inventory SET status = ? WHERE id = ?", (new_status, medicine_id))

    # Insert Sale
    conn.execute("INSERT INTO sales (id, medicine_id, sale_date, qty, sold_by) VALUES (?, ?, ?, ?, ?)",
                 (next_sale_id, medicine_id, sale_date_str, qty, sold_by))
    
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
