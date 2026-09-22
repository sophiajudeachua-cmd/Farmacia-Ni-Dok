import sqlite3
import os
import sys

def get_db_path():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, 'farmacia.db')

def get_db_connection():
    conn = sqlite3.connect(get_db_path())
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    # Create tables
    c.executescript('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            name TEXT NOT NULL,
            user_id TEXT,
            contact TEXT,
            email TEXT,
            license_no TEXT,
            joined_date TEXT,
            address TEXT,
            is_online INTEGER DEFAULT 0,
            last_login TEXT DEFAULT NULL,
            last_logout TEXT DEFAULT NULL,
            last_seen TEXT DEFAULT NULL,
            is_active INTEGER DEFAULT 1,
            deactivated_at TEXT DEFAULT NULL,
            permissions TEXT DEFAULT NULL
        );
        CREATE TABLE IF NOT EXISTS suppliers (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            address TEXT NOT NULL,
            contact TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS inventory (
            id TEXT PRIMARY KEY,
            generic TEXT,
            brand TEXT NOT NULL,
            category TEXT NOT NULL,
            stock INTEGER NOT NULL,
            reorder_point INTEGER NOT NULL,
            status TEXT NOT NULL,
            product_name TEXT NOT NULL,
            description TEXT NOT NULL,
            unit_of_measure TEXT NOT NULL,
            purchase_price REAL NOT NULL,
            supplier TEXT NOT NULL,
            archived_at TEXT DEFAULT NULL,
            deleted_at TEXT DEFAULT NULL
        );
        CREATE TABLE IF NOT EXISTS batches (
            id TEXT PRIMARY KEY,
            medicine_id TEXT NOT NULL,
            expiry_date TEXT,
            current_qty INTEGER NOT NULL,
            status TEXT NOT NULL,
            expiry_pending INTEGER DEFAULT 0,
            FOREIGN KEY (medicine_id) REFERENCES inventory (id)
        );
        CREATE TABLE IF NOT EXISTS purchase_orders (
            id TEXT PRIMARY KEY,
            supplier_id TEXT NOT NULL,
            prepared_by TEXT NOT NULL,
            order_date TEXT NOT NULL,
            status TEXT NOT NULL,
            notes TEXT,
            received_by TEXT,
            FOREIGN KEY (supplier_id) REFERENCES suppliers (id)
        );
        CREATE TABLE IF NOT EXISTS purchase_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            purchase_order_id TEXT NOT NULL,
            medicine_id TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            FOREIGN KEY (purchase_order_id) REFERENCES purchase_orders (id),
            FOREIGN KEY (medicine_id) REFERENCES inventory (id)
        );
        CREATE TABLE IF NOT EXISTS sales (
            id TEXT PRIMARY KEY,
            medicine_id TEXT NOT NULL,
            sale_date TEXT NOT NULL,
            qty INTEGER NOT NULL,
            sold_by TEXT NOT NULL,
            FOREIGN KEY (medicine_id) REFERENCES inventory (id)
        );
        CREATE TABLE IF NOT EXISTS stock_movements (
            id TEXT PRIMARY KEY,
            type TEXT NOT NULL,
            medicine_id TEXT NOT NULL,
            batch_id TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            movement_date TEXT NOT NULL,
            reference TEXT NOT NULL,
            FOREIGN KEY (medicine_id) REFERENCES inventory (id),
            FOREIGN KEY (batch_id) REFERENCES batches (id)
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            title TEXT NOT NULL,
            subtitle TEXT NOT NULL,
            color TEXT NOT NULL,
            created_at TEXT NOT NULL,
            seen INTEGER DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT NOT NULL,
            target_type TEXT NOT NULL,
            target_id TEXT NOT NULL,
            target_name TEXT NOT NULL,
            performed_by TEXT NOT NULL,
            performed_at TEXT NOT NULL,
            details TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS disposals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            batch_id TEXT NOT NULL,
            medicine_id TEXT NOT NULL,
            medicine_name TEXT NOT NULL,
            qty_disposed INTEGER NOT NULL,
            reason TEXT DEFAULT 'Expired',
            disposed_by TEXT NOT NULL,
            disposed_at TEXT NOT NULL,
            notes TEXT DEFAULT ''
        );
    ''')

    # Seed data if empty
    c.execute('SELECT COUNT(*) FROM users')
    if c.fetchone()[0] == 0:
        c.execute("""
            INSERT INTO users (username, password, role, name, user_id, contact, email, license_no, joined_date, address) 
            VALUES ('admin', 'admin', 'Superadmin', 'Super Admin', 'USR-001', '0917-888-8888', 'owner@farmacianidok.com', 'RPH-2023-0001', '1/10/2023', 'San Marcos, Calumpit, Bulacan')
        """)
        c.execute("""
            INSERT INTO users (username, password, role, name, user_id, contact, email, license_no, joined_date, address) 
            VALUES ('assistant', 'assistant', 'Admin', 'Assistant Pharmacist', 'USR-002', '0917-777-7777', 'staff@farmacianidok.com', 'N/A', '6/1/2024', 'Calumpit, Bulacan')
        """)
        
        # Suppliers
        suppliers = [
            ('SUP-001', 'ABC Pharma Supply', 'Calumpit, Bulacan', '0917-111-1111'),
            ('SUP-002', 'MediCore Distributors', 'Malolos, Bulacan', '0917-222-2222'),
        ]
        c.executemany("INSERT INTO suppliers VALUES (?, ?, ?, ?)", suppliers)

        # Inventory
        inventory = [
            # Medicines
            ('MED-001', 'Paracetamol', 'Biogesic', 'Medicines', 30, 50, 'Low Stock', 'Paracetamol', 'Strength: 500 mg, Dosage Form: Tablet, Pack Size: Box of 100 tablets, Rx Required: No', 'Box', 150.00, 'ABC Pharma Supply'),
            ('MED-002', 'Amoxicillin', 'Amoxil', 'Medicines', 18, 40, 'Low Stock', 'Amoxicillin', 'Strength: 500 mg, Dosage Form: Capsule, Pack Size: Box of 100 capsules, Rx Required: Yes', 'Box', 350.00, 'MediCore Distributors'),
            ('MED-003', 'Carbocisteine', 'Solmux', 'Medicines', 12, 30, 'Low Stock', 'Carbocisteine', 'Strength: 500 mg, Dosage Form: Capsule, Pack Size: Box of 100 capsules, Rx Required: No', 'Box', 220.00, 'ABC Pharma Supply'),
            ('MED-004', 'Salbutamol', 'Ventolin', 'Medicines', 60, 20, 'Good', 'Salbutamol', 'Strength: 100 mcg, Dosage Form: Inhaler, Pack Size: 200 doses, Rx Required: Yes', 'Piece', 450.00, 'MediCore Distributors'),
            
            # Vitamins & Supplements
            ('MED-005', 'Multivitamins', 'Centrum', 'Vitamins & Supplements', 80, 30, 'Good', 'Multivitamins', 'Supplement Type: Multivitamins, Strength: N/A, Dosage Form: Tablet, Pack Size: Bottle of 60, Age Group: Adult', 'Bottle', 650.00, 'ABC Pharma Supply'),
            ('MED-006', 'Multivitamins + Vitamin C', 'Enervon', 'Vitamins & Supplements', 45, 20, 'Good', 'Multivitamins + Vitamin C', 'Supplement Type: Multivitamins + Vitamin C, Strength: N/A, Dosage Form: Tablet, Pack Size: Box of 30, Age Group: Adult', 'Box', 300.00, 'MediCore Distributors'),
            ('MED-007', 'Sodium Ascorbate', 'Fern-C', 'Vitamins & Supplements', 10, 50, 'Low Stock', 'Sodium Ascorbate', 'Supplement Type: Sodium Ascorbate, Strength: 500 mg, Dosage Form: Capsule, Pack Size: Box of 100, Age Group: Adult', 'Box', 420.00, 'ABC Pharma Supply'),
            ('MED-008', 'Vitamin C', 'Ceelin', 'Vitamins & Supplements', 25, 15, 'Good', 'Vitamin C', 'Supplement Type: Vitamin C, Strength: 100 mg/5 mL, Dosage Form: Syrup, Pack Size: 120 mL, Age Group: Children', 'Bottle', 150.00, 'MediCore Distributors'),

            # Medical Supplies
            ('MED-009', 'N/A', 'BD', 'Medical Supplies', 150, 50, 'Good', 'Syringe', 'Material: Plastic, Sterile/Non-Sterile: Yes, Size: 3 mL, Unit Type: Piece, Disposable: Yes', 'Piece', 15.00, 'ABC Pharma Supply'),
            ('MED-010', 'N/A', 'Surgitech', 'Medical Supplies', 40, 20, 'Good', 'Surgical Gloves', 'Material: Latex, Sterile/Non-Sterile: Yes, Size: Medium, Unit Type: Box (100 pcs), Disposable: Yes', 'Box', 380.00, 'MediCore Distributors'),
            ('MED-011', 'N/A', 'Mediplast', 'Medical Supplies', 8, 15, 'Low Stock', 'Gauze Pads', 'Material: Cotton, Sterile/Non-Sterile: Yes, Size: 4x4 in, Unit Type: Pack, Disposable: Yes', 'Pack', 50.00, 'ABC Pharma Supply'),
            ('MED-012', 'N/A', '3M', 'Medical Supplies', 35, 10, 'Good', 'Micropore Tape', 'Material: Paper, Sterile/Non-Sterile: N/A, Size: 1 inch, Unit Type: Roll, Disposable: Yes', 'Roll', 80.00, 'MediCore Distributors'),

            # First Aid
            ('MED-013', 'Povidone-Iodine', 'Betadine', 'First Aid', 100, 30, 'Good', 'Povidone-Iodine', 'Product Type: Antiseptic, Volume/Size: 120 mL, Sterile: Yes, Intended Use: Wound Cleaning', 'Bottle', 140.00, 'ABC Pharma Supply'),
            ('MED-014', 'Burn Gel', 'Burnaid', 'First Aid', 5, 10, 'Low Stock', 'Burn Gel', 'Product Type: Burn Gel, Volume/Size: 25 g, Sterile: Yes, Intended Use: Burn Relief', 'Tube', 280.00, 'MediCore Distributors'),
            ('MED-015', 'Adhesive Bandages', 'Band-Aid', 'First Aid', 200, 50, 'Good', 'Adhesive Bandages', 'Product Type: Adhesive Bandages, Volume/Size: Standard, Sterile: Yes, Intended Use: Minor Cuts', 'Pack', 95.00, 'ABC Pharma Supply'),

            # Personal Care
            ('MED-016', 'Toothpaste', 'Colgate', 'Personal Care', 60, 20, 'Good', 'Toothpaste', 'Variant/Scent: Total, Volume/Weight: 150 g, Skin/Hair Type: N/A, Pack Size: Tube', 'Tube', 120.00, 'ABC Pharma Supply'),
            ('MED-017', 'Shampoo', 'Head & Shoulders', 'Personal Care', 35, 15, 'Good', 'Shampoo', 'Variant/Scent: Cool Menthol, Volume/Weight: 170 mL, Skin/Hair Type: Normal Hair, Pack Size: Bottle', 'Bottle', 160.00, 'MediCore Distributors'),
            ('MED-018', 'Deodorant', 'Rexona', 'Personal Care', 8, 20, 'Low Stock', 'Deodorant', 'Variant/Scent: Powder Dry, Volume/Weight: 50 mL, Skin/Hair Type: N/A, Pack Size: Roll-on', 'Piece', 95.00, 'ABC Pharma Supply'),

            # Baby Care
            ('MED-019', 'Baby Diapers', 'Pampers', 'Baby Care', 50, 15, 'Good', 'Baby Diapers', 'Age Range: Newborn, Size: Small, Weight/Volume: 3-8 kg, Pack Quantity: 44 pcs', 'Pack', 450.00, 'MediCore Distributors'),
            ('MED-020', 'Baby Powder', "Johnson's", 'Baby Care', 30, 10, 'Good', 'Baby Powder', 'Age Range: 0+ months, Size: S, Weight/Volume: 100 g, Pack Quantity: 1 bottle', 'Bottle', 85.00, 'ABC Pharma Supply'),
            ('MED-021', 'Growing-Up Milk', 'Enfagrow', 'Baby Care', 4, 10, 'Low Stock', 'Growing-Up Milk', 'Age Range: 1-3 years, Size: Stage 3, Weight/Volume: 900 g, Pack Quantity: 1 can', 'Piece', 1150.00, 'MediCore Distributors'),

            # Medical Devices
            ('MED-022', 'Digital Thermometer', 'Omron', 'Medical Devices', 15, 5, 'Good', 'Digital Thermometer', 'Device Type: Thermometer, Model: MC-246, Measurement Range: 32-42 C, Power Source: Battery, Warranty: 1 Year', 'Piece', 299.00, 'ABC Pharma Supply'),
            ('MED-023', 'Blood Pressure Monitor', 'Omron', 'Medical Devices', 6, 4, 'Good', 'Blood Pressure Monitor', 'Device Type: Blood Pressure Monitor, Model: HEM-7120, Measurement Range: Standard, Power Source: Battery, Warranty: 2 Years', 'Piece', 2490.00, 'MediCore Distributors'),
            ('MED-024', 'Pulse Oximeter', 'Rossmax', 'Medical Devices', 2, 5, 'Low Stock', 'Pulse Oximeter', 'Device Type: Pulse Oximeter, Model: SB220, Measurement Range: 35-100%, Power Source: Battery, Warranty: 1 Year', 'Piece', 1250.00, 'ABC Pharma Supply'),

            # Skin Care
            ('MED-025', 'Acne Gel', 'Benzac', 'Skin Care', 15, 8, 'Good', 'Acne Gel', 'Skin Type: Oily Skin, Active Ingredient: Benzoyl Peroxide 5%, SPF: N/A, Volume/Weight: 30 g', 'Tube', 850.00, 'MediCore Distributors'),
            ('MED-026', 'Sunscreen', 'Belo', 'Skin Care', 22, 10, 'Good', 'Sunscreen', 'Skin Type: All Skin Types, Active Ingredient: SPF 40, SPF: 40, Volume/Weight: 50 mL', 'Tube', 380.00, 'ABC Pharma Supply'),
            ('MED-027', 'Moisturizing Cream', 'Cetaphil', 'Skin Care', 3, 6, 'Low Stock', 'Moisturizing Cream', 'Skin Type: Dry Skin, Active Ingredient: Glycerin, SPF: N/A, Volume/Weight: 100 g', 'Tube', 650.00, 'MediCore Distributors'),

            # Eye & Ear Care
            ('MED-028', 'Eye Drops', 'Rohto', 'Eye & Ear Care', 40, 15, 'Good', 'Eye Drops', 'Application Area: Eye, Sterile: Yes, Volume: 13 mL, Multi-dose/Single-dose: Multi-dose', 'Bottle', 180.00, 'ABC Pharma Supply'),
            ('MED-029', 'Ear Drops', 'Otex', 'Eye & Ear Care', 18, 10, 'Good', 'Ear Drops', 'Application Area: Ear, Sterile: Yes, Volume: 10 mL, Multi-dose/Single-dose: Multi-dose', 'Bottle', 350.00, 'MediCore Distributors'),
            ('MED-030', 'Artificial Tears', 'Refresh', 'Eye & Ear Care', 4, 12, 'Low Stock', 'Artificial Tears', 'Application Area: Eye, Sterile: Yes, Volume: 15 mL, Multi-dose/Single-dose: Multi-dose', 'Bottle', 480.00, 'ABC Pharma Supply'),

            # Health & Wellness
            ('MED-031', 'Nutrition Shake', 'Ensure', 'Health & Wellness', 24, 10, 'Good', 'Nutrition Shake', 'Product Type: Nutrition Shake, Flavor/Variant: Vanilla, Weight/Volume: 400 g, Age Group: Adult', 'Piece', 950.00, 'MediCore Distributors'),
            ('MED-032', 'Hand Sanitizer', 'Green Cross', 'Health & Wellness', 80, 20, 'Good', 'Hand Sanitizer', 'Product Type: Hand Sanitizer, Flavor/Variant: Lemon, Weight/Volume: 500 mL, Age Group: General', 'Bottle', 120.00, 'ABC Pharma Supply'),
            ('MED-033', 'ORS', 'Hydrite', 'Health & Wellness', 5, 30, 'Low Stock', 'ORS', 'Product Type: Oral Rehydration Salts, Flavor/Variant: Orange, Weight/Volume: 5.6 g Sachet, Age Group: General', 'Sachet', 15.00, 'MediCore Distributors'),
            
            # Additional Diaper Products
            ('MED-034', 'Baby Diapers', 'Huggies', 'Baby Care', 45, 10, 'Good', 'Baby Diapers', 'Size: M', 'Pack', 380.00, 'ABC Pharma Supply'),
            ('MED-035', 'Baby Diapers', 'Huggies', 'Baby Care', 30, 10, 'Good', 'Baby Diapers', 'Size: L', 'Pack', 410.00, 'MediCore Distributors'),
            ('MED-036', 'Baby Diapers', 'Huggies', 'Baby Care', 25, 10, 'Good', 'Baby Diapers', 'Size: XL', 'Pack', 450.00, 'ABC Pharma Supply'),
            
            # Additional Powder Products
            ('MED-037', 'Baby Powder', "Johnson's", 'Baby Care', 40, 10, 'Good', 'Baby Powder', 'Age Range: 0+ months, Size: M, Weight/Volume: 200 g, Pack Quantity: 1 bottle', 'Bottle', 150.00, 'ABC Pharma Supply'),
            ('MED-038', 'Baby Powder', "Johnson's", 'Baby Care', 20, 10, 'Good', 'Baby Powder', 'Age Range: 0+ months, Size: L, Weight/Volume: 500 g, Pack Quantity: 1 bottle', 'Bottle', 280.00, 'MediCore Distributors')
        ]
        # Reorder point is 100 for all products; status evaluated against low_stock_threshold (40)
        inventory_extended = [tuple(list(row[:5]) + [100, ('No Stock' if row[4] == 0 else ('Low Stock' if row[4] <= 40 else 'Good'))] + list(row[7:]) + [None, None]) for row in inventory]
        c.executemany("INSERT INTO inventory VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", inventory_extended)

        # Batches
        batches = [
            ('BAT-001', 'MED-001', '9/15/2026', 30, 'Good'),
            ('BAT-002', 'MED-002', '8/10/2026', 18, 'Near Expiry'),
            ('BAT-003', 'MED-003', '9/20/2026', 12, 'Near Expiry'),
            ('BAT-004', 'MED-004', '11/12/2026', 60, 'Good'),
            ('BAT-005', 'MED-005', '12/24/2027', 80, 'Good'),
            ('BAT-006', 'MED-006', '8/15/2027', 45, 'Good'),
            ('BAT-007', 'MED-007', '10/10/2026', 10, 'Near Expiry'),
            ('BAT-008', 'MED-008', '10/30/2026', 25, 'Good'),
            ('BAT-009', 'MED-009', '1/15/2028', 150, 'Good'),
            ('BAT-010', 'MED-010', '9/20/2027', 40, 'Good'),
            ('BAT-011', 'MED-011', '8/15/2026', 8, 'Near Expiry'),
            ('BAT-012', 'MED-012', '12/10/2027', 35, 'Good'),
            ('BAT-013', 'MED-013', '7/18/2027', 100, 'Good'),
            ('BAT-014', 'MED-014', '11/12/2026', 5, 'Near Expiry'),
            ('BAT-015', 'MED-015', '11/25/2027', 200, 'Good'),
            ('BAT-016', 'MED-016', '10/10/2027', 60, 'Good'),
            ('BAT-017', 'MED-017', '8/12/2027', 35, 'Good'),
            ('BAT-018', 'MED-018', '9/18/2026', 8, 'Near Expiry'),
            ('BAT-019', 'MED-019', '12/15/2027', 50, 'Good'),
            ('BAT-020', 'MED-020', '9/10/2027', 30, 'Good'),
            ('BAT-021', 'MED-021', '10/28/2026', 4, 'Near Expiry'),
            ('BAT-022', 'MED-022', '6/15/2028', 15, 'Good'),
            ('BAT-023', 'MED-023', '5/20/2028', 6, 'Good'),
            ('BAT-024', 'MED-024', '8/24/2026', 2, 'Near Expiry'),
            ('BAT-025', 'MED-025', '11/15/2027', 15, 'Good'),
            ('BAT-026', 'MED-026', '8/22/2027', 22, 'Good'),
            ('BAT-027', 'MED-027', '11/05/2026', 3, 'Near Expiry'),
            ('BAT-028', 'MED-028', '12/10/2027', 40, 'Good'),
            ('BAT-029', 'MED-029', '9/15/2027', 18, 'Good'),
            ('BAT-030', 'MED-030', '9/30/2026', 4, 'Near Expiry'),
            ('BAT-031', 'MED-031', '10/15/2027', 24, 'Good'),
            ('BAT-032', 'MED-032', '7/20/2027', 80, 'Good'),
            ('BAT-033', 'MED-033', '10/01/2026', 5, 'Near Expiry'),
            ('BAT-034', 'MED-034', '12/20/2027', 45, 'Good'),
            ('BAT-035', 'MED-035', '12/20/2027', 30, 'Good'),
            ('BAT-036', 'MED-036', '12/20/2027', 25, 'Good'),
            ('BAT-037', 'MED-037', '12/20/2027', 40, 'Good'),
            ('BAT-038', 'MED-038', '12/20/2027', 20, 'Good')
        ]
        batches_extended = [tuple(list(row) + [0]) for row in batches]
        c.executemany("INSERT INTO batches VALUES (?, ?, ?, ?, ?, ?)", batches_extended)
        
        # Purchase Orders
        purchase_orders = [
            ('PO-001', 'SUP-001', 'Owner', '3/20/2026', 'For Receiving', 'Urgent stock update', None),
            ('PO-002', 'SUP-002', 'Owner', '3/19/2026', 'Received', 'Regular monthly replenishment', 'Owner')
        ]
        c.executemany("INSERT INTO purchase_orders VALUES (?, ?, ?, ?, ?, ?, ?)", purchase_orders)

        # Purchase Order Items
        po_items = [
            (None, 'PO-001', 'MED-001', 100),
            (None, 'PO-001', 'MED-002', 50),
            (None, 'PO-002', 'MED-003', 80)
        ]
        c.executemany("INSERT INTO purchase_order_items VALUES (?, ?, ?, ?)", po_items)

        # Sales
        sales = [
            ('SAL-001', 'MED-009', '3/20/2026', 5, 'Assistant Pharmacist'),
            ('SAL-002', 'MED-015', '3/21/2026', 12, 'Assistant Pharmacist'),
            ('SAL-003', 'MED-009', '3/22/2026', 8, 'Assistant Pharmacist'),
            ('SAL-004', 'MED-015', '3/23/2026', 15, 'Assistant Pharmacist'),
            ('SAL-005', 'MED-015', '3/24/2026', 20, 'Assistant Pharmacist'),
            ('SAL-006', 'MED-009', '3/25/2026', 18, 'Assistant Pharmacist'),
            ('SAL-007', 'MED-009', '2/15/2026', 15, 'Assistant Pharmacist'),
            ('SAL-008', 'MED-015', '1/10/2026', 10, 'Assistant Pharmacist'),
            ('SAL-009', 'MED-015', '11/20/2025', 25, 'Assistant Pharmacist'),
            ('SAL-010', 'MED-015', '5/15/2024', 30, 'Assistant Pharmacist')
        ]
        c.executemany("INSERT INTO sales VALUES (?, ?, ?, ?, ?)", sales)

        # Stock Movements
        movements = [
            ('TRN-001', 'Stock-Out', 'MED-009', 'BAT-009', -5, '2026-03-20 10:15 AM', 'Sale SAL-001'),
            ('TRN-002', 'Stock-Out', 'MED-015', 'BAT-015', -12, '2026-03-21 02:30 PM', 'Sale SAL-002'),
            ('TRN-003', 'Stock-Out', 'MED-009', 'BAT-009', -8, '2026-03-22 11:00 AM', 'Sale SAL-003'),
            ('TRN-004', 'Stock-Out', 'MED-015', 'BAT-015', -15, '2026-03-23 04:15 PM', 'Sale SAL-004'),
            ('TRN-005', 'Stock-Out', 'MED-015', 'BAT-015', -20, '2026-03-24 09:45 AM', 'Sale SAL-005'),
            ('TRN-006', 'Stock-Out', 'MED-009', 'BAT-009', -18, '2026-03-25 01:20 PM', 'Sale SAL-006'),
            ('TRN-007', 'Stock-Out', 'MED-009', 'BAT-009', -15, '2026-02-15 10:15 AM', 'Sale SAL-007'),
            ('TRN-008', 'Stock-Out', 'MED-015', 'BAT-015', -10, '2026-01-10 02:30 PM', 'Sale SAL-008'),
            ('TRN-009', 'Stock-Out', 'MED-015', 'BAT-015', -25, '2025-11-20 11:00 AM', 'Sale SAL-009'),
            ('TRN-010', 'Stock-Out', 'MED-015', 'BAT-015', -30, '2024-05-15 04:15 PM', 'Sale SAL-010')
        ]
        c.executemany("INSERT INTO stock_movements VALUES (?, ?, ?, ?, ?, ?, ?)", movements)
        
        # Settings
        default_settings = [
            ('pharmacy_name', 'Farmacia ni Dok'),
            ('address', 'San Marcos, Calumpit, Bulacan'),
            ('contact', '0917-000-0000'),
            ('email', 'farmacianidok@gmail.com'),
            ('dark_mode', 'false'),
            ('font_size', 'Medium'),
            ('high_contrast', 'false'),
            ('reduce_motion', 'false'),
            ('screen_reader', 'false'),
            ('show_generic', 'true'),
            ('low_stock_alerts', 'true'),
            ('expiry_alerts', 'true'),
            ('po_alerts', 'true'),
            ('sales_notifications', 'false'),
            ('low_stock_threshold', '40'),
            ('near_expiry_warning', '90'),
            ('default_currency', 'PHP'),
            ('auto_logout', '30'),
            ('require_password', 'false')
        ]
        c.executemany("INSERT INTO settings VALUES (?, ?)", default_settings)
        
        # Default Notifications
        default_notifications = [
            (None, 'low_stock', 'Biogesic is Low Stock', 'Only 30 units left · Reorder point: 50', '#ef4444', '2026-03-22 10:00 AM', 0),
            (None, 'low_stock', 'Amoxil is Low Stock', 'Only 18 units left · Reorder point: 40', '#ef4444', '2026-03-22 10:01 AM', 0),
            (None, 'low_stock', 'Zyrtec is Low Stock', 'Only 12 units left · Reorder point: 30', '#ef4444', '2026-03-22 10:02 AM', 0),
            (None, 'expiry', 'Amoxil — Near Expiry', 'Expires in 71 days · Batch BAT-002', '#ffb800', '2026-03-22 10:03 AM', 0),
            (None, 'expiry', 'Zyrtec — Near Expiry', 'Expires in 29 days · Batch BAT-003', '#ffb800', '2026-03-22 10:04 AM', 0),
            (None, 'po', 'PO PO-001 pending receipt', 'From ABC Pharma Supply · 2 items · For Receiving', '#3b82f6', '2026-03-22 10:05 AM', 0)
        ]
        c.executemany("INSERT INTO notifications VALUES (?, ?, ?, ?, ?, ?, ?)", default_notifications)

    # Automatic column migrations for existing databases
    try:
        user_cols = [r[1] for r in c.execute("PRAGMA table_info(users)").fetchall()]
        if 'is_online' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN is_online INTEGER DEFAULT 0")
        if 'last_login' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN last_login TEXT DEFAULT NULL")
        if 'last_logout' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN last_logout TEXT DEFAULT NULL")
        if 'last_seen' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN last_seen TEXT DEFAULT NULL")
        if 'is_active' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN is_active INTEGER DEFAULT 1")
        if 'deactivated_at' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN deactivated_at TEXT DEFAULT NULL")
        if 'permissions' not in user_cols:
            c.execute("ALTER TABLE users ADD COLUMN permissions TEXT DEFAULT NULL")

        act_cols = [r[1] for r in c.execute("PRAGMA table_info(activity_log)").fetchall()]
        if 'details' not in act_cols:
            c.execute("ALTER TABLE activity_log ADD COLUMN details TEXT DEFAULT ''")
    except Exception as e:
        print("Schema migration error:", e)

    conn.commit()
    conn.close()

if __name__ == '__main__':
    init_db()
    print("Database initialized.")
