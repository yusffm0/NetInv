"""
Network Inventory Management System
Professional Flask application with audit logging, Excel import/export,
duplicate detection, and persistent SQLite storage.
"""

import os
import io
import csv
import json
import uuid
import sqlite3
import logging
from datetime import datetime
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, jsonify, send_file, flash, g
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import openpyxl

# ---------------------------------------------------------------------------
# App Configuration
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "nims-secret-key-change-in-production")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit

DATABASE = os.path.join(os.path.dirname(__file__), "inventory.db")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

# Undo/Redo / history storage (temporary, lightweight in-memory cache)
action_history = {}
undo_stack = []
redo_stack = []

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Database helpers
# ---------------------------------------------------------------------------

def get_db():
    """Open a new database connection for the current request."""
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
        g.db.execute("PRAGMA foreign_keys=ON")
    return g.db


@app.teardown_appcontext
def close_db(exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    """Create tables and seed default data if the database is fresh."""
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row

    db.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            username  TEXT UNIQUE NOT NULL,
            password  TEXT NOT NULL,
            role      TEXT NOT NULL DEFAULT 'viewer'
        );

        CREATE TABLE IF NOT EXISTS inventory (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            serial_number TEXT,
            asset_tag     TEXT,
            device_name   TEXT NOT NULL,
            device_type   TEXT,
            brand         TEXT,
            model         TEXT,
            quantity      INTEGER DEFAULT 1,
            ip_address    TEXT,
            mac_address   TEXT,
            location      TEXT,
            status        TEXT DEFAULT 'Active',
            notes         TEXT,
            created_at    TEXT NOT NULL,
            updated_at    TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp   TEXT NOT NULL,
            username    TEXT NOT NULL,
            display_name TEXT,
            ip_address  TEXT,
            mac_address TEXT,
            action      TEXT NOT NULL,
            item_id     TEXT,
            item_name   TEXT,
            reason      TEXT,
            details     TEXT
        );

        CREATE TABLE IF NOT EXISTS quantity_changes (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id     INTEGER NOT NULL,
            old_quantity INTEGER NOT NULL,
            new_quantity INTEGER NOT NULL,
            reason      TEXT,
            username    TEXT NOT NULL,
            timestamp   TEXT NOT NULL,
            FOREIGN KEY (item_id) REFERENCES inventory (id)
        );
    """)

    existing_columns = [row[1] for row in db.execute("PRAGMA table_info(audit_log)").fetchall()]
    if "mac_address" not in existing_columns:
        try:
            db.execute("ALTER TABLE audit_log ADD COLUMN mac_address TEXT")
        except sqlite3.OperationalError:
            pass

    # Seed admin user
    existing = db.execute("SELECT id FROM users WHERE username='yusf'").fetchone()
    if not existing:
        db.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            ("yusf", generate_password_hash("123"), "admin"),
        )

    # Seed initial inventory from the Excel data provided
    count = db.execute("SELECT COUNT(*) FROM inventory").fetchone()[0]
    if count == 0:
        seed_items = [
            ("طقم مفكات",        "Tools",         "",         "",   1),
            ("كرتونة كبل Ethernet","Cable",        "Ethernet", "",  10),
            ("Router TP-Link TL-MR6400","Router",  "TP-Link","TL-MR6400", 51),
            ("علبة كهرباء بوش",   "Electrical",   "Bosch",    "",   2),
            ("جهاز Vitec",        "Device",       "Vitec",    "",   1),
            ("شاشات قاعة",        "Display",      "",         "",   7),
            ("علبة RG",           "Accessory",    "",         "",  24),
            ("سماعات بوش",        "Audio",        "Bosch",    "",   5),
            ("كاميرات Hanwha",    "CCTV",         "Hanwha",   "",   7),
            ("كاميرا بوش",        "CCTV",         "Bosch",    "",   1),
            ("Access Point TP-Link","Access Point","TP-Link",  "",  44),
            ("حامل Rack",         "Rack",         "",         "",   1),
            ("كرتونة قعدة كاميرا","CCTV",         "",         "",   3),
            ("Pole Adapter Large","Accessory",    "",         "",   3),
            ("ستاند كاميرا صغيره","CCTV",         "",         "",   3),
            ("طبق فايبر",         "Fiber",        "",         "",   2),
            ("Huawei Switch",     "Switch",       "Huawei",   "",   1),
            ("Power Beam",        "Wireless",     "",         "",   1),
            ("Cisco Power Supply","Power Supply", "Cisco",    "",   3),
            ("ريموت Samsung",     "Remote",       "Samsung",  "",   6),
            ("Remote Bosch كبير", "Remote",       "Bosch",    "",   1),
            ("Remote Bosch صغير", "Remote",       "Bosch",    "",   2),
            ("كابل فايبر CH-GT",  "Fiber Cable",  "",         "",  15),
            ("كابل فايبر LC04",   "Fiber Cable",  "",         "",  20),
            ("كابل فايبر IF0.90M","Fiber Cable",  "",         "",  26),
            ("SFP Fiber Cable",   "Fiber Cable",  "",         "",   1),
            ("كابل فايبر",        "Fiber Cable",  "",         "",   4),
            ("Adaptor Power Beam","Accessory",    "",         "",   2),
            ("Cable Ethernet 1m Green","Cable",   "",         "", 100),
            ("Cable Ethernet 1m White","Cable",   "",         "",  12),
            ("Cable Ethernet 3m",  "Cable",       "",         "",  85),
            ("Switch 24 Port Cisco","Switch",     "Cisco",    "",   4),
            ("Switch 24 Port TP-Link","Switch",   "TP-Link",  "",   6),
            ("كرتونة تاب وقلم",   "Accessory",   "",         "",   1),
            ("Switch 10 Port TP-Link","Switch",   "TP-Link",  "",  11),
            ("Power Supply Cisco", "Power Supply","Cisco",    "",   5),
            ("Access Point Cisco", "Access Point","Cisco",    "",   7),
            ("Cable Power",        "Cable",       "",         "",   2),
            ("قاعدة Access Point", "Accessory",   "",         "",   7),
            ("Mouse Pad",          "Peripheral",  "",         "",   1),
            ("Mouse Wireless",     "Peripheral",  "",         "",   2),
            ("Phase Split",        "Electrical",  "",         "",   4),
            ("Firewall Modicon",   "Firewall",    "Modicon",  "",   2),
            ("Switch Shelf Small", "Accessory",   "",         "",  19),
            ("Switch Shelf Large", "Accessory",   "",         "",   3),
            ("ريموت Samsung",     "Remote",       "Samsung",  "",   1),
        ]
        now = datetime.utcnow().isoformat()
        for name, dtype, brand, model, qty in seed_items:
            db.execute(
                """INSERT INTO inventory
                   (device_name, device_type, brand, model, quantity,
                    status, created_at, updated_at)
                   VALUES (?,?,?,?,?,'Active',?,?)""",
                (name, dtype, brand, model, qty, now, now),
            )

    db.commit()
    db.close()


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            return jsonify({"error": "Admin access required"}), 403
        return f(*args, **kwargs)
    return decorated


def record_history(audit_id, action, item_id=None, old_data=None, new_data=None):
    record = {
        "audit_id": audit_id,
        "action": action,
        "item_id": item_id,
        "old_data": old_data,
        "new_data": new_data,
        "timestamp": datetime.utcnow().isoformat(),
    }
    action_history[audit_id] = record
    undo_stack.append(audit_id)
    redo_stack.clear()


def get_history(audit_id):
    return action_history.get(audit_id)


def push_redo(audit_id):
    if audit_id and audit_id not in redo_stack:
        redo_stack.append(audit_id)


def pop_redo():
    return redo_stack.pop() if redo_stack else None


def log_action(action, item_id=None, item_name=None, details=None, mac_address=None):
    """Insert an audit log row using session context."""
    db = get_db()
    cur = db.execute(
        """INSERT INTO audit_log
           (timestamp, username, display_name, ip_address, mac_address,
            action, item_id, item_name, reason, details)
           VALUES (?,?,?,?,?,?,?,?,?,?)""",
        (
            datetime.utcnow().isoformat(),
            session.get("username", "unknown"),
            session.get("display_name", ""),
            request.remote_addr,
            mac_address,
            action,
            str(item_id) if item_id else None,
            item_name,
            session.get("audit_reason", ""),
            details,
        ),
    )
    db.commit()
    return cur.lastrowid


# ---------------------------------------------------------------------------
# Routes – Auth
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    if "user_id" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


def get_dashboard_data():
    """Helper function to get dashboard statistics."""
    db = get_db()
    total_users = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    total_devices = db.execute("SELECT COUNT(*) FROM inventory").fetchone()[0]
    total_quantity = db.execute("SELECT SUM(quantity) FROM inventory").fetchone()[0] or 0
    last_login = db.execute("SELECT username, timestamp FROM audit_log WHERE action='LOGIN' ORDER BY timestamp DESC LIMIT 1").fetchone()
    last_login_info = last_login and f"{last_login['username']} on {last_login['timestamp'][:10]}" or "None"
    device_types = db.execute("SELECT device_type, COUNT(*) as count FROM inventory GROUP BY device_type").fetchall()
    chart_data = [{"type": (r["device_type"] or "Other"), "count": int(r["count"])} for r in device_types]
    
    return {
        "total_users": int(total_users),
        "total_devices": int(total_devices),
        "total_quantity": int(total_quantity),
        "last_login": str(last_login_info),
        "chart_data": chart_data if chart_data else []
    }


@app.route("/dashboard")
@login_required
def dashboard():
    stats = get_dashboard_data()
    return render_template("inventory.html", user_role=session.get("role"), 
                           show_add_user=(session.get("role") == "admin"),
                           **stats)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE username=?", (username,)
        ).fetchone()
        if user and check_password_hash(user["password"], password):
            session["user_id"]  = user["id"]
            session["username"] = user["username"]
            session["role"]     = user["role"]
            session["display_name"] = user["username"]
            log_action("LOGIN", details="User logged in")
            return jsonify({"ok": True})
        return jsonify({"ok": False, "error": "Invalid credentials"}), 401
    return render_template("login.html")


@app.route("/logout")
def logout():
    log_action("LOGOUT")
    session.clear()
    return redirect(url_for("login"))


@app.route("/api/audit/<int:audit_id>/undo", methods=["POST"])
@admin_required
def undo_audit(audit_id):
    """Undo a previous audit action using stored history."""
    db = get_db()
    history = get_history(audit_id)
    if not history:
        return jsonify({"error": "History not available for this action"}), 404

    action = history["action"]
    item_id = history["item_id"]

    if action in ("ADD", "IMPORT_ADD") and item_id:
        db.execute("DELETE FROM inventory WHERE id=?", (int(item_id),))
        db.commit()
        log_action("UNDO", item_id=item_id, item_name=history.get("new_data", {}).get("device_name"),
                   details=f"Undid {action} action")
        push_redo(audit_id)
        return jsonify({"ok": True, "message": f"Undid {action} for {history.get('new_data', {}).get('device_name','item')}"})

    if action in ("DELETE",) and item_id:
        old_data = history.get("old_data") or {}
        if not old_data:
            return jsonify({"error": "Cannot restore deleted data"}), 400
        db.execute(
            """INSERT OR REPLACE INTO inventory
               (id, serial_number, asset_tag, device_name, device_type, brand, model,
                quantity, ip_address, mac_address, location, status, notes,
                created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                item_id,
                old_data.get("serial_number"), old_data.get("asset_tag"), old_data.get("device_name"),
                old_data.get("device_type"), old_data.get("brand"), old_data.get("model"),
                old_data.get("quantity"), old_data.get("ip_address"), old_data.get("mac_address"),
                old_data.get("location"), old_data.get("status"), old_data.get("notes"),
                old_data.get("created_at"), old_data.get("updated_at"),
            ),
        )
        db.commit()
        log_action("UNDO", item_id=item_id, item_name=old_data.get("device_name"),
                   details=f"Restored deleted item")
        push_redo(audit_id)
        return jsonify({"ok": True, "message": f"Restored deleted item {old_data.get('device_name','item')}"})

    if action in ("EDIT", "IMPORT_MERGE", "IMPORT_REPLACE") and item_id:
        old_data = history.get("old_data") or {}
        if not old_data:
            return jsonify({"error": "No prior state available"}), 400
        db.execute(
            """UPDATE inventory SET
               serial_number=?, asset_tag=?, device_name=?, device_type=?, brand=?, model=?,
               quantity=?, ip_address=?, mac_address=?, notes=?, updated_at=?
               WHERE id=?""",
            (
                old_data.get("serial_number"), old_data.get("asset_tag"), old_data.get("device_name"),
                old_data.get("device_type"), old_data.get("brand"), old_data.get("model"),
                old_data.get("quantity"), old_data.get("ip_address"), old_data.get("mac_address"),
                old_data.get("notes"), datetime.utcnow().isoformat(), item_id,
            ),
        )
        db.commit()
        log_action("UNDO", item_id=item_id, item_name=old_data.get("device_name"),
                   details=f"Reverted {action}")
        push_redo(audit_id)
        return jsonify({"ok": True, "message": f"Reverted {action} for {old_data.get('device_name','item')}"})

    return jsonify({"error": "Cannot undo this action"}), 400


@app.route("/api/audit/redo", methods=["POST"])
@admin_required
def redo_audit():
    audit_id = pop_redo()
    if not audit_id:
        return jsonify({"error": "Nothing to redo"}), 400
    history = get_history(audit_id)
    if not history:
        return jsonify({"error": "No history available for redo"}), 404

    db = get_db()
    action = history["action"]
    item_id = history["item_id"]
    new_data = history.get("new_data") or {}

    if action in ("ADD", "IMPORT_ADD") and new_data:
        db.execute(
            """INSERT OR REPLACE INTO inventory
               (id, serial_number, asset_tag, device_name, device_type, brand, model,
                quantity, ip_address, mac_address, notes, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                item_id,
                new_data.get("serial_number"), new_data.get("asset_tag"), new_data.get("device_name"),
                new_data.get("device_type"), new_data.get("brand"), new_data.get("model"),
                new_data.get("quantity"), new_data.get("ip_address"), new_data.get("mac_address"),
                new_data.get("notes"), datetime.utcnow().isoformat(), datetime.utcnow().isoformat(),
            ),
        )
        db.commit()
        log_action("REDO", item_id=item_id, item_name=new_data.get("device_name"),
                   details=f"Redid {action}")
        return jsonify({"ok": True, "message": f"Redid {action} for {new_data.get('device_name','item')}"})

    if action == "DELETE" and item_id:
        db.execute("DELETE FROM inventory WHERE id=?", (item_id,))
        db.commit()
        log_action("REDO", item_id=item_id, item_name=history.get("old_data", {}).get("device_name"),
                   details="Redid DELETE")
        return jsonify({"ok": True, "message": "Redid DELETE"})

    if action in ("EDIT", "IMPORT_MERGE", "IMPORT_REPLACE") and item_id:
        db.execute(
            """UPDATE inventory SET
               serial_number=?, asset_tag=?, device_name=?, device_type=?, brand=?, model=?,
               quantity=?, ip_address=?, mac_address=?, notes=?, updated_at=?
               WHERE id=?""",
            (
                new_data.get("serial_number"), new_data.get("asset_tag"), new_data.get("device_name"),
                new_data.get("device_type"), new_data.get("brand"), new_data.get("model"),
                new_data.get("quantity"), new_data.get("ip_address"), new_data.get("mac_address"),
                new_data.get("notes"), datetime.utcnow().isoformat(), item_id,
            ),
        )
        db.commit()
        log_action("REDO", item_id=item_id, item_name=new_data.get("device_name"),
                   details=f"Redid {action}")
        return jsonify({"ok": True, "message": f"Redid {action} for {new_data.get('device_name','item')}"})

    return jsonify({"error": "Cannot redo this action"}), 400


# ---------------------------------------------------------------------------
# Routes – Inventory
# ---------------------------------------------------------------------------

@app.route("/inventory")
@login_required
def inventory():
    return render_template(
        "inventory.html",
        show_add_user=(session.get("role") == "admin"),
        user_role=session.get("role"),
        total_users=0,
        total_devices=0,
        total_quantity=0,
        last_login="—",
        chart_data=[]
    )


@app.route("/logs")
@login_required
def logs_page():
    return render_template(
        "inventory.html",
        show_add_user=(session.get("role") == "admin"),
        user_role=session.get("role"),
        total_users=0,
        total_devices=0,
        total_quantity=0,
        last_login="—",
        chart_data=[]
    )


@app.route("/movements")
@login_required
def movements():
    return render_template(
        "inventory.html",
        show_add_user=(session.get("role") == "admin"),
        user_role=session.get("role"),
        total_users=0,
        total_devices=0,
        total_quantity=0,
        last_login="—",
        chart_data=[]
    )


@app.route("/api/dashboard")
@login_required
def api_dashboard():
    stats = get_dashboard_data()
    return jsonify(stats)


@app.route("/api/inventory", methods=["GET"])
@login_required
def api_get_inventory():
    db   = get_db()
    rows = db.execute("SELECT * FROM inventory ORDER BY id").fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/inventory", methods=["POST"])
@admin_required
def api_add_item():
    data = request.get_json()
    now  = datetime.utcnow().isoformat()
    db   = get_db()
    item_data = {
        "serial_number": data.get("serial_number"),
        "asset_tag":     data.get("asset_tag"),
        "device_name":   data.get("device_name"),
        "device_type":   data.get("device_type"),
        "brand":         data.get("brand"),
        "model":         data.get("model"),
        "quantity":      int(data.get("quantity", 1) or 1),
        "ip_address":    data.get("ip_address"),
        "mac_address":   data.get("mac_address"),
        "notes":         data.get("notes"),
    }
    cur  = db.execute(
        """INSERT INTO inventory
           (serial_number, asset_tag, device_name, device_type, brand, model,
            quantity, ip_address, mac_address, notes, created_at, updated_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            item_data["serial_number"], item_data["asset_tag"],
            item_data["device_name"],   item_data["device_type"],
            item_data["brand"],         item_data["model"],
            item_data["quantity"],      item_data["ip_address"],
            item_data["mac_address"],   item_data["notes"], now, now,
        ),
    )
    db.commit()
    item_id = cur.lastrowid
    details = f"MAC={item_data['mac_address'] or ''} | IP={item_data['ip_address'] or ''}"
    audit_id = log_action("ADD", item_id=item_id, item_name=item_data.get("device_name"), details=details, mac_address=item_data.get("mac_address"))
    record_history(audit_id, "ADD", item_id=item_id, old_data=None, new_data=item_data)
    return jsonify({"ok": True, "id": item_id}), 201


@app.route("/api/inventory/<int:item_id>", methods=["PUT"])
@admin_required
def api_update_item(item_id):
    data = request.get_json()
    now  = datetime.utcnow().isoformat()
    db   = get_db()
    existing = db.execute("SELECT * FROM inventory WHERE id=?", (item_id,)).fetchone()
    if not existing:
        return jsonify({"error": "Not found"}), 404

    old_item = {k: existing[k] for k in [
        "serial_number", "asset_tag", "device_name", "device_type",
        "brand", "model", "quantity", "ip_address", "mac_address", "notes"
    ]}

    updated_item = {
        "serial_number": data.get("serial_number", existing["serial_number"]),
        "asset_tag":     data.get("asset_tag", existing["asset_tag"]),
        "device_name":   data.get("device_name", existing["device_name"]),
        "device_type":   data.get("device_type", existing["device_type"]),
        "brand":         data.get("brand", existing["brand"]),
        "model":         data.get("model", existing["model"]),
        "quantity":      int(data.get("quantity", existing["quantity"] or 1) or 1),
        "ip_address":    data.get("ip_address", existing["ip_address"]),
        "mac_address":   data.get("mac_address", existing["mac_address"]),
        "notes":         data.get("notes", existing["notes"]),
    }

    changes = []
    labels = {
        "serial_number": "Serial Number",
        "asset_tag": "Asset Tag",
        "device_name": "Device Name",
        "device_type": "Type",
        "brand": "Brand",
        "model": "Model",
        "quantity": "Quantity",
        "ip_address": "IP Address",
        "mac_address": "MAC Address",
        "notes": "Notes",
    }
    for field, label in labels.items():
        if str(old_item[field] or "") != str(updated_item[field] or ""):
            changes.append(f"{label}: {old_item[field] or ''} → {updated_item[field] or ''}")

    db.execute(
        """UPDATE inventory SET
           serial_number=?, asset_tag=?, device_name=?, device_type=?,
           brand=?, model=?, quantity=?, ip_address=?, mac_address=?,
           notes=?, updated_at=?
           WHERE id=?""",
        (
            updated_item["serial_number"], updated_item["asset_tag"],
            updated_item["device_name"],   updated_item["device_type"],
            updated_item["brand"],         updated_item["model"],
            updated_item["quantity"],      updated_item["ip_address"],
            updated_item["mac_address"],   updated_item["notes"], now, item_id,
        ),
    )
    db.commit()

    # Record quantity change if applicable
    if old_item["quantity"] != updated_item["quantity"]:
        reason = data.get("reason", "")
        db.execute(
            """INSERT INTO quantity_changes
               (item_id, old_quantity, new_quantity, reason, username, timestamp)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                item_id,
                old_item["quantity"],
                updated_item["quantity"],
                reason,
                session.get("username"),
                now,
            ),
        )
        db.commit()

    detail_text = changes and "; ".join(changes) or "No changes"
    audit_id = log_action(
        "EDIT",
        item_id=item_id,
        item_name=updated_item["device_name"],
        details=detail_text,
        mac_address=updated_item.get("mac_address"),
    )
    record_history(audit_id, "EDIT", item_id=item_id, old_data=old_item, new_data=updated_item)
    return jsonify({"ok": True})


@app.route("/api/inventory/<int:item_id>", methods=["DELETE"])
@admin_required
def api_delete_item(item_id):
    db   = get_db()
    item = db.execute("SELECT * FROM inventory WHERE id=?", (item_id,)).fetchone()
    if not item:
        return jsonify({"error": "Not found"}), 404
    deleted_item = {k: item[k] for k in item.keys()}
    db.execute("DELETE FROM inventory WHERE id=?", (item_id,))
    db.commit()
    audit_id = log_action(
        "DELETE",
        item_id=item_id,
        item_name=item["device_name"],
        details=f"Deleted device MAC={item['mac_address']}",
        mac_address=item["mac_address"],
    )
    record_history(audit_id, "DELETE", item_id=item_id, old_data=deleted_item, new_data=None)
    return jsonify({"ok": True})


@app.route("/api/users", methods=["POST"])
@admin_required
def api_create_user():
    data = request.get_json() or {}
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""
    role = (data.get("role") or "viewer").strip().lower()

    if not username or not password:
        return jsonify({"error": "Username and password are required"}), 400
    if role not in ("viewer", "editor", "admin"):
        role = "viewer"

    db = get_db()
    try:
        db.execute(
            "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
            (username, generate_password_hash(password), role),
        )
        db.commit()
    except sqlite3.IntegrityError:
        return jsonify({"error": "Username already exists"}), 400

    log_action(
        "USER_ADD",
        item_name=username,
        details=f"Created user {username} with role {role}",
    )
    return jsonify({"ok": True, "username": username, "role": role}), 201


@app.route("/api/users", methods=["GET"])
@admin_required
def api_list_users():
    db = get_db()
    rows = db.execute("SELECT id, username, role FROM users ORDER BY username").fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/users/<int:user_id>", methods=["PUT"])
@admin_required
def api_update_user(user_id):
    data = request.get_json() or {}
    username = (data.get("username") or "").strip()
    password = data.get("password")
    role = (data.get("role") or "viewer").strip().lower()
    if role not in ("viewer", "editor", "admin"):
        role = "viewer"
    if not username:
        return jsonify({"error": "Username is required"}), 400

    db = get_db()
    existing = db.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
    if not existing:
        return jsonify({"error": "User not found"}), 404

    params = [username, role, user_id]
    query = "UPDATE users SET username=?, role=?"
    if password:
        query += ", password=?"
        params.insert(2, generate_password_hash(password))
    query += " WHERE id=?"

    try:
        db.execute(query, tuple(params))
        db.commit()
    except sqlite3.IntegrityError:
        return jsonify({"error": "Username already exists"}), 400

    log_action(
        "USER_EDIT",
        item_name=username,
        details=f"Updated user {username} role={role}",
    )
    return jsonify({"ok": True})


@app.route("/api/users/<int:user_id>", methods=["DELETE"])
@admin_required
def api_delete_user(user_id):
    db = get_db()
    user = db.execute("SELECT username FROM users WHERE id=?", (user_id,)).fetchone()
    if not user:
        return jsonify({"error": "User not found"}), 404
    db.execute("DELETE FROM users WHERE id=?", (user_id,))
    db.commit()
    log_action(
        "USER_DELETE",
        item_name=user["username"],
        details=f"Deleted user {user['username']}",
    )
    return jsonify({"ok": True})


@app.route("/users")
@admin_required
def users_page():
    return render_template("users.html", user_role=session.get("role"))


@app.route("/api/movements")
@login_required
def api_movements():
    db = get_db()
    rows = db.execute("""
        SELECT qc.*, i.device_name, i.serial_number, i.asset_tag
        FROM quantity_changes qc
        JOIN inventory i ON qc.item_id = i.id
        ORDER BY qc.timestamp DESC
    """).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/movements/clear", methods=["POST"])
@login_required
def api_clear_movements():
    db = get_db()
    db.execute("DELETE FROM quantity_changes")
    db.commit()
    log_action("CLEAR", details="Quantity movements cleared by user")
    return jsonify({"ok": True, "message": "Quantity movements cleared"})


# ---------------------------------------------------------------------------
# Routes – Export
# ---------------------------------------------------------------------------

@app.route("/export/inventory")
@login_required
def export_inventory():
    db = get_db()
    search = request.args.get('search', '').strip()
    type_filter = request.args.get('type', '').strip()
    
    query = "SELECT * FROM inventory WHERE 1=1"
    params = []
    
    if search:
        query += " AND (device_name LIKE ? OR brand LIKE ? OR model LIKE ? OR serial_number LIKE ? OR ip_address LIKE ? OR mac_address LIKE ?)"
        like_search = f"%{search}%"
        params.extend([like_search] * 6)
    
    if type_filter:
        query += " AND device_type = ?"
        params.append(type_filter)
    
    query += " ORDER BY id"
    
    rows = db.execute(query, params).fetchall()
    si = io.StringIO()
    writer = csv.writer(si)

    include_network = session.get("role") != "viewer"
    headers = [
        "ID", "Serial Number", "Asset Tag", "Device Name", "Type",
        "Brand", "Model", "Quantity",
    ]
    if include_network:
        headers.extend(["IP Address", "MAC Address"])
    headers.extend(["Notes", "Created At", "Updated At"])
    writer.writerow(headers)

    for r in rows:
        row = [
            r["id"], r["serial_number"], r["asset_tag"], r["device_name"],
            r["device_type"], r["brand"], r["model"], r["quantity"],
        ]
        if include_network:
            row.extend([r["ip_address"], r["mac_address"]])
        row.extend([r["notes"], r["created_at"], r["updated_at"]])
        writer.writerow(row)
    log_action("EXPORT", details=f"Inventory CSV export ({len(rows)} items)")
    output = io.BytesIO()
    output.write(si.getvalue().encode("utf-8-sig"))
    output.seek(0)
    return send_file(
        output,
        mimetype="text/csv",
        as_attachment=True,
        download_name=f"inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    )


@app.route("/export/audit")
@login_required
def export_audit():
    db   = get_db()
    rows = db.execute("SELECT * FROM audit_log ORDER BY id DESC").fetchall()
    si   = io.StringIO()
    writer = csv.writer(si)
    writer.writerow([
        "ID", "Timestamp", "Username", "Display Name", "IP Address", "MAC Address",
        "Action", "Item ID", "Item Name", "Reason", "Details",
    ])
    for r in rows:
        writer.writerow([
            r["id"], r["timestamp"], r["username"], r["display_name"],
            r["ip_address"], r.get("mac_address", ""), r["action"], r["item_id"], r["item_name"],
            r["reason"], r["details"],
        ])
    output = io.BytesIO()
    output.write(si.getvalue().encode("utf-8-sig"))
    output.seek(0)
    return send_file(
        output,
        mimetype="text/csv",
        as_attachment=True,
        download_name=f"audit_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    )


@app.route("/api/audit/clear", methods=["POST"])
@admin_required
def api_clear_audit():
    db = get_db()
    db.execute("DELETE FROM audit_log")
    db.commit()
    log_action("CLEAR", details="Audit log cleared by admin")
    return jsonify({"ok": True, "message": "Audit log cleared"})


# ---------------------------------------------------------------------------
# Routes – Import
# ---------------------------------------------------------------------------

def _normalize_text(value):
    return str(value).strip() if value is not None else ""


def _find_duplicate(db, record: dict):
    """Find matching inventory item by priority: serial, asset tag, ip, device name."""
    search_keys = ["serial_number", "asset_tag", "ip_address", "device_name"]
    for key in search_keys:
        value = record.get(key)
        if not value:
            continue
        if key == "device_name":
            existing = db.execute(
                "SELECT * FROM inventory WHERE LOWER(device_name)=LOWER(?)", (value.lower(),)
            ).fetchone()
        else:
            existing = db.execute(
                f"SELECT * FROM inventory WHERE {key}=?", (value,)
            ).fetchone()
        if existing:
            return existing
    return None


def _parse_import_file(f):
    try:
        wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        raise ValueError(f"Failed to read Excel: {e}")
    if not rows:
        raise ValueError("Empty spreadsheet")

    # Detect header row
    header_row_idx = 0
    for i, row in enumerate(rows):
        if any(cell is not None for cell in row):
            header_row_idx = i
            break

    headers = [str(c).strip().lower() if c else "" for c in rows[header_row_idx]]
    field_map = {
        "device_name":   ["device name", "device_name", "name", "اسم الجهاز", "النوع", "الجهاز"],
        "device_type":   ["type", "device type", "device_type", "النوع", "الصنف"],
        "brand":         ["brand", "manufacturer", "الماركة", "الشركة"],
        "model":         ["model", "model number", "الموديل"],
        "quantity":      ["quantity", "qty", "count", "الكمية", "كمية"],
        "serial_number": ["serial", "serial number", "serial_number", "s/n", "الرقم التسلسلي"],
        "asset_tag":     ["asset tag", "asset_tag", "tag", "رقم الأصل"],
        "ip_address":    ["ip", "ip address", "ip_address", "عنوان ip"],
        "mac_address":   ["mac", "mac address", "mac_address", "عنوان mac"],
        "notes":         ["notes", "remarks", "comment", "ملاحظات"],
    }

    col_index = {}
    for field, aliases in field_map.items():
        for i, h in enumerate(headers):
            if h in aliases:
                col_index[field] = i
                break
    if not col_index:
        col_index = {"serial_number": 0, "quantity": 1, "device_name": 2}

    def get_field(row, field, default=""):
        idx = col_index.get(field)
        if idx is None or idx >= len(row):
            return default
        return _normalize_text(row[idx])

    records = []
    for i, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
        if all(cell is None for cell in row):
            continue
        device_name = get_field(row, "device_name")
        if not device_name:
            continue
        quantity = get_field(row, "quantity", "1")
        try:
            quantity = int(float(quantity))
        except (TypeError, ValueError):
            quantity = 1

        records.append({
            "row": i,
            "device_name": device_name,
            "device_type": get_field(row, "device_type"),
            "brand": get_field(row, "brand"),
            "model": get_field(row, "model"),
            "quantity": quantity,
            "serial_number": get_field(row, "serial_number"),
            "asset_tag": get_field(row, "asset_tag"),
            "ip_address": get_field(row, "ip_address"),
            "mac_address": get_field(row, "mac_address"),
            "notes": get_field(row, "notes"),
        })
    return records


@app.route("/import/excel", methods=["POST"])
@admin_required
def import_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    f = request.files["file"]
    if not f.filename or not f.filename.rsplit(".", 1)[-1].lower() in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Only .xlsx files are accepted"}), 400

    try:
        records = _parse_import_file(f)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    db = get_db()
    preview_new = []
    preview_duplicates = []
    for record in records:
        existing = _find_duplicate(db, record)
        if existing:
            existing_data = {k: existing[k] for k in [
                "id", "serial_number", "asset_tag", "device_name", "device_type",
                "brand", "model", "quantity", "ip_address", "mac_address", "notes"
            ]}
            preview_duplicates.append({
                "row": record["row"],
                "existing": existing_data,
                "incoming": record,
            })
        else:
            preview_new.append(record)

    return jsonify({"ok": True, "new_items": preview_new, "duplicates": preview_duplicates})


@app.route("/import/excel/apply", methods=["POST"])
@admin_required
def import_excel_apply():
    data = request.get_json() or {}
    new_items = data.get("new_items", [])
    duplicate_actions = data.get("duplicate_actions", [])
    db = get_db()
    now = datetime.utcnow().isoformat()

    added = 0
    updated = 0
    skipped = 0
    errors = []

    for item in new_items:
        try:
            cur = db.execute(
                """INSERT INTO inventory
                   (serial_number, asset_tag, device_name, device_type, brand, model,
                    quantity, ip_address, mac_address, notes, created_at, updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    item.get("serial_number"), item.get("asset_tag"), item.get("device_name"),
                    item.get("device_type"), item.get("brand"), item.get("model"),
                    int(item.get("quantity") or 1), item.get("ip_address"), item.get("mac_address"),
                    item.get("notes"), now, now,
                ),
            )
            db.commit()
            item_id = cur.lastrowid
            audit_id = log_action("IMPORT_ADD", item_id=item_id, item_name=item.get("device_name"),
                                  details=f"Imported new item", mac_address=item.get("mac_address"))
            record_history(audit_id, "IMPORT_ADD", item_id=item_id, old_data=None, new_data=item)
            added += 1
        except Exception as exc:
            errors.append(f"Create row {item.get('row', '')}: {exc}")

    for action_entry in duplicate_actions:
        action = action_entry.get("action")
        if action not in ("merge", "replace"):
            skipped += 1
            continue
        existing_id = action_entry.get("existing_id")
        incoming = action_entry.get("incoming", {})
        existing = db.execute("SELECT * FROM inventory WHERE id=?", (existing_id,)).fetchone()
        if not existing:
            errors.append(f"Existing item {existing_id} not found")
            continue

        old_data = {k: existing[k] for k in [
            "serial_number", "asset_tag", "device_name", "device_type",
            "brand", "model", "quantity", "ip_address", "mac_address", "notes"
        ]}
        updated_item = {
            "serial_number": incoming.get("serial_number") or old_data["serial_number"],
            "asset_tag":     incoming.get("asset_tag") or old_data["asset_tag"],
            "device_name":   incoming.get("device_name") or old_data["device_name"],
            "device_type":   incoming.get("device_type") or old_data["device_type"],
            "brand":         incoming.get("brand") or old_data["brand"],
            "model":         incoming.get("model") or old_data["model"],
            "ip_address":    incoming.get("ip_address") or old_data["ip_address"],
            "mac_address":   incoming.get("mac_address") or old_data["mac_address"],
            "notes":         incoming.get("notes") or old_data["notes"],
        }
        incoming_qty = int(incoming.get("quantity") or 1)
        if action == "merge":
            updated_item["quantity"] = (existing["quantity"] or 0) + incoming_qty
            action_label = "IMPORT_MERGE"
            details = f"Merged quantity {existing['quantity']} + {incoming_qty}"
        else:
            updated_item["quantity"] = incoming_qty
            action_label = "IMPORT_REPLACE"
            details = f"Replaced item with imported data"

        try:
            db.execute(
                """UPDATE inventory SET
                   serial_number=?, asset_tag=?, device_name=?, device_type=?, brand=?, model=?,
                   quantity=?, ip_address=?, mac_address=?, notes=?, updated_at=?
                   WHERE id=?""",
                (
                    updated_item["serial_number"], updated_item["asset_tag"],
                    updated_item["device_name"], updated_item["device_type"],
                    updated_item["brand"], updated_item["model"],
                    updated_item["quantity"], updated_item["ip_address"],
                    updated_item["mac_address"], updated_item["notes"], now, existing_id,
                ),
            )
            db.commit()
            audit_id = log_action(action_label, item_id=existing_id, item_name=updated_item.get("device_name"),
                                  details=details, mac_address=updated_item.get("mac_address"))
            record_history(audit_id, action_label, item_id=existing_id, old_data=old_data, new_data=updated_item)
            updated += 1
        except Exception as exc:
            errors.append(f"Update item {existing_id}: {exc}")

    return jsonify({"ok": True, "added": added, "updated": updated, "skipped": skipped, "errors": errors})


# ---------------------------------------------------------------------------
# Routes – Audit Log API
# ---------------------------------------------------------------------------

@app.route("/api/audit", methods=["GET"])
@login_required
def api_get_audit():
    db    = get_db()
    rows  = db.execute("SELECT * FROM audit_log ORDER BY id DESC").fetchall()
    return jsonify([dict(r) for r in rows])


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5000)
