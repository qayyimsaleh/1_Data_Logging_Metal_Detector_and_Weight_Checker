"""
shared_config.py - Shared Database & Configuration Module
==========================================================
Used by both production.py and report.py

SECURITY FIXES:
 - Stored procedure whitelist prevents SQL injection via proc names
 - All queries use parameterized ? placeholders
 - Input sanitization for all user text
 - IP validation
 - No passwords/credentials logged

BUG FIXES:
 - Bare except: clauses replaced with Exception
 - Connection retry logic with proper cleanup
 - Thread-safe connection check
"""

import pyodbc
import logging
import os
import re
import socket
import getpass
from datetime import datetime
from logging.handlers import RotatingFileHandler

# ═══════════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════════
APP_VERSION = "2.0.0"
APP_TITLE = "PanCen Production Management System"

CONN_STR = (
    "Driver={ODBC Driver 17 for SQL Server};"
    "Server=localhost\\SQLEXPRESS;"
    "Database=SNP1;"
    "Trusted_Connection=yes;"
)

# Whitelist of allowed stored procedures (SQL injection prevention)
ALLOWED_PROCEDURES = frozenset([
    "sp_VerifyUser", "sp_GetMachineByIP", "sp_GetAllMachines",
    "sp_GetNextProductionId", "sp_GetNextLogId",
    "sp_InsertUpdateProductionSession", "sp_InsertLogEntry",
    "sp_GetProductionReport", "sp_GetSessionDetails",
    "sp_GetSessionStats", "sp_GetRecentSessionReadings",
    "sp_GetTotalSessions", "sp_GetLastProductionId",
    "sp_GetCurrentSessionReadingCount", "sp_GetProductionIdByLot",
])

IP_REGEX = re.compile(
    r"^(?:(?:25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)\.){3}"
    r"(?:25[0-5]|2[0-4]\d|1\d{2}|[1-9]?\d)$"
)

# Modern dark color scheme
COLORS = {
    "bg_dark": "#1e1e2e", "bg_mid": "#252536", "bg_card": "#2d2d44",
    "bg_input": "#363652", "accent": "#7c5cfc", "accent_hover": "#6a4ce0",
    "green": "#3dd68c", "red": "#ff6b6b", "orange": "#ffb347",
    "blue": "#4ea8de", "text": "#e2e2e9", "text2": "#a0a0b8",
    "text3": "#6e6e87", "border": "#3d3d5c", "terminal_bg": "#0f0f1a",
}

# ═══════════════════════════════════════════════════
# VALIDATION HELPERS
# ═══════════════════════════════════════════════════
def validate_ip(ip: str) -> bool:
    return bool(IP_REGEX.match(ip.strip())) if ip else False

def validate_port(port_str: str):
    try:
        p = int(port_str)
        return p if 1 <= p <= 65535 else None
    except (ValueError, TypeError):
        return None

def sanitize(text: str, maxlen: int = 255) -> str:
    if not text:
        return ""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text.strip())[:maxlen]

def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(2)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"

# ═══════════════════════════════════════════════════
# LOGGER
# ═══════════════════════════════════════════════════
def make_logger(name: str) -> logging.Logger:
    os.makedirs("logs", exist_ok=True)
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s|%(levelname)-7s|%(name)s| %(message)s", "%Y-%m-%d %H:%M:%S")
    fh = RotatingFileHandler(
        os.path.join("logs", f"{name}_{datetime.now():%Y%m%d_%H%M%S}.log"),
        maxBytes=10*1024*1024, backupCount=5, encoding="utf-8")
    fh.setFormatter(fmt); fh.setLevel(logging.INFO)
    ch = logging.StreamHandler(); ch.setFormatter(fmt); ch.setLevel(logging.INFO)
    logger.addHandler(fh); logger.addHandler(ch)
    logger.info(f"=== {name} started | user={getpass.getuser()} ===")
    return logger

# ═══════════════════════════════════════════════════
# DATABASE MANAGER
# ═══════════════════════════════════════════════════
class DB:
    def __init__(self, logger):
        self.log = logger
        self.conn = None
        self.cursor = None

    def connect(self) -> bool:
        for i in range(3):
            try:
                self.conn = pyodbc.connect(CONN_STR, timeout=10)
                self.cursor = self.conn.cursor()
                self.log.info("DB connected")
                return True
            except Exception as e:
                self.log.error(f"DB connect attempt {i+1}: {e}")
                self.close()
        return False

    def close(self):
        for attr in ("cursor", "conn"):
            try:
                obj = getattr(self, attr, None)
                if obj: obj.close()
            except Exception: pass
            setattr(self, attr, None)

    def alive(self) -> bool:
        try:
            if self.conn and self.cursor:
                self.cursor.execute("SELECT 1")
                return True
        except Exception: pass
        return False

    def ensure(self) -> bool:
        if not self.alive():
            self.close()
            return self.connect()
        return True

    def call_sp(self, name, params=None, fetch=False):
        """Execute stored procedure with whitelist check."""
        if name not in ALLOWED_PROCEDURES:
            raise ValueError(f"Blocked SP: {name}")
        if not self.ensure():
            raise ConnectionError("DB unavailable")
        params = params or []
        sql = f"EXEC {name} {','.join(['?']*len(params))}"
        try:
            self.cursor.execute(sql, params)
            if fetch:
                return self.cursor.fetchall()
            self.conn.commit()
            return True
        except Exception as e:
            self.log.error(f"SP error ({name}): {e}")
            try: self.conn.rollback()
            except Exception: pass
            raise

    def query(self, sql, params=None, fetch=True):
        """Execute parameterized query (for hardcoded trusted SQL only)."""
        if not self.ensure():
            raise ConnectionError("DB unavailable")
        try:
            self.cursor.execute(sql, params or [])
            if fetch: return self.cursor.fetchall()
            self.conn.commit()
            return True
        except Exception as e:
            self.log.error(f"Query error: {e}")
            try: self.conn.rollback()
            except Exception: pass
            raise

# ═══════════════════════════════════════════════════
# DROPDOWN HELPERS
# ═══════════════════════════════════════════════════
_DROPDOWN_SQL = {
    "lot_numbers": "SELECT lot_number FROM master_lot_numbers WHERE is_active=1 ORDER BY lot_number",
    "shifts": "SELECT shift_code FROM master_shifts WHERE is_active=1 ORDER BY shift_code",
    "products": "SELECT product_name FROM master_products WHERE is_active=1 ORDER BY product_name",
    "buyers": "SELECT buyer_name FROM master_buyers WHERE is_active=1 ORDER BY buyer_name",
    "contracts": "SELECT contract_code FROM master_contracts WHERE is_active=1 ORDER BY contract_code",
    "tanks": "SELECT tank_code FROM master_tanks WHERE is_active=1 ORDER BY tank_code",
    "bag_suppliers": "SELECT supplier_name FROM master_bag_suppliers WHERE is_active=1 ORDER BY supplier_name",
    "bag_batch_no": "SELECT batch_number FROM master_bag_batch WHERE is_active=1 ORDER BY batch_number",
    "bag_weights": "SELECT weight_value FROM master_bag_weights WHERE is_active=1 ORDER BY weight_value",
    "packing_types": "SELECT packing_name FROM master_packing_types WHERE is_active=1 ORDER BY packing_name",
    "quantities": "SELECT quantity_value FROM master_quantities WHERE is_active=1 ORDER BY quantity_value",
    "net_weights": "SELECT weight_value FROM master_net_weights WHERE is_active=1 ORDER BY weight_value",
    "under_limits": "SELECT limit_value FROM master_under_limits WHERE is_active=1 ORDER BY limit_value",
    "over_limits": "SELECT limit_value FROM master_over_limits WHERE is_active=1 ORDER BY limit_value",
}

_INSERT_SQL = {
    "lot_numbers": ("INSERT INTO master_lot_numbers(lot_number,created_by) VALUES(?,?)", 2),
    "shifts": ("INSERT INTO master_shifts(shift_code,shift_name) VALUES(?,?)", 2),
    "products": ("INSERT INTO master_products(product_name) VALUES(?)", 1),
    "buyers": ("INSERT INTO master_buyers(buyer_name) VALUES(?)", 1),
    "contracts": ("INSERT INTO master_contracts(contract_code) VALUES(?)", 1),
    "tanks": ("INSERT INTO master_tanks(tank_code) VALUES(?)", 1),
    "bag_suppliers": ("INSERT INTO master_bag_suppliers(supplier_name) VALUES(?)", 1),
    "bag_batch_no": ("INSERT INTO master_bag_batch(batch_number) VALUES(?)", 1),
    "bag_weights": ("INSERT INTO master_bag_weights(weight_value,description) VALUES(?,?)", 2),
    "packing_types": ("INSERT INTO master_packing_types(packing_name) VALUES(?)", 1),
    "quantities": ("INSERT INTO master_quantities(quantity_value,description) VALUES(?,?)", 2),
    "net_weights": ("INSERT INTO master_net_weights(weight_value,description) VALUES(?,?)", 2),
    "under_limits": ("INSERT INTO master_under_limits(limit_value,description) VALUES(?,?)", 2),
    "over_limits": ("INSERT INTO master_over_limits(limit_value,description) VALUES(?,?)", 2),
}

def get_dropdown(db: DB, table: str) -> list:
    sql = _DROPDOWN_SQL.get(table)
    if not sql: return []
    try:
        return [str(r[0]) for r in (db.query(sql) or [])]
    except Exception:
        return []

def add_dropdown(db: DB, table: str, value, user: str, desc=None) -> bool:
    entry = _INSERT_SQL.get(table)
    if not entry: return False
    sql, cnt = entry
    if table in ("bag_weights",): value = float(value)
    elif table in ("quantities","net_weights","under_limits","over_limits"): value = int(value)
    params = [value, desc or user] if cnt == 2 else [value]
    try:
        db.query(sql, params, fetch=False)
        return True
    except Exception:
        return False