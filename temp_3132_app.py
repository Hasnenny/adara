import sys
import os
import sqlite3
import hashlib
import json
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, 
                            QComboBox, QTabWidget, QFileDialog, QMessageBox, QCheckBox, 
                            QGroupBox, QFormLayout, QSpinBox, QDoubleSpinBox, QDialog, 
                            QScrollArea, QSplitter, QFrame, QHeaderView, QSystemTrayIcon,
                            QMenu, QAction, QStyle, QInputDialog, QDateEdit, QTimeEdit,
                            QGridLayout, QRadioButton, QButtonGroup, QSlider, QTextEdit,
                            QToolBar, QStatusBar, QProgressBar, QToolButton, QCalendarWidget,
                            QDesktopWidget, QDialogButtonBox, QSizePolicy, QStackedWidget,
                            QListWidget, QListWidgetItem, QAbstractItemView, QTreeView,
                            QTreeWidget, QTreeWidgetItem, QSpacerItem, QDial, QLCDNumber)
from PyQt5.QtCore import Qt, QTimer, QDateTime, QDate, QTime, QSize, QPoint, QRect, pyqtSignal, QThread, QSettings, QLocale, QTranslator, QMarginsF
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QPixmap, QPainter, QBrush, QPen, QTextCharFormat, QCursor, QFontDatabase, QTextDocument, QTextCursor, QPdfWriter, QPageLayout, QPageSize
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
import barcode
from barcode.writer import ImageWriter
import pyqtgraph as pg
import qrcode
from PIL import Image, ImageDraw, ImageFont
import io
import base64
import tempfile
import shutil
import zipfile
import platform
import subprocess
import ctypes
import uuid
import locale
import re
import serial
import usb.core
import usb.util
from fpdf import FPDF
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

# إعدادات اللغة العربية
locale.setlocale(locale.LC_ALL, 'ar_SA.utf8')

# قائمة العملات العربية والدولار
CURRENCIES = {
    "دولار أمريكي": "$",
    "درهم إماراتي": "د.إ",
    "دينار أردني": "د.أ",
    "دينار بحريني": "د.ب",
    "دينار جزائري": "د.ج",
    "ريال سعودي": "ر.س",
    "دينار عراقي": "د.ع",
    "ريال قطري": "ر.ق",
    "دينار كويتي": "د.ك",
    "دينار ليبي": "د.ل",
    "درهم مغربي": "د.م",
    "أوقية موريتانية": "أ.م",
    "ريال عماني": "ر.ع",
    "دينار تونسي": "د.ت",
    "ريال يمني": "ر.ي",
    "جنيه مصري": "ج.م",
    "ليرة سورية": "ل.س",
    "ليرة لبنانية": "ل.ل",
    "جنيه سوداني": "ج.س",
    "شلن صومالي": "ش.ص",
    "فرنك جيبوتي": "ف.ج",
    "ريال إيراني": "ر.إ",
    "ليرة تركية": "ل.ت",
    "شيكل إسرائيلي": "ش.إ",
    "دينار لبي": "د.ل",
    "درهم مغربي": "د.م",
    "أفغاني": "أفغ",
    "روبية باكستانية": "ر.ب"
}

# إنشاء قاعدة البيانات
def create_database():
    conn = sqlite3.connect('xnonedbs.db')
    cursor = conn.cursor()
    
    # جدول المستخدمين
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL,
        full_name TEXT,
        email TEXT,
        phone TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # جدول المنتجات
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        barcode TEXT UNIQUE,
        name TEXT NOT NULL,
        description TEXT,
        category TEXT,
        quantity REAL NOT NULL,
        price REAL NOT NULL,
        cost_price REAL,
        min_quantity REAL DEFAULT 0,
        expiry_date DATE,
        supplier TEXT,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # جدول الفواتير
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_number TEXT UNIQUE NOT NULL,
        customer_name TEXT,
        customer_phone TEXT,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        subtotal REAL NOT NULL,
        tax REAL DEFAULT 0,
        discount REAL DEFAULT 0,
        total REAL NOT NULL,
        payment_method TEXT,
        notes TEXT,
        user_id INTEGER,
        FOREIGN KEY (user_id) REFERENCES users (id)
    )
    ''')
    
    # جدول تفاصيل الفواتير
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS invoice_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER,
        product_id INTEGER,
        quantity REAL NOT NULL,
        price REAL NOT NULL,
        total REAL NOT NULL,
        FOREIGN KEY (invoice_id) REFERENCES invoices (id),
        FOREIGN KEY (product_id) REFERENCES products (id)
    )
    ''')
    
    # جدول الإعدادات
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        setting_key TEXT UNIQUE NOT NULL,
        setting_value TEXT NOT NULL
    )
    ''')
    
    # جدول التصنيفات
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        description TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # جدول الموردين
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        contact_person TEXT,
        phone TEXT,
        email TEXT,
        address TEXT,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # جدول العملاء
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        email TEXT,
        address TEXT,
        notes TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # جدول النسخ الاحتياطي
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS backups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        file_path TEXT NOT NULL,
        date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        size REAL,
        notes TEXT
    )
    ''')
    
    # جدول السجلات
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT NOT NULL,
        details TEXT,
        user_id INTEGER,
        timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users (id)
    )
    ''')
    
    # التحقق من وجود مستخدم مدير
    cursor.execute("SELECT * FROM users WHERE role = 'admin'")
    if not cursor.fetchone():
        # إنشاء مستخدم مدير افتراضي
        hashed_password = hashlib.sha256("admin123".encode()).hexdigest()
        cursor.execute("INSERT INTO users (username, password, role, full_name) VALUES (?, ?, ?, ?)",
                      ("admin", hashed_password, "admin", "مدير النظام"))
    
    # إضافة الإعدادات الافتراضية
    default_settings = [
        ("store_name", "متجري"),
        ("currency", "ريال سعودي"),
        ("currency_symbol", "ر.س"),
        ("tax_rate", "0.15"),
        ("invoice_prefix", "فاتورة-"),
        ("low_stock_threshold", "10"),
        ("backup_auto", "true"),
        ("backup_frequency", "7"),
        ("notifications", "true"),
        ("language", "ar"),
        ("theme", "light"),
        ("logo_path", ""),
        ("address", ""),
        ("phone", ""),
        ("email", ""),
        ("website", ""),
        ("facebook", ""),
        ("twitter", ""),
        ("instagram", ""),
        ("invoice_header", ""),
        ("invoice_footer", "شكراً لتعاملكم معنا"),
        ("receipt_printer", "default"),
        ("barcode_reader", "default"),
        ("show_profit_margin", "true"),
        ("auto_backup_path", ""),
        ("last_backup_date", ""),
        ("invoice_template", "default"),
        ("decimal_places", "2"),
        ("date_format", "dd/mm/yyyy"),
        ("time_format", "24h"),
        ("start_week_on", "sunday"),
        ("fiscal_year_start", "01/01"),
        ("enable_taxes", "true"),
        ("enable_discounts", "true"),
        ("enable_customer_management", "true"),
        ("enable_supplier_management", "true"),
        ("enable_expenses", "true"),
        ("enable_reports", "true"),
        ("enable_inventory", "true"),
        ("enable_invoices", "true"),
        ("enable_users", "true"),
        ("enable_settings", "true"),
        ("enable_backups", "true"),
        ("enable_logs", "true"),
        ("enable_categories", "true"),
        ("enable_products", "true"),
        ("enable_customers", "true"),
        ("enable_suppliers", "true"),
        ("enable_expenses", "true"),
        ("enable_reports", "true"),
        ("enable_invoices", "true"),
        ("enable_users", "true"),
        ("enable_settings", "true"),
        ("enable_backups", "true"),
        ("enable_logs", "true"),
        ("enable_categories", "true"),
        ("enable_products", "true"),
        ("enable_customers", "true"),
        ("enable_suppliers", "true"),
        ("enable_expenses", "true"),
        ("enable_reports", "true"),
        ("enable_invoices", "true"),
        ("enable_users", "true"),
        ("enable_settings", "true"),
        ("enable_backups", "true"),
        ("enable_logs", "true")
    ]
    
    for key, value in default_settings:
        cursor.execute("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", (key, value))
    
    conn.commit()
    conn.close()

# دالة لتشفير كلمة المرور
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# دالة للتحقق من كلمة المرور
def verify_password(password, hashed_password):
    return hashlib.sha256(password.encode()).hexdigest() == hashed_password

# دالة للحصول على الإعدادات
def get_setting(key, default=None):
    conn = sqlite3.connect('xnonedbs.db')
    cursor = conn.cursor()
    cursor.execute("SELECT setting_value FROM settings WHERE setting_key = ?", (key,))
    result = cursor.fetchone()
    conn.close()
    return result[0] if result else default

# دالة لتحديث الإعدادات
def update_setting(key, value):
    conn = sqlite3.connect('xnonedbs.db')
    cursor = conn.cursor()
    cursor.execute("UPDATE settings SET setting_value = ? WHERE setting_key = ?", (value, key))
    conn.commit()
    conn.close()

# دالة لإضافة سجل
def add_log(action, details, user_id=None):
    conn = sqlite3.connect('xnonedbs.db')
    cursor = conn.cursor()
    cursor.execute("INSERT INTO logs (action, details, user_id) VALUES (?, ?, ?)", (action, details, user_id))
    conn.commit()
    conn.close()

# واجهة تسجيل الدخول
class LoginWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("تسجيل الدخول - XnoneDBS")
        self.setFixedSize(400, 300)
        self.setWindowIcon(QIcon("icon.png"))
        
        # تصميم الواجهة
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                font-size: 14px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton#secondary {
                background-color: #f44336;
            }
            QPushButton#secondary:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        # شعار التطبيق
        logo_label = QLabel()
        logo_pixmap = QPixmap("logo.png")
        if not logo_pixmap.isNull():
            logo_label.setPixmap(logo_pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(logo_label)
        
        # عنوان التطبيق
        title_label = QLabel("XnoneDBS")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #2196F3; margin: 10px;")
        layout.addWidget(title_label)
        
        # نموذج تسجيل الدخول
        form_layout = QFormLayout()
        
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("اسم المستخدم")
        form_layout.addRow("اسم المستخدم:", self.username_input)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("كلمة المرور")
        self.password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("كلمة المرور:", self.password_input)
        
        layout.addLayout(form_layout)
        
        # أزرار تسجيل الدخول وإنشاء حساب جديد
        buttons_layout = QHBoxLayout()
        
        login_button = QPushButton("تسجيل الدخول")
        login_button.clicked.connect(self.login)
        buttons_layout.addWidget(login_button)
        
        register_button = QPushButton("إنشاء حساب جديد")
        register_button.setObjectName("secondary")
        register_button.clicked.connect(self.open_register_dialog)
        buttons_layout.addWidget(register_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # التركيز على حقل اسم المستخدم
        self.username_input.setFocus()
        
        # محاولة تسجيل الدخول بالضغط على Enter
        self.password_input.returnPressed.connect(self.login)
    
    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()
        
        if not username or not password:
            self.error_label.setText("يرجى إدخال اسم المستخدم وكلمة المرور")
            return
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, password, role, full_name FROM users WHERE username = ?", (username,))
        user = cursor.fetchone()
        conn.close()
        
        if user and verify_password(password, user[2]):
            # تسجيل الدخول بنجاح
            self.user_id = user[0]
            self.username = user[1]
            self.role = user[3]
            self.full_name = user[4]
            self.accept()
        else:
            self.error_label.setText("اسم المستخدم أو كلمة المرور غير صحيحة")
    
    def open_register_dialog(self):
        dialog = RegisterDialog(self)
        if dialog.exec_():
            self.username_input.setText(dialog.username)
            self.password_input.setText("")
            self.password_input.setFocus()

# نافذة إنشاء حساب جديد
class RegisterDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("إنشاء حساب جديد")
        self.setFixedSize(400, 400)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QComboBox {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                font-size: 14px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("اسم المستخدم")
        form_layout.addRow("اسم المستخدم:", self.username_input)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("كلمة المرور")
        self.password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("كلمة المرور:", self.password_input)
        
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setPlaceholderText("تأكيد كلمة المرور")
        self.confirm_password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("تأكيد كلمة المرور:", self.confirm_password_input)
        
        self.full_name_input = QLineEdit()
        self.full_name_input.setPlaceholderText("الاسم الكامل")
        form_layout.addRow("الاسم الكامل:", self.full_name_input)
        
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("البريد الإلكتروني")
        form_layout.addRow("البريد الإلكتروني:", self.email_input)
        
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("رقم الهاتف")
        form_layout.addRow("رقم الهاتف:", self.phone_input)
        
        self.role_combo = QComboBox()
        self.role_combo.addItems(["مستخدم عادي", "مدير"])
        form_layout.addRow("الصلاحية:", self.role_combo)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.register)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # التركيز على حقل اسم المستخدم
        self.username_input.setFocus()
    
    def register(self):
        username = self.username_input.text()
        password = self.password_input.text()
        confirm_password = self.confirm_password_input.text()
        full_name = self.full_name_input.text()
        email = self.email_input.text()
        phone = self.phone_input.text()
        role = "admin" if self.role_combo.currentText() == "مدير" else "user"
        
        # التحقق من الحقول المطلوبة
        if not username or not password or not confirm_password:
            self.error_label.setText("يرجى ملء الحقول المطلوبة")
            return
        
        # التحقق من تطابق كلمة المرور
        if password != confirm_password:
            self.error_label.setText("كلمة المرور غير متطابقة")
            return
        
        # التحقق من طول كلمة المرور
        if len(password) < 6:
            self.error_label.setText("كلمة المرور يجب أن تكون 6 أحرف على الأقل")
            return
        
        # التحقق من عدم وجود اسم المستخدم مسبقاً
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT username FROM users WHERE username = ?", (username,))
        if cursor.fetchone():
            conn.close()
            self.error_label.setText("اسم المستخدم موجود بالفعل")
            return
        
        # إضافة المستخدم الجديد
        hashed_password = hash_password(password)
        try:
            cursor.execute(
                "INSERT INTO users (username, password, role, full_name, email, phone) VALUES (?, ?, ?, ?, ?, ?)",
                (username, hashed_password, role, full_name, email, phone)
            )
            conn.commit()
            self.username = username
            conn.close()
            self.accept()
        except sqlite3.Error as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# النافذة الرئيسية للتطبيق
class MainWindow(QMainWindow):
    def __init__(self, user_id, username, role, full_name):
        super().__init__()
        self.user_id = user_id
        self.username = username
        self.role = role
        self.full_name = full_name
        
        # إعدادات التطبيق
        self.store_name = get_setting("store_name", "متجري")
        self.currency = get_setting("currency", "ريال سعودي")
        self.currency_symbol = get_setting("currency_symbol", "ر.س")
        self.tax_rate = float(get_setting("tax_rate", "0.15"))
        self.invoice_prefix = get_setting("invoice_prefix", "فاتورة-")
        self.low_stock_threshold = int(get_setting("low_stock_threshold", "10"))
        self.backup_auto = get_setting("backup_auto", "true") == "true"
        self.backup_frequency = int(get_setting("backup_frequency", "7"))
        self.notifications = get_setting("notifications", "true") == "true"
        self.theme = get_setting("theme", "light")
        self.decimal_places = int(get_setting("decimal_places", "2"))
        
        # إعداد الواجهة الرئيسية
        self.setWindowTitle(f"{self.store_name} - XnoneDBS")
        self.setWindowIcon(QIcon("icon.png"))
        self.setMinimumSize(1024, 768)
        
        # تحديد الثيم
        if self.theme == "dark":
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #2d2d2d;
                    color: #f0f0f0;
                }
                QWidget {
                    background-color: #2d2d2d;
                    color: #f0f0f0;
                }
                QTabWidget::pane {
                    border: 1px solid #3d3d3d;
                }
                QTabWidget::tab-bar {
                    left: 5px;
                }
                QTabBar::tab {
                    background-color: #3d3d3d;
                    color: #f0f0f0;
                    padding: 8px 16px;
                    border: 1px solid #3d3d3d;
                    margin-right: 2px;
                }
                QTabBar::tab:selected {
                    background-color: #4d4d4d;
                    border-bottom: 2px solid #2196F3;
                }
                QTableWidget {
                    background-color: #3d3d3d;
                    alternate-background-color: #4d4d4d;
                    gridline-color: #5d5d5d;
                    color: #f0f0f0;
                    border: 1px solid #3d3d3d;
                }
                QTableWidget::item {
                    padding: 5px;
                }
                QTableWidget::item:selected {
                    background-color: #2196F3;
                }
                QHeaderView::section {
                    background-color: #4d4d4d;
                    color: #f0f0f0;
                    padding: 5px;
                    border: 1px solid #5d5d5d;
                    font-weight: bold;
                }
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 4px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1e88e5;
                }
                QPushButton#danger {
                    background-color: #f44336;
                }
                QPushButton#danger:hover {
                    background-color: #d32f2f;
                }
                QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit, QTimeEdit {
                    padding: 8px;
                    border: 1px solid #5d5d5d;
                    border-radius: 4px;
                    background-color: #4d4d4d;
                    color: #f0f0f0;
                    font-size: 14px;
                }
                QLabel {
                    color: #f0f0f0;
                    font-size: 14px;
                }
                QGroupBox {
                    border: 1px solid #5d5d5d;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 10px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    padding: 0 5px;
                    color: #2196F3;
                }
                QStatusBar {
                    background-color: #3d3d3d;
                    color: #f0f0f0;
                }
                QMenuBar {
                    background-color: #3d3d3d;
                    color: #f0f0f0;
                }
                QMenuBar::item:selected {
                    background-color: #2196F3;
                }
                QMenu {
                    background-color: #4d4d4d;
                    color: #f0f0f0;
                }
                QMenu::item:selected {
                    background-color: #2196F3;
                }
                QToolBar {
                    background-color: #3d3d3d;
                    border: 1px solid #5d5d5d;
                }
                QToolButton {
                    background-color: transparent;
                    color: #f0f0f0;
                    border: none;
                    padding: 5px;
                }
                QToolButton:hover {
                    background-color: #2196F3;
                }
                QMessageBox {
                    background-color: #3d3d3d;
                }
                QMessageBox QPushButton {
                    min-width: 80px;
                }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #f5f5f5;
                    color: #333;
                }
                QWidget {
                    background-color: #f5f5f5;
                    color: #333;
                }
                QTabWidget::pane {
                    border: 1px solid #ddd;
                }
                QTabWidget::tab-bar {
                    left: 5px;
                }
                QTabBar::tab {
                    background-color: #e0e0e0;
                    color: #333;
                    padding: 8px 16px;
                    border: 1px solid #ddd;
                    margin-right: 2px;
                }
                QTabBar::tab:selected {
                    background-color: #fff;
                    border-bottom: 2px solid #2196F3;
                }
                QTableWidget {
                    background-color: white;
                    alternate-background-color: #f9f9f9;
                    gridline-color: #ddd;
                    color: #333;
                    border: 1px solid #ddd;
                }
                QTableWidget::item {
                    padding: 5px;
                }
                QTableWidget::item:selected {
                    background-color: #2196F3;
                    color: white;
                }
                QHeaderView::section {
                    background-color: #f0f0f0;
                    color: #333;
                    padding: 5px;
                    border: 1px solid #ddd;
                    font-weight: bold;
                }
                QPushButton {
                    background-color: #2196F3;
                    color: white;
                    border: none;
                    padding: 8px 16px;
                    border-radius: 4px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #1e88e5;
                }
                QPushButton#danger {
                    background-color: #f44336;
                }
                QPushButton#danger:hover {
                    background-color: #d32f2f;
                }
                QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit, QTimeEdit {
                    padding: 8px;
                    border: 1px solid #ddd;
                    border-radius: 4px;
                    background-color: white;
                    color: #333;
                    font-size: 14px;
                }
                QLabel {
                    color: #333;
                    font-size: 14px;
                }
                QGroupBox {
                    border: 1px solid #ddd;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 10px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    padding: 0 5px;
                    color: #2196F3;
                }
                QStatusBar {
                    background-color: #f0f0f0;
                    color: #333;
                }
                QMenuBar {
                    background-color: #f0f0f0;
                    color: #333;
                }
                QMenuBar::item:selected {
                    background-color: #2196F3;
                    color: white;
                }
                QMenu {
                    background-color: white;
                    color: #333;
                }
                QMenu::item:selected {
                    background-color: #2196F3;
                    color: white;
                }
                QToolBar {
                    background-color: #f0f0f0;
                    border: 1px solid #ddd;
                }
                QToolButton {
                    background-color: transparent;
                    color: #333;
                    border: none;
                    padding: 5px;
                }
                QToolButton:hover {
                    background-color: #2196F3;
                    color: white;
                }
                QMessageBox {
                    background-color: white;
                }
                QMessageBox QPushButton {
                    min-width: 80px;
                }
            """)
        
        # إنشاء شريط القوائم
        self.create_menu_bar()
        
        # إنشاء شريط الأدوات
        self.create_toolbar()
        
        # إنشاء واجهة التبويبات
        self.create_tab_widget()
        
        # إنشاء شريط الحالة
        self.create_status_bar()
        
        # إعداد أيقونة النظام
        self.setup_tray_icon()
        
        # إعداد المؤقت للنسخ الاحتياطي التلقائي
        self.setup_backup_timer()
        
        # إعداد المؤقت للتحقق من المخزون المنخفض
        self.setup_low_stock_timer()
        
        # تسجيل الدخول
        add_log("تسجيل الدخول", f"قام المستخدم {self.username} بتسجيل الدخول", self.user_id)
        
        # عرض رسالة ترحيب
        self.show_welcome_message()
    
    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # قائمة الملف
        file_menu = menubar.addMenu("ملف")
        
        backup_action = QAction("نسخ احتياطي", self)
        backup_action.triggered.connect(self.backup_data)
        file_menu.addAction(backup_action)
        
        restore_action = QAction("استرجاع", self)
        restore_action.triggered.connect(self.restore_data)
        file_menu.addAction(restore_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("خروج", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # قائمة المخزون
        inventory_menu = menubar.addMenu("المخزون")
        
        products_action = QAction("المنتجات", self)
        products_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.products_tab))
        inventory_menu.addAction(products_action)
        
        categories_action = QAction("التصنيفات", self)
        categories_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.categories_tab))
        inventory_menu.addAction(categories_action)
        
        low_stock_action = QAction("المنتجات منخفضة المخزون", self)
        low_stock_action.triggered.connect(self.show_low_stock_products)
        inventory_menu.addAction(low_stock_action)
        
        # قائمة المبيعات
        sales_menu = menubar.addMenu("المبيعات")
        
        new_invoice_action = QAction("فاتورة جديدة", self)
        new_invoice_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.invoices_tab))
        sales_menu.addAction(new_invoice_action)
        
        invoices_list_action = QAction("قائمة الفواتير", self)
        invoices_list_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.invoices_list_tab))
        sales_menu.addAction(invoices_list_action)
        
        # قائمة العملاء
        customers_menu = menubar.addMenu("العملاء")
        
        customers_action = QAction("إدارة العملاء", self)
        customers_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.customers_tab))
        customers_menu.addAction(customers_action)
        
        # قائمة الموردين
        suppliers_menu = menubar.addMenu("الموردين")
        
        suppliers_action = QAction("إدارة الموردين", self)
        suppliers_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.suppliers_tab))
        suppliers_menu.addAction(suppliers_action)
        
        # قائمة التقارير
        reports_menu = menubar.addMenu("التقارير")
        
        sales_report_action = QAction("تقرير المبيعات", self)
        sales_report_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.reports_tab))
        reports_menu.addAction(sales_report_action)
        
        inventory_report_action = QAction("تقرير المخزون", self)
        inventory_report_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.inventory_report_tab))
        reports_menu.addAction(inventory_report_action)
        
        profit_report_action = QAction("تقرير الأرباح", self)
        profit_report_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.profit_report_tab))
        reports_menu.addAction(profit_report_action)
        
        # قائمة الإعدادات
        settings_menu = menubar.addMenu("الإعدادات")
        
        general_settings_action = QAction("الإعدادات العامة", self)
        general_settings_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.settings_tab))
        settings_menu.addAction(general_settings_action)
        
        currency_settings_action = QAction("إعدادات العملة", self)
        currency_settings_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.currency_tab))
        settings_menu.addAction(currency_settings_action)
        
        users_settings_action = QAction("إدارة المستخدمين", self)
        users_settings_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.users_tab))
        if self.role == "admin":
            settings_menu.addAction(users_settings_action)
        
        # قائمة المساعدة
        help_menu = menubar.addMenu("مساعدة")
        
        about_action = QAction("حول", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
        help_action = QAction("مساعدة", self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)
    
    def create_toolbar(self):
        toolbar = QToolBar()
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        
        # زر المنتجات
        products_action = QAction("المنتجات", self)
        products_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.products_tab))
        toolbar.addAction(products_action)
        
        # زر الفواتير
        invoices_action = QAction("الفواتير", self)
        invoices_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.invoices_tab))
        toolbar.addAction(invoices_action)
        
        # زر العملاء
        customers_action = QAction("العملاء", self)
        customers_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.customers_tab))
        toolbar.addAction(customers_action)
        
        # زر التقارير
        reports_action = QAction("التقارير", self)
        reports_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.reports_tab))
        toolbar.addAction(reports_action)
        
        # زر الإعدادات
        settings_action = QAction("الإعدادات", self)
        settings_action.triggered.connect(lambda: self.tab_widget.setCurrentWidget(self.settings_tab))
        toolbar.addAction(settings_action)
        
        toolbar.addSeparator()
        
        # زر البحث
        search_action = QAction("بحث", self)
        search_action.triggered.connect(self.show_search_dialog)
        toolbar.addAction(search_action)
        
        # زر المساعدة
        help_action = QAction("مساعدة", self)
        help_action.triggered.connect(self.show_help)
        toolbar.addAction(help_action)
    
    def create_tab_widget(self):
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)
        
        # تبويب المنتجات
        self.products_tab = QWidget()
        self.create_products_tab()
        self.tab_widget.addTab(self.products_tab, "المنتجات")
        
        # تبويب الفواتير
        self.invoices_tab = QWidget()
        self.create_invoices_tab()
        self.tab_widget.addTab(self.invoices_tab, "الفواتير")
        
        # تبويب قائمة الفواتير
        self.invoices_list_tab = QWidget()
        self.create_invoices_list_tab()
        self.tab_widget.addTab(self.invoices_list_tab, "قائمة الفواتير")
        
        # تبويب العملاء
        self.customers_tab = QWidget()
        self.create_customers_tab()
        self.tab_widget.addTab(self.customers_tab, "العملاء")
        
        # تبويب الموردين
        self.suppliers_tab = QWidget()
        self.create_suppliers_tab()
        self.tab_widget.addTab(self.suppliers_tab, "الموردين")
        
        # تبويب التصنيفات
        self.categories_tab = QWidget()
        self.create_categories_tab()
        self.tab_widget.addTab(self.categories_tab, "التصنيفات")
        
        # تبويب التقارير
        self.reports_tab = QWidget()
        self.create_reports_tab()
        self.tab_widget.addTab(self.reports_tab, "التقارير")
        
        # تبويب تقرير المخزون
        self.inventory_report_tab = QWidget()
        self.create_inventory_report_tab()
        self.tab_widget.addTab(self.inventory_report_tab, "تقرير المخزون")
        
        # تبويب تقرير الأرباح
        self.profit_report_tab = QWidget()
        self.create_profit_report_tab()
        self.tab_widget.addTab(self.profit_report_tab, "تقرير الأرباح")
        
        # تبويب إعدادات العملة
        self.currency_tab = QWidget()
        self.create_currency_tab()
        self.tab_widget.addTab(self.currency_tab, "إعدادات العملة")
        
        # تبويب الإعدادات العامة
        self.settings_tab = QWidget()
        self.create_settings_tab()
        self.tab_widget.addTab(self.settings_tab, "الإعدادات")
        
        # تبويب إدارة المستخدمين (للمدير فقط)
        if self.role == "admin":
            self.users_tab = QWidget()
            self.create_users_tab()
            self.tab_widget.addTab(self.users_tab, "المستخدمون")
    
    def create_status_bar(self):
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        
        # عرض اسم المستخدم الحالي
        user_label = QLabel(f"المستخدم: {self.full_name} ({self.username})")
        status_bar.addPermanentWidget(user_label)
        
        # عرض التاريخ والوقت
        self.datetime_label = QLabel()
        self.update_datetime()
        status_bar.addPermanentWidget(self.datetime_label)
        
        # تحديث التاريخ والوقت كل ثانية
        self.datetime_timer = QTimer(self)
        self.datetime_timer.timeout.connect(self.update_datetime)
        self.datetime_timer.start(1000)
        
        # عرض عدد المنتجات منخفضة المخزون
        self.low_stock_label = QLabel()
        self.update_low_stock_count()
        status_bar.addWidget(self.low_stock_label)
    
    def update_datetime(self):
        now = QDateTime.currentDateTime()
        date_format = get_setting("date_format", "dd/mm/yyyy")
        time_format = get_setting("time_format", "24h")
        
        if date_format == "dd/mm/yyyy":
            date_str = now.toString("dd/MM/yyyy")
        elif date_format == "mm/dd/yyyy":
            date_str = now.toString("MM/dd/yyyy")
        else:  # yyyy/mm/dd
            date_str = now.toString("yyyy/MM/dd")
        
        if time_format == "12h":
            time_str = now.toString("hh:mm AP")
        else:  # 24h
            time_str = now.toString("hh:mm")
        
        self.datetime_label.setText(f"{date_str} {time_str}")
    
    def update_low_stock_count(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM products WHERE quantity <= ?", (self.low_stock_threshold,))
        count = cursor.fetchone()[0]
        conn.close()
        
        if count > 0:
            self.low_stock_label.setText(f"تنبيه: هناك {count} منتجات منخفضة المخزون")
            self.low_stock_label.setStyleSheet("color: red; font-weight: bold;")
        else:
            self.low_stock_label.setText("")
    
    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("icon.png"))
        
        # إنشاء قائمة النظام
        tray_menu = QMenu()
        
        show_action = QAction("إظهار", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)
        
        hide_action = QAction("إخفاء", self)
        hide_action.triggered.connect(self.hide)
        tray_menu.addAction(hide_action)
        
        tray_menu.addSeparator()
        
        exit_action = QAction("خروج", self)
        exit_action.triggered.connect(self.close)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
        # النقر المزدوج على الأيقونة لإظهار النافذة
        self.tray_icon.activated.connect(self.tray_icon_activated)
    
    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()
            self.activateWindow()
    
    def setup_backup_timer(self):
        if self.backup_auto:
            # تحويل الأيام إلى مللي ثانية
            interval = self.backup_frequency * 24 * 60 * 60 * 1000
            self.backup_timer = QTimer(self)
            self.backup_timer.timeout.connect(self.auto_backup)
            self.backup_timer.start(interval)
    
    def setup_low_stock_timer(self):
        # تحقق كل ساعة من المنتجات منخفضة المخزون
        self.low_stock_timer = QTimer(self)
        self.low_stock_timer.timeout.connect(self.check_low_stock)
        self.low_stock_timer.start(60 * 60 * 1000)  # ساعة واحدة بالمللي ثانية
    
    def auto_backup(self):
        try:
            # إنشاء مجلد النسخ الاحتياطي إذا لم يكن موجوداً
            backup_dir = get_setting("auto_backup_path", "")
            if not backup_dir:
                backup_dir = os.path.join(os.path.expanduser("~"), "XnoneDBS_Backups")
                os.makedirs(backup_dir, exist_ok=True)
            
            # إنشاء اسم ملف النسخ الاحتياطي
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"xnonedbs_backup_{timestamp}.db")
            
            # نسخ قاعدة البيانات
            shutil.copy2("xnonedbs.db", backup_file)
            
            # تسجيل النسخ الاحتياطي
            file_size = os.path.getsize(backup_file) / (1024 * 1024)  # حجم الملف بالميجابايت
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("INSERT INTO backups (file_path, size, notes) VALUES (?, ?, ?)",
                          (backup_file, file_size, "نسخ احتياطي تلقائي"))
            conn.commit()
            conn.close()
            
            # تحديث تاريخ آخر نسخ احتياطي
            update_setting("last_backup_date", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            
            # إظهار إشعار
            if self.notifications:
                self.tray_icon.showMessage(
                    "نسخ احتياطي",
                    "تم إتمام النسخ الاحتياطي التلقائي بنجاح",
                    QSystemTrayIcon.Information,
                    3000
                )
        except Exception as e:
            print(f"Error in auto backup: {str(e)}")
            if self.notifications:
                self.tray_icon.showMessage(
                    "خطأ في النسخ الاحتياطي",
                    f"حدث خطأ أثناء النسخ الاحتياطي: {str(e)}",
                    QSystemTrayIcon.Warning,
                    5000
                )
    
    def check_low_stock(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT name, quantity FROM products WHERE quantity <= ?", (self.low_stock_threshold,))
        low_stock_products = cursor.fetchall()
        conn.close()
        
        if low_stock_products and self.notifications:
            products_text = "\n".join([f"{name}: {quantity}" for name, quantity in low_stock_products])
            self.tray_icon.showMessage(
                "تنبيه المخزون",
                f"المنتجات التالية منخفضة المخزون:\n{products_text}",
                QSystemTrayIcon.Warning,
                5000
            )
        
        self.update_low_stock_count()
    
    def show_welcome_message(self):
        welcome_msg = f"مرحباً {self.full_name}!\n"
        welcome_msg += f"آخر تسجيل دخول: {self.get_last_login()}"
        
        QMessageBox.information(self, "مرحباً", welcome_msg)
    
    def get_last_login(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT timestamp FROM logs WHERE user_id = ? AND action = 'تسجيل الدخول' ORDER BY timestamp DESC LIMIT 1, 1", (self.user_id,))
        result = cursor.fetchone()
        conn.close()
        
        if result:
            last_login = datetime.strptime(result[0], "%Y-%m-%d %H:%M:%S")
            date_format = get_setting("date_format", "dd/mm/yyyy")
            time_format = get_setting("time_format", "24h")
            
            if date_format == "dd/mm/yyyy":
                date_str = last_login.strftime("%d/%m/%Y")
            elif date_format == "mm/dd/yyyy":
                date_str = last_login.strftime("%m/%d/%Y")
            else:  # yyyy/mm/dd
                date_str = last_login.strftime("%Y/%m/%d")
            
            if time_format == "12h":
                time_str = last_login.strftime("%I:%M %p")
            else:  # 24h
                time_str = last_login.strftime("%H:%M")
            
            return f"{date_str} {time_str}"
        else:
            "هذه هي المرة الأولى التي تسجل فيها الدخول"
    
    def create_products_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر إضافة منتج
        add_product_button = QPushButton("إضافة منتج")
        add_product_button.clicked.connect(self.add_product)
        toolbar_layout.addWidget(add_product_button)
        
        # زر تعديل منتج
        edit_product_button = QPushButton("تعديل منتج")
        edit_product_button.clicked.connect(self.edit_product)
        toolbar_layout.addWidget(edit_product_button)
        
        # زر حذف منتج
        delete_product_button = QPushButton("حذف منتج")
        delete_product_button.clicked.connect(self.delete_product)
        delete_product_button.setObjectName("danger")
        toolbar_layout.addWidget(delete_product_button)
        
        # زر استيراد المنتجات
        import_button = QPushButton("استيراد")
        import_button.clicked.connect(self.import_products)
        toolbar_layout.addWidget(import_button)
        
        # زر تصدير المنتجات
        export_button = QPushButton("تصدير")
        export_button.clicked.connect(self.export_products)
        toolbar_layout.addWidget(export_button)
        
        # زر تحديث الجدول
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_products_table)
        toolbar_layout.addWidget(refresh_button)
        
        # حقل البحث
        search_label = QLabel("بحث:")
        toolbar_layout.addWidget(search_label)
        
        self.product_search_input = QLineEdit()
        self.product_search_input.setPlaceholderText("ابحث عن منتج...")
        self.product_search_input.textChanged.connect(self.filter_products)
        toolbar_layout.addWidget(self.product_search_input)
        
        # قائمة التصفية
        filter_label = QLabel("التصفية:")
        toolbar_layout.addWidget(filter_label)
        
        self.product_filter_combo = QComboBox()
        self.product_filter_combo.addItems(["الكل", "منخفض المخزون", "منتهية الصلاحية", "فئة معينة"])
        self.product_filter_combo.currentIndexChanged.connect(self.filter_products)
        toolbar_layout.addWidget(self.product_filter_combo)
        
        layout.addLayout(toolbar_layout)
        
        # جدول المنتجات
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(10)
        self.products_table.setHorizontalHeaderLabels([
            "الباركود", "اسم المنتج", "الوصف", "التصنيف", "الكمية", 
            "سعر البيع", "سعر الشراء", "الحد الأدنى", "تاريخ الانتهاء", "المورد"
        ])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.products_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.products_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.products_table.setSortingEnabled(True)
        self.products_table.doubleClicked.connect(self.edit_product)
        layout.addWidget(self.products_table)
        
        # إخفاء بعض الأعمدة افتراضياً
        self.products_table.setColumnHidden(2, True)  # الوصف
        self.products_table.setColumnHidden(7, True)  # الحد الأدنى
        self.products_table.setColumnHidden(8, True)  # تاريخ الانتهاء
        self.products_table.setColumnHidden(9, True)  # المورد
        
        # زر تخصيص الأعمدة
        customize_columns_button = QPushButton("تخصيص الأعمدة")
        customize_columns_button.clicked.connect(self.customize_products_columns)
        layout.addWidget(customize_columns_button)
        
        self.products_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_products_table()
    
    def refresh_products_table(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products ORDER BY name")
        products = cursor.fetchall()
        conn.close()
        
        self.products_table.setRowCount(len(products))
        
        for row, product in enumerate(products):
            # الباركود
            self.products_table.setItem(row, 0, QTableWidgetItem(product[1] or ""))
            
            # اسم المنتج
            self.products_table.setItem(row, 1, QTableWidgetItem(product[2]))
            
            # الوصف
            self.products_table.setItem(row, 2, QTableWidgetItem(product[3] or ""))
            
            # التصنيف
            self.products_table.setItem(row, 3, QTableWidgetItem(product[4] or ""))
            
            # الكمية
            quantity_item = QTableWidgetItem(str(product[5]))
            if product[5] <= self.low_stock_threshold:
                quantity_item.setBackground(QColor(255, 200, 200))
            self.products_table.setItem(row, 4, quantity_item)
            
            # سعر البيع
            price_item = QTableWidgetItem(f"{product[6]:.{self.decimal_places}f} {self.currency_symbol}")
            self.products_table.setItem(row, 5, price_item)
            
            # سعر الشراء
            cost_price_item = QTableWidgetItem(f"{product[7]:.{self.decimal_places}f} {self.currency_symbol}" if product[7] else "")
            self.products_table.setItem(row, 6, cost_price_item)
            
            # الحد الأدنى
            self.products_table.setItem(row, 7, QTableWidgetItem(str(product[8]) if product[8] else ""))
            
            # تاريخ الانتهاء
            expiry_item = QTableWidgetItem(product[9] if product[9] else "")
            if product[9] and product[9] < datetime.now().strftime("%Y-%m-%d"):
                expiry_item.setBackground(QColor(255, 200, 200))
            self.products_table.setItem(row, 8, expiry_item)
            
            # المورد
            self.products_table.setItem(row, 9, QTableWidgetItem(product[10] or ""))
    
    def filter_products(self):
        search_text = self.product_search_input.text().lower()
        filter_type = self.product_filter_combo.currentText()
        
        for row in range(self.products_table.rowCount()):
            match = True
            
            # تطبيق البحث
            if search_text:
                row_match = False
                for col in range(self.products_table.columnCount()):
                    item = self.products_table.item(row, col)
                    if item and search_text in item.text().lower():
                        row_match = True
                        break
                match = row_match
            
            # تطبيق التصفية
            if match and filter_type != "الكل":
                quantity_item = self.products_table.item(row, 4)
                expiry_item = self.products_table.item(row, 8)
                
                if filter_type == "منخفض المخزون":
                    if quantity_item:
                        try:
                            quantity = float(quantity_item.text())
                            match = quantity <= self.low_stock_threshold
                        except:
                            match = False
                    else:
                        match = False
                
                elif filter_type == "منتهية الصلاحية":
                    if expiry_item and expiry_item.text():
                        try:
                            expiry_date = datetime.strptime(expiry_item.text(), "%Y-%m-%d")
                            match = expiry_date < datetime.now()
                        except:
                            match = False
                    else:
                        match = False
                
                elif filter_type == "فئة معينة":
                    # عرض مربع حوار لاختيار الفئة
                    categories = self.get_categories()
                    if categories:
                        category, ok = QInputDialog.getItem(
                            self, "اختر الفئة", "اختر الفئة للتصفية:", categories, 0, False
                        )
                        if ok and category:
                            category_item = self.products_table.item(row, 3)
                            match = category_item and category_item.text() == category
                        else:
                            match = False
                    else:
                        match = False
            
            self.products_table.setRowHidden(row, not match)
    
    def get_categories(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM categories ORDER BY name")
        categories = [row[0] for row in cursor.fetchall()]
        conn.close()
        return categories
    
    def customize_products_columns(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("تخصيص أعمدة المنتجات")
        dialog.setMinimumSize(400, 300)
        
        layout = QVBoxLayout()
        
        # قائمة الأعمدة
        columns_list = QListWidget()
        columns_list.setSelectionMode(QListWidget.MultiSelection)
        
        # إضافة الأعمدة مع حالتها الحالية
        columns = [
            ("الباركود", 0),
            ("اسم المنتج", 1),
            ("الوصف", 2),
            ("التصنيف", 3),
            ("الكمية", 4),
            ("سعر البيع", 5),
            ("سعر الشراء", 6),
            ("الحد الأدنى", 7),
            ("تاريخ الانتهاء", 8),
            ("المورد", 9)
        ]
        
        for name, index in columns:
            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, index)
            item.setSelected(not self.products_table.isColumnHidden(index))
            columns_list.addItem(item)
        
        layout.addWidget(columns_list)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(lambda: self.save_products_columns_settings(columns_list, dialog))
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.clicked.connect(dialog.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        dialog.setLayout(layout)
        dialog.exec_()
    
    def save_products_columns_settings(self, columns_list, dialog):
        for i in range(columns_list.count()):
            item = columns_list.item(i)
            column_index = item.data(Qt.UserRole)
            is_selected = item.isSelected()
            self.products_table.setColumnHidden(column_index, not is_selected)
        
        dialog.accept()
    
    def add_product(self):
        dialog = ProductDialog(self)
        if dialog.exec_():
            self.refresh_products_table()
            add_log("إضافة منتج", f"تم إضافة المنتج: {dialog.product_name}", self.user_id)
    
    def edit_product(self):
        selected_rows = self.products_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار منتج للتعديل")
            return
        
        row = selected_rows[0].row()
        barcode = self.products_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products WHERE barcode = ?", (barcode,))
        product = cursor.fetchone()
        conn.close()
        
        if product:
            dialog = ProductDialog(self, product)
            if dialog.exec_():
                self.refresh_products_table()
                add_log("تعديل منتج", f"تم تعديل المنتج: {dialog.product_name}", self.user_id)
    
    def delete_product(self):
        selected_rows = self.products_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار منتج للحذف")
            return
        
        row = selected_rows[0].row()
        barcode = self.products_table.item(row, 0).text()
        product_name = self.products_table.item(row, 1).text()
        
        reply = QMessageBox.question(
            self, "تأكيد الحذف", 
            f"هل أنت متأكد من حذف المنتج '{product_name}'؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("DELETE FROM products WHERE barcode = ?", (barcode,))
            conn.commit()
            conn.close()
            
            self.refresh_products_table()
            add_log("حذف منتج", f"تم حذف المنتج: {product_name}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم حذف المنتج بنجاح")
    
    def import_products(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "استيراد المنتجات", "", "ملفات Excel (*.xlsx);;ملفات CSV (*.csv)"
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)
            
            # التحقق من الأعمدة المطلوبة
            required_columns = ['name', 'quantity', 'price']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                QMessageBox.warning(
                    self, "خطأ", 
                    f"الملف المحدد يفتقر إلى الأعمدة المطلوبة: {', '.join(missing_columns)}"
                )
                return
            
            # إضافة المنتجات إلى قاعدة البيانات
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            imported_count = 0
            for _, row in df.iterrows():
                try:
                    # إنشاء باركود إذا لم يكن موجوداً
                    barcode = row.get('barcode', str(uuid.uuid4()))
                    
                    cursor.execute(
                        """
                        INSERT OR REPLACE INTO products 
                        (barcode, name, description, category, quantity, price, cost_price, min_quantity, expiry_date, supplier, notes)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            barcode,
                            row['name'],
                            row.get('description', ''),
                            row.get('category', ''),
                            row['quantity'],
                            row['price'],
                            row.get('cost_price', 0),
                            row.get('min_quantity', 0),
                            row.get('expiry_date', ''),
                            row.get('supplier', ''),
                            row.get('notes', '')
                        )
                    )
                    imported_count += 1
                except Exception as e:
                    print(f"Error importing product {row.get('name', 'Unknown')}: {str(e)}")
            
            conn.commit()
            conn.close()
            
            self.refresh_products_table()
            add_log("استيراد منتجات", f"تم استيراد {imported_count} منتج", self.user_id)
            QMessageBox.information(self, "نجاح", f"تم استيراد {imported_count} منتج بنجاح")
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء استيراد المنتجات: {str(e)}")
    
    def export_products(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "تصدير المنتجات", "", "ملفات Excel (*.xlsx);;ملفات CSV (*.csv)"
        )
        
        if not file_path:
            return
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            
            if file_path.endswith('.xlsx'):
                # استخدام pandas للتصدير إلى Excel
                df = pd.read_sql_query("SELECT * FROM products", conn)
                
                # إنشاء كاتب Excel
                writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='المنتجات', index=False)
                
                # الحصول على كائن ورقة العمل
                worksheet = writer.sheets['المنتجات']
                
                # إضافة تنسيق
                header_format = writer.book.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # كتابة العناوين مع التنسيق
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # ضبط عرض الأعمدة
                for i, col in enumerate(df.columns):
                    max_len = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2
                    worksheet.set_column(i, i, max_len)
                
                writer.close()
            else:
                # استخدام pandas للتصدير إلى CSV
                df = pd.read_sql_query("SELECT * FROM products", conn)
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            
            conn.close()
            
            add_log("تصدير منتجات", "تم تصدير المنتجات", self.user_id)
            QMessageBox.information(self, "نجاح", "تم تصدير المنتجات بنجاح")
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير المنتجات: {str(e)}")
    
    def show_low_stock_products(self):
        self.tab_widget.setCurrentWidget(self.products_tab)
        self.product_filter_combo.setCurrentText("منخفض المخزون")
        self.filter_products()
    
    def create_invoices_tab(self):
        layout = QVBoxLayout()
        
        # معلومات العميل
        customer_group = QGroupBox("معلومات العميل")
        customer_layout = QFormLayout()
        
        self.customer_name_input = QLineEdit()
        self.customer_name_input.setPlaceholderText("اسم العميل")
        customer_layout.addRow("اسم العميل:", self.customer_name_input)
        
        self.customer_phone_input = QLineEdit()
        self.customer_phone_input.setPlaceholderText("رقم الهاتف")
        customer_layout.addRow("رقم الهاتف:", self.customer_phone_input)
        
        customer_group.setLayout(customer_layout)
        layout.addWidget(customer_group)
        
        # تفاصيل الفاتورة
        invoice_details_group = QGroupBox("تفاصيل الفاتورة")
        invoice_details_layout = QFormLayout()
        
        self.invoice_number_label = QLabel(self.generate_invoice_number())
        invoice_details_layout.addRow("رقم الفاتورة:", self.invoice_number_label)
        
        self.invoice_date_input = QDateEdit()
        self.invoice_date_input.setDate(QDate.currentDate())
        self.invoice_date_input.setCalendarPopup(True)
        invoice_details_layout.addRow("التاريخ:", self.invoice_date_input)
        
        self.invoice_time_input = QTimeEdit()
        self.invoice_time_input.setTime(QTime.currentTime())
        invoice_details_layout.addRow("الوقت:", self.invoice_time_input)
        
        self.payment_method_combo = QComboBox()
        self.payment_method_combo.addItems(["نقدي", "بطاقة ائتمان", "تحويل بنكي", "أخرى"])
        invoice_details_layout.addRow("طريقة الدفع:", self.payment_method_combo)
        
        invoice_details_group.setLayout(invoice_details_layout)
        layout.addWidget(invoice_details_group)
        
        # إضافة المنتجات
        add_product_group = QGroupBox("إضافة منتج")
        add_product_layout = QHBoxLayout()
        
        # حقل الباركود
        self.barcode_input = QLineEdit()
        self.barcode_input.setPlaceholderText("مسح الباركود أو إدخاله يدوياً")
        self.barcode_input.returnPressed.connect(self.add_product_by_barcode)
        add_product_layout.addWidget(self.barcode_input)
        
        # زر البحث
        search_product_button = QPushButton("بحث")
        search_product_button.clicked.connect(self.search_product_for_invoice)
        add_product_layout.addWidget(search_product_button)
        
        add_product_group.setLayout(add_product_layout)
        layout.addWidget(add_product_group)
        
        # جدول منتجات الفاتورة
        self.invoice_items_table = QTableWidget()
        self.invoice_items_table.setColumnCount(5)
        self.invoice_items_table.setHorizontalHeaderLabels([
            "الباركود", "اسم المنتج", "الكمية", "السعر", "الإجمالي"
        ])
        self.invoice_items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.invoice_items_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.invoice_items_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.invoice_items_table)
        
        # أزرار التعديل والحذف
        items_buttons_layout = QHBoxLayout()
        
        edit_item_button = QPushButton("تعديل الكمية")
        edit_item_button.clicked.connect(self.edit_invoice_item)
        items_buttons_layout.addWidget(edit_item_button)
        
        remove_item_button = QPushButton("حذف من الفاتورة")
        remove_item_button.clicked.connect(self.remove_invoice_item)
        remove_item_button.setObjectName("danger")
        items_buttons_layout.addWidget(remove_item_button)
        
        layout.addLayout(items_buttons_layout)
        
        # ملخص الفاتورة
        summary_group = QGroupBox("ملخص الفاتورة")
        summary_layout = QFormLayout()
        
        self.subtotal_label = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addRow("المجموع:", self.subtotal_label)
        
        self.tax_input = QDoubleSpinBox()
        self.tax_input.setRange(0, 100)
        self.tax_input.setValue(self.tax_rate * 100)
        self.tax_input.setSuffix(" %")
        self.tax_input.valueChanged.connect(self.update_invoice_totals)
        summary_layout.addRow("الضريبة:", self.tax_input)
        
        self.tax_amount_label = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addRow("قيمة الضريبة:", self.tax_amount_label)
        
        self.discount_input = QDoubleSpinBox()
        self.discount_input.setRange(0, 1000000)
        self.discount_input.setValue(0)
        self.discount_input.setSuffix(f" {self.currency_symbol}")
        self.discount_input.valueChanged.connect(self.update_invoice_totals)
        summary_layout.addRow("الخصم:", self.discount_input)
        
        self.total_label = QLabel(f"0.00 {self.currency_symbol}")
        total_font = QFont()
        total_font.setBold(True)
        total_font.setPointSize(14)
        self.total_label.setFont(total_font)
        summary_layout.addRow("الإجمالي:", self.total_label)
        
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        # ملاحظات الفاتورة
        notes_group = QGroupBox("ملاحظات")
        notes_layout = QVBoxLayout()
        
        self.invoice_notes_input = QTextEdit()
        self.invoice_notes_input.setPlaceholderText("أدخل ملاحظات الفاتورة هنا...")
        notes_layout.addWidget(self.invoice_notes_input)
        
        notes_group.setLayout(notes_layout)
        layout.addWidget(notes_group)
        
        # أزرار الحفظ والطباعة
        actions_layout = QHBoxLayout()
        
        save_invoice_button = QPushButton("حفظ الفاتورة")
        save_invoice_button.clicked.connect(self.save_invoice)
        actions_layout.addWidget(save_invoice_button)
        
        print_invoice_button = QPushButton("طباعة الفاتورة")
        print_invoice_button.clicked.connect(self.print_invoice)
        actions_layout.addWidget(print_invoice_button)
        
        clear_invoice_button = QPushButton("مسح الفاتورة")
        clear_invoice_button.clicked.connect(self.clear_invoice)
        clear_invoice_button.setObjectName("danger")
        actions_layout.addWidget(clear_invoice_button)
        
        layout.addLayout(actions_layout)
        
        self.invoices_tab.setLayout(layout)
        
        # إعداد قارئ الباركود
        self.setup_barcode_reader()
        
        # تهيئة الفاتورة
        self.invoice_items = []
        self.update_invoice_totals()
    
    def generate_invoice_number(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM invoices")
        count = cursor.fetchone()[0]
        conn.close()
        
        return f"{self.invoice_prefix}{count + 1:06d}"
    
    def setup_barcode_reader(self):
        # إعداد مؤقت للباركود
        self.barcode_timer = QTimer(self)
        self.barcode_timer.setSingleShot(True)
        self.barcode_timer.timeout.connect(self.process_barcode)
        
        self.barcode_buffer = ""
        
        # ربط حقل الباركود بالحدث
        self.barcode_input.textChanged.connect(self.on_barcode_text_changed)
    
    def on_barcode_text_changed(self, text):
        # إضافة النص إلى المخزن المؤقت
        self.barcode_buffer += text
        
        # إعادة تعيين المؤقت
        self.barcode_timer.start(100)  # 100 مللي ثانية
    
    def process_barcode(self):
        barcode = self.barcode_buffer.strip()
        self.barcode_buffer = ""
        
        if barcode:
            self.add_product_by_barcode(barcode)
    
    def add_product_by_barcode(self, barcode=None):
        if not barcode:
            barcode = self.barcode_input.text().strip()
        
        if not barcode:
            return
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products WHERE barcode = ?", (barcode,))
        product = cursor.fetchone()
        conn.close()
        
        if product:
            # التحقق من وجود المنتج في الفاتورة مسبقاً
            for item in self.invoice_items:
                if item['barcode'] == barcode:
                    # زيادة الكمية
                    item['quantity'] += 1
                    item['total'] = item['quantity'] * item['price']
                    self.update_invoice_items_table()
                    self.update_invoice_totals()
                    self.barcode_input.clear()
                    return
            
            # إضافة منتج جديد للفاتورة
            self.invoice_items.append({
                'id': product[0],
                'barcode': product[1],
                'name': product[2],
                'price': product[6],
                'quantity': 1,
                'total': product[6]
            })
            
            self.update_invoice_items_table()
            self.update_invoice_totals()
            self.barcode_input.clear()
        else:
            QMessageBox.warning(self, "تحذير", "المنتج غير موجود في المخزون")
    
    def search_product_for_invoice(self):
        dialog = ProductSearchDialog(self)
        if dialog.exec_() and dialog.selected_product:
            product = dialog.selected_product
            
            # التحقق من وجود المنتج في الفاتورة مسبقاً
            for item in self.invoice_items:
                if item['barcode'] == product['barcode']:
                    # زيادة الكمية
                    item['quantity'] += 1
                    item['total'] = item['quantity'] * item['price']
                    self.update_invoice_items_table()
                    self.update_invoice_totals()
                    return
            
            # إضافة منتج جديد للفاتورة
            self.invoice_items.append({
                'id': product['id'],
                'barcode': product['barcode'],
                'name': product['name'],
                'price': product['price'],
                'quantity': 1,
                'total': product['price']
            })
            
            self.update_invoice_items_table()
            self.update_invoice_totals()
    
    def update_invoice_items_table(self):
        self.invoice_items_table.setRowCount(len(self.invoice_items))
        
        for row, item in enumerate(self.invoice_items):
            # الباركود
            self.invoice_items_table.setItem(row, 0, QTableWidgetItem(item['barcode']))
            
            # اسم المنتج
            self.invoice_items_table.setItem(row, 1, QTableWidgetItem(item['name']))
            
            # الكمية
            self.invoice_items_table.setItem(row, 2, QTableWidgetItem(str(item['quantity'])))
            
            # السعر
            self.invoice_items_table.setItem(row, 3, QTableWidgetItem(f"{item['price']:.{self.decimal_places}f} {self.currency_symbol}"))
            
            # الإجمالي
            self.invoice_items_table.setItem(row, 4, QTableWidgetItem(f"{item['total']:.{self.decimal_places}f} {self.currency_symbol}"))
    
    def update_invoice_totals(self):
        subtotal = sum(item['total'] for item in self.invoice_items)
        tax_rate = self.tax_input.value() / 100
        tax_amount = subtotal * tax_rate
        discount = self.discount_input.value()
        total = subtotal + tax_amount - discount
        
        self.subtotal_label.setText(f"{subtotal:.{self.decimal_places}f} {self.currency_symbol}")
        self.tax_amount_label.setText(f"{tax_amount:.{self.decimal_places}f} {self.currency_symbol}")
        self.total_label.setText(f"{total:.{self.decimal_places}f} {self.currency_symbol}")
    
    def edit_invoice_item(self):
        selected_rows = self.invoice_items_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار منتج لتعديل الكمية")
            return
        
        row = selected_rows[0].row()
        item = self.invoice_items[row]
        
        quantity, ok = QInputDialog.getDouble(
            self, "تعديل الكمية", 
            "أدخل الكمية الجديدة:", 
            item['quantity'], 0.01, 1000000, 2
        )
        
        if ok:
            item['quantity'] = quantity
            item['total'] = item['quantity'] * item['price']
            self.update_invoice_items_table()
            self.update_invoice_totals()
    
    def remove_invoice_item(self):
        selected_rows = self.invoice_items_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار منتج للحذف")
            return
        
        row = selected_rows[0].row()
        del self.invoice_items[row]
        
        self.update_invoice_items_table()
        self.update_invoice_totals()
    
    def save_invoice(self):
        if not self.invoice_items:
            QMessageBox.warning(self, "تحذير", "لا توجد منتجات في الفاتورة")
            return
        
        # الحصول على بيانات الفاتورة
        invoice_number = self.invoice_number_label.text()
        customer_name = self.customer_name_input.text()
        customer_phone = self.customer_phone_input.text()
        invoice_date = self.invoice_date_input.date().toString("yyyy-MM-dd")
        invoice_time = self.invoice_time_input.time().toString("HH:mm:ss")
        payment_method = self.payment_method_combo.currentText()
        notes = self.invoice_notes_input.toPlainText()
        subtotal = sum(item['total'] for item in self.invoice_items)
        tax_rate = self.tax_input.value() / 100
        tax_amount = subtotal * tax_rate
        discount = self.discount_input.value()
        total = subtotal + tax_amount - discount
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            # حفظ الفاتورة
            cursor.execute(
                """
                INSERT INTO invoices 
                (invoice_number, customer_name, customer_phone, date, subtotal, tax, discount, total, payment_method, notes, user_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    invoice_number,
                    customer_name,
                    customer_phone,
                    f"{invoice_date} {invoice_time}",
                    subtotal,
                    tax_amount,
                    discount,
                    total,
                    payment_method,
                    notes,
                    self.user_id
                )
            )
            
            invoice_id = cursor.lastrowid
            
            # حفظ تفاصيل الفاتورة
            for item in self.invoice_items:
                cursor.execute(
                    """
                    INSERT INTO invoice_items 
                    (invoice_id, product_id, quantity, price, total)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (
                        invoice_id,
                        item['id'],
                        item['quantity'],
                        item['price'],
                        item['total']
                    )
                )
                
                # تحديث كمية المنتج في المخزون
                cursor.execute(
                    "UPDATE products SET quantity = quantity - ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                    (item['quantity'], item['id'])
                )
            
            conn.commit()
            conn.close()
            
            add_log("حفظ فاتورة", f"تم حفظ الفاتورة: {invoice_number}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم حفظ الفاتورة بنجاح")
            
            # مسح الفاتورة الحالية
            self.clear_invoice()
            
            # تحديث جدول المنتجات
            self.refresh_products_table()
            
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء حفظ الفاتورة: {str(e)}")
    
    def print_invoice(self):
        if not self.invoice_items:
            QMessageBox.warning(self, "تحذير", "لا توجد منتجات في الفاتورة")
            return
        
        # حفظ الفاتورة أولاً
        self.save_invoice()
        
        # الحصول على بيانات الفاتورة
        invoice_number = self.invoice_number_label.text()
        customer_name = self.customer_name_input.text()
        customer_phone = self.customer_phone_input.text()
        invoice_date = self.invoice_date_input.date().toString("yyyy-MM-dd")
        invoice_time = self.invoice_time_input.time().toString("HH:mm:ss")
        payment_method = self.payment_method_combo.currentText()
        notes = self.invoice_notes_input.toPlainText()
        subtotal = sum(item['total'] for item in self.invoice_items)
        tax_rate = self.tax_input.value() / 100
        tax_amount = subtotal * tax_rate
        discount = self.discount_input.value()
        total = subtotal + tax_amount - discount
        
        # إنشاء ملف PDF
        file_path, _ = QFileDialog.getSaveFileName(
            self, "حفظ الفاتورة", f"{invoice_number}.pdf", "ملفات PDF (*.pdf)"
        )
        
        if not file_path:
            return
        
        try:
            # إنشاء مستند PDF
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # إنشاء قائمة العناصر
            story = []
            
            # إضافة الشعار إذا كان موجوداً
            logo_path = get_setting("logo_path", "")
            if logo_path and os.path.exists(logo_path):
                logo = Image(logo_path, width=100, height=100)
                story.append(logo)
            
            # إضافة عنوان الفاتورة
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph(f"فاتورة مبيعات - {self.store_name}", title_style))
            
            # إضافة معلومات المتجر
            store_info = []
            store_info.append(Paragraph(f"<b>اسم المتجر:</b> {self.store_name}", styles['Normal']))
            store_info.append(Paragraph(f"<b>العنوان:</b> {get_setting('address', '')}", styles['Normal']))
            store_info.append(Paragraph(f"<b>الهاتف:</b> {get_setting('phone', '')}", styles['Normal']))
            store_info.append(Paragraph(f"<b>البريد الإلكتروني:</b> {get_setting('email', '')}", styles['Normal']))
            
            store_table = Table([store_info], colWidths=[3*inch])
            store_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(store_table)
            story.append(Spacer(1, 12))
            
            # إضافة معلومات الفاتورة
            invoice_info = []
            invoice_info.append([Paragraph("<b>رقم الفاتورة:</b>", styles['Normal']), Paragraph(invoice_number, styles['Normal'])])
            invoice_info.append([Paragraph("<b>التاريخ:</b>", styles['Normal']), Paragraph(f"{invoice_date} {invoice_time}", styles['Normal'])])
            invoice_info.append([Paragraph("<b>اسم العميل:</b>", styles['Normal']), Paragraph(customer_name, styles['Normal'])])
            invoice_info.append([Paragraph("<b>رقم الهاتف:</b>", styles['Normal']), Paragraph(customer_phone, styles['Normal'])])
            invoice_info.append([Paragraph("<b>طريقة الدفع:</b>", styles['Normal']), Paragraph(payment_method, styles['Normal'])])
            
            invoice_table = Table(invoice_info, colWidths=[1.5*inch, 3*inch])
            invoice_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(invoice_table)
            story.append(Spacer(1, 12))
            
            # إضافة جدول المنتجات
            items_data = [['#', 'الباركود', 'اسم المنتج', 'الكمية', 'السعر', 'الإجمالي']]
            
            for i, item in enumerate(self.invoice_items, 1):
                items_data.append([
                    str(i),
                    item['barcode'],
                    item['name'],
                    f"{item['quantity']:.{self.decimal_places}f}",
                    f"{item['price']:.{self.decimal_places}f} {self.currency_symbol}",
                    f"{item['total']:.{self.decimal_places}f} {self.currency_symbol}"
                ])
            
            items_table = Table(items_data, colWidths=[0.5*inch, 1*inch, 2*inch, 0.8*inch, 1*inch, 1*inch])
            items_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ]))
            story.append(items_table)
            story.append(Spacer(1, 12))
            
            # إضافة ملخص الفاتورة
            summary_data = [
                [Paragraph("<b>المجموع:</b>", styles['Normal']), Paragraph(f"{subtotal:.{self.decimal_places}f} {self.currency_symbol}", styles['Normal'])],
                [Paragraph("<b>الضريبة:</b>", styles['Normal']), Paragraph(f"{tax_amount:.{self.decimal_places}f} {self.currency_symbol}", styles['Normal'])],
                [Paragraph("<b>الخصم:</b>", styles['Normal']), Paragraph(f"{discount:.{self.decimal_places}f} {self.currency_symbol}", styles['Normal'])],
                [Paragraph("<b>الإجمالي:</b>", styles['Normal']), Paragraph(f"<b>{total:.{self.decimal_places}f} {self.currency_symbol}</b>", styles['Normal'])]
            ]
            
            summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # إضافة الملاحظات
            if notes:
                story.append(Paragraph("<b>ملاحظات:</b>", styles['Normal']))
                story.append(Paragraph(notes, styles['Normal']))
                story.append(Spacer(1, 12))
            
            # إضافة تذييل الفاتورة
            footer_text = get_setting("invoice_footer", "شكراً لتعاملكم معنا")
            story.append(Paragraph(footer_text, styles['Normal']))
            
            # بنشاء ملف PDF
            doc.build(story)
            
            # فتح الملف
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            
            add_log("طباعة فاتورة", f"تم طباعة الفاتورة: {invoice_number}", self.user_id)
            
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء طباعة الفاتورة: {str(e)}")
    
    def clear_invoice(self):
        self.invoice_items = []
        self.update_invoice_items_table()
        self.update_invoice_totals()
        
        self.customer_name_input.clear()
        self.customer_phone_input.clear()
        self.invoice_number_label.setText(self.generate_invoice_number())
        self.invoice_date_input.setDate(QDate.currentDate())
        self.invoice_time_input.setTime(QTime.currentTime())
        self.payment_method_combo.setCurrentIndex(0)
        self.invoice_notes_input.clear()
        self.tax_input.setValue(self.tax_rate * 100)
        self.discount_input.setValue(0)
    
    def create_invoices_list_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر تحديث الجدول
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_invoices_list)
        toolbar_layout.addWidget(refresh_button)
        
        # زر عرض الفاتورة
        view_button = QPushButton("عرض الفاتورة")
        view_button.clicked.connect(self.view_invoice)
        toolbar_layout.addWidget(view_button)
        
        # زر طباعة الفاتورة
        print_button = QPushButton("طباعة الفاتورة")
        print_button.clicked.connect(self.print_selected_invoice)
        toolbar_layout.addWidget(print_button)
        
        # حقل البحث
        search_label = QLabel("بحث:")
        toolbar_layout.addWidget(search_label)
        
        self.invoice_search_input = QLineEdit()
        self.invoice_search_input.setPlaceholderText("ابحث عن فاتورة...")
        self.invoice_search_input.textChanged.connect(self.filter_invoices)
        toolbar_layout.addWidget(self.invoice_search_input)
        
        # قائمة التصفية
        filter_label = QLabel("التصفية:")
        toolbar_layout.addWidget(filter_label)
        
        self.invoice_filter_combo = QComboBox()
        self.invoice_filter_combo.addItems(["الكل", "اليوم", "هذا الأسبوع", "هذا الشهر", "هذه السنة"])
        self.invoice_filter_combo.currentIndexChanged.connect(self.filter_invoices)
        toolbar_layout.addWidget(self.invoice_filter_combo)
        
        layout.addLayout(toolbar_layout)
        
        # جدول الفواتير
        self.invoices_table = QTableWidget()
        self.invoices_table.setColumnCount(7)
        self.invoices_table.setHorizontalHeaderLabels([
            "رقم الفاتورة", "اسم العميل", "التاريخ", "المبلغ", "طريقة الدفع", "المستخدم", "ملاحظات"
        ])
        self.invoices_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.invoices_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.invoices_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.invoices_table.setSortingEnabled(True)
        self.invoices_table.doubleClicked.connect(self.view_invoice)
        layout.addWidget(self.invoices_table)
        
        self.invoices_list_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_invoices_list()
    
    def refresh_invoices_list(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("""
            SELECT i.invoice_number, i.customer_name, i.date, i.total, i.payment_method, u.full_name, i.notes
            FROM invoices i
            JOIN users u ON i.user_id = u.id
            ORDER BY i.date DESC
        """)
        invoices = cursor.fetchall()
        conn.close()
        
        self.invoices_table.setRowCount(len(invoices))
        
        for row, invoice in enumerate(invoices):
            # رقم الفاتورة
            self.invoices_table.setItem(row, 0, QTableWidgetItem(invoice[0]))
            
            # اسم العميل
            self.invoices_table.setItem(row, 1, QTableWidgetItem(invoice[1] or ""))
            
            # التاريخ
            date_str = datetime.strptime(invoice[2], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M")
            self.invoices_table.setItem(row, 2, QTableWidgetItem(date_str))
            
            # المبلغ
            amount_item = QTableWidgetItem(f"{invoice[3]:.{self.decimal_places}f} {self.currency_symbol}")
            self.invoices_table.setItem(row, 3, amount_item)
            
            # طريقة الدفع
            self.invoices_table.setItem(row, 4, QTableWidgetItem(invoice[4]))
            
            # المستخدم
            self.invoices_table.setItem(row, 5, QTableWidgetItem(invoice[5]))
            
            # ملاحظات
            self.invoices_table.setItem(row, 6, QTableWidgetItem(invoice[6] or ""))
    
    def filter_invoices(self):
        search_text = self.invoice_search_input.text().lower()
        filter_type = self.invoice_filter_combo.currentText()
        
        for row in range(self.invoices_table.rowCount()):
            match = True
            
            # تطبيق البحث
            if search_text:
                row_match = False
                for col in range(self.invoices_table.columnCount()):
                    item = self.invoices_table.item(row, col)
                    if item and search_text in item.text().lower():
                        row_match = True
                        break
                match = row_match
            
            # تطبيق التصفية
            if match and filter_type != "الكل":
                date_item = self.invoices_table.item(row, 2)
                if date_item:
                    try:
                        invoice_date = datetime.strptime(date_item.text(), "%Y-%m-%d %H:%M")
                        today = datetime.now()
                        
                        if filter_type == "اليوم":
                            match = invoice_date.date() == today.date()
                        elif filter_type == "هذا الأسبوع":
                            start_of_week = today.date() - today.weekday()
                            match = invoice_date.date() >= start_of_week
                        elif filter_type == "هذا الشهر":
                            match = invoice_date.month == today.month and invoice_date.year == today.year
                        elif filter_type == "هذه السنة":
                            match = invoice_date.year == today.year
                    except:
                        match = False
                else:
                    match = False
            
            self.invoices_table.setRowHidden(row, not match)
    
    def view_invoice(self):
        selected_rows = self.invoices_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار فاتورة للعرض")
            return
        
        row = selected_rows[0].row()
        invoice_number = self.invoices_table.item(row, 0).text()
        
        dialog = InvoiceViewDialog(self, invoice_number)
        dialog.exec_()
    
    def print_selected_invoice(self):
        selected_rows = self.invoices_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار فاتورة للطباعة")
            return
        
        row = selected_rows[0].row()
        invoice_number = self.invoices_table.item(row, 0).text()
        
        dialog = InvoiceViewDialog(self, invoice_number)
        dialog.print_invoice()
    
    def create_customers_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر إضافة عميل
        add_customer_button = QPushButton("إضافة عميل")
        add_customer_button.clicked.connect(self.add_customer)
        toolbar_layout.addWidget(add_customer_button)
        
        # زر تعديل عميل
        edit_customer_button = QPushButton("تعديل عميل")
        edit_customer_button.clicked.connect(self.edit_customer)
        toolbar_layout.addWidget(edit_customer_button)
        
        # زر حذف عميل
        delete_customer_button = QPushButton("حذف عميل")
        delete_customer_button.clicked.connect(self.delete_customer)
        delete_customer_button.setObjectName("danger")
        toolbar_layout.addWidget(delete_customer_button)
        
        # زر تحديث الجدول
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_customers_table)
        toolbar_layout.addWidget(refresh_button)
        
        # حقل البحث
        search_label = QLabel("بحث:")
        toolbar_layout.addWidget(search_label)
        
        self.customer_search_input = QLineEdit()
        self.customer_search_input.setPlaceholderText("ابحث عن عميل...")
        self.customer_search_input.textChanged.connect(self.filter_customers)
        toolbar_layout.addWidget(self.customer_search_input)
        
        layout.addLayout(toolbar_layout)
        
        # جدول العملاء
        self.customers_table = QTableWidget()
        self.customers_table.setColumnCount(5)
        self.customers_table.setHorizontalHeaderLabels([
            "الاسم", "رقم الهاتف", "البريد الإلكتروني", "العنوان", "ملاحظات"
        ])
        self.customers_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.customers_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.customers_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.customers_table.setSortingEnabled(True)
        self.customers_table.doubleClicked.connect(self.edit_customer)
        layout.addWidget(self.customers_table)
        
        self.customers_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_customers_table()
    
    def refresh_customers_table(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers ORDER BY name")
        customers = cursor.fetchall()
        conn.close()
        
        self.customers_table.setRowCount(len(customers))
        
        for row, customer in enumerate(customers):
            # الاسم
            self.customers_table.setItem(row, 0, QTableWidgetItem(customer[1]))
            
            # رقم الهاتف
            self.customers_table.setItem(row, 1, QTableWidgetItem(customer[2] or ""))
            
            # البريد الإلكتروني
            self.customers_table.setItem(row, 2, QTableWidgetItem(customer[3] or ""))
            
            # العنوان
            self.customers_table.setItem(row, 3, QTableWidgetItem(customer[4] or ""))
            
            # ملاحظات
            self.customers_table.setItem(row, 4, QTableWidgetItem(customer[5] or ""))
    
    def filter_customers(self):
        search_text = self.customer_search_input.text().lower()
        
        for row in range(self.customers_table.rowCount()):
            match = False
            
            for col in range(self.customers_table.columnCount()):
                item = self.customers_table.item(row, col)
                if item and search_text in item.text().lower():
                    match = True
                    break
            
            self.customers_table.setRowHidden(row, not match)
    
    def add_customer(self):
        dialog = CustomerDialog(self)
        if dialog.exec_():
            self.refresh_customers_table()
            add_log("إضافة عميل", f"تم إضافة العميل: {dialog.customer_name}", self.user_id)
    
    def edit_customer(self):
        selected_rows = self.customers_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار عميل للتعديل")
            return
        
        row = selected_rows[0].row()
        customer_name = self.customers_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers WHERE name = ?", (customer_name,))
        customer = cursor.fetchone()
        conn.close()
        
        if customer:
            dialog = CustomerDialog(self, customer)
            if dialog.exec_():
                self.refresh_customers_table()
                add_log("تعديل عميل", f"تم تعديل العميل: {dialog.customer_name}", self.user_id)
    
    def delete_customer(self):
        selected_rows = self.customers_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار عميل للحذف")
            return
        
        row = selected_rows[0].row()
        customer_name = self.customers_table.item(row, 0).text()
        
        reply = QMessageBox.question(
            self, "تأكيد الحذف", 
            f"هل أنت متأكد من حذف العميل '{customer_name}'؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("DELETE FROM customers WHERE name = ?", (customer_name,))
            conn.commit()
            conn.close()
            
            self.refresh_customers_table()
            add_log("حذف عميل", f"تم حذف العميل: {customer_name}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم حذف العميل بنجاح")
    
    def create_suppliers_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر إضافة مورد
        add_supplier_button = QPushButton("إضافة مورد")
        add_supplier_button.clicked.connect(self.add_supplier)
        toolbar_layout.addWidget(add_supplier_button)
        
        # زر تعديل مورد
        edit_supplier_button = QPushButton("تعديل مورد")
        edit_supplier_button.clicked.connect(self.edit_supplier)
        toolbar_layout.addWidget(edit_supplier_button)
        
        # زر حذف مورد
        delete_supplier_button = QPushButton("حذف مورد")
        delete_supplier_button.clicked.connect(self.delete_supplier)
        delete_supplier_button.setObjectName("danger")
        toolbar_layout.addWidget(delete_supplier_button)
        
        # زر تحديث الجدول
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_suppliers_table)
        toolbar_layout.addWidget(refresh_button)
        
        # حقل البحث
        search_label = QLabel("بحث:")
        toolbar_layout.addWidget(search_label)
        
        self.supplier_search_input = QLineEdit()
        self.supplier_search_input.setPlaceholderText("ابحث عن مورد...")
        self.supplier_search_input.textChanged.connect(self.filter_suppliers)
        toolbar_layout.addWidget(self.supplier_search_input)
        
        layout.addLayout(toolbar_layout)
        
        # جدول الموردين
        self.suppliers_table = QTableWidget()
        self.suppliers_table.setColumnCount(6)
        self.suppliers_table.setHorizontalHeaderLabels([
            "الاسم", "شخص الاتصال", "رقم الهاتف", "البريد الإلكتروني", "العنوان", "ملاحظات"
        ])
        self.suppliers_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.suppliers_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.suppliers_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.suppliers_table.setSortingEnabled(True)
        self.suppliers_table.doubleClicked.connect(self.edit_supplier)
        layout.addWidget(self.suppliers_table)
        
        self.suppliers_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_suppliers_table()
    
    def refresh_suppliers_table(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM suppliers ORDER BY name")
        suppliers = cursor.fetchall()
        conn.close()
        
        self.suppliers_table.setRowCount(len(suppliers))
        
        for row, supplier in enumerate(suppliers):
            # الاسم
            self.suppliers_table.setItem(row, 0, QTableWidgetItem(supplier[1]))
            
            # شخص الاتصال
            self.suppliers_table.setItem(row, 1, QTableWidgetItem(supplier[2] or ""))
            
            # رقم الهاتف
            self.suppliers_table.setItem(row, 2, QTableWidgetItem(supplier[3] or ""))
            
            # البريد الإلكتروني
            self.suppliers_table.setItem(row, 3, QTableWidgetItem(supplier[4] or ""))
            
            # العنوان
            self.suppliers_table.setItem(row, 4, QTableWidgetItem(supplier[5] or ""))
            
            # ملاحظات
            self.suppliers_table.setItem(row, 5, QTableWidgetItem(supplier[6] or ""))
    
    def filter_suppliers(self):
        search_text = self.supplier_search_input.text().lower()
        
        for row in range(self.suppliers_table.rowCount()):
            match = False
            
            for col in range(self.suppliers_table.columnCount()):
                item = self.suppliers_table.item(row, col)
                if item and search_text in item.text().lower():
                    match = True
                    break
            
            self.suppliers_table.setRowHidden(row, not match)
    
    def add_supplier(self):
        dialog = SupplierDialog(self)
        if dialog.exec_():
            self.refresh_suppliers_table()
            add_log("إضافة مورد", f"تم إضافة المورد: {dialog.supplier_name}", self.user_id)
    
    def edit_supplier(self):
        selected_rows = self.suppliers_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار مورد للتعديل")
            return
        
        row = selected_rows[0].row()
        supplier_name = self.suppliers_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM suppliers WHERE name = ?", (supplier_name,))
        supplier = cursor.fetchone()
        conn.close()
        
        if supplier:
            dialog = SupplierDialog(self, supplier)
            if dialog.exec_():
                self.refresh_suppliers_table()
                add_log("تعديل مورد", f"تم تعديل المورد: {dialog.supplier_name}", self.user_id)
    
    def delete_supplier(self):
        selected_rows = self.suppliers_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار مورد للحذف")
            return
        
        row = selected_rows[0].row()
        supplier_name = self.suppliers_table.item(row, 0).text()
        
        reply = QMessageBox.question(
            self, "تأكيد الحذف", 
            f"هل أنت متأكد من حذف المورد '{supplier_name}'؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("DELETE FROM suppliers WHERE name = ?", (supplier_name,))
            conn.commit()
            conn.close()
            
            self.refresh_suppliers_table()
            add_log("حذف مورد", f"تم حذف المورد: {supplier_name}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم حذف المورد بنجاح")
    
    def create_categories_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر إضافة تصنيف
        add_category_button = QPushButton("إضافة تصنيف")
        add_category_button.clicked.connect(self.add_category)
        toolbar_layout.addWidget(add_category_button)
        
        # زر تعديل تصنيف
        edit_category_button = QPushButton("تعديل تصنيف")
        edit_category_button.clicked.connect(self.edit_category)
        toolbar_layout.addWidget(edit_category_button)
        
        # زر حذف تصنيف
        delete_category_button = QPushButton("حذف تصنيف")
        delete_category_button.clicked.connect(self.delete_category)
        delete_category_button.setObjectName("danger")
        toolbar_layout.addWidget(delete_category_button)
        
        # زر تحديث الجدول
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_categories_table)
        toolbar_layout.addWidget(refresh_button)
        
        # حقل البحث
        search_label = QLabel("بحث:")
        toolbar_layout.addWidget(search_label)
        
        self.category_search_input = QLineEdit()
        self.category_search_input.setPlaceholderText("ابحث عن تصنيف...")
        self.category_search_input.textChanged.connect(self.filter_categories)
        toolbar_layout.addWidget(self.category_search_input)
        
        layout.addLayout(toolbar_layout)
        
        # جدول التصنيفات
        self.categories_table = QTableWidget()
        self.categories_table.setColumnCount(3)
        self.categories_table.setHorizontalHeaderLabels([
            "الاسم", "الوصف", "تاريخ الإنشاء"
        ])
        self.categories_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.categories_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.categories_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.categories_table.setSortingEnabled(True)
        self.categories_table.doubleClicked.connect(self.edit_category)
        layout.addWidget(self.categories_table)
        
        self.categories_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_categories_table()
    
    def refresh_categories_table(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM categories ORDER BY name")
        categories = cursor.fetchall()
        conn.close()
        
        self.categories_table.setRowCount(len(categories))
        
        for row, category in enumerate(categories):
            # الاسم
            self.categories_table.setItem(row, 0, QTableWidgetItem(category[1]))
            
            # الوصف
            self.categories_table.setItem(row, 1, QTableWidgetItem(category[2] or ""))
            
            # تاريخ الإنشاء
            created_at = datetime.strptime(category[3], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
            self.categories_table.setItem(row, 2, QTableWidgetItem(created_at))
    
    def filter_categories(self):
        search_text = self.category_search_input.text().lower()
        
        for row in range(self.categories_table.rowCount()):
            match = False
            
            for col in range(self.categories_table.columnCount()):
                item = self.categories_table.item(row, col)
                if item and search_text in item.text().lower():
                    match = True
                    break
            
            self.categories_table.setRowHidden(row, not match)
    
    def add_category(self):
        dialog = CategoryDialog(self)
        if dialog.exec_():
            self.refresh_categories_table()
            add_log("إضافة تصنيف", f"تم إضافة التصنيف: {dialog.category_name}", self.user_id)
    
    def edit_category(self):
        selected_rows = self.categories_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار تصنيف للتعديل")
            return
        
        row = selected_rows[0].row()
        category_name = self.categories_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM categories WHERE name = ?", (category_name,))
        category = cursor.fetchone()
        conn.close()
        
        if category:
            dialog = CategoryDialog(self, category)
            if dialog.exec_():
                self.refresh_categories_table()
                add_log("تعديل تصنيف", f"تم تعديل التصنيف: {dialog.category_name}", self.user_id)
    
    def delete_category(self):
        selected_rows = self.categories_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار تصنيف للحذف")
            return
        
        row = selected_rows[0].row()
        category_name = self.categories_table.item(row, 0).text()
        
        reply = QMessageBox.question(
            self, "تأكيد الحذف", 
            f"هل أنت متأكد من حذف التصنيف '{category_name}'؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("DELETE FROM categories WHERE name = ?", (category_name,))
            conn.commit()
            conn.close()
            
            self.refresh_categories_table()
            add_log("حذف تصنيف", f"تم حذف التصنيف: {category_name}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم حذف التصنيف بنجاح")
    
    def create_reports_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر تحديث التقرير
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_sales_report)
        toolbar_layout.addWidget(refresh_button)
        
        # زر تصدير التقرير
        export_button = QPushButton("تصدير")
        export_button.clicked.connect(self.export_sales_report)
        toolbar_layout.addWidget(export_button)
        
        # زر طباعة التقرير
        print_button = QPushButton("طباعة")
        print_button.clicked.connect(self.print_sales_report)
        toolbar_layout.addWidget(print_button)
        
        layout.addLayout(toolbar_layout)
        
        # إعدادات التقرير
        settings_group = QGroupBox("إعدادات التقرير")
        settings_layout = QFormLayout()
        
        # نوع التقرير
        self.report_type_combo = QComboBox()
        self.report_type_combo.addItems(["يومي", "أسبوعي", "شهري", "سنوي"])
        self.report_type_combo.currentIndexChanged.connect(self.refresh_sales_report)
        settings_layout.addRow("نوع التقرير:", self.report_type_combo)
        
        # من تاريخ
        self.from_date_input = QDateEdit()
        self.from_date_input.setDate(QDate.currentDate().addMonths(-1))
        self.from_date_input.setCalendarPopup(True)
        self.from_date_input.dateChanged.connect(self.refresh_sales_report)
        settings_layout.addRow("من تاريخ:", self.from_date_input)
        
        # إلى تاريخ
        self.to_date_input = QDateEdit()
        self.to_date_input.setDate(QDate.currentDate())
        self.to_date_input.setCalendarPopup(True)
        self.to_date_input.dateChanged.connect(self.refresh_sales_report)
        settings_layout.addRow("إلى تاريخ:", self.to_date_input)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # ملخص التقرير
        summary_group = QGroupBox("ملخص التقرير")
        summary_layout = QGridLayout()
        
        # إجمالي المبيعات
        total_sales_label = QLabel("إجمالي المبيعات:")
        summary_layout.addWidget(total_sales_label, 0, 0)
        
        self.total_sales_value = QLabel(f"0.00 {self.currency_symbol}")
        total_sales_font = QFont()
        total_sales_font.setBold(True)
        self.total_sales_value.setFont(total_sales_font)
        summary_layout.addWidget(self.total_sales_value, 0, 1)
        
        # عدد الفواتير
        invoices_count_label = QLabel("عدد الفواتير:")
        summary_layout.addWidget(invoices_count_label, 1, 0)
        
        self.invoices_count_value = QLabel("0")
        summary_layout.addWidget(self.invoices_count_value, 1, 1)
        
        # متوسط قيمة الفاتورة
        avg_invoice_label = QLabel("متوسط قيمة الفاتورة:")
        summary_layout.addWidget(avg_invoice_label, 2, 0)
        
        self.avg_invoice_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.avg_invoice_value, 2, 1)
        
        # إجمالي الضرائب
        total_tax_label = QLabel("إجمالي الضرائب:")
        summary_layout.addWidget(total_tax_label, 0, 2)
        
        self.total_tax_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.total_tax_value, 0, 3)
        
        # إجمالي الخصومات
        total_discount_label = QLabel("إجمالي الخصومات:")
        summary_layout.addWidget(total_discount_label, 1, 2)
        
        self.total_discount_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.total_discount_value, 1, 3)
        
        # صافي المبيعات
        net_sales_label = QLabel("صافي المبيعات:")
        summary_layout.addWidget(net_sales_label, 2, 2)
        
        self.net_sales_value = QLabel(f"0.00 {self.currency_symbol}")
        net_sales_font = QFont()
        net_sales_font.setBold(True)
        self.net_sales_value.setFont(net_sales_font)
        summary_layout.addWidget(self.net_sales_value, 2, 3)
        
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        # الرسم البياني
        self.sales_figure = Figure(figsize=(10, 6), dpi=100)
        self.sales_canvas = FigureCanvas(self.sales_figure)
        layout.addWidget(self.sales_canvas)
        
        # جدول المبيعات
        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(5)
        self.sales_table.setHorizontalHeaderLabels([
            "التاريخ", "عدد الفواتير", "إجمالي المبيعات", "إجمالي الضرائب", "صافي المبيعات"
        ])
        self.sales_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.sales_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.sales_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.sales_table.setSortingEnabled(True)
        layout.addWidget(self.sales_table)
        
        self.reports_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_sales_report()
    
    def refresh_sales_report(self):
        report_type = self.report_type_combo.currentText()
        from_date = self.from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.to_date_input.date().toString("yyyy-MM-dd")
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        
        # الحصول على بيانات المبيعات
        if report_type == "يومي":
            cursor.execute("""
                SELECT 
                    DATE(date) as date,
                    COUNT(*) as invoices_count,
                    SUM(total) as total_sales,
                    SUM(tax) as total_tax,
                    SUM(discount) as total_discount
                FROM invoices
                WHERE DATE(date) BETWEEN ? AND ?
                GROUP BY DATE(date)
                ORDER BY DATE(date)
            """, (from_date, to_date))
        elif report_type == "أسبوعي":
            cursor.execute("""
                SELECT 
                    strftime('%Y-%W', date) as week,
                    COUNT(*) as invoices_count,
                    SUM(total) as total_sales,
                    SUM(tax) as total_tax,
                    SUM(discount) as total_discount
                FROM invoices
                WHERE DATE(date) BETWEEN ? AND ?
                GROUP BY strftime('%Y-%W', date)
                ORDER BY strftime('%Y-%W', date)
            """, (from_date, to_date))
        elif report_type == "شهري":
            cursor.execute("""
                SELECT 
                    strftime('%Y-%m', date) as month,
                    COUNT(*) as invoices_count,
                    SUM(total) as total_sales,
                    SUM(tax) as total_tax,
                    SUM(discount) as total_discount
                FROM invoices
                WHERE DATE(date) BETWEEN ? AND ?
                GROUP BY strftime('%Y-%m', date)
                ORDER BY strftime('%Y-%m', date)
            """, (from_date, to_date))
        else:  # سنوي
            cursor.execute("""
                SELECT 
                    strftime('%Y', date) as year,
                    COUNT(*) as invoices_count,
                    SUM(total) as total_sales,
                    SUM(tax) as total_tax,
                    SUM(discount) as total_discount
                FROM invoices
                WHERE DATE(date) BETWEEN ? AND ?
                GROUP BY strftime('%Y', date)
                ORDER BY strftime('%Y', date)
            """, (from_date, to_date))
        
        sales_data = cursor.fetchall()
        
        # الحصول على إجماليات
        cursor.execute("""
            SELECT 
                COUNT(*) as invoices_count,
                SUM(total) as total_sales,
                SUM(tax) as total_tax,
                SUM(discount) as total_discount
            FROM invoices
            WHERE DATE(date) BETWEEN ? AND ?
        """, (from_date, to_date))
        
        totals = cursor.fetchone()
        conn.close()
        
        # تحديث ملخص التقرير
        if totals and totals[1]:
            self.total_sales_value.setText(f"{totals[1]:.{self.decimal_places}f} {self.currency_symbol}")
            self.invoices_count_value.setText(str(totals[0]))
            self.avg_invoice_value.setText(f"{(totals[1] / totals[0]):.{self.decimal_places}f} {self.currency_symbol}" if totals[0] > 0 else f"0.00 {self.currency_symbol}")
            self.total_tax_value.setText(f"{totals[2]:.{self.decimal_places}f} {self.currency_symbol}" if totals[2] else f"0.00 {self.currency_symbol}")
            self.total_discount_value.setText(f"{totals[3]:.{self.decimal_places}f} {self.currency_symbol}" if totals[3] else f"0.00 {self.currency_symbol}")
            net_sales = totals[1] - (totals[3] if totals[3] else 0)
            self.net_sales_value.setText(f"{net_sales:.{self.decimal_places}f} {self.currency_symbol}")
        else:
            self.total_sales_value.setText(f"0.00 {self.currency_symbol}")
            self.invoices_count_value.setText("0")
            self.avg_invoice_value.setText(f"0.00 {self.currency_symbol}")
            self.total_tax_value.setText(f"0.00 {self.currency_symbol}")
            self.total_discount_value.setText(f"0.00 {self.currency_symbol}")
            self.net_sales_value.setText(f"0.00 {self.currency_symbol}")
        
        # تحديث الرسم البياني
        self.sales_figure.clear()
        ax = self.sales_figure.add_subplot(111)
        
        if sales_data:
            dates = [row[0] for row in sales_data]
            sales = [row[2] for row in sales_data]
            
            ax.bar(dates, sales, color='skyblue')
            ax.set_xlabel('التاريخ')
            ax.set_ylabel(f'المبيعات ({self.currency_symbol})')
            ax.set_title('تقرير المبيعات')
            
            # تدوير التواريخ
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
            
            # ضبط التباعد
            self.sales_figure.tight_layout()
        else:
            ax.text(0.5, 0.5, 'لا توجد بيانات للعرض', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
        
        self.sales_canvas.draw()
        
        # تحديث جدول المبيعات
        self.sales_table.setRowCount(len(sales_data))
        
        for row, data in enumerate(sales_data):
            # التاريخ
            self.sales_table.setItem(row, 0, QTableWidgetItem(data[0]))
            
            # عدد الفواتير
            self.sales_table.setItem(row, 1, QTableWidgetItem(str(data[1])))
            
            # إجمالي المبيعات
            self.sales_table.setItem(row, 2, QTableWidgetItem(f"{data[2]:.{self.decimal_places}f} {self.currency_symbol}"))
            
            # إجمالي الضرائب
            self.sales_table.setItem(row, 3, QTableWidgetItem(f"{data[3]:.{self.decimal_places}f} {self.currency_symbol}" if data[3] else f"0.00 {self.currency_symbol}"))
            
            # صافي المبيعات
            net_sales = data[2] - (data[4] if data[4] else 0)
            self.sales_table.setItem(row, 4, QTableWidgetItem(f"{net_sales:.{self.decimal_places}f} {self.currency_symbol}"))
    
    def export_sales_report(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "تصدير تقرير المبيعات", "", "ملفات Excel (*.xlsx);;ملفات PDF (*.pdf)"
        )
        
        if not file_path:
            return
        
        report_type = self.report_type_combo.currentText()
        from_date = self.from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.to_date_input.date().toString("yyyy-MM-dd")
        
        try:
            if file_path.endswith('.xlsx'):
                # إنشاء ملف Excel
                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet('تقرير المبيعات')
                
                # إضافة التنسيق
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
                
                # كتابة العناوين
                worksheet.write(0, 0, 'تقرير المبيعات', header_format)
                worksheet.write(1, 0, f'نوع التقرير: {report_type}')
                worksheet.write(2, 0, f'من تاريخ: {from_date}')
                worksheet.write(3, 0, f'إلى تاريخ: {to_date}')
                
                # كتابة ملخص التقرير
                worksheet.write(5, 0, 'ملخص التقرير', header_format)
                worksheet.write(6, 0, 'إجمالي المبيعات')
                worksheet.write(6, 1, self.total_sales_value.text())
                worksheet.write(7, 0, 'عدد الفواتير')
                worksheet.write(7, 1, self.invoices_count_value.text())
                worksheet.write(8, 0, 'متوسط قيمة الفاتورة')
                worksheet.write(8, 1, self.avg_invoice_value.text())
                worksheet.write(9, 0, 'إجمالي الضرائب')
                worksheet.write(9, 1, self.total_tax_value.text())
                worksheet.write(10, 0, 'إجمالي الخصومات')
                worksheet.write(10, 1, self.total_discount_value.text())
                worksheet.write(11, 0, 'صافي المبيعات')
                worksheet.write(11, 1, self.net_sales_value.text())
                
                # كتابة جدول المبيعات
                worksheet.write(13, 0, 'تفاصيل المبيعات', header_format)
                
                # عناوين الأعمدة
                headers = ['التاريخ', 'عدد الفواتير', 'إجمالي المبيعات', 'إجمالي الضرائب', 'صافي المبيعات']
                for col, header in enumerate(headers):
                    worksheet.write(14, col, header, header_format)
                
                # بيانات الجدول
                for row in range(self.sales_table.rowCount()):
                    for col in range(self.sales_table.columnCount()):
                        item = self.sales_table.item(row, col)
                        if item:
                            worksheet.write(row + 15, col, item.text())
                
                # ضبط عرض الأعمدة
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 4, 15)
                
                workbook.close()
            else:
                # إنشاء ملف PDF
                doc = SimpleDocTemplate(
                    file_path,
                    pagesize=A4,
                    rightMargin=72,
                    leftMargin=72,
                    topMargin=72,
                    bottomMargin=18
                )
                
                # إنشاء قائمة العناصر
                story = []
                
                # إضافة عنوان التقرير
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=18,
                    spaceAfter=30,
                    alignment=TA_CENTER
                )
                story.append(Paragraph("تقرير المبيعات", title_style))
                
                # إضافة معلومات التقرير
                report_info = []
                report_info.append(Paragraph(f"<b>نوع التقرير:</b> {report_type}", styles['Normal']))
                report_info.append(Paragraph(f"<b>من تاريخ:</b> {from_date}", styles['Normal']))
                report_info.append(Paragraph(f"<b>إلى تاريخ:</b> {to_date}", styles['Normal']))
                
                report_table = Table([report_info], colWidths=[3*inch])
                report_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                story.append(report_table)
                story.append(Spacer(1, 12))
                
                # إضافة ملخص التقرير
                summary_data = [
                    [Paragraph("<b>إجمالي المبيعات:</b>", styles['Normal']), Paragraph(self.total_sales_value.text(), styles['Normal'])],
                    [Paragraph("<b>عدد الفواتير:</b>", styles['Normal']), Paragraph(self.invoices_count_value.text(), styles['Normal'])],
                    [Paragraph("<b>متوسط قيمة الفاتورة:</b>", styles['Normal']), Paragraph(self.avg_invoice_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي الضرائب:</b>", styles['Normal']), Paragraph(self.total_tax_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي الخصومات:</b>", styles['Normal']), Paragraph(self.total_discount_value.text(), styles['Normal'])],
                    [Paragraph("<b>صافي المبيعات:</b>", styles['Normal']), Paragraph(self.net_sales_value.text(), styles['Normal'])]
                ]
                
                summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
                summary_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 12))
                
                # إضافة جدول المبيعات
                sales_data = [['التاريخ', 'عدد الفواتير', 'إجمالي المبيعات', 'إجمالي الضرائب', 'صافي المبيعات']]
                
                for row in range(self.sales_table.rowCount()):
                    row_data = []
                    for col in range(self.sales_table.columnCount()):
                        item = self.sales_table.item(row, col)
                        row_data.append(item.text() if item else "")
                    sales_data.append(row_data)
                
                sales_table = Table(sales_data, colWidths=[1*inch, 1*inch, 1*inch, 1*inch, 1*inch])
                sales_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ]))
                story.append(sales_table)
                
                # بناء ملف PDF
                doc.build(story)
            
            add_log("تصدير تقرير", "تم تصدير تقرير المبيعات", self.user_id)
            QMessageBox.information(self, "نجاح", "تم تصدير التقرير بنجاح")
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير التقرير: {str(e)}")
    
    def print_sales_report(self):
        # إنشاء ملف PDF مؤقت
        temp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        temp_file.close()
        
        # تصدير التقرير إلى الملف المؤقت
        file_path = temp_file.name
        
        report_type = self.report_type_combo.currentText()
        from_date = self.from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.to_date_input.date().toString("yyyy-MM-dd")
        
        try:
            # إنشاء ملف PDF
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # إنشاء قائمة العناصر
            story = []
            
            # إضافة عنوان التقرير
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph("تقرير المبيعات", title_style))
            
            # إضافة معلومات التقرير
            report_info = []
            report_info.append(Paragraph(f"<b>نوع التقرير:</b> {report_type}", styles['Normal']))
            report_info.append(Paragraph(f"<b>من تاريخ:</b> {from_date}", styles['Normal']))
            report_info.append(Paragraph(f"<b>إلى تاريخ:</b> {to_date}", styles['Normal']))
            
            report_table = Table([report_info], colWidths=[3*inch])
            report_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(report_table)
            story.append(Spacer(1, 12))
            
            # إضافة ملخص التقرير
            summary_data = [
                [Paragraph("<b>إجمالي المبيعات:</b>", styles['Normal']), Paragraph(self.total_sales_value.text(), styles['Normal'])],
                [Paragraph("<b>عدد الفواتير:</b>", styles['Normal']), Paragraph(self.invoices_count_value.text(), styles['Normal'])],
                [Paragraph("<b>متوسط قيمة الفاتورة:</b>", styles['Normal']), Paragraph(self.avg_invoice_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي الضرائب:</b>", styles['Normal']), Paragraph(self.total_tax_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي الخصومات:</b>", styles['Normal']), Paragraph(self.total_discount_value.text(), styles['Normal'])],
                [Paragraph("<b>صافي المبيعات:</b>", styles['Normal']), Paragraph(self.net_sales_value.text(), styles['Normal'])]
            ]
            
            summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # إضافة جدول المبيعات
            sales_data = [['التاريخ', 'عدد الفواتير', 'إجمالي المبيعات', 'إجمالي الضرائب', 'صافي المبيعات']]
            
            for row in range(self.sales_table.rowCount()):
                row_data = []
                for col in range(self.sales_table.columnCount()):
                    item = self.sales_table.item(row, col)
                    row_data.append(item.text() if item else "")
                sales_data.append(row_data)
            
            sales_table = Table(sales_data, colWidths=[1*inch, 1*inch, 1*inch, 1*inch, 1*inch])
            sales_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ]))
            story.append(sales_table)
            
            # بناء ملف PDF
            doc.build(story)
            
            # فتح الملف
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            
            add_log("طباعة تقرير", "تم طباعة تقرير المبيعات", self.user_id)
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء طباعة التقرير: {str(e)}")
        
        finally:
            # حذف الملف المؤقت بعد فتحه
            try:
                os.unlink(file_path)
            except:
                pass
    
    def create_inventory_report_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر تحديث التقرير
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_inventory_report)
        toolbar_layout.addWidget(refresh_button)
        
        # زر تصدير التقرير
        export_button = QPushButton("تصدير")
        export_button.clicked.connect(self.export_inventory_report)
        toolbar_layout.addWidget(export_button)
        
        # زر طباعة التقرير
        print_button = QPushButton("طباعة")
        print_button.clicked.connect(self.print_inventory_report)
        toolbar_layout.addWidget(print_button)
        
        layout.addLayout(toolbar_layout)
        
        # ملخص التقرير
        summary_group = QGroupBox("ملخص التقرير")
        summary_layout = QGridLayout()
        
        # عدد المنتجات
        products_count_label = QLabel("عدد المنتجات:")
        summary_layout.addWidget(products_count_label, 0, 0)
        
        self.products_count_value = QLabel("0")
        summary_layout.addWidget(self.products_count_value, 0, 1)
        
        # إجمالي قيمة المخزون
        inventory_value_label = QLabel("إجمالي قيمة المخزون:")
        summary_layout.addWidget(inventory_value_label, 1, 0)
        
        self.inventory_value_value = QLabel(f"0.00 {self.currency_symbol}")
        inventory_value_font = QFont()
        inventory_value_font.setBold(True)
        self.inventory_value_value.setFont(inventory_value_font)
        summary_layout.addWidget(self.inventory_value_value, 1, 1)
        
        # عدد المنتجات منخفضة المخزون
        low_stock_count_label = QLabel("عدد المنتجات منخفضة المخزون:")
        summary_layout.addWidget(low_stock_count_label, 2, 0)
        
        self.low_stock_count_value = QLabel("0")
        low_stock_font = QFont()
        low_stock_font.setBold(True)
        self.low_stock_count_value.setFont(low_stock_font)
        self.low_stock_count_value.setStyleSheet("color: red;")
        summary_layout.addWidget(self.low_stock_count_value, 2, 1)
        
        # عدد المنتجات منتهية الصلاحية
        expired_count_label = QLabel("عدد المنتجات منتهية الصلاحية:")
        summary_layout.addWidget(expired_count_label, 0, 2)
        
        self.expired_count_value = QLabel("0")
        expired_font = QFont()
        expired_font.setBold(True)
        self.expired_count_value.setFont(expired_font)
        self.expired_count_value.setStyleSheet("color: red;")
        summary_layout.addWidget(self.expired_count_value, 0, 3)
        
        # إجمالي قيمة المخزون منخفض
        low_stock_value_label = QLabel("إجمالي قيمة المخزون منخفض:")
        summary_layout.addWidget(low_stock_value_label, 1, 2)
        
        self.low_stock_value_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.low_stock_value_value, 1, 3)
        
        # إجمالي قيمة المخزون منتهي
        expired_value_label = QLabel("إجمالي قيمة المخزون منتهي:")
        summary_layout.addWidget(expired_value_label, 2, 2)
        
        self.expired_value_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.expired_value_value, 2, 3)
        
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        # الرسم البياني
        self.inventory_figure = Figure(figsize=(10, 6), dpi=100)
        self.inventory_canvas = FigureCanvas(self.inventory_figure)
        layout.addWidget(self.inventory_canvas)
        
        # جدول المخزون
        self.inventory_table = QTableWidget()
        self.inventory_table.setColumnCount(7)
        self.inventory_table.setHorizontalHeaderLabels([
            "الباركود", "اسم المنتج", "التصنيف", "الكمية", "سعر الشراء", "قيمة المخزون", "الحالة"
        ])
        self.inventory_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.inventory_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.inventory_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.inventory_table.setSortingEnabled(True)
        layout.addWidget(self.inventory_table)
        
        self.inventory_report_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_inventory_report()
    
    def refresh_inventory_report(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        
        # الحصول على بيانات المخزون
        cursor.execute("""
            SELECT 
                barcode,
                name,
                category,
                quantity,
                cost_price,
                expiry_date,
                min_quantity
            FROM products
            ORDER BY name
        """)
        
        products_data = cursor.fetchall()
        
        # حساب الإجماليات
        products_count = len(products_data)
        inventory_value = 0
        low_stock_count = 0
        low_stock_value = 0
        expired_count = 0
        expired_value = 0
        
        today = datetime.now().strftime("%Y-%m-%d")
        
        for product in products_data:
            barcode, name, category, quantity, cost_price, expiry_date, min_quantity = product
            
            # حساب قيمة المخزون
            if cost_price and quantity:
                product_value = cost_price * quantity
                inventory_value += product_value
                
                # التحقق من المخزون المنخفض
                if quantity <= min_quantity:
                    low_stock_count += 1
                    low_stock_value += product_value
                
                # التحقق من المنتجات منتهية الصلاحية
                if expiry_date and expiry_date < today:
                    expired_count += 1
                    expired_value += product_value
        
        conn.close()
        
        # تحديث ملخص التقرير
        self.products_count_value.setText(str(products_count))
        self.inventory_value_value.setText(f"{inventory_value:.{self.decimal_places}f} {self.currency_symbol}")
        self.low_stock_count_value.setText(str(low_stock_count))
        self.low_stock_value_value.setText(f"{low_stock_value:.{self.decimal_places}f} {self.currency_symbol}")
        self.expired_count_value.setText(str(expired_count))
        self.expired_value_value.setText(f"{expired_value:.{self.decimal_places}f} {self.currency_symbol}")
        
        # تحديث الرسم البياني
        self.inventory_figure.clear()
        ax = self.inventory_figure.add_subplot(111)
        
        # بيانات الرسم البياني
        categories = ['المنتجات', 'منخفض المخزون', 'منتهية الصلاحية']
        values = [products_count, low_stock_count, expired_count]
        colors = ['skyblue', 'orange', 'red']
        
        ax.bar(categories, values, color=colors)
        ax.set_xlabel('النوع')
        ax.set_ylabel('العدد')
        ax.set_title('تقرير المخزون')
        
        # إضافة القيم فوق الأعمدة
        for i, v in enumerate(values):
            ax.text(i, v + 0.5, str(v), ha='center')
        
        self.inventory_canvas.draw()
        
        # تحديث جدول المخزون
        self.inventory_table.setRowCount(len(products_data))
        
        for row, product in enumerate(products_data):
            barcode, name, category, quantity, cost_price, expiry_date, min_quantity = product
            
            # الباركود
            self.inventory_table.setItem(row, 0, QTableWidgetItem(barcode or ""))
            
            # اسم المنتج
            self.inventory_table.setItem(row, 1, QTableWidgetItem(name))
            
            # التصنيف
            self.inventory_table.setItem(row, 2, QTableWidgetItem(category or ""))
            
            # الكمية
            quantity_item = QTableWidgetItem(str(quantity))
            if quantity <= min_quantity:
                quantity_item.setBackground(QColor(255, 200, 200))
            self.inventory_table.setItem(row, 3, quantity_item)
            
            # سعر الشراء
            cost_price_item = QTableWidgetItem(f"{cost_price:.{self.decimal_places}f} {self.currency_symbol}" if cost_price else "")
            self.inventory_table.setItem(row, 4, cost_price_item)
            
            # قيمة المخزون
            if cost_price and quantity:
                inventory_value_item = QTableWidgetItem(f"{(cost_price * quantity):.{self.decimal_places}f} {self.currency_symbol}")
            else:
                inventory_value_item = QTableWidgetItem("")
            self.inventory_table.setItem(row, 5, inventory_value_item)
            
            # الحالة
            status = ""
            status_item = QTableWidgetItem(status)
            
            if quantity <= min_quantity:
                status = "منخفض المخزون"
                status_item.setBackground(QColor(255, 200, 200))
            
            if expiry_date and expiry_date < today:
                if status:
                    status += "، منتهي الصلاحية"
                else:
                    status = "منتهي الصلاحية"
                status_item.setBackground(QColor(255, 200, 200))
            
            if not status:
                status = "جيد"
                status_item.setBackground(QColor(200, 255, 200))
            
            self.inventory_table.setItem(row, 6, QTableWidgetItem(status))
    
    def export_inventory_report(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "تصدير تقرير المخزون", "", "ملفات Excel (*.xlsx);;ملفات PDF (*.pdf)"
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.xlsx'):
                # إنشاء ملف Excel
                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet('تقرير المخزون')
                
                # إضافة التنسيق
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # كتابة العناوين
                worksheet.write(0, 0, 'تقرير المخزون', header_format)
                
                # كتابة ملخص التقرير
                worksheet.write(1, 0, 'ملخص التقرير', header_format)
                worksheet.write(2, 0, 'عدد المنتجات')
                worksheet.write(2, 1, self.products_count_value.text())
                worksheet.write(3, 0, 'إجمالي قيمة المخزون')
                worksheet.write(3, 1, self.inventory_value_value.text())
                worksheet.write(4, 0, 'عدد المنتجات منخفضة المخزون')
                worksheet.write(4, 1, self.low_stock_count_value.text())
                worksheet.write(5, 0, 'إجمالي قيمة المخزون منخفض')
                worksheet.write(5, 1, self.low_stock_value_value.text())
                worksheet.write(6, 0, 'عدد المنتجات منتهية الصلاحية')
                worksheet.write(6, 1, self.expired_count_value.text())
                worksheet.write(7, 0, 'إجمالي قيمة المخزون منتهي')
                worksheet.write(7, 1, self.expired_value_value.text())
                
                # كتابة جدول المخزون
                worksheet.write(9, 0, 'تفاصيل المخزون', header_format)
                
                # عناوين الأعمدة
                headers = ['الباركود', 'اسم المنتج', 'التصنيف', 'الكمية', 'سعر الشراء', 'قيمة المخزون', 'الحالة']
                for col, header in enumerate(headers):
                    worksheet.write(10, col, header, header_format)
                
                # بيانات الجدول
                for row in range(self.inventory_table.rowCount()):
                    for col in range(self.inventory_table.columnCount()):
                        item = self.inventory_table.item(row, col)
                        if item:
                            worksheet.write(row + 11, col, item.text())
                
                # ضبط عرض الأعمدة
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 1, 20)
                worksheet.set_column(2, 2, 15)
                worksheet.set_column(3, 3, 10)
                worksheet.set_column(4, 4, 15)
                worksheet.set_column(5, 5, 15)
                worksheet.set_column(6, 6, 15)
                
                workbook.close()
            else:
                # إنشاء ملف PDF
                doc = SimpleDocTemplate(
                    file_path,
                    pagesize=A4,
                    rightMargin=72,
                    leftMargin=72,
                    topMargin=72,
                    bottomMargin=18
                )
                
                # إنشاء قائمة العناصر
                story = []
                
                # إضافة عنوان التقرير
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=18,
                    spaceAfter=30,
                    alignment=TA_CENTER
                )
                story.append(Paragraph("تقرير المخزون", title_style))
                
                # إضافة ملخص التقرير
                summary_data = [
                    [Paragraph("<b>عدد المنتجات:</b>", styles['Normal']), Paragraph(self.products_count_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي قيمة المخزون:</b>", styles['Normal']), Paragraph(self.inventory_value_value.text(), styles['Normal'])],
                    [Paragraph("<b>عدد المنتجات منخفضة المخزون:</b>", styles['Normal']), Paragraph(self.low_stock_count_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي قيمة المخزون منخفض:</b>", styles['Normal']), Paragraph(self.low_stock_value_value.text(), styles['Normal'])],
                    [Paragraph("<b>عدد المنتجات منتهية الصلاحية:</b>", styles['Normal']), Paragraph(self.expired_count_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي قيمة المخزون منتهي:</b>", styles['Normal']), Paragraph(self.expired_value_value.text(), styles['Normal'])]
                ]
                
                summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
                summary_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 12))
                
                # إضافة جدول المخزون
                inventory_data = [['الباركود', 'اسم المنتج', 'التصنيف', 'الكمية', 'سعر الشراء', 'قيمة المخزون', 'الحالة']]
                
                for row in range(self.inventory_table.rowCount()):
                    row_data = []
                    for col in range(self.inventory_table.columnCount()):
                        item = self.inventory_table.item(row, col)
                        row_data.append(item.text() if item else "")
                    inventory_data.append(row_data)
                
                inventory_table = Table(inventory_data, colWidths=[1*inch, 2*inch, 1*inch, 1*inch, 1*inch, 1*inch, 1*inch])
                inventory_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ]))
                story.append(inventory_table)
                
                # بناء ملف PDF
                doc.build(story)
            
            add_log("تصدير تقرير", "تم تصدير تقرير المخزون", self.user_id)
            QMessageBox.information(self, "نجاح", "تم تصدير التقرير بنجاح")
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير التقرير: {str(e)}")
    
    def print_inventory_report(self):
        # إنشاء ملف PDF مؤقت
        temp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        temp_file.close()
        
        # تصدير التقرير إلى الملف المؤقت
        file_path = temp_file.name
        
        try:
            # إنشاء ملف PDF
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # إنشاء قائمة العناصر
            story = []
            
            # إضافة عنوان التقرير
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph("تقرير المخزون", title_style))
            
            # إضافة ملخص التقرير
            summary_data = [
                [Paragraph("<b>عدد المنتجات:</b>", styles['Normal']), Paragraph(self.products_count_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي قيمة المخزون:</b>", styles['Normal']), Paragraph(self.inventory_value_value.text(), styles['Normal'])],
                [Paragraph("<b>عدد المنتجات منخفضة المخزون:</b>", styles['Normal']), Paragraph(self.low_stock_count_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي قيمة المخزون منخفض:</b>", styles['Normal']), Paragraph(self.low_stock_value_value.text(), styles['Normal'])],
                [Paragraph("<b>عدد المنتجات منتهية الصلاحية:</b>", styles['Normal']), Paragraph(self.expired_count_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي قيمة المخزون منتهي:</b>", styles['Normal']), Paragraph(self.expired_value_value.text(), styles['Normal'])]
            ]
            
            summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # إضافة جدول المخزون
            inventory_data = [['الباركود', 'اسم المنتج', 'التصنيف', 'الكمية', 'سعر الشراء', 'قيمة المخزون', 'الحالة']]
            
            for row in range(self.inventory_table.rowCount()):
                row_data = []
                for col in range(self.inventory_table.columnCount()):
                    item = self.inventory_table.item(row, col)
                    row_data.append(item.text() if item else "")
                inventory_data.append(row_data)
            
            inventory_table = Table(inventory_data, colWidths=[1*inch, 2*inch, 1*inch, 1*inch, 1*inch, 1*inch, 1*inch])
            inventory_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ]))
            story.append(inventory_table)
            
            # بناء ملف PDF
            doc.build(story)
            
            # فتح الملف
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            
            add_log("طباعة تقرير", "تم طباعة تقرير المخزون", self.user_id)
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء طباعة التقرير: {str(e)}")
        
        finally:
            # حذف الملف المؤقت بعد فتحه
            try:
                os.unlink(file_path)
            except:
                pass
    
    def create_profit_report_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر تحديث التقرير
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_profit_report)
        toolbar_layout.addWidget(refresh_button)
        
        # زر تصدير التقرير
        export_button = QPushButton("تصدير")
        export_button.clicked.connect(self.export_profit_report)
        toolbar_layout.addWidget(export_button)
        
        # زر طباعة التقرير
        print_button = QPushButton("طباعة")
        print_button.clicked.connect(self.print_profit_report)
        toolbar_layout.addWidget(print_button)
        
        layout.addLayout(toolbar_layout)
        
        # إعدادات التقرير
        settings_group = QGroupBox("إعدادات التقرير")
        settings_layout = QFormLayout()
        
        # من تاريخ
        self.profit_from_date_input = QDateEdit()
        self.profit_from_date_input.setDate(QDate.currentDate().addMonths(-1))
        self.profit_from_date_input.setCalendarPopup(True)
        self.profit_from_date_input.dateChanged.connect(self.refresh_profit_report)
        settings_layout.addRow("من تاريخ:", self.profit_from_date_input)
        
        # إلى تاريخ
        self.profit_to_date_input = QDateEdit()
        self.profit_to_date_input.setDate(QDate.currentDate())
        self.profit_to_date_input.setCalendarPopup(True)
        self.profit_to_date_input.dateChanged.connect(self.refresh_profit_report)
        settings_layout.addRow("إلى تاريخ:", self.profit_to_date_input)
        
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)
        
        # ملخص التقرير
        summary_group = QGroupBox("ملخص التقرير")
        summary_layout = QGridLayout()
        
        # إجمالي المبيعات
        total_sales_label = QLabel("إجمالي المبيعات:")
        summary_layout.addWidget(total_sales_label, 0, 0)
        
        self.profit_total_sales_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.profit_total_sales_value, 0, 1)
        
        # إجمالي التكلفة
        total_cost_label = QLabel("إجمالي التكلفة:")
        summary_layout.addWidget(total_cost_label, 1, 0)
        
        self.profit_total_cost_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.profit_total_cost_value, 1, 1)
        
        # إجمالي الربح
        total_profit_label = QLabel("إجمالي الربح:")
        summary_layout.addWidget(total_profit_label, 2, 0)
        
        self.profit_total_profit_value = QLabel(f"0.00 {self.currency_symbol}")
        profit_font = QFont()
        profit_font.setBold(True)
        self.profit_total_profit_value.setFont(profit_font)
        summary_layout.addWidget(self.profit_total_profit_value, 2, 1)
        
        # نسبة الربح
        profit_margin_label = QLabel("نسبة الربح:")
        summary_layout.addWidget(profit_margin_label, 0, 2)
        
        self.profit_margin_value = QLabel("0.00%")
        summary_layout.addWidget(self.profit_margin_value, 0, 3)
        
        # إجمالي الضرائب
        total_tax_label = QLabel("إجمالي الضرائب:")
        summary_layout.addWidget(total_tax_label, 1, 2)
        
        self.profit_total_tax_value = QLabel(f"0.00 {self.currency_symbol}")
        summary_layout.addWidget(self.profit_total_tax_value, 1, 3)
        
        # صافي الربح
        net_profit_label = QLabel("صافي الربح:")
        summary_layout.addWidget(net_profit_label, 2, 2)
        
        self.profit_net_profit_value = QLabel(f"0.00 {self.currency_symbol}")
        net_profit_font = QFont()
        net_profit_font.setBold(True)
        self.profit_net_profit_value.setFont(net_profit_font)
        summary_layout.addWidget(self.profit_net_profit_value, 2, 3)
        
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        # الرسم البياني
        self.profit_figure = Figure(figsize=(10, 6), dpi=100)
        self.profit_canvas = FigureCanvas(self.profit_figure)
        layout.addWidget(self.profit_canvas)
        
        # جدول الأرباح
        self.profit_table = QTableWidget()
        self.profit_table.setColumnCount(6)
        self.profit_table.setHorizontalHeaderLabels([
            "التاريخ", "عدد الفواتير", "إجمالي المبيعات", "إجمالي التكلفة", "إجمالي الربح", "نسبة الربح"
        ])
        self.profit_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.profit_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.profit_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.profit_table.setSortingEnabled(True)
        layout.addWidget(self.profit_table)
        
        self.profit_report_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_profit_report()
    
    def refresh_profit_report(self):
        from_date = self.profit_from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.profit_to_date_input.date().toString("yyyy-MM-dd")
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        
        # الحصول على بيانات المبيعات
        cursor.execute("""
            SELECT 
                DATE(date) as date,
                COUNT(*) as invoices_count,
                SUM(total) as total_sales,
                SUM(tax) as total_tax
            FROM invoices
            WHERE DATE(date) BETWEEN ? AND ?
            GROUP BY DATE(date)
            ORDER BY DATE(date)
        """, (from_date, to_date))
        
        sales_data = cursor.fetchall()
        
        # الحصول على بيانات التكلفة
        profit_data = []
        
        for sale in sales_data:
            date, invoices_count, total_sales, total_tax = sale
            
            # الحصول على تكلفة المنتجات المباعة في هذا اليوم
            cursor.execute("""
                SELECT SUM(ii.quantity * p.cost_price) as total_cost
                FROM invoice_items ii
                JOIN products p ON ii.product_id = p.id
                JOIN invoices i ON ii.invoice_id = i.id
                WHERE DATE(i.date) = ?
            """, (date,))
            
            cost_result = cursor.fetchone()
            total_cost = cost_result[0] if cost_result and cost_result[0] else 0
            
            # حساب الربح
            total_profit = total_sales - total_cost
            profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
            
            profit_data.append({
                'date': date,
                'invoices_count': invoices_count,
                'total_sales': total_sales,
                'total_cost': total_cost,
                'total_profit': total_profit,
                'profit_margin': profit_margin,
                'total_tax': total_tax
            })
        
        # حساب الإجماليات
        total_sales = sum(item['total_sales'] for item in profit_data)
        total_cost = sum(item['total_cost'] for item in profit_data)
        total_profit = sum(item['total_profit'] for item in profit_data)
        total_tax = sum(item['total_tax'] for item in profit_data)
        net_profit = total_profit - total_tax
        profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
        
        conn.close()
        
        # تحديث ملخص التقرير
        self.profit_total_sales_value.setText(f"{total_sales:.{self.decimal_places}f} {self.currency_symbol}")
        self.profit_total_cost_value.setText(f"{total_cost:.{self.decimal_places}f} {self.currency_symbol}")
        self.profit_total_profit_value.setText(f"{total_profit:.{self.decimal_places}f} {self.currency_symbol}")
        self.profit_margin_value.setText(f"{profit_margin:.{self.decimal_places}f}%")
        self.profit_total_tax_value.setText(f"{total_tax:.{self.decimal_places}f} {self.currency_symbol}")
        self.profit_net_profit_value.setText(f"{net_profit:.{self.decimal_places}f} {self.currency_symbol}")
        
        # تحديث الرسم البياني
        self.profit_figure.clear()
        ax = self.profit_figure.add_subplot(111)
        
        if profit_data:
            dates = [item['date'] for item in profit_data]
            profits = [item['total_profit'] for item in profit_data]
            
            ax.bar(dates, profits, color='green')
            ax.set_xlabel('التاريخ')
            ax.set_ylabel(f'الربح ({self.currency_symbol})')
            ax.set_title('تقرير الأرباح')
            
            # تدوير التواريخ
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right')
            
            # ضبط التباعد
            self.profit_figure.tight_layout()
        else:
            ax.text(0.5, 0.5, 'لا توجد بيانات للعرض', horizontalalignment='center', verticalalignment='center', transform=ax.transAxes)
        
        self.profit_canvas.draw()
        
        # تحديث جدول الأرباح
        self.profit_table.setRowCount(len(profit_data))
        
        for row, data in enumerate(profit_data):
            # التاريخ
            self.profit_table.setItem(row, 0, QTableWidgetItem(data['date']))
            
            # عدد الفواتير
            self.profit_table.setItem(row, 1, QTableWidgetItem(str(data['invoices_count'])))
            
            # إجمالي المبيعات
            self.profit_table.setItem(row, 2, QTableWidgetItem(f"{data['total_sales']:.{self.decimal_places}f} {self.currency_symbol}"))
            
            # إجمالي التكلفة
            self.profit_table.setItem(row, 3, QTableWidgetItem(f"{data['total_cost']:.{self.decimal_places}f} {self.currency_symbol}"))
            
            # إجمالي الربح
            profit_item = QTableWidgetItem(f"{data['total_profit']:.{self.decimal_places}f} {self.currency_symbol}")
            if data['total_profit'] < 0:
                profit_item.setBackground(QColor(255, 200, 200))
            self.profit_table.setItem(row, 4, profit_item)
            
            # نسبة الربح
            margin_item = QTableWidgetItem(f"{data['profit_margin']:.{self.decimal_places}f}%")
            if data['profit_margin'] < 0:
                margin_item.setBackground(QColor(255, 200, 200))
            self.profit_table.setItem(row, 5, margin_item)
    
    def export_profit_report(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "تصدير تقرير الأرباح", "", "ملفات Excel (*.xlsx);;ملفات PDF (*.pdf)"
        )
        
        if not file_path:
            return
        
        from_date = self.profit_from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.profit_to_date_input.date().toString("yyyy-MM-dd")
        
        try:
            if file_path.endswith('.xlsx'):
                # إنشاء ملف Excel
                workbook = xlsxwriter.Workbook(file_path)
                worksheet = workbook.add_worksheet('تقرير الأرباح')
                
                # إضافة التنسيق
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                profit_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#C6EFCE',
                    'border': 1
                })
                
                loss_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#FFC7CE',
                    'border': 1
                })
                
                # كتابة العناوين
                worksheet.write(0, 0, 'تقرير الأرباح', header_format)
                worksheet.write(1, 0, f'من تاريخ: {from_date}')
                worksheet.write(2, 0, f'إلى تاريخ: {to_date}')
                
                # كتابة ملخص التقرير
                worksheet.write(4, 0, 'ملخص التقرير', header_format)
                worksheet.write(5, 0, 'إجمالي المبيعات')
                worksheet.write(5, 1, self.profit_total_sales_value.text())
                worksheet.write(6, 0, 'إجمالي التكلفة')
                worksheet.write(6, 1, self.profit_total_cost_value.text())
                worksheet.write(7, 0, 'إجمالي الربح')
                worksheet.write(7, 1, self.profit_total_profit_value.text())
                worksheet.write(8, 0, 'نسبة الربح')
                worksheet.write(8, 1, self.profit_margin_value.text())
                worksheet.write(9, 0, 'إجمالي الضرائب')
                worksheet.write(9, 1, self.profit_total_tax_value.text())
                worksheet.write(10, 0, 'صافي الربح')
                worksheet.write(10, 1, self.profit_net_profit_value.text())
                
                # كتابة جدول الأرباح
                worksheet.write(12, 0, 'تفاصيل الأرباح', header_format)
                
                # عناوين الأعمدة
                headers = ['التاريخ', 'عدد الفواتير', 'إجمالي المبيعات', 'إجمالي التكلفة', 'إجمالي الربح', 'نسبة الربح']
                for col, header in enumerate(headers):
                    worksheet.write(13, col, header, header_format)
                
                # بيانات الجدول
                for row in range(self.profit_table.rowCount()):
                    for col in range(self.profit_table.columnCount()):
                        item = self.profit_table.item(row, col)
                        if item:
                            # تنسيق خلايا الربح والخسارة
                            if col == 4 or col == 5:  # عمود الربح ونسبة الربح
                                value_text = item.text()
                                if self.currency_symbol in value_text:
                                    value = float(value_text.replace(self.currency_symbol, '').strip())
                                elif '%' in value_text:
                                    value = float(value_text.replace('%', '').strip())
                                else:
                                    value = 0
                                
                                if value < 0:
                                    worksheet.write(row + 14, col, value_text, loss_format)
                                else:
                                    worksheet.write(row + 14, col, value_text, profit_format)
                            else:
                                worksheet.write(row + 14, col, item.text())
                
                # ضبط عرض الأعمدة
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 1, 15)
                worksheet.set_column(2, 5, 15)
                
                workbook.close()
            else:
                # إنشاء ملف PDF
                doc = SimpleDocTemplate(
                    file_path,
                    pagesize=A4,
                    rightMargin=72,
                    leftMargin=72,
                    topMargin=72,
                    bottomMargin=18
                )
                
                # إنشاء قائمة العناصر
                story = []
                
                # إضافة عنوان التقرير
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=18,
                    spaceAfter=30,
                    alignment=TA_CENTER
                )
                story.append(Paragraph("تقرير الأرباح", title_style))
                
                # إضافة معلومات التقرير
                report_info = []
                report_info.append(Paragraph(f"<b>من تاريخ:</b> {from_date}", styles['Normal']))
                report_info.append(Paragraph(f"<b>إلى تاريخ:</b> {to_date}", styles['Normal']))
                
                report_table = Table([report_info], colWidths=[3*inch])
                report_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                story.append(report_table)
                story.append(Spacer(1, 12))
                
                # إضافة ملخص التقرير
                summary_data = [
                    [Paragraph("<b>إجمالي المبيعات:</b>", styles['Normal']), Paragraph(self.profit_total_sales_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي التكلفة:</b>", styles['Normal']), Paragraph(self.profit_total_cost_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي الربح:</b>", styles['Normal']), Paragraph(self.profit_total_profit_value.text(), styles['Normal'])],
                    [Paragraph("<b>نسبة الربح:</b>", styles['Normal']), Paragraph(self.profit_margin_value.text(), styles['Normal'])],
                    [Paragraph("<b>إجمالي الضرائب:</b>", styles['Normal']), Paragraph(self.profit_total_tax_value.text(), styles['Normal'])],
                    [Paragraph("<b>صافي الربح:</b>", styles['Normal']), Paragraph(self.profit_net_profit_value.text(), styles['Normal'])]
                ]
                
                summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
                summary_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                ]))
                story.append(summary_table)
                story.append(Spacer(1, 12))
                
                # إضافة جدول الأرباح
                profit_data = [['التاريخ', 'عدد الفواتير', 'إجمالي المبيعات', 'إجمالي التكلفة', 'إجمالي الربح', 'نسبة الربح']]
                
                for row in range(self.profit_table.rowCount()):
                    row_data = []
                    for col in range(self.profit_table.columnCount()):
                        item = self.profit_table.item(row, col)
                        row_data.append(item.text() if item else "")
                    profit_data.append(row_data)
                
                profit_table = Table(profit_data, colWidths=[1*inch, 1*inch, 1*inch, 1*inch, 1*inch, 1*inch])
                profit_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ]))
                story.append(profit_table)
                
                # بناء ملف PDF
                doc.build(story)
            
            add_log("تصدير تقرير", "تم تصدير تقرير الأرباح", self.user_id)
            QMessageBox.information(self, "نجاح", "تم تصدير التقرير بنجاح")
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء تصدير التقرير: {str(e)}")
    
    def print_profit_report(self):
        # إنشاء ملف PDF مؤقت
        temp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        temp_file.close()
        
        # تصدير التقرير إلى الملف المؤقت
        file_path = temp_file.name
        
        from_date = self.profit_from_date_input.date().toString("yyyy-MM-dd")
        to_date = self.profit_to_date_input.date().toString("yyyy-MM-dd")
        
        try:
            # إنشاء ملف PDF
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # إنشاء قائمة العناصر
            story = []
            
            # إضافة عنوان التقرير
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph("تقرير الأرباح", title_style))
            
            # إضافة معلومات التقرير
            report_info = []
            report_info.append(Paragraph(f"<b>من تاريخ:</b> {from_date}", styles['Normal']))
            report_info.append(Paragraph(f"<b>إلى تاريخ:</b> {to_date}", styles['Normal']))
            
            report_table = Table([report_info], colWidths=[3*inch])
            report_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(report_table)
            story.append(Spacer(1, 12))
            
            # إضافة ملخص التقرير
            summary_data = [
                [Paragraph("<b>إجمالي المبيعات:</b>", styles['Normal']), Paragraph(self.profit_total_sales_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي التكلفة:</b>", styles['Normal']), Paragraph(self.profit_total_cost_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي الربح:</b>", styles['Normal']), Paragraph(self.profit_total_profit_value.text(), styles['Normal'])],
                [Paragraph("<b>نسبة الربح:</b>", styles['Normal']), Paragraph(self.profit_margin_value.text(), styles['Normal'])],
                [Paragraph("<b>إجمالي الضرائب:</b>", styles['Normal']), Paragraph(self.profit_total_tax_value.text(), styles['Normal'])],
                [Paragraph("<b>صافي الربح:</b>", styles['Normal']), Paragraph(self.profit_net_profit_value.text(), styles['Normal'])]
            ]
            
            summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # إضافة جدول الأرباح
            profit_data = [['التاريخ', 'عدد الفواتير', 'إجمالي المبيعات', 'إجمالي التكلفة', 'إجمالي الربح', 'نسبة الربح']]
            
            for row in range(self.profit_table.rowCount()):
                row_data = []
                for col in range(self.profit_table.columnCount()):
                    item = self.profit_table.item(row, col)
                    row_data.append(item.text() if item else "")
                profit_data.append(row_data)
            
            profit_table = Table(profit_data, colWidths=[1*inch, 1*inch, 1*inch, 1*inch, 1*inch, 1*inch])
            profit_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ]))
            story.append(profit_table)
            
            # بناء ملف PDF
            doc.build(story)
            
            # فتح الملف
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            
            add_log("طباعة تقرير", "تم طباعة تقرير الأرباح", self.user_id)
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء طباعة التقرير: {str(e)}")
        
        finally:
            # حذف الملف المؤقت بعد فتحه
            try:
                os.unlink(file_path)
            except:
                pass
    
    def create_currency_tab(self):
        layout = QVBoxLayout()
        
        # مجموعة العملة
        currency_group = QGroupBox("إعدادات العملة")
        currency_layout = QFormLayout()
        
        # اختيار العملة
        self.currency_combo = QComboBox()
        for currency in CURRENCIES:
            self.currency_combo.addItem(currency)
        
        # تعيين العملة الحالية
        current_currency = get_setting("currency", "ريال سعودي")
        index = self.currency_combo.findText(current_currency)
        if index >= 0:
            self.currency_combo.setCurrentIndex(index)
        
        self.currency_combo.currentTextChanged.connect(self.on_currency_changed)
        currency_layout.addRow("العملة:", self.currency_combo)
        
        # رمز العملة
        self.currency_symbol_input = QLineEdit()
        self.currency_symbol_input.setText(get_setting("currency_symbol", "ر.س"))
        currency_layout.addRow("رمز العملة:", self.currency_symbol_input)
        
        # عدد الأرقام العشرية
        self.decimal_places_spin = QSpinBox()
        self.decimal_places_spin.setRange(0, 6)
        self.decimal_places_spin.setValue(int(get_setting("decimal_places", "2")))
        currency_layout.addRow("عدد الأرقام العشرية:", self.decimal_places_spin)
        
        currency_group.setLayout(currency_layout)
        layout.addWidget(currency_group)
        
        # مجموعة العملات المتاحة
        available_currencies_group = QGroupBox("العملات المتاحة")
        available_currencies_layout = QVBoxLayout()
        
        # جدول العملات
        self.currencies_table = QTableWidget()
        self.currencies_table.setColumnCount(2)
        self.currencies_table.setHorizontalHeaderLabels(["العملة", "الرمز"])
        self.currencies_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.currencies_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.currencies_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.currencies_table.setSortingEnabled(True)
        
        # ملء جدول العملات
        self.currencies_table.setRowCount(len(CURRENCIES))
        for row, (currency, symbol) in enumerate(CURRENCIES.items()):
            self.currencies_table.setItem(row, 0, QTableWidgetItem(currency))
            self.currencies_table.setItem(row, 1, QTableWidgetItem(symbol))
        
        available_currencies_layout.addWidget(self.currencies_table)
        available_currencies_group.setLayout(available_currencies_layout)
        layout.addWidget(available_currencies_group)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_currency_settings)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.clicked.connect(self.reject_currency_settings)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        self.currency_tab.setLayout(layout)
    
    def on_currency_changed(self, currency_name):
        self.currency_symbol_input.setText(CURRENCIES.get(currency_name, ""))
    
    def save_currency_settings(self):
        currency = self.currency_combo.currentText()
        currency_symbol = self.currency_symbol_input.text()
        decimal_places = self.decimal_places_spin.value()
        
        update_setting("currency", currency)
        update_setting("currency_symbol", currency_symbol)
        update_setting("decimal_places", str(decimal_places))
        
        # تحديث المتغيرات المحلية
        self.currency = currency
        self.currency_symbol = currency_symbol
        self.decimal_places = decimal_places
        
        # تحديث الواجهة
        self.refresh_products_table()
        self.refresh_invoices_list()
        self.refresh_sales_report()
        self.refresh_inventory_report()
        self.refresh_profit_report()
        
        add_log("تعديل إعدادات العملة", f"تم تغيير العملة إلى {currency} ({currency_symbol})", self.user_id)
        QMessageBox.information(self, "نجاح", "تم حفظ إعدادات العملة بنجاح")
    
    def reject_currency_settings(self):
        # استعادة الإعدادات الأصلية
        self.currency = get_setting("currency", "ريال سعودي")
        self.currency_symbol = get_setting("currency_symbol", "ر.س")
        self.decimal_places = int(get_setting("decimal_places", "2"))
        
        # تحديث الواجهة
        index = self.currency_combo.findText(self.currency)
        if index >= 0:
            self.currency_combo.setCurrentIndex(index)
        
        self.currency_symbol_input.setText(self.currency_symbol)
        self.decimal_places_spin.setValue(self.decimal_places)
    
    def create_settings_tab(self):
        layout = QVBoxLayout()
        
        # مجموعة المعلومات الأساسية
        basic_info_group = QGroupBox("المعلومات الأساسية")
        basic_info_layout = QFormLayout()
        
        # اسم المتجر
        self.store_name_input = QLineEdit()
        self.store_name_input.setText(get_setting("store_name", "متجري"))
        basic_info_layout.addRow("اسم المتجر:", self.store_name_input)
        
        # العنوان
        self.store_address_input = QLineEdit()
        self.store_address_input.setText(get_setting("address", ""))
        basic_info_layout.addRow("العنوان:", self.store_address_input)
        
        # الهاتف
        self.store_phone_input = QLineEdit()
        self.store_phone_input.setText(get_setting("phone", ""))
        basic_info_layout.addRow("الهاتف:", self.store_phone_input)
        
        # البريد الإلكتروني
        self.store_email_input = QLineEdit()
        self.store_email_input.setText(get_setting("email", ""))
        basic_info_layout.addRow("البريد الإلكتروني:", self.store_email_input)
        
        # الموقع الإلكتروني
        self.store_website_input = QLineEdit()
        self.store_website_input.setText(get_setting("website", ""))
        basic_info_layout.addRow("الموقع الإلكتروني:", self.store_website_input)
        
        basic_info_group.setLayout(basic_info_layout)
        layout.addWidget(basic_info_group)
        
        # مجموعة الفواتير
        invoices_group = QGroupBox("إعدادات الفواتير")
        invoices_layout = QFormLayout()
        
        # بادئة رقم الفاتورة
        self.invoice_prefix_input = QLineEdit()
        self.invoice_prefix_input.setText(get_setting("invoice_prefix", "فاتورة-"))
        invoices_layout.addRow("بادئة رقم الفاتورة:", self.invoice_prefix_input)
        
        # نسبة الضريبة
        self.tax_rate_input = QDoubleSpinBox()
        self.tax_rate_input.setRange(0, 100)
        self.tax_rate_input.setValue(float(get_setting("tax_rate", "0.15")) * 100)
        self.tax_rate_input.setSuffix(" %")
        invoices_layout.addRow("نسبة الضريبة:", self.tax_rate_input)
        
        # تذييل الفاتورة
        self.invoice_footer_input = QTextEdit()
        self.invoice_footer_input.setText(get_setting("invoice_footer", "شكراً لتعاملكم معنا"))
        invoices_layout.addRow("تذييل الفاتورة:", self.invoice_footer_input)
        
        # رفع الشعار
        logo_layout = QHBoxLayout()
        
        self.logo_path_label = QLabel(get_setting("logo_path", "") or "لم يتم اختيار شعار")
        logo_layout.addWidget(self.logo_path_label)
        
        browse_logo_button = QPushButton("استعراض")
        browse_logo_button.clicked.connect(self.browse_logo)
        logo_layout.addWidget(browse_logo_button)
        
        clear_logo_button = QPushButton("مسح")
        clear_logo_button.clicked.connect(self.clear_logo)
        clear_logo_button.setObjectName("danger")
        logo_layout.addWidget(clear_logo_button)
        
        invoices_layout.addRow("شعار المتجر:", logo_layout)
        
        invoices_group.setLayout(invoices_layout)
        layout.addWidget(invoices_group)
        
        # مجموعة النسخ الاحتياطي
        backup_group = QGroupBox("إعدادات النسخ الاحتياطي")
        backup_layout = QFormLayout()
        
        # النسخ الاحتياطي التلقائي
        self.backup_auto_check = QCheckBox()
        self.backup_auto_check.setChecked(get_setting("backup_auto", "true") == "true")
        backup_layout.addRow("النسخ الاحتياطي التلقائي:", self.backup_auto_check)
        
        # تكرار النسخ الاحتياطي
        self.backup_frequency_spin = QSpinBox()
        self.backup_frequency_spin.setRange(1, 30)
        self.backup_frequency_spin.setValue(int(get_setting("backup_frequency", "7")))
        self.backup_frequency_spin.setSuffix(" أيام")
        backup_layout.addRow("تكرار النسخ الاحتياطي:", self.backup_frequency_spin)
        
        # مسار النسخ الاحتياطي
        backup_path_layout = QHBoxLayout()
        
        self.backup_path_label = QLabel(get_setting("auto_backup_path", "") or "المسار الافتراضي")
        backup_path_layout.addWidget(self.backup_path_label)
        
        browse_backup_path_button = QPushButton("استعراض")
        browse_backup_path_button.clicked.connect(self.browse_backup_path)
        backup_path_layout.addWidget(browse_backup_path_button)
        
        backup_layout.addRow("مسار النسخ الاحتياطي:", backup_path_layout)
        
        # أزرار النسخ الاحتياطي
        backup_buttons_layout = QHBoxLayout()
        
        backup_now_button = QPushButton("نسخ احتياطي الآن")
        backup_now_button.clicked.connect(self.backup_data)
        backup_buttons_layout.addWidget(backup_now_button)
        
        restore_button = QPushButton("استرجاع")
        restore_button.clicked.connect(self.restore_data)
        backup_buttons_layout.addWidget(restore_button)
        
        backup_layout.addRow("", backup_buttons_layout)
        
        backup_group.setLayout(backup_layout)
        layout.addWidget(backup_group)
        
        # مجموعة الإشعارات
        notifications_group = QGroupBox("إعدادات الإشعارات")
        notifications_layout = QFormLayout()
        
        # تفعيل الإشعارات
        self.notifications_check = QCheckBox()
        self.notifications_check.setChecked(get_setting("notifications", "true") == "true")
        notifications_layout.addRow("تفعيل الإشعارات:", self.notifications_check)
        
        # حد المخزون المنخفض
        self.low_stock_threshold_spin = QSpinBox()
        self.low_stock_threshold_spin.setRange(1, 1000)
        self.low_stock_threshold_spin.setValue(int(get_setting("low_stock_threshold", "10")))
        notifications_layout.addRow("حد المخزون المنخفض:", self.low_stock_threshold_spin)
        
        notifications_group.setLayout(notifications_layout)
        layout.addWidget(notifications_group)
        
        # مجموعة المظهر
        appearance_group = QGroupBox("إعدادات المظهر")
        appearance_layout = QFormLayout()
        
        # الثيم
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["فاتح", "داكن"])
        self.theme_combo.setCurrentIndex(0 if get_setting("theme", "light") == "light" else 1)
        appearance_layout.addRow("الثيم:", self.theme_combo)
        
        # تنسيق التاريخ
        self.date_format_combo = QComboBox()
        self.date_format_combo.addItems(["يوم/شهر/سنة", "شهر/يوم/سنة", "سنة/شهر/يوم"])
        current_format = get_setting("date_format", "dd/mm/yyyy")
        if current_format == "dd/mm/yyyy":
            self.date_format_combo.setCurrentIndex(0)
        elif current_format == "mm/dd/yyyy":
            self.date_format_combo.setCurrentIndex(1)
        else:
            self.date_format_combo.setCurrentIndex(2)
        appearance_layout.addRow("تنسيق التاريخ:", self.date_format_combo)
        
        # تنسيق الوقت
        self.time_format_combo = QComboBox()
        self.time_format_combo.addItems(["24 ساعة", "12 ساعة"])
        self.time_format_combo.setCurrentIndex(0 if get_setting("time_format", "24h") == "24h" else 1)
        appearance_layout.addRow("تنسيق الوقت:", self.time_format_combo)
        
        appearance_group.setLayout(appearance_layout)
        layout.addWidget(appearance_group)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_settings)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.clicked.connect(self.reject_settings)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        self.settings_tab.setLayout(layout)
    
    def browse_logo(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "اختر شعار المتجر", "", "صور (*.png *.jpg *.jpeg *.bmp *.gif)"
        )
        
        if file_path:
            self.logo_path_label.setText(file_path)
    
    def clear_logo(self):
        self.logo_path_label.setText("")
    
    def browse_backup_path(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "اختر مسار النسخ الاحتياطي"
        )
        
        if dir_path:
            self.backup_path_label.setText(dir_path)
    
    def save_settings(self):
        # حفظ الإعدادات الأساسية
        update_setting("store_name", self.store_name_input.text())
        update_setting("address", self.store_address_input.text())
        update_setting("phone", self.store_phone_input.text())
        update_setting("email", self.store_email_input.text())
        update_setting("website", self.store_website_input.text())
        
        # حفظ إعدادات الفواتير
        update_setting("invoice_prefix", self.invoice_prefix_input.text())
        update_setting("tax_rate", str(self.tax_rate_input.value() / 100))
        update_setting("invoice_footer", self.invoice_footer_input.toPlainText())
        update_setting("logo_path", self.logo_path_label.text())
        
        # حفظ إعدادات النسخ الاحتياطي
        update_setting("backup_auto", "true" if self.backup_auto_check.isChecked() else "false")
        update_setting("backup_frequency", str(self.backup_frequency_spin.value()))
        update_setting("auto_backup_path", self.backup_path_label.text())
        
        # حفظ إعدادات الإشعارات
        update_setting("notifications", "true" if self.notifications_check.isChecked() else "false")
        update_setting("low_stock_threshold", str(self.low_stock_threshold_spin.value()))
        
        # حفظ إعدادات المظهر
        theme = "light" if self.theme_combo.currentIndex() == 0 else "dark"
        update_setting("theme", theme)
        
        if self.date_format_combo.currentIndex() == 0:
            date_format = "dd/mm/yyyy"
        elif self.date_format_combo.currentIndex() == 1:
            date_format = "mm/dd/yyyy"
        else:
            date_format = "yyyy/mm/dd"
        update_setting("date_format", date_format)
        
        time_format = "24h" if self.time_format_combo.currentIndex() == 0 else "12h"
        update_setting("time_format", time_format)
        
        # تحديث المتغيرات المحلية
        self.store_name = self.store_name_input.text()
        self.tax_rate = self.tax_rate_input.value() / 100
        self.invoice_prefix = self.invoice_prefix_input.text()
        self.backup_auto = self.backup_auto_check.isChecked()
        self.backup_frequency = self.backup_frequency_spin.value()
        self.notifications = self.notifications_check.isChecked()
        self.low_stock_threshold = self.low_stock_threshold_spin.value()
        self.theme = theme
        
        # تحديث الواجهة
        self.setWindowTitle(f"{self.store_name} - XnoneDBS")
        self.refresh_products_table()
        self.refresh_invoices_list()
        self.refresh_sales_report()
        self.refresh_inventory_report()
        self.refresh_profit_report()
        self.update_low_stock_count()
        
        # إعادة تشغيل المؤقتات إذا لزم الأمر
        if hasattr(self, 'backup_timer'):
            self.backup_timer.stop()
            if self.backup_auto:
                interval = self.backup_frequency * 24 * 60 * 60 * 1000
                self.backup_timer.start(interval)
        
        if hasattr(self, 'low_stock_timer'):
            self.low_stock_timer.stop()
            self.low_stock_timer.start(60 * 60 * 1000)  # ساعة واحدة بالمللي ثانية
        
        add_log("تعديل الإعدادات", "تم تعديل الإعدادات العامة", self.user_id)
        QMessageBox.information(self, "نجاح", "تم حفظ الإعدادات بنجاح")
    
    def reject_settings(self):
        # استعادة الإعدادات الأصلية
        self.store_name_input.setText(get_setting("store_name", "متجري"))
        self.store_address_input.setText(get_setting("address", ""))
        self.store_phone_input.setText(get_setting("phone", ""))
        self.store_email_input.setText(get_setting("email", ""))
        self.store_website_input.setText(get_setting("website", ""))
        self.invoice_prefix_input.setText(get_setting("invoice_prefix", "فاتورة-"))
        self.tax_rate_input.setValue(float(get_setting("tax_rate", "0.15")) * 100)
        self.invoice_footer_input.setText(get_setting("invoice_footer", "شكراً لتعاملكم معنا"))
        self.logo_path_label.setText(get_setting("logo_path", "") or "لم يتم اختيار شعار")
        self.backup_auto_check.setChecked(get_setting("backup_auto", "true") == "true")
        self.backup_frequency_spin.setValue(int(get_setting("backup_frequency", "7")))
        self.backup_path_label.setText(get_setting("auto_backup_path", "") or "المسار الافتراضي")
        self.notifications_check.setChecked(get_setting("notifications", "true") == "true")
        self.low_stock_threshold_spin.setValue(int(get_setting("low_stock_threshold", "10")))
        self.theme_combo.setCurrentIndex(0 if get_setting("theme", "light") == "light" else 1)
        
        current_format = get_setting("date_format", "dd/mm/yyyy")
        if current_format == "dd/mm/yyyy":
            self.date_format_combo.setCurrentIndex(0)
        elif current_format == "mm/dd/yyyy":
            self.date_format_combo.setCurrentIndex(1)
        else:
            self.date_format_combo.setCurrentIndex(2)
        
        self.time_format_combo.setCurrentIndex(0 if get_setting("time_format", "24h") == "24h" else 1)
    
    def create_users_tab(self):
        layout = QVBoxLayout()
        
        # شريط الأدوات
        toolbar_layout = QHBoxLayout()
        
        # زر إضافة مستخدم
        add_user_button = QPushButton("إضافة مستخدم")
        add_user_button.clicked.connect(self.add_user)
        toolbar_layout.addWidget(add_user_button)
        
        # زر تعديل مستخدم
        edit_user_button = QPushButton("تعديل مستخدم")
        edit_user_button.clicked.connect(self.edit_user)
        toolbar_layout.addWidget(edit_user_button)
        
        # زر حذف مستخدم
        delete_user_button = QPushButton("حذف مستخدم")
        delete_user_button.clicked.connect(self.delete_user)
        delete_user_button.setObjectName("danger")
        toolbar_layout.addWidget(delete_user_button)
        
        # زر تغيير كلمة المرور
        change_password_button = QPushButton("تغيير كلمة المرور")
        change_password_button.clicked.connect(self.change_user_password)
        toolbar_layout.addWidget(change_password_button)
        
        # زر تحديث الجدول
        refresh_button = QPushButton("تحديث")
        refresh_button.clicked.connect(self.refresh_users_table)
        toolbar_layout.addWidget(refresh_button)
        
        # حقل البحث
        search_label = QLabel("بحث:")
        toolbar_layout.addWidget(search_label)
        
        self.user_search_input = QLineEdit()
        self.user_search_input.setPlaceholderText("ابحث عن مستخدم...")
        self.user_search_input.textChanged.connect(self.filter_users)
        toolbar_layout.addWidget(self.user_search_input)
        
        layout.addLayout(toolbar_layout)
        
        # جدول المستخدمين
        self.users_table = QTableWidget()
        self.users_table.setColumnCount(6)
        self.users_table.setHorizontalHeaderLabels([
            "اسم المستخدم", "الاسم الكامل", "البريد الإلكتروني", "رقم الهاتف", "الصلاحية", "تاريخ الإنشاء"
        ])
        self.users_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.users_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.users_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.users_table.setSortingEnabled(True)
        self.users_table.doubleClicked.connect(self.edit_user)
        layout.addWidget(self.users_table)
        
        self.users_tab.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_users_table()
    
    def refresh_users_table(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id, username, full_name, email, phone, role, created_at FROM users ORDER BY username")
        users = cursor.fetchall()
        conn.close()
        
        self.users_table.setRowCount(len(users))
        
        for row, user in enumerate(users):
            # اسم المستخدم
            self.users_table.setItem(row, 0, QTableWidgetItem(user[1]))
            
            # الاسم الكامل
            self.users_table.setItem(row, 1, QTableWidgetItem(user[2] or ""))
            
            # البريد الإلكتروني
            self.users_table.setItem(row, 2, QTableWidgetItem(user[3] or ""))
            
            # رقم الهاتف
            self.users_table.setItem(row, 3, QTableWidgetItem(user[4] or ""))
            
            # الصلاحية
            role_text = "مدير" if user[5] == "admin" else "مستخدم عادي"
            self.users_table.setItem(row, 4, QTableWidgetItem(role_text))
            
            # تاريخ الإنشاء
            created_at = datetime.strptime(user[6], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
            self.users_table.setItem(row, 5, QTableWidgetItem(created_at))
    
    def filter_users(self):
        search_text = self.user_search_input.text().lower()
        
        for row in range(self.users_table.rowCount()):
            match = False
            
            for col in range(self.users_table.columnCount()):
                item = self.users_table.item(row, col)
                if item and search_text in item.text().lower():
                    match = True
                    break
            
            self.users_table.setRowHidden(row, not match)
    
    def add_user(self):
        dialog = UserDialog(self)
        if dialog.exec_():
            self.refresh_users_table()
            add_log("إضافة مستخدم", f"تم إضافة المستخدم: {dialog.username}", self.user_id)
    
    def edit_user(self):
        selected_rows = self.users_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار مستخدم للتعديل")
            return
        
        row = selected_rows[0].row()
        username = self.users_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ?", (username,))
        user = cursor.fetchone()
        conn.close()
        
        if user:
            dialog = UserDialog(self, user)
            if dialog.exec_():
                self.refresh_users_table()
                add_log("تعديل مستخدم", f"تم تعديل المستخدم: {dialog.username}", self.user_id)
    
    def delete_user(self):
        selected_rows = self.users_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار مستخدم للحذف")
            return
        
        row = selected_rows[0].row()
        username = self.users_table.item(row, 0).text()
        
        # لا يمكن حذف المستخدم الحالي
        if username == self.username:
            QMessageBox.warning(self, "تحذير", "لا يمكنك حذف حسابك الحالي")
            return
        
        reply = QMessageBox.question(
            self, "تأكيد الحذف", 
            f"هل أنت متأكد من حذف المستخدم '{username}'؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("DELETE FROM users WHERE username = ?", (username,))
            conn.commit()
            conn.close()
            
            self.refresh_users_table()
            add_log("حذف مستخدم", f"تم حذف المستخدم: {username}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم حذف المستخدم بنجاح")
    
    def change_user_password(self):
        selected_rows = self.users_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار مستخدم لتغيير كلمة المرور")
            return
        
        row = selected_rows[0].row()
        username = self.users_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
        user_id = cursor.fetchone()[0]
        conn.close()
        
        dialog = ChangePasswordDialog(self, user_id)
        if dialog.exec_():
            add_log("تغيير كلمة المرور", f"تم تغيير كلمة مرور المستخدم: {username}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم تغيير كلمة المرور بنجاح")
    
    def backup_data(self):
        # إنشاء مجلد النسخ الاحتياطي إذا لم يكن موجوداً
        backup_dir = QFileDialog.getExistingDirectory(
            self, "اختر مجلد لحفظ النسخ الاحتياطي"
        )
        
        if not backup_dir:
            return
        
        try:
            # إنشاء اسم ملف النسخ الاحتياطي
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"xnonedbs_backup_{timestamp}.db")
            
            # نسخ قاعدة البيانات
            shutil.copy2("xnonedbs.db", backup_file)
            
            # تسجيل النسخ الاحتياطي
            file_size = os.path.getsize(backup_file) / (1024 * 1024)  # حجم الملف بالميجابايت
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("INSERT INTO backups (file_path, size, notes) VALUES (?, ?, ?)",
                          (backup_file, file_size, "نسخ احتياطي يدوي"))
            conn.commit()
            conn.close()
            
            # تحديث تاريخ آخر نسخ احتياطي
            update_setting("last_backup_date", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            
            add_log("نسخ احتياطي", f"تم إنشاء نسخة احتياطية في {backup_file}", self.user_id)
            QMessageBox.information(self, "نجاح", "تم إنشاء النسخة الاحتياطية بنجاح")
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء إنشاء النسخة الاحتياطية: {str(e)}")
    
    def restore_data(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف النسخ الاحتياطي", "", "ملفات قاعدة البيانات (*.db)"
        )
        
        if not file_path:
            return
        
        reply = QMessageBox.question(
            self, "تأكيد الاسترجاع", 
            "هل أنت متأكد من استرجاع البيانات من النسخة الاحتياطية؟ سيتم استبدال جميع البيانات الحالية.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                # إنشاء نسخة احتياطية من قاعدة البيانات الحالية
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = f"xnonedbs_backup_before_restore_{timestamp}.db"
                shutil.copy2("xnonedbs.db", backup_file)
                
                # استرجاع قاعدة البيانات
                shutil.copy2(file_path, "xnonedbs.db")
                
                add_log("استرجاع بيانات", f"تم استرجاع البيانات من {file_path}", self.user_id)
                QMessageBox.information(self, "نجاح", "تم استرجاع البيانات بنجاح")
                
                # إعادة تشغيل التطبيق
                QMessageBox.information(self, "إعادة التشغيل", "سيتم إعادة تشغيل التطبيق لتطبيق التغييرات")
                QApplication.quit()
                os.execl(sys.executable, sys.executable, *sys.argv)
            
            except Exception as e:
                QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء استرجاع البيانات: {str(e)}")
    
    def show_search_dialog(self):
        dialog = SearchDialog(self)
        dialog.exec_()
    
    def show_about(self):
        about_text = """
        <h2>XnoneDBS</h2>
        <p>نظام إدارة المحلات التجارية المتكامل</p>
        <p>الإصدار: 1.0.0</p>
        <p>تم تطويره باستخدام PyQt5 و SQLite</p>
        <p>جميع الحقوق محفوظة &copy; 2023</p>
        """
        
        QMessageBox.about(self, "حول البرنامج", about_text)
    
    def show_help(self):
        help_text = """
        <h2>مساعدة XnoneDBS</h2>
        <h3>المنتجات</h3>
        <p>يمكنك إضافة وتعديل وحذف المنتجات من تبويب المنتجات. يمكنك أيضاً استيراد وتصدير المنتجات من وإلى ملفات Excel.</p>
        
        <h3>الفواتير</h3>
        <p>يمكنك إنشاء فواتير جديدة من تبويب الفواتير. يمكنك إضافة المنتجات عن طريق مسح الباركود أو البحث عن المنتج يدوياً.</p>
        
        <h3>العملاء</h3>
        <p>يمكنك إدارة العملاء من تبويب العملاء. يمكنك إضافة وتعديل وحذف العملاء.</p>
        
        <h3>الموردين</h3>
        <p>يمكنك إدارة الموردين من تبويب الموردين. يمكنك إضافة وتعديل وحذف الموردين.</p>
        
        <h3>التقارير</h3>
        <p>يمكنك عرض تقارير المبيعات والمخزون والأرباح من تبويبات التقارير. يمكنك تصدير التقارير إلى ملفات Excel أو PDF.</p>
        
        <h3>الإعدادات</h3>
        <p>يمكنك تغيير إعدادات المتجر والعملة والنسخ الاحتياطي من تبويب الإعدادات.</p>
        
        <h3>المستخدمون</h3>
        <p>يمكنك إدارة المستخدمين من تبويب المستخدمين (للمدير فقط). يمكنك إضافة وتعديل وحذف المستخدمين وتغيير كلمات المرور.</p>
        """
        
        QMessageBox.information(self, "مساعدة", help_text)
    
    def closeEvent(self, event):
        # تسجيل الخروج
        add_log("تسجيل الخروج", f"قام المستخدم {self.username} بتسجيل الخروج", self.user_id)
        
        # إخفاء أيقونة النظام
        self.tray_icon.hide()
        
        # قبول الحدث
        event.accept()

# حوار إضافة/تعديل المنتج
class ProductDialog(QDialog):
    def __init__(self, parent=None, product=None):
        super().__init__(parent)
        self.product = product
        self.product_name = ""
        
        if product:
            self.setWindowTitle("تعديل منتج")
        else:
            self.setWindowTitle("إضافة منتج")
        
        self.setFixedSize(500, 600)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QTextEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
            QPushButton#generate {
                background-color: #4CAF50;
            }
            QPushButton#generate:hover {
                background-color: #45a049;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # الباركود
        barcode_layout = QHBoxLayout()
        
        self.barcode_input = QLineEdit()
        self.barcode_input.setPlaceholderText("أدخل الباركود أو اتركه فارغاً للتوليد التلقائي")
        barcode_layout.addWidget(self.barcode_input)
        
        generate_barcode_button = QPushButton("توليد")
        generate_barcode_button.setObjectName("generate")
        generate_barcode_button.clicked.connect(self.generate_barcode)
        barcode_layout.addWidget(generate_barcode_button)
        
        form_layout.addRow("الباركود:", barcode_layout)
        
        # اسم المنتج
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("أدخل اسم المنتج")
        form_layout.addRow("اسم المنتج:", self.name_input)
        
        # الوصف
        self.description_input = QTextEdit()
        self.description_input.setPlaceholderText("أدخل وصف المنتج")
        form_layout.addRow("الوصف:", self.description_input)
        
        # التصنيف
        self.category_combo = QComboBox()
        self.category_combo.setEditable(True)
        self.load_categories()
        form_layout.addRow("التصنيف:", self.category_combo)
        
        # الكمية
        self.quantity_input = QDoubleSpinBox()
        self.quantity_input.setRange(0, 1000000)
        self.quantity_input.setValue(1)
        self.quantity_input.setDecimals(2)
        form_layout.addRow("الكمية:", self.quantity_input)
        
        # سعر البيع
        self.price_input = QDoubleSpinBox()
        self.price_input.setRange(0, 1000000)
        self.price_input.setValue(0)
        self.price_input.setDecimals(2)
        self.price_input.setSuffix(f" {self.parent().currency_symbol}")
        form_layout.addRow("سعر البيع:", self.price_input)
        
        # سعر الشراء
        self.cost_price_input = QDoubleSpinBox()
        self.cost_price_input.setRange(0, 1000000)
        self.cost_price_input.setValue(0)
        self.cost_price_input.setDecimals(2)
        self.cost_price_input.setSuffix(f" {self.parent().currency_symbol}")
        form_layout.addRow("سعر الشراء:", self.cost_price_input)
        
        # الحد الأدنى للكمية
        self.min_quantity_input = QDoubleSpinBox()
        self.min_quantity_input.setRange(0, 1000000)
        self.min_quantity_input.setValue(0)
        self.min_quantity_input.setDecimals(2)
        form_layout.addRow("الحد الأدنى للكمية:", self.min_quantity_input)
        
        # تاريخ الانتهاء
        self.expiry_date_input = QDateEdit()
        self.expiry_date_input.setCalendarPopup(True)
        self.expiry_date_input.setDate(QDate.currentDate().addMonths(1))
        form_layout.addRow("تاريخ الانتهاء:", self.expiry_date_input)
        
        # المورد
        self.supplier_combo = QComboBox()
        self.supplier_combo.setEditable(True)
        self.load_suppliers()
        form_layout.addRow("المورد:", self.supplier_combo)
        
        # ملاحظات
        self.notes_input = QTextEdit()
        self.notes_input.setPlaceholderText("أدخل ملاحظات إضافية")
        form_layout.addRow("ملاحظات:", self.notes_input)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_product)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # ملء الحقول إذا كان تعديل منتج
        if product:
            self.barcode_input.setText(product[1] or "")
            self.name_input.setText(product[2])
            self.description_input.setText(product[3] or "")
            
            category_index = self.category_combo.findText(product[4] or "")
            if category_index >= 0:
                self.category_combo.setCurrentIndex(category_index)
            else:
                self.category_combo.setCurrentText(product[4] or "")
            
            self.quantity_input.setValue(product[5])
            self.price_input.setValue(product[6])
            self.cost_price_input.setValue(product[7] or 0)
            self.min_quantity_input.setValue(product[8] or 0)
            
            if product[9]:
                expiry_date = QDate.fromString(product[9], "yyyy-MM-dd")
                self.expiry_date_input.setDate(expiry_date)
                            
            supplier_index = self.supplier_combo.findText(product[10] or "")
            if supplier_index >= 0:
                self.supplier_combo.setCurrentIndex(supplier_index)
            else:
                self.supplier_combo.setCurrentText(product[10] or "")
            
            self.notes_input.setText(product[11] or "")
        
        # التركيز على حقل اسم المنتج
        self.name_input.setFocus()
    
    def load_categories(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM categories ORDER BY name")
        categories = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        self.category_combo.addItems(categories)
    
    def load_suppliers(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM suppliers ORDER BY name")
        suppliers = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        self.supplier_combo.addItems(suppliers)
    
    def generate_barcode(self):
        # توليد باركود فريد
        barcode = str(uuid.uuid4().int)[:12]
        self.barcode_input.setText(barcode)
    
    def save_product(self):
        # الحصول على البيانات من الحقول
        barcode = self.barcode_input.text().strip()
        name = self.name_input.text().strip()
        description = self.description_input.toPlainText().strip()
        category = self.category_combo.currentText().strip()
        quantity = self.quantity_input.value()
        price = self.price_input.value()
        cost_price = self.cost_price_input.value()
        min_quantity = self.min_quantity_input.value()
        expiry_date = self.expiry_date_input.date().toString("yyyy-MM-dd")
        supplier = self.supplier_combo.currentText().strip()
        notes = self.notes_input.toPlainText().strip()
        
        # التحقق من الحقول المطلوبة
        if not name:
            self.error_label.setText("يرجى إدخال اسم المنتج")
            return
        
        if price <= 0:
            self.error_label.setText("يرجى إدخال سعر بيع صحيح")
            return
        
        if quantity <= 0:
            self.error_label.setText("يرجى إدخال كمية صحيحة")
            return
        
        # إذا لم يتم إدخال باركود، قم بتوليد واحد
        if not barcode:
            barcode = str(uuid.uuid4().int)[:12]
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            if self.product:
                # تحديث منتج موجود
                cursor.execute(
                    """
                    UPDATE products SET
                    barcode = ?, name = ?, description = ?, category = ?, quantity = ?,
                    price = ?, cost_price = ?, min_quantity = ?, expiry_date = ?,
                    supplier = ?, notes = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (
                        barcode, name, description, category, quantity, price,
                        cost_price, min_quantity, expiry_date, supplier, notes,
                        self.product[0]
                    )
                )
            else:
                # إضافة منتج جديد
                cursor.execute(
                    """
                    INSERT INTO products
                    (barcode, name, description, category, quantity, price, cost_price, min_quantity, expiry_date, supplier, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (barcode, name, description, category, quantity, price, cost_price, min_quantity, expiry_date, supplier, notes)
                )
                
                # إذا كان التصنيف جديداً، أضفه إلى جدول التصنيفات
                if category:
                    cursor.execute("SELECT COUNT(*) FROM categories WHERE name = ?", (category,))
                    if cursor.fetchone()[0] == 0:
                        cursor.execute("INSERT INTO categories (name) VALUES (?)", (category,))
                
                # إذا كان المورد جديداً، أضفه إلى جدول الموردين
                if supplier:
                    cursor.execute("SELECT COUNT(*) FROM suppliers WHERE name = ?", (supplier,))
                    if cursor.fetchone()[0] == 0:
                        cursor.execute("INSERT INTO suppliers (name) VALUES (?)", (supplier,))
            
            conn.commit()
            conn.close()
            
            self.product_name = name
            self.accept()
        
        except sqlite3.IntegrityError:
            conn.close()
            self.error_label.setText("الباركود مستخدم بالفعل")
        except Exception as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# حوار البحث عن منتج للفاتورة
class ProductSearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("بحث عن منتج")
        self.setFixedSize(600, 400)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9f9f9;
                gridline-color: #ddd;
                color: #333;
                border: 1px solid #ddd;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: #2196F3;
                color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #333;
                padding: 5px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        # حقل البحث
        search_layout = QHBoxLayout()
        
        search_label = QLabel("بحث:")
        search_layout.addWidget(search_label)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("ابحث عن منتج بالاسم أو الباركود...")
        self.search_input.textChanged.connect(self.filter_products)
        search_layout.addWidget(self.search_input)
        
        layout.addLayout(search_layout)
        
        # جدول المنتجات
        self.products_table = QTableWidget()
        self.products_table.setColumnCount(5)
        self.products_table.setHorizontalHeaderLabels([
            "الباركود", "اسم المنتج", "التصنيف", "الكمية", "السعر"
        ])
        self.products_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.products_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.products_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.products_table.setSortingEnabled(True)
        self.products_table.doubleClicked.connect(self.select_product)
        layout.addWidget(self.products_table)
        
        # أزرار الاختيار والإلغاء
        buttons_layout = QHBoxLayout()
        
        select_button = QPushButton("اختيار")
        select_button.clicked.connect(self.select_product)
        buttons_layout.addWidget(select_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
        
        # تحميل البيانات
        self.refresh_products_table()
        
        # التركيز على حقل البحث
        self.search_input.setFocus()
        
        self.selected_product = None
    
    def refresh_products_table(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products ORDER BY name")
        products = cursor.fetchall()
        conn.close()
        
        self.products_table.setRowCount(len(products))
        
        for row, product in enumerate(products):
            # الباركود
            self.products_table.setItem(row, 0, QTableWidgetItem(product[1] or ""))
            
            # اسم المنتج
            self.products_table.setItem(row, 1, QTableWidgetItem(product[2]))
            
            # التصنيف
            self.products_table.setItem(row, 2, QTableWidgetItem(product[4] or ""))
            
            # الكمية
            quantity_item = QTableWidgetItem(str(product[5]))
            if product[5] <= 10:  # حد المخزون المنخفض
                quantity_item.setBackground(QColor(255, 200, 200))
            self.products_table.setItem(row, 3, quantity_item)
            
            # السعر
            price_item = QTableWidgetItem(f"{product[6]:.2f} {self.parent().currency_symbol}")
            self.products_table.setItem(row, 4, price_item)
    
    def filter_products(self):
        search_text = self.search_input.text().lower()
        
        for row in range(self.products_table.rowCount()):
            match = False
            
            for col in range(self.products_table.columnCount()):
                item = self.products_table.item(row, col)
                if item and search_text in item.text().lower():
                    match = True
                    break
            
            self.products_table.setRowHidden(row, not match)
    
    def select_product(self):
        selected_rows = self.products_table.selectedItems()
        if not selected_rows:
            QMessageBox.warning(self, "تحذير", "يرجى اختيار منتج")
            return
        
        row = selected_rows[0].row()
        barcode = self.products_table.item(row, 0).text()
        
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products WHERE barcode = ?", (barcode,))
        product = cursor.fetchone()
        conn.close()
        
        if product:
            self.selected_product = {
                'id': product[0],
                'barcode': product[1],
                'name': product[2],
                'price': product[6]
            }
            self.accept()

# حوار عرض الفاتورة
class InvoiceViewDialog(QDialog):
    def __init__(self, parent=None, invoice_number=None):
        super().__init__(parent)
        self.invoice_number = invoice_number
        self.setWindowTitle(f"عرض الفاتورة - {invoice_number}")
        self.setFixedSize(800, 600)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9f9f9;
                gridline-color: #ddd;
                color: #333;
                border: 1px solid #ddd;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #333;
                padding: 5px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#close {
                background-color: #f44336;
            }
            QPushButton#close:hover {
                background-color: #d32f2f;
            }
            QPushButton#print {
                background-color: #4CAF50;
            }
            QPushButton#print:hover {
                background-color: #45a049;
            }
        """)
        
        layout = QVBoxLayout()
        
        # معلومات الفاتورة
        info_group = QGroupBox("معلومات الفاتورة")
        info_layout = QFormLayout()
        
        # رقم الفاتورة
        self.invoice_number_label = QLabel()
        info_layout.addRow("رقم الفاتورة:", self.invoice_number_label)
        
        # التاريخ
        self.invoice_date_label = QLabel()
        info_layout.addRow("التاريخ:", self.invoice_date_label)
        
        # اسم العميل
        self.customer_name_label = QLabel()
        info_layout.addRow("اسم العميل:", self.customer_name_label)
        
        # رقم هاتف العميل
        self.customer_phone_label = QLabel()
        info_layout.addRow("رقم هاتف العميل:", self.customer_phone_label)
        
        # طريقة الدفع
        self.payment_method_label = QLabel()
        info_layout.addRow("طريقة الدفع:", self.payment_method_label)
        
        # المستخدم
        self.user_label = QLabel()
        info_layout.addRow("المستخدم:", self.user_label)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # جدول منتجات الفاتورة
        self.items_table = QTableWidget()
        self.items_table.setColumnCount(5)
        self.items_table.setHorizontalHeaderLabels([
            "الباركود", "اسم المنتج", "الكمية", "السعر", "الإجمالي"
        ])
        self.items_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.items_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.items_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.items_table)
        
        # ملخص الفاتورة
        summary_group = QGroupBox("ملخص الفاتورة")
        summary_layout = QFormLayout()
        
        # المجموع
        self.subtotal_label = QLabel()
        summary_layout.addRow("المجموع:", self.subtotal_label)
        
        # الضريبة
        self.tax_label = QLabel()
        summary_layout.addRow("الضريبة:", self.tax_label)
        
        # الخصم
        self.discount_label = QLabel()
        summary_layout.addRow("الخصم:", self.discount_label)
        
        # الإجمالي
        self.total_label = QLabel()
        total_font = QFont()
        total_font.setBold(True)
        total_font.setPointSize(14)
        self.total_label.setFont(total_font)
        summary_layout.addRow("الإجمالي:", self.total_label)
        
        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)
        
        # ملاحظات الفاتورة
        notes_group = QGroupBox("ملاحظات")
        notes_layout = QVBoxLayout()
        
        self.notes_label = QLabel()
        self.notes_label.setWordWrap(True)
        notes_layout.addWidget(self.notes_label)
        
        notes_group.setLayout(notes_layout)
        layout.addWidget(notes_group)
        
        # أزرار الطباعة والإغلاق
        buttons_layout = QHBoxLayout()
        
        print_button = QPushButton("طباعة")
        print_button.setObjectName("print")
        print_button.clicked.connect(self.print_invoice)
        buttons_layout.addWidget(print_button)
        
        close_button = QPushButton("إغلاق")
        close_button.setObjectName("close")
        close_button.clicked.connect(self.accept)
        buttons_layout.addWidget(close_button)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
        
        # تحميل بيانات الفاتورة
        self.load_invoice_data()
    
    def load_invoice_data(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        
        # الحصول على بيانات الفاتورة
        cursor.execute("""
            SELECT i.invoice_number, i.date, i.customer_name, i.customer_phone, 
                   i.payment_method, u.full_name, i.subtotal, i.tax, i.discount, 
                   i.total, i.notes
            FROM invoices i
            JOIN users u ON i.user_id = u.id
            WHERE i.invoice_number = ?
        """, (self.invoice_number,))
        
        invoice = cursor.fetchone()
        
        if invoice:
            # تحديث معلومات الفاتورة
            self.invoice_number_label.setText(invoice[0])
            
            date_str = datetime.strptime(invoice[1], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M")
            self.invoice_date_label.setText(date_str)
            
            self.customer_name_label.setText(invoice[2] or "")
            self.customer_phone_label.setText(invoice[3] or "")
            self.payment_method_label.setText(invoice[4])
            self.user_label.setText(invoice[5])
            
            # تحديث ملخص الفاتورة
            self.subtotal_label.setText(f"{invoice[6]:.2f} {self.parent().currency_symbol}")
            self.tax_label.setText(f"{invoice[7]:.2f} {self.parent().currency_symbol}")
            self.discount_label.setText(f"{invoice[8]:.2f} {self.parent().currency_symbol}")
            self.total_label.setText(f"{invoice[9]:.2f} {self.parent().currency_symbol}")
            
            # تحديث الملاحظات
            self.notes_label.setText(invoice[10] or "")
            
            # الحصول على منتجات الفاتورة
            cursor.execute("""
                SELECT ii.quantity, ii.price, ii.total, p.barcode, p.name
                FROM invoice_items ii
                JOIN products p ON ii.product_id = p.id
                JOIN invoices i ON ii.invoice_id = i.id
                WHERE i.invoice_number = ?
                ORDER BY p.name
            """, (self.invoice_number,))
            
            items = cursor.fetchall()
            
            # تحديث جدول المنتجات
            self.items_table.setRowCount(len(items))
            
            for row, item in enumerate(items):
                # الباركود
                self.items_table.setItem(row, 0, QTableWidgetItem(item[3] or ""))
                
                # اسم المنتج
                self.items_table.setItem(row, 1, QTableWidgetItem(item[4]))
                
                # الكمية
                self.items_table.setItem(row, 2, QTableWidgetItem(str(item[0])))
                
                # السعر
                self.items_table.setItem(row, 3, QTableWidgetItem(f"{item[1]:.2f} {self.parent().currency_symbol}"))
                
                # الإجمالي
                self.items_table.setItem(row, 4, QTableWidgetItem(f"{item[2]:.2f} {self.parent().currency_symbol}"))
        
        conn.close()
    
    def print_invoice(self):
        # إنشاء ملف PDF مؤقت
        temp_file = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        temp_file.close()
        
        # تصدير الفاتورة إلى الملف المؤقت
        file_path = temp_file.name
        
        try:
            # الحصول على بيانات الفاتورة
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT i.invoice_number, i.date, i.customer_name, i.customer_phone, 
                       i.payment_method, u.full_name, i.subtotal, i.tax, i.discount, 
                       i.total, i.notes
                FROM invoices i
                JOIN users u ON i.user_id = u.id
                WHERE i.invoice_number = ?
            """, (self.invoice_number,))
            
            invoice = cursor.fetchone()
            
            cursor.execute("""
                SELECT ii.quantity, ii.price, ii.total, p.barcode, p.name
                FROM invoice_items ii
                JOIN products p ON ii.product_id = p.id
                JOIN invoices i ON ii.invoice_id = i.id
                WHERE i.invoice_number = ?
                ORDER BY p.name
            """, (self.invoice_number,))
            
            items = cursor.fetchall()
            conn.close()
            
            # إنشاء مستند PDF
            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # إنشاء قائمة العناصر
            story = []
            
            # إضافة الشعار إذا كان موجوداً
            logo_path = get_setting("logo_path", "")
            if logo_path and os.path.exists(logo_path):
                logo = Image(logo_path, width=100, height=100)
                story.append(logo)
            
            # إضافة عنوان الفاتورة
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=TA_CENTER
            )
            story.append(Paragraph(f"فاتورة مبيعات - {get_setting('store_name', 'متجري')}", title_style))
            
            # إضافة معلومات المتجر
            store_info = []
            store_info.append(Paragraph(f"<b>اسم المتجر:</b> {get_setting('store_name', '')}", styles['Normal']))
            store_info.append(Paragraph(f"<b>العنوان:</b> {get_setting('address', '')}", styles['Normal']))
            store_info.append(Paragraph(f"<b>الهاتف:</b> {get_setting('phone', '')}", styles['Normal']))
            store_info.append(Paragraph(f"<b>البريد الإلكتروني:</b> {get_setting('email', '')}", styles['Normal']))
            
            store_table = Table([store_info], colWidths=[3*inch])
            store_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(store_table)
            story.append(Spacer(1, 12))
            
            # إضافة معلومات الفاتورة
            invoice_info = []
            invoice_info.append([Paragraph("<b>رقم الفاتورة:</b>", styles['Normal']), Paragraph(invoice[0], styles['Normal'])])
            
            date_str = datetime.strptime(invoice[1], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M")
            invoice_info.append([Paragraph("<b>التاريخ:</b>", styles['Normal']), Paragraph(date_str, styles['Normal'])])
            invoice_info.append([Paragraph("<b>اسم العميل:</b>", styles['Normal']), Paragraph(invoice[2] or "", styles['Normal'])])
            invoice_info.append([Paragraph("<b>رقم الهاتف:</b>", styles['Normal']), Paragraph(invoice[3] or "", styles['Normal'])])
            invoice_info.append([Paragraph("<b>طريقة الدفع:</b>", styles['Normal']), Paragraph(invoice[4], styles['Normal'])])
            
            invoice_table = Table(invoice_info, colWidths=[1.5*inch, 3*inch])
            invoice_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            story.append(invoice_table)
            story.append(Spacer(1, 12))
            
            # إضافة جدول المنتجات
            items_data = [['#', 'الباركود', 'اسم المنتج', 'الكمية', 'السعر', 'الإجمالي']]
            
            for i, item in enumerate(items, 1):
                items_data.append([
                    str(i),
                    item[3],
                    item[4],
                    f"{item[0]:.2f}",
                    f"{item[1]:.2f} {self.parent().currency_symbol}",
                    f"{item[2]:.2f} {self.parent().currency_symbol}"
                ])
            
            items_table = Table(items_data, colWidths=[0.5*inch, 1*inch, 2*inch, 0.8*inch, 1*inch, 1*inch])
            items_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ]))
            story.append(items_table)
            story.append(Spacer(1, 12))
            
            # إضافة ملخص الفاتورة
            summary_data = [
                [Paragraph("<b>المجموع:</b>", styles['Normal']), Paragraph(f"{invoice[6]:.2f} {self.parent().currency_symbol}", styles['Normal'])],
                [Paragraph("<b>الضريبة:</b>", styles['Normal']), Paragraph(f"{invoice[7]:.2f} {self.parent().currency_symbol}", styles['Normal'])],
                [Paragraph("<b>الخصم:</b>", styles['Normal']), Paragraph(f"{invoice[8]:.2f} {self.parent().currency_symbol}", styles['Normal'])],
                [Paragraph("<b>الإجمالي:</b>", styles['Normal']), Paragraph(f"<b>{invoice[9]:.2f} {self.parent().currency_symbol}</b>", styles['Normal'])]
            ]
            
            summary_table = Table(summary_data, colWidths=[1.5*inch, 1.5*inch])
            summary_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, -1), 'Arial'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('LINEBELOW', (0, -1), (-1, -1), 1, colors.black),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # إضافة الملاحظات
            if invoice[10]:
                story.append(Paragraph("<b>ملاحظات:</b>", styles['Normal']))
                story.append(Paragraph(invoice[10], styles['Normal']))
                story.append(Spacer(1, 12))
            
            # إضافة تذييل الفاتورة
            footer_text = get_setting("invoice_footer", "شكراً لتعاملكم معنا")
            story.append(Paragraph(footer_text, styles['Normal']))
            
            # بناء ملف PDF
            doc.build(story)
            
            # فتح الملف
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            
            add_log("طباعة فاتورة", f"تم طباعة الفاتورة: {self.invoice_number}", self.parent().user_id)
        
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء طباعة الفاتورة: {str(e)}")
        
        finally:
            # حذف الملف المؤقت بعد فتحه
            try:
                os.unlink(file_path)
            except:
                pass

# حوار إضافة/تعديل العميل
class CustomerDialog(QDialog):
    def __init__(self, parent=None, customer=None):
        super().__init__(parent)
        self.customer = customer
        self.customer_name = ""
        
        if customer:
            self.setWindowTitle("تعديل عميل")
        else:
            self.setWindowTitle("إضافة عميل")
        
        self.setFixedSize(400, 400)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QTextEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # الاسم
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("أدخل اسم العميل")
        form_layout.addRow("الاسم:", self.name_input)
        
        # رقم الهاتف
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("أدخل رقم الهاتف")
        form_layout.addRow("رقم الهاتف:", self.phone_input)
        
        # البريد الإلكتروني
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("أدخل البريد الإلكتروني")
        form_layout.addRow("البريد الإلكتروني:", self.email_input)
        
        # العنوان
        self.address_input = QTextEdit()
        self.address_input.setPlaceholderText("أدخل العنوان")
        form_layout.addRow("العنوان:", self.address_input)
        
        # ملاحظات
        self.notes_input = QTextEdit()
        self.notes_input.setPlaceholderText("أدخل ملاحظات إضافية")
        form_layout.addRow("ملاحظات:", self.notes_input)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_customer)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # ملء الحقول إذا كان تعديل عميل
        if customer:
            self.name_input.setText(customer[1])
            self.phone_input.setText(customer[2] or "")
            self.email_input.setText(customer[3] or "")
            self.address_input.setText(customer[4] or "")
            self.notes_input.setText(customer[5] or "")
        
        # التركيز على حقل الاسم
        self.name_input.setFocus()
    
    def save_customer(self):
        # الحصول على البيانات من الحقول
        name = self.name_input.text().strip()
        phone = self.phone_input.text().strip()
        email = self.email_input.text().strip()
        address = self.address_input.toPlainText().strip()
        notes = self.notes_input.toPlainText().strip()
        
        # التحقق من الحقول المطلوبة
        if not name:
            self.error_label.setText("يرجى إدخال اسم العميل")
            return
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            if self.customer:
                # تحديث عميل موجود
                cursor.execute(
                    """
                    UPDATE customers SET
                    name = ?, phone = ?, email = ?, address = ?, notes = ?
                    WHERE id = ?
                    """,
                    (name, phone, email, address, notes, self.customer[0])
                )
            else:
                # إضافة عميل جديد
                cursor.execute(
                    """
                    INSERT INTO customers (name, phone, email, address, notes)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (name, phone, email, address, notes)
                )
            
            conn.commit()
            conn.close()
            
            self.customer_name = name
            self.accept()
        
        except sqlite3.IntegrityError:
            conn.close()
            self.error_label.setText("العميل موجود بالفعل")
        except Exception as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# حوار إضافة/تعديل المورد
class SupplierDialog(QDialog):
    def __init__(self, parent=None, supplier=None):
        super().__init__(parent)
        self.supplier = supplier
        self.supplier_name = ""
        
        if supplier:
            self.setWindowTitle("تعديل مورد")
        else:
            self.setWindowTitle("إضافة مورد")
        
        self.setFixedSize(400, 400)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QTextEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # الاسم
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("أدخل اسم المورد")
        form_layout.addRow("الاسم:", self.name_input)
        
        # شخص الاتصال
        self.contact_person_input = QLineEdit()
        self.contact_person_input.setPlaceholderText("أدخل شخص الاتصال")
        form_layout.addRow("شخص الاتصال:", self.contact_person_input)
        
        # رقم الهاتف
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("أدخل رقم الهاتف")
        form_layout.addRow("رقم الهاتف:", self.phone_input)
        
        # البريد الإلكتروني
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("أدخل البريد الإلكتروني")
        form_layout.addRow("البريد الإلكتروني:", self.email_input)
        
        # العنوان
        self.address_input = QTextEdit()
        self.address_input.setPlaceholderText("أدخل العنوان")
        form_layout.addRow("العنوان:", self.address_input)
        
        # ملاحظات
        self.notes_input = QTextEdit()
        self.notes_input.setPlaceholderText("أدخل ملاحظات إضافية")
        form_layout.addRow("ملاحظات:", self.notes_input)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_supplier)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # ملء الحقول إذا كان تعديل مورد
        if supplier:
            self.name_input.setText(supplier[1])
            self.contact_person_input.setText(supplier[2] or "")
            self.phone_input.setText(supplier[3] or "")
            self.email_input.setText(supplier[4] or "")
            self.address_input.setText(supplier[5] or "")
            self.notes_input.setText(supplier[6] or "")
        
        # التركيز على حقل الاسم
        self.name_input.setFocus()
    
    def save_supplier(self):
        # الحصول على البيانات من الحقول
        name = self.name_input.text().strip()
        contact_person = self.contact_person_input.text().strip()
        phone = self.phone_input.text().strip()
        email = self.email_input.text().strip()
        address = self.address_input.toPlainText().strip()
        notes = self.notes_input.toPlainText().strip()
        
        # التحقق من الحقول المطلوبة
        if not name:
            self.error_label.setText("يرجى إدخال اسم المورد")
            return
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            if self.supplier:
                # تحديث مورد موجود
                cursor.execute(
                    """
                    UPDATE suppliers SET
                    name = ?, contact_person = ?, phone = ?, email = ?, address = ?, notes = ?
                    WHERE id = ?
                    """,
                    (name, contact_person, phone, email, address, notes, self.supplier[0])
                )
            else:
                # إضافة مورد جديد
                cursor.execute(
                    """
                    INSERT INTO suppliers (name, contact_person, phone, email, address, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (name, contact_person, phone, email, address, notes)
                )
            
            conn.commit()
            conn.close()
            
            self.supplier_name = name
            self.accept()
        
        except sqlite3.IntegrityError:
            conn.close()
            self.error_label.setText("المورد موجود بالفعل")
        except Exception as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# حوار إضافة/تعديل التصنيف
class CategoryDialog(QDialog):
    def __init__(self, parent=None, category=None):
        super().__init__(parent)
        self.category = category
        self.category_name = ""
        
        if category:
            self.setWindowTitle("تعديل تصنيف")
        else:
            self.setWindowTitle("إضافة تصنيف")
        
        self.setFixedSize(400, 300)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QTextEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # الاسم
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("أدخل اسم التصنيف")
        form_layout.addRow("الاسم:", self.name_input)
        
        # الوصف
        self.description_input = QTextEdit()
        self.description_input.setPlaceholderText("أدخل وصف التصنيف")
        form_layout.addRow("الوصف:", self.description_input)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_category)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # ملء الحقول إذا كان تعديل تصنيف
        if category:
            self.name_input.setText(category[1])
            self.description_input.setText(category[2] or "")
        
        # التركيز على حقل الاسم
        self.name_input.setFocus()
    
    def save_category(self):
        # الحصول على البيانات من الحقول
        name = self.name_input.text().strip()
        description = self.description_input.toPlainText().strip()
        
        # التحقق من الحقول المطلوبة
        if not name:
            self.error_label.setText("يرجى إدخال اسم التصنيف")
            return
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            if self.category:
                # تحديث تصنيف موجود
                cursor.execute(
                    """
                    UPDATE categories SET
                    name = ?, description = ?
                    WHERE id = ?
                    """,
                    (name, description, self.category[0])
                )
            else:
                # إضافة تصنيف جديد
                cursor.execute(
                    """
                    INSERT INTO categories (name, description)
                    VALUES (?, ?)
                    """,
                    (name, description)
                )
            
            conn.commit()
            conn.close()
            
            self.category_name = name
            self.accept()
        
        except sqlite3.IntegrityError:
            conn.close()
            self.error_label.setText("التصنيف موجود بالفعل")
        except Exception as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# حوار إضافة/تعديل المستخدم
class UserDialog(QDialog):
    def __init__(self, parent=None, user=None):
        super().__init__(parent)
        self.user = user
        self.username = ""
        
        if user:
            self.setWindowTitle("تعديل مستخدم")
        else:
            self.setWindowTitle("إضافة مستخدم")
        
        self.setFixedSize(400, 400)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QComboBox, QTextEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # اسم المستخدم
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("أدخل اسم المستخدم")
        form_layout.addRow("اسم المستخدم:", self.username_input)
        
        # كلمة المرور
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("أدخل كلمة المرور")
        self.password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("كلمة المرور:", self.password_input)
        
        # تأكيد كلمة المرور
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setPlaceholderText("أدخل كلمة المرور مرة أخرى")
        self.confirm_password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("تأكيد كلمة المرور:", self.confirm_password_input)
        
        # الاسم الكامل
        self.full_name_input = QLineEdit()
        self.full_name_input.setPlaceholderText("أدخل الاسم الكامل")
        form_layout.addRow("الاسم الكامل:", self.full_name_input)
        
        # البريد الإلكتروني
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("أدخل البريد الإلكتروني")
        form_layout.addRow("البريد الإلكتروني:", self.email_input)
        
        # رقم الهاتف
        self.phone_input = QLineEdit()
        self.phone_input.setPlaceholderText("أدخل رقم الهاتف")
        form_layout.addRow("رقم الهاتف:", self.phone_input)
        
        # الصلاحية
        self.role_combo = QComboBox()
        self.role_combo.addItems(["مستخدم عادي", "مدير"])
        form_layout.addRow("الصلاحية:", self.role_combo)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.save_user)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # ملء الحقول إذا كان تعديل مستخدم
        if user:
            self.username_input.setText(user[1])
            self.full_name_input.setText(user[3] or "")
            self.email_input.setText(user[4] or "")
            self.phone_input.setText(user[5] or "")
            self.role_combo.setCurrentIndex(0 if user[2] == "user" else 1)
        
        # التركيز على حقل اسم المستخدم
        self.username_input.setFocus()
    
    def save_user(self):
        # الحصول على البيانات من الحقول
        username = self.username_input.text().strip()
        password = self.password_input.text()
        confirm_password = self.confirm_password_input.text()
        full_name = self.full_name_input.text().strip()
        email = self.email_input.text().strip()
        phone = self.phone_input.text().strip()
        role = "admin" if self.role_combo.currentIndex() == 1 else "user"
        
        # التحقق من الحقول المطلوبة
        if not username:
            self.error_label.setText("يرجى إدخال اسم المستخدم")
            return
        
        # إذا كان مستخدم جديد، تحقق من كلمة المرور
        if not self.user:
            if not password:
                self.error_label.setText("يرجى إدخال كلمة المرور")
                return
            
            if password != confirm_password:
                self.error_label.setText("كلمة المرور غير متطابقة")
                return
            
            if len(password) < 6:
                self.error_label.setText("كلمة المرور يجب أن تكون 6 أحرف على الأقل")
                return
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            if self.user:
                # تحديث مستخدم موجود
                update_fields = ["full_name = ?", "email = ?", "phone = ?", "role = ?"]
                params = [full_name, email, phone, role, self.user[0]]
                
                # تحديث كلمة المرور فقط إذا تم إدخالها
                if password:
                    hashed_password = hash_password(password)
                    update_fields.append("password = ?")
                    params.insert(-1, hashed_password)
                
                cursor.execute(
                    f"""
                    UPDATE users SET
                    {', '.join(update_fields)}
                    WHERE id = ?
                    """,
                    params
                )
            else:
                # إضافة مستخدم جديد
                hashed_password = hash_password(password)
                cursor.execute(
                    """
                    INSERT INTO users (username, password, role, full_name, email, phone)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (username, hashed_password, role, full_name, email, phone)
                )
            
            conn.commit()
            conn.close()
            
            self.username = username
            self.accept()
        
        except sqlite3.IntegrityError:
            conn.close()
            self.error_label.setText("اسم المستخدم موجود بالفعل")
        except Exception as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# حوار تغيير كلمة المرور
class ChangePasswordDialog(QDialog):
    def __init__(self, parent=None, user_id=None):
        super().__init__(parent)
        self.user_id = user_id
        
        self.setWindowTitle("تغيير كلمة المرور")
        self.setFixedSize(400, 300)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#cancel {
                background-color: #f44336;
            }
            QPushButton#cancel:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # كلمة المرور الجديدة
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("أدخل كلمة المرور الجديدة")
        self.password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("كلمة المرور الجديدة:", self.password_input)
        
        # تأكيد كلمة المرور
        self.confirm_password_input = QLineEdit()
        self.confirm_password_input.setPlaceholderText("أدخل كلمة المرور مرة أخرى")
        self.confirm_password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("تأكيد كلمة المرور:", self.confirm_password_input)
        
        layout.addLayout(form_layout)
        
        # أزرار الحفظ والإلغاء
        buttons_layout = QHBoxLayout()
        
        save_button = QPushButton("حفظ")
        save_button.clicked.connect(self.change_password)
        buttons_layout.addWidget(save_button)
        
        cancel_button = QPushButton("إلغاء")
        cancel_button.setObjectName("cancel")
        cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_button)
        
        layout.addLayout(buttons_layout)
        
        # رسالة الخطأ
        self.error_label = QLabel()
        self.error_label.setAlignment(Qt.AlignCenter)
        self.error_label.setStyleSheet("color: red; margin-top: 10px;")
        layout.addWidget(self.error_label)
        
        self.setLayout(layout)
        
        # التركيز على حقل كلمة المرور
        self.password_input.setFocus()
    
    def change_password(self):
        # الحصول على البيانات من الحقول
        password = self.password_input.text()
        confirm_password = self.confirm_password_input.text()
        
        # التحقق من الحقول المطلوبة
        if not password:
            self.error_label.setText("يرجى إدخال كلمة المرور")
            return
        
        if password != confirm_password:
            self.error_label.setText("كلمة المرور غير متطابقة")
            return
        
        if len(password) < 6:
            self.error_label.setText("كلمة المرور يجب أن تكون 6 أحرف على الأقل")
            return
        
        try:
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            
            # تحديث كلمة المرور
            hashed_password = hash_password(password)
            cursor.execute(
                "UPDATE users SET password = ? WHERE id = ?",
                (hashed_password, self.user_id)
            )
            
            conn.commit()
            conn.close()
            
            self.accept()
        
        except Exception as e:
            conn.close()
            self.error_label.setText(f"خطأ في قاعدة البيانات: {str(e)}")

# حوار البحث العام
class SearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("بحث")
        self.setFixedSize(600, 400)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QLabel {
                color: #333;
                font-size: 14px;
            }
            QLineEdit, QComboBox {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9f9f9;
                gridline-color: #ddd;
                color: #333;
                border: 1px solid #ddd;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item:selected {
                background-color: #2196F3;
                color: white;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #333;
                padding: 5px;
                border: 1px solid #ddd;
                font-weight: bold;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1e88e5;
            }
            QPushButton#close {
                background-color: #f44336;
            }
            QPushButton#close:hover {
                background-color: #d32f2f;
            }
        """)
        
        layout = QVBoxLayout()
        
        # إعدادات البحث
        search_settings_layout = QHBoxLayout()
        
        # نوع البحث
        type_label = QLabel("نوع البحث:")
        search_settings_layout.addWidget(type_label)
        
        self.search_type_combo = QComboBox()
        self.search_type_combo.addItems(["المنتجات", "العملاء", "الموردين", "الفواتير"])
        self.search_type_combo.currentIndexChanged.connect(self.change_search_type)
        search_settings_layout.addWidget(self.search_type_combo)
        
        search_settings_layout.addStretch()
        
        layout.addLayout(search_settings_layout)
        
        # حقل البحث
        search_layout = QHBoxLayout()
        
        search_label = QLabel("بحث:")
        search_layout.addWidget(search_label)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("أدخل نص البحث...")
        self.search_input.textChanged.connect(self.perform_search)
        search_layout.addWidget(self.search_input)
        
        layout.addLayout(search_layout)
        
        # جدول النتائج
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(5)
        self.results_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.results_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.results_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.results_table.setSortingEnabled(True)
        self.results_table.doubleClicked.connect(self.open_result)
        layout.addWidget(self.results_table)
        
        # أزرار الإغلاق
        buttons_layout = QHBoxLayout()
        
        close_button = QPushButton("إغلاق")
        close_button.setObjectName("close")
        close_button.clicked.connect(self.accept)
        buttons_layout.addWidget(close_button)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
        
        # تحميل النتائج الأولية
        self.change_search_type()
        
        # التركيز على حقل البحث
        self.search_input.setFocus()
    
    def change_search_type(self):
        search_type = self.search_type_combo.currentText()
        
        if search_type == "المنتجات":
            self.results_table.setHorizontalHeaderLabels([
                "الباركود", "اسم المنتج", "التصنيف", "الكمية", "السعر"
            ])
            self.load_products()
        elif search_type == "العملاء":
            self.results_table.setHorizontalHeaderLabels([
                "الاسم", "رقم الهاتف", "البريد الإلكتروني", "العنوان", "ملاحظات"
            ])
            self.load_customers()
        elif search_type == "الموردين":
            self.results_table.setHorizontalHeaderLabels([
                "الاسم", "شخص الاتصال", "رقم الهاتف", "البريد الإلكتروني", "العنوان"
            ])
            self.load_suppliers()
        elif search_type == "الفواتير":
            self.results_table.setHorizontalHeaderLabels([
                "رقم الفاتورة", "اسم العميل", "التاريخ", "المبلغ", "طريقة الدفع"
            ])
            self.load_invoices()
    
    def load_products(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products ORDER BY name")
        products = cursor.fetchall()
        conn.close()
        
        self.results_table.setRowCount(len(products))
        
        for row, product in enumerate(products):
            # الباركود
            self.results_table.setItem(row, 0, QTableWidgetItem(product[1] or ""))
            
            # اسم المنتج
            self.results_table.setItem(row, 1, QTableWidgetItem(product[2]))
            
            # التصنيف
            self.results_table.setItem(row, 2, QTableWidgetItem(product[4] or ""))
            
            # الكمية
            self.results_table.setItem(row, 3, QTableWidgetItem(str(product[5])))
            
            # السعر
            self.results_table.setItem(row, 4, QTableWidgetItem(f"{product[6]:.2f} {self.parent().currency_symbol}"))
    
    def load_customers(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers ORDER BY name")
        customers = cursor.fetchall()
        conn.close()
        
        self.results_table.setRowCount(len(customers))
        
        for row, customer in enumerate(customers):
            # الاسم
            self.results_table.setItem(row, 0, QTableWidgetItem(customer[1]))
            
            # رقم الهاتف
            self.results_table.setItem(row, 1, QTableWidgetItem(customer[2] or ""))
            
            # البريد الإلكتروني
            self.results_table.setItem(row, 2, QTableWidgetItem(customer[3] or ""))
            
            # العنوان
            self.results_table.setItem(row, 3, QTableWidgetItem(customer[4] or ""))
            
            # ملاحظات
            self.results_table.setItem(row, 4, QTableWidgetItem(customer[5] or ""))
    
    def load_suppliers(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM suppliers ORDER BY name")
        suppliers = cursor.fetchall()
        conn.close()
        
        self.results_table.setRowCount(len(suppliers))
        
        for row, supplier in enumerate(suppliers):
            # الاسم
            self.results_table.setItem(row, 0, QTableWidgetItem(supplier[1]))
            
            # شخص الاتصال
            self.results_table.setItem(row, 1, QTableWidgetItem(supplier[2] or ""))
            
            # رقم الهاتف
            self.results_table.setItem(row, 2, QTableWidgetItem(supplier[3] or ""))
            
            # البريد الإلكتروني
            self.results_table.setItem(row, 3, QTableWidgetItem(supplier[4] or ""))
            
            # العنوان
            self.results_table.setItem(row, 4, QTableWidgetItem(supplier[5] or ""))
    
    def load_invoices(self):
        conn = sqlite3.connect('xnonedbs.db')
        cursor = conn.cursor()
        cursor.execute("""
            SELECT i.invoice_number, i.customer_name, i.date, i.total, i.payment_method
            FROM invoices i
            ORDER BY i.date DESC
        """)
        invoices = cursor.fetchall()
        conn.close()
        
        self.results_table.setRowCount(len(invoices))
        
        for row, invoice in enumerate(invoices):
            # رقم الفاتورة
            self.results_table.setItem(row, 0, QTableWidgetItem(invoice[0]))
            
            # اسم العميل
            self.results_table.setItem(row, 1, QTableWidgetItem(invoice[1] or ""))
            
            # التاريخ
            date_str = datetime.strptime(invoice[2], "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M")
            self.results_table.setItem(row, 2, QTableWidgetItem(date_str))
            
            # المبلغ
            self.results_table.setItem(row, 3, QTableWidgetItem(f"{invoice[3]:.2f} {self.parent().currency_symbol}"))
            
            # طريقة الدفع
            self.results_table.setItem(row, 4, QTableWidgetItem(invoice[4]))
    
    def perform_search(self):
        search_text = self.search_input.text().lower()
        
        for row in range(self.results_table.rowCount()):
            match = False
            
            for col in range(self.results_table.columnCount()):
                item = self.results_table.item(row, col)
                if item and search_text in item.text().lower():
                    match = True
                    break
            
            self.results_table.setRowHidden(row, not match)
    
    def open_result(self):
        selected_rows = self.results_table.selectedItems()
        if not selected_rows:
            return
        
        row = selected_rows[0].row()
        search_type = self.search_type_combo.currentText()
        
        if search_type == "المنتجات":
            # الانتقال إلى تبويب المنتجات وفتح حوار التعديل
            self.parent().tab_widget.setCurrentWidget(self.parent().products_tab)
            
            barcode = self.results_table.item(row, 0).text()
            
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM products WHERE barcode = ?", (barcode,))
            product = cursor.fetchone()
            conn.close()
            
            if product:
                dialog = ProductDialog(self.parent(), product)
                dialog.exec_()
        
        elif search_type == "العملاء":
            # الانتقال إلى تبويب العملاء وفتح حوار التعديل
            self.parent().tab_widget.setCurrentWidget(self.parent().customers_tab)
            
            customer_name = self.results_table.item(row, 0).text()
            
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM customers WHERE name = ?", (customer_name,))
            customer = cursor.fetchone()
            conn.close()
            
            if customer:
                dialog = CustomerDialog(self.parent(), customer)
                dialog.exec_()
        
        elif search_type == "الموردين":
            # الانتقال إلى تبويب الموردين وفتح حوار التعديل
            self.parent().tab_widget.setCurrentWidget(self.parent().suppliers_tab)
            
            supplier_name = self.results_table.item(row, 0).text()
            
            conn = sqlite3.connect('xnonedbs.db')
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM suppliers WHERE name = ?", (supplier_name,))
            supplier = cursor.fetchone()
            conn.close()
            
            if supplier:
                dialog = SupplierDialog(self.parent(), supplier)
                dialog.exec_()
        
        elif search_type == "الفواتير":
            # فتح حوار عرض الفاتورة
            invoice_number = self.results_table.item(row, 0).text()
            dialog = InvoiceViewDialog(self.parent(), invoice_number)
            dialog.exec_()
        
        self.accept()

# نقطة الدخول الرئيسية
if __name__ == "__main__":
    # إنشاء قاعدة البيانات إذا لم تكن موجودة
    create_database()
    
    # إنشاء التطبيق
    app = QApplication(sys.argv)
    
    # إعداد الخطوط العربية
    QFontDatabase.addApplicationFont("fonts/NotoSansArabic-Regular.ttf")
    app.setFont(QFont("Noto Sans Arabic", 10))
    
    # إعداد اتجاه الواجهة من اليمين لليسار
    app.setLayoutDirection(Qt.RightToLeft)
    
    # إنشاء نافذة تسجيل الدخول
    login_window = LoginWindow()
    
    if login_window.exec_() == QDialog.Accepted:
        # إنشاء النافذة الرئيسية
        main_window = MainWindow(
            login_window.user_id,
            login_window.username,
            login_window.role,
            login_window.full_name
        )
        main_window.show()
        sys.exit(app.exec_())
    else:
        sys.exit(0)