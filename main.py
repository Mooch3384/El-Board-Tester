"""
Board Tester Pro - Desktop Application
Electrical board testing software with Arduino

Author: [Your Name]
Version: 3.0
"""

import sys
import json
import time
import sqlite3
import os
from datetime import datetime
from collections import deque
from typing import Optional, List, Dict
import csv

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

import serial
import serial.tools.list_ports
import pyqtgraph as pg
from pyqtgraph import PlotWidget
import numpy as np


# --------------------------------
# Stylesheet for dark theme
# --------------------------------
STYLESHEET = """
QMainWindow, QWidget {
    background-color: #1a1a2e;
    color: #eee;
    font-family: 'Segoe UI', Arial;
    font-size: 11px;
}
QGroupBox {
    border: 2px solid #0f3460;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 10px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 15px;
    padding: 0 8px;
    color: #00d9ff;
}
QPushButton {
    background-color: #0f3460;
    border: none;
    border-radius: 6px;
    padding: 10px 20px;
    color: white;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #1a5490;
}
QPushButton:pressed {
    background-color: #0a2540;
}
QPushButton:disabled {
    background-color: #333;
    color: #666;
}
QComboBox, QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox {
    background-color: #16213e;
    border: 2px solid #0f3460;
    border-radius: 5px;
    padding: 8px;
}
QTabWidget::pane {
    border: 2px solid #0f3460;
    border-radius: 8px;
}
QTabBar::tab {
    background-color: #16213e;
    border: 2px solid #0f3460;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 10px 20px;
    margin-right: 2px;
}
QTabBar::tab:selected {
    background-color: #0f3460;
    color: #00d9ff;
}
QTableWidget {
    background-color: #16213e;
    border: 2px solid #0f3460;
    border-radius: 5px;
    gridline-color: #0f3460;
}
QHeaderView::section {
    background-color: #0f3460;
    padding: 8px;
    border: none;
    font-weight: bold;
}
QScrollBar:vertical {
    background-color: #16213e;
    width: 12px;
}
QScrollBar::handle:vertical {
    background-color: #0f3460;
    border-radius: 6px;
}
QStatusBar {
    background-color: #0f3460;
    color: #00d9ff;
}
QListWidget {
    background-color: #16213e;
    border: 2px solid #0f3460;
    border-radius: 5px;
}
QListWidget::item {
    padding: 10px;
    border-radius: 4px;
    margin: 2px;
}
QListWidget::item:selected {
    background-color: #0f3460;
}
"""


# --------------------------------
# Database handler
# --------------------------------
class Database:
    def __init__(self, path="test_records.db"):
        self.path = path
        self._init_tables()
    
    def _init_tables(self):
        conn = sqlite3.connect(self.path)
        c = conn.cursor()
        
        c.execute('''CREATE TABLE IF NOT EXISTS tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            board TEXT,
            serial_num TEXT,
            operator TEXT,
            start_time TEXT,
            end_time TEXT,
            duration REAL,
            status TEXT,
            v_min REAL, v_max REAL, v_avg REAL,
            i_min REAL, i_max REAL, i_avg REAL,
            p_min REAL, p_max REAL, p_avg REAL,
            f_min REAL, f_max REAL, f_avg REAL,
            v_violations INTEGER,
            i_violations INTEGER,
            f_violations INTEGER,
            notes TEXT,
            raw_data TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )''')
        
        c.execute('''CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            board_type TEXT,
            v_min REAL, v_max REAL,
            i_min REAL, i_max REAL,
            f_min REAL, f_max REAL,
            description TEXT
        )''')
        
        conn.commit()
        conn.close()
    
    def save_test(self, data: dict) -> int:
        conn = sqlite3.connect(self.path)
        c = conn.cursor()
        
        c.execute('''INSERT INTO tests 
            (name, board, serial_num, operator, start_time, end_time, duration, status,
             v_min, v_max, v_avg, i_min, i_max, i_avg, p_min, p_max, p_avg,
             f_min, f_max, f_avg, v_violations, i_violations, f_violations, notes, raw_data)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (data['name'], data['board'], data['serial_num'], data['operator'],
             data['start_time'], data['end_time'], data['duration'], data['status'],
             data['v_min'], data['v_max'], data['v_avg'],
             data['i_min'], data['i_max'], data['i_avg'],
             data['p_min'], data['p_max'], data['p_avg'],
             data['f_min'], data['f_max'], data['f_avg'],
             data['v_violations'], data['i_violations'], data['f_violations'],
             data['notes'], data['raw_data']))
        
        test_id = c.lastrowid
        conn.commit()
        conn.close()
        return test_id
    
    def get_all_tests(self) -> List[Dict]:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM tests ORDER BY id DESC')
        rows = c.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    
    def get_test(self, test_id: int) -> Optional[Dict]:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM tests WHERE id = ?', (test_id,))
        row = c.fetchone()
        conn.close()
        return dict(row) if row else None
    
    def delete_test(self, test_id: int):
        conn = sqlite3.connect(self.path)
        c = conn.cursor()
        c.execute('DELETE FROM tests WHERE id = ?', (test_id,))
        conn.commit()
        conn.close()
    
    def search(self, query: str) -> List[Dict]:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        q = f'%{query}%'
        c.execute('''SELECT * FROM tests 
                     WHERE board LIKE ? OR serial_num LIKE ? OR name LIKE ? OR operator LIKE ?
                     ORDER BY id DESC''', (q, q, q, q))
        rows = c.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    
    def filter_by_status(self, status: str) -> List[Dict]:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM tests WHERE status = ? ORDER BY id DESC', (status,))
        rows = c.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    
    def get_stats(self) -> Dict:
        conn = sqlite3.connect(self.path)
        c = conn.cursor()
        
        c.execute('SELECT COUNT(*) FROM tests')
        total = c.fetchone()[0]
        
        c.execute('SELECT COUNT(*) FROM tests WHERE status = "PASS"')
        passed = c.fetchone()[0]
        
        c.execute('SELECT COUNT(*) FROM tests WHERE status = "FAIL"')
        failed = c.fetchone()[0]
        
        conn.close()
        
        rate = (passed / total * 100) if total > 0 else 0
        return {'total': total, 'passed': passed, 'failed': failed, 'pass_rate': rate}
    
    # template stuff
    def save_template(self, t: dict):
        conn = sqlite3.connect(self.path)
        c = conn.cursor()
        c.execute('''INSERT OR REPLACE INTO templates 
                     (name, board_type, v_min, v_max, i_min, i_max, f_min, f_max, description)
                     VALUES (?,?,?,?,?,?,?,?,?)''',
                  (t['name'], t['board_type'], t['v_min'], t['v_max'],
                   t['i_min'], t['i_max'], t['f_min'], t['f_max'], t['description']))
        conn.commit()
        conn.close()
    
    def get_templates(self) -> List[Dict]:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute('SELECT * FROM templates ORDER BY name')
        rows = c.fetchall()
        conn.close()
        return [dict(r) for r in rows]
    
    def delete_template(self, tid: int):
        conn = sqlite3.connect(self.path)
        c = conn.cursor()
        c.execute('DELETE FROM templates WHERE id = ?', (tid,))
        conn.commit()
        conn.close()


# --------------------------------
# Serial communication thread
# --------------------------------
class SerialWorker(QThread):
    data_received = pyqtSignal(dict)
    status_changed = pyqtSignal(bool, str)
    error = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.port = None
        self.baud = 115200
        self.running = False
        self.ser = None
    
    def connect_to(self, port, baud=115200):
        self.port = port
        self.baud = baud
        self.running = True
        self.start()
    
    def disconnect(self):
        self.running = False
        self.wait(1000)
        if self.ser and self.ser.is_open:
            self.ser.close()
        self.status_changed.emit(False, "Disconnected")
    
    def send(self, cmd):
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(f"{cmd}\n".encode())
            except Exception as e:
                self.error.emit(str(e))
    
    def run(self):
        try:
            self.ser = serial.Serial(self.port, self.baud, timeout=0.1)
            time.sleep(2)  # arduino reset delay
            self.status_changed.emit(True, f"Connected: {self.port}")
            
            buf = ""
            while self.running:
                if self.ser.in_waiting:
                    try:
                        data = self.ser.read(self.ser.in_waiting).decode('utf-8', errors='ignore')
                        buf += data
                        
                        while '\n' in buf:
                            line, buf = buf.split('\n', 1)
                            line = line.strip()
                            if line.startswith('{') and line.endswith('}'):
                                try:
                                    self.data_received.emit(json.loads(line))
                                except:
                                    pass
                    except Exception as e:
                        self.error.emit(str(e))
                else:
                    time.sleep(0.01)
        except serial.SerialException as e:
            self.status_changed.emit(False, f"Failed: {e}")
            self.error.emit(str(e))


# --------------------------------
# Custom meter widget
# --------------------------------
class Meter(QFrame):
    def __init__(self, title, unit, color="#00d9ff"):
        super().__init__()
        self.color = color
        self.min_th = None
        self.max_th = None
        self.value = 0
        self.violation = False
        
        self.setStyleSheet(f"""
            Meter {{
                background-color: #16213e;
                border: 2px solid {color};
                border-radius: 10px;
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(5)
        
        self.title_lbl = QLabel(title)
        self.title_lbl.setAlignment(Qt.AlignCenter)
        self.title_lbl.setStyleSheet(f"color: {color}; font-size: 12px; font-weight: bold;")
        layout.addWidget(self.title_lbl)
        
        self.value_lbl = QLabel("---")
        self.value_lbl.setAlignment(Qt.AlignCenter)
        self.value_lbl.setStyleSheet("font-size: 28px; font-weight: bold; font-family: Consolas;")
        layout.addWidget(self.value_lbl)
        
        self.unit_lbl = QLabel(unit)
        self.unit_lbl.setAlignment(Qt.AlignCenter)
        self.unit_lbl.setStyleSheet(f"color: {color}; font-size: 14px;")
        layout.addWidget(self.unit_lbl)
        
        self.status_lbl = QLabel("NORMAL")
        self.status_lbl.setAlignment(Qt.AlignCenter)
        self.status_lbl.setStyleSheet("""
            background-color: #00a86b; color: white;
            padding: 3px 10px; border-radius: 3px;
            font-size: 10px; font-weight: bold;
        """)
        layout.addWidget(self.status_lbl)
    
    def set_value(self, val, decimals=3):
        self.value = val
        self.value_lbl.setText(f"{val:.{decimals}f}")
        self._check()
    
    def set_thresholds(self, min_v, max_v):
        self.min_th = min_v
        self.max_th = max_v
    
    def _check(self):
        if self.min_th is None or self.max_th is None:
            self.violation = False
            return
        
        if self.value < self.min_th:
            self.violation = True
            self.status_lbl.setText("LOW")
            self.status_lbl.setStyleSheet("""
                background-color: #ffa500; color: black;
                padding: 3px 10px; border-radius: 3px;
                font-size: 10px; font-weight: bold;
            """)
        elif self.value > self.max_th:
            self.violation = True
            self.status_lbl.setText("HIGH")
            self.status_lbl.setStyleSheet("""
                background-color: #e94560; color: white;
                padding: 3px 10px; border-radius: 3px;
                font-size: 10px; font-weight: bold;
            """)
        else:
            self.violation = False
            self.status_lbl.setText("NORMAL")
            self.status_lbl.setStyleSheet("""
                background-color: #00a86b; color: white;
                padding: 3px 10px; border-radius: 3px;
                font-size: 10px; font-weight: bold;
            """)


# --------------------------------
# Status indicator
# --------------------------------
class StatusLight(QWidget):
    def __init__(self):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.dot = QLabel("●")
        self.dot.setStyleSheet("color: #e94560; font-size: 20px;")
        layout.addWidget(self.dot)
        
        self.text = QLabel("Disconnected")
        self.text.setStyleSheet("color: #888;")
        layout.addWidget(self.text)
    
    def set_connected(self, connected, msg=""):
        if connected:
            self.dot.setStyleSheet("color: #00a86b; font-size: 20px;")
            self.text.setStyleSheet("color: #00a86b;")
            self.text.setText(msg or "Connected")
        else:
            self.dot.setStyleSheet("color: #e94560; font-size: 20px;")
            self.text.setStyleSheet("color: #888;")
            self.text.setText(msg or "Disconnected")


# --------------------------------
# Test record card
# --------------------------------
class TestCard(QFrame):
    clicked = pyqtSignal(int)
    
    def __init__(self, record):
        super().__init__()
        self.record = record
        self.record_id = record['id']
        
        status = record.get('status', 'PENDING')
        colors = {'PASS': '#00a86b', 'FAIL': '#e94560', 'ABORTED': '#ffa500', 'PENDING': '#888'}
        color = colors.get(status, '#888')
        
        self.setStyleSheet(f"""
            TestCard {{
                background-color: #16213e;
                border: 2px solid {color};
                border-radius: 8px;
                padding: 10px;
            }}
            TestCard:hover {{
                background-color: #1a3a5c;
            }}
        """)
        self.setCursor(Qt.PointingHandCursor)
        
        layout = QVBoxLayout(self)
        
        # header row
        header = QHBoxLayout()
        title = QLabel(f"#{record['id']} - {record.get('name', 'Untitled')}")
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #00d9ff;")
        header.addWidget(title)
        header.addStretch()
        
        status_lbl = QLabel(status)
        status_lbl.setStyleSheet(f"background-color: {color}; color: white; padding: 3px 10px; border-radius: 3px; font-weight: bold;")
        header.addWidget(status_lbl)
        layout.addLayout(header)
        
        # info row
        info = QHBoxLayout()
        info.addWidget(QLabel(f"Board: {record.get('board', 'N/A')}"))
        info.addWidget(QLabel(f"Serial: {record.get('serial_num', 'N/A')}"))
        info.addWidget(QLabel(f"Duration: {record.get('duration', 0):.1f}s"))
        info.addStretch()
        layout.addLayout(info)
    
    def mousePressEvent(self, event):
        self.clicked.emit(self.record_id)


# --------------------------------
# Dialog: New Test
# --------------------------------
class NewTestDialog(QDialog):
    def __init__(self, parent=None, templates=None):
        super().__init__(parent)
        self.setWindowTitle("New Test")
        self.setMinimumWidth(450)
        self.templates = templates or []
        
        layout = QVBoxLayout(self)
        
        # info section
        info_grp = QGroupBox("Test Info")
        info_layout = QFormLayout(info_grp)
        
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Test name...")
        info_layout.addRow("Name:", self.name_edit)
        
        self.board_edit = QLineEdit()
        self.board_edit.setPlaceholderText("Board name...")
        info_layout.addRow("Board:", self.board_edit)
        
        self.serial_edit = QLineEdit()
        self.serial_edit.setPlaceholderText("Serial number...")
        info_layout.addRow("Serial:", self.serial_edit)
        
        self.operator_edit = QLineEdit()
        self.operator_edit.setPlaceholderText("Your name...")
        info_layout.addRow("Operator:", self.operator_edit)
        
        layout.addWidget(info_grp)
        
        # template
        if self.templates:
            tmpl_grp = QGroupBox("Template")
            tmpl_layout = QVBoxLayout(tmpl_grp)
            self.tmpl_combo = QComboBox()
            self.tmpl_combo.addItem("-- None --", None)
            for t in self.templates:
                self.tmpl_combo.addItem(t['name'], t)
            self.tmpl_combo.currentIndexChanged.connect(self._on_template)
            tmpl_layout.addWidget(self.tmpl_combo)
            layout.addWidget(tmpl_grp)
        
        # thresholds
        th_grp = QGroupBox("Thresholds")
        th_layout = QGridLayout(th_grp)
        
        th_layout.addWidget(QLabel("V Min:"), 0, 0)
        self.v_min = QDoubleSpinBox()
        self.v_min.setRange(0, 1000)
        th_layout.addWidget(self.v_min, 0, 1)
        
        th_layout.addWidget(QLabel("V Max:"), 0, 2)
        self.v_max = QDoubleSpinBox()
        self.v_max.setRange(0, 1000)
        self.v_max.setValue(50)
        th_layout.addWidget(self.v_max, 0, 3)
        
        th_layout.addWidget(QLabel("I Min:"), 1, 0)
        self.i_min = QDoubleSpinBox()
        self.i_min.setRange(0, 100)
        self.i_min.setDecimals(3)
        th_layout.addWidget(self.i_min, 1, 1)
        
        th_layout.addWidget(QLabel("I Max:"), 1, 2)
        self.i_max = QDoubleSpinBox()
        self.i_max.setRange(0, 100)
        self.i_max.setDecimals(3)
        self.i_max.setValue(5)
        th_layout.addWidget(self.i_max, 1, 3)
        
        th_layout.addWidget(QLabel("F Min:"), 2, 0)
        self.f_min = QDoubleSpinBox()
        self.f_min.setRange(0, 1e6)
        th_layout.addWidget(self.f_min, 2, 1)
        
        th_layout.addWidget(QLabel("F Max:"), 2, 2)
        self.f_max = QDoubleSpinBox()
        self.f_max.setRange(0, 1e6)
        self.f_max.setValue(100000)
        th_layout.addWidget(self.f_max, 2, 3)
        
        layout.addWidget(th_grp)
        
        # notes
        notes_grp = QGroupBox("Notes")
        notes_layout = QVBoxLayout(notes_grp)
        self.notes_edit = QTextEdit()
        self.notes_edit.setMaximumHeight(60)
        notes_layout.addWidget(self.notes_edit)
        layout.addWidget(notes_grp)
        
        # buttons
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
    
    def _on_template(self, idx):
        t = self.tmpl_combo.currentData()
        if t:
            self.v_min.setValue(t.get('v_min', 0))
            self.v_max.setValue(t.get('v_max', 50))
            self.i_min.setValue(t.get('i_min', 0))
            self.i_max.setValue(t.get('i_max', 5))
            self.f_min.setValue(t.get('f_min', 0))
            self.f_max.setValue(t.get('f_max', 100000))
    
    def get_data(self):
        return {
            'name': self.name_edit.text() or f"Test_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
            'board': self.board_edit.text(),
            'serial': self.serial_edit.text(),
            'operator': self.operator_edit.text(),
            'v_min': self.v_min.value(),
            'v_max': self.v_max.value(),
            'i_min': self.i_min.value(),
            'i_max': self.i_max.value(),
            'f_min': self.f_min.value(),
            'f_max': self.f_max.value(),
            'notes': self.notes_edit.toPlainText()
        }


# --------------------------------
# Dialog: Test Details
# --------------------------------
class TestDetailsDialog(QDialog):
    def __init__(self, record, parent=None):
        super().__init__(parent)
        self.record = record
        self.setWindowTitle(f"Test #{record['id']}")
        self.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(self)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        
        # header
        status = record.get('status', 'PENDING')
        colors = {'PASS': '#00a86b', 'FAIL': '#e94560', 'ABORTED': '#ffa500', 'PENDING': '#888'}
        
        header = QHBoxLayout()
        title = QLabel(f"Test #{record['id']}: {record.get('name', 'Untitled')}")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #00d9ff;")
        header.addWidget(title)
        header.addStretch()
        
        status_lbl = QLabel(status)
        status_lbl.setStyleSheet(f"background-color: {colors.get(status)}; color: white; padding: 8px 16px; border-radius: 5px; font-weight: bold;")
        header.addWidget(status_lbl)
        content_layout.addLayout(header)
        
        # info
        info_grp = QGroupBox("Info")
        info_layout = QGridLayout(info_grp)
        
        fields = [
            ("Board:", record.get('board', 'N/A')),
            ("Serial:", record.get('serial_num', 'N/A')),
            ("Operator:", record.get('operator', 'N/A')),
            ("Start:", record.get('start_time', 'N/A')[:19] if record.get('start_time') else 'N/A'),
            ("End:", record.get('end_time', 'N/A')[:19] if record.get('end_time') else 'N/A'),
            ("Duration:", f"{record.get('duration', 0):.1f}s"),
        ]
        
        for i, (lbl, val) in enumerate(fields):
            r, c = i // 2, (i % 2) * 2
            info_layout.addWidget(QLabel(lbl), r, c)
            v = QLabel(str(val))
            v.setStyleSheet("color: #00d9ff;")
            info_layout.addWidget(v, r, c + 1)
        
        content_layout.addWidget(info_grp)
        
        # measurements
        meas_grp = QGroupBox("Measurements")
        meas_layout = QGridLayout(meas_grp)
        
        headers = ["", "Min", "Max", "Avg", "Violations"]
        for c, h in enumerate(headers):
            lbl = QLabel(h)
            lbl.setStyleSheet("font-weight: bold;")
            meas_layout.addWidget(lbl, 0, c)
        
        params = [("Voltage", 'v'), ("Current", 'i'), ("Power", 'p'), ("Frequency", 'f')]
        for r, (name, key) in enumerate(params, 1):
            meas_layout.addWidget(QLabel(name), r, 0)
            meas_layout.addWidget(QLabel(f"{record.get(f'{key}_min', 0):.4f}"), r, 1)
            meas_layout.addWidget(QLabel(f"{record.get(f'{key}_max', 0):.4f}"), r, 2)
            meas_layout.addWidget(QLabel(f"{record.get(f'{key}_avg', 0):.4f}"), r, 3)
            
            viol = record.get(f'{key}_violations', 0) if key != 'p' else '-'
            viol_lbl = QLabel(str(viol))
            if isinstance(viol, int) and viol > 0:
                viol_lbl.setStyleSheet("color: #e94560; font-weight: bold;")
            meas_layout.addWidget(viol_lbl, r, 4)
        
        content_layout.addWidget(meas_grp)
        
        # notes
        if record.get('notes'):
            notes_grp = QGroupBox("Notes")
            notes_layout = QVBoxLayout(notes_grp)
            notes = QTextEdit()
            notes.setPlainText(record.get('notes', ''))
            notes.setReadOnly(True)
            notes.setMaximumHeight(80)
            notes_layout.addWidget(notes)
            content_layout.addWidget(notes_grp)
        
        # graph if we have data
        raw = record.get('raw_data')
        if raw and raw != '[]':
            try:
                data = json.loads(raw)
                if data:
                    graph_grp = QGroupBox("Graph")
                    graph_layout = QVBoxLayout(graph_grp)
                    
                    plot = PlotWidget()
                    plot.setBackground('#1a1a2e')
                    plot.showGrid(x=True, y=True, alpha=0.3)
                    plot.addLegend()
                    
                    times = [d.get('time', i) for i, d in enumerate(data)]
                    volts = [d.get('voltage', 0) for d in data]
                    amps = [d.get('current', 0) for d in data]
                    
                    plot.plot(times, volts, pen=pg.mkPen('#ff6b6b', width=2), name='V')
                    plot.plot(times, amps, pen=pg.mkPen('#4ecdc4', width=2), name='I')
                    plot.setMinimumHeight(250)
                    
                    graph_layout.addWidget(plot)
                    content_layout.addWidget(graph_grp)
            except:
                pass
        
        content_layout.addStretch()
        scroll.setWidget(content)
        layout.addWidget(scroll)
        
        # bottom buttons
        btn_layout = QHBoxLayout()
        
        export_btn = QPushButton("Export Report")
        export_btn.clicked.connect(self._export)
        btn_layout.addWidget(export_btn)
        
        btn_layout.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        
        layout.addLayout(btn_layout)
    
    def _export(self):
        fname, _ = QFileDialog.getSaveFileName(
            self, "Export",
            f"test_report_{self.record['id']}.txt",
            "Text Files (*.txt);;HTML Files (*.html)"
        )
        
        if not fname:
            return
        
        r = self.record
        
        if fname.endswith('.html'):
            status = r.get('status', 'PENDING')
            colors = {'PASS': '#00a86b', 'FAIL': '#e94560', 'ABORTED': '#ffa500', 'PENDING': '#888'}
            
            html = f"""<!DOCTYPE html>
<html>
<head><title>Test #{r['id']}</title>
<style>
body {{ font-family: Arial; margin: 40px; }}
.container {{ max-width: 800px; margin: auto; }}
h1 {{ color: #333; }}
.status {{ display: inline-block; padding: 8px 16px; border-radius: 5px; color: white; background: {colors.get(status)}; }}
table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
th, td {{ padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }}
th {{ background: #f5f5f5; }}
</style>
</head>
<body>
<div class="container">
<h1>Test Report #{r['id']}</h1>
<p><span class="status">{status}</span></p>
<h2>Info</h2>
<table>
<tr><th>Name</th><td>{r.get('name', 'N/A')}</td></tr>
<tr><th>Board</th><td>{r.get('board', 'N/A')}</td></tr>
<tr><th>Serial</th><td>{r.get('serial_num', 'N/A')}</td></tr>
<tr><th>Operator</th><td>{r.get('operator', 'N/A')}</td></tr>
<tr><th>Start</th><td>{r.get('start_time', 'N/A')}</td></tr>
<tr><th>End</th><td>{r.get('end_time', 'N/A')}</td></tr>
<tr><th>Duration</th><td>{r.get('duration', 0):.2f}s</td></tr>
</table>
<h2>Measurements</h2>
<table>
<tr><th></th><th>Min</th><th>Max</th><th>Avg</th></tr>
<tr><td>Voltage</td><td>{r.get('v_min', 0):.4f}</td><td>{r.get('v_max', 0):.4f}</td><td>{r.get('v_avg', 0):.4f}</td></tr>
<tr><td>Current</td><td>{r.get('i_min', 0):.4f}</td><td>{r.get('i_max', 0):.4f}</td><td>{r.get('i_avg', 0):.4f}</td></tr>
<tr><td>Power</td><td>{r.get('p_min', 0):.4f}</td><td>{r.get('p_max', 0):.4f}</td><td>{r.get('p_avg', 0):.4f}</td></tr>
<tr><td>Frequency</td><td>{r.get('f_min', 0):.4f}</td><td>{r.get('f_max', 0):.4f}</td><td>{r.get('f_avg', 0):.4f}</td></tr>
</table>
<h2>Notes</h2>
<p>{r.get('notes', 'None')}</p>
<hr>
<p><small>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</small></p>
</div>
</body>
</html>"""
            
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(html)
        else:
            txt = f"""TEST REPORT #{r['id']}
{'='*40}
Status: {r.get('status', 'PENDING')}

INFO
----
Name: {r.get('name', 'N/A')}
Board: {r.get('board', 'N/A')}
Serial: {r.get('serial_num', 'N/A')}
Operator: {r.get('operator', 'N/A')}
Start: {r.get('start_time', 'N/A')}
End: {r.get('end_time', 'N/A')}
Duration: {r.get('duration', 0):.2f}s

MEASUREMENTS
------------
Voltage: {r.get('v_min', 0):.4f} - {r.get('v_max', 0):.4f} (avg: {r.get('v_avg', 0):.4f})
Current: {r.get('i_min', 0):.4f} - {r.get('i_max', 0):.4f} (avg: {r.get('i_avg', 0):.4f})
Power: {r.get('p_min', 0):.4f} - {r.get('p_max', 0):.4f} (avg: {r.get('p_avg', 0):.4f})
Frequency: {r.get('f_min', 0):.4f} - {r.get('f_max', 0):.4f} (avg: {r.get('f_avg', 0):.4f})

NOTES
-----
{r.get('notes', 'None')}

Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
"""
            with open(fname, 'w', encoding='utf-8') as f:
                f.write(txt)
        
        QMessageBox.information(self, "Done", f"Saved to {fname}")


# --------------------------------
# Dialog: Templates
# --------------------------------
class TemplatesDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Templates")
        self.setMinimumSize(500, 400)
        
        layout = QVBoxLayout(self)
        
        self.list = QListWidget()
        self.list.itemClicked.connect(self._on_select)
        layout.addWidget(self.list)
        
        form_grp = QGroupBox("Template")
        form = QFormLayout(form_grp)
        
        self.name_edit = QLineEdit()
        form.addRow("Name:", self.name_edit)
        
        self.type_edit = QLineEdit()
        form.addRow("Board Type:", self.type_edit)
        
        th_widget = QWidget()
        th_layout = QGridLayout(th_widget)
        th_layout.setContentsMargins(0, 0, 0, 0)
        
        self.v_min = QDoubleSpinBox()
        self.v_min.setRange(0, 1000)
        th_layout.addWidget(QLabel("V:"), 0, 0)
        th_layout.addWidget(self.v_min, 0, 1)
        
        self.v_max = QDoubleSpinBox()
        self.v_max.setRange(0, 1000)
        self.v_max.setValue(50)
        th_layout.addWidget(QLabel("-"), 0, 2)
        th_layout.addWidget(self.v_max, 0, 3)
        
        self.i_min = QDoubleSpinBox()
        self.i_min.setRange(0, 100)
        self.i_min.setDecimals(3)
        th_layout.addWidget(QLabel("I:"), 1, 0)
        th_layout.addWidget(self.i_min, 1, 1)
        
        self.i_max = QDoubleSpinBox()
        self.i_max.setRange(0, 100)
        self.i_max.setDecimals(3)
        self.i_max.setValue(5)
        th_layout.addWidget(QLabel("-"), 1, 2)
        th_layout.addWidget(self.i_max, 1, 3)
        
        form.addRow("Thresholds:", th_widget)
        
        self.desc_edit = QTextEdit()
        self.desc_edit.setMaximumHeight(50)
        form.addRow("Description:", self.desc_edit)
        
        layout.addWidget(form_grp)
        
        btns = QHBoxLayout()
        
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self._save)
        btns.addWidget(save_btn)
        
        del_btn = QPushButton("Delete")
        del_btn.clicked.connect(self._delete)
        btns.addWidget(del_btn)
        
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self._clear)
        btns.addWidget(clear_btn)
        
        btns.addStretch()
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btns.addWidget(close_btn)
        
        layout.addLayout(btns)
        
        self._load()
    
    def _load(self):
        self.list.clear()
        for t in self.db.get_templates():
            item = QListWidgetItem(f"{t['name']} ({t['board_type']})")
            item.setData(Qt.UserRole, t)
            self.list.addItem(item)
    
    def _on_select(self, item):
        t = item.data(Qt.UserRole)
        self.name_edit.setText(t['name'])
        self.type_edit.setText(t.get('board_type', ''))
        self.v_min.setValue(t.get('v_min', 0))
        self.v_max.setValue(t.get('v_max', 50))
        self.i_min.setValue(t.get('i_min', 0))
        self.i_max.setValue(t.get('i_max', 5))
        self.desc_edit.setPlainText(t.get('description', ''))
    
    def _save(self):
        if not self.name_edit.text():
            QMessageBox.warning(self, "Error", "Enter a name")
            return
        
        self.db.save_template({
            'name': self.name_edit.text(),
            'board_type': self.type_edit.text(),
            'v_min': self.v_min.value(),
            'v_max': self.v_max.value(),
            'i_min': self.i_min.value(),
            'i_max': self.i_max.value(),
            'f_min': 0,
            'f_max': 100000,
            'description': self.desc_edit.toPlainText()
        })
        self._load()
        self._clear()
    
    def _delete(self):
        item = self.list.currentItem()
        if item:
            t = item.data(Qt.UserRole)
            if QMessageBox.question(self, "Delete", f"Delete '{t['name']}'?") == QMessageBox.Yes:
                self.db.delete_template(t['id'])
                self._load()
                self._clear()
    
    def _clear(self):
        self.name_edit.clear()
        self.type_edit.clear()
        self.v_min.setValue(0)
        self.v_max.setValue(50)
        self.i_min.setValue(0)
        self.i_max.setValue(5)
        self.desc_edit.clear()


# --------------------------------
# Main window
# --------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Board Tester Pro")
        self.setMinimumSize(1300, 850)
        
        self.db = Database()
        
        # data buffers
        self.buf_size = 500
        self.time_buf = deque(maxlen=self.buf_size)
        self.v_buf = deque(maxlen=self.buf_size)
        self.i_buf = deque(maxlen=self.buf_size)
        self.p_buf = deque(maxlen=self.buf_size)
        self.f_buf = deque(maxlen=self.buf_size)
        
        self.t0 = time.time()
        
        # test state
        self.testing = False
        self.test_data = []
        self.test_start = None
        self.test_info = {}
        
        # violation counts
        self.v_viols = 0
        self.i_viols = 0
        self.f_viols = 0
        
        # serial
        self.serial = SerialWorker()
        self.serial.data_received.connect(self._on_data)
        self.serial.status_changed.connect(self._on_status)
        self.serial.error.connect(self._on_error)
        
        self._setup_ui()
        self.setStyleSheet(STYLESHEET)
        
        # update timer
        self.timer = QTimer()
        self.timer.timeout.connect(self._update_plots)
        self.timer.start(50)
        
        # test timer
        self.test_timer = QTimer()
        self.test_timer.timeout.connect(self._update_duration)
    
    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        
        # top bar
        top = self._create_top_bar()
        main_layout.addWidget(top)
        
        # content
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._create_left_panel())
        splitter.addWidget(self._create_right_panel())
        splitter.setSizes([380, 900])
        main_layout.addWidget(splitter, 1)
        
        # status bar
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Ready")
    
    def _create_top_bar(self):
        bar = QWidget()
        layout = QHBoxLayout(bar)
        
        title = QLabel("Board Tester Pro")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #00d9ff;")
        layout.addWidget(title)
        
        layout.addStretch()
        
        layout.addWidget(QLabel("Port:"))
        self.port_cb = QComboBox()
        self.port_cb.setMinimumWidth(140)
        layout.addWidget(self.port_cb)
        
        refresh = QPushButton("↻")
        refresh.setMaximumWidth(35)
        refresh.clicked.connect(self._refresh_ports)
        layout.addWidget(refresh)
        
        layout.addWidget(QLabel("Baud:"))
        self.baud_cb = QComboBox()
        self.baud_cb.addItems(["9600", "19200", "57600", "115200"])
        self.baud_cb.setCurrentText("115200")
        layout.addWidget(self.baud_cb)
        
        self.conn_btn = QPushButton("Connect")
        self.conn_btn.clicked.connect(self._toggle_connection)
        layout.addWidget(self.conn_btn)
        
        self.status_light = StatusLight()
        layout.addWidget(self.status_light)
        
        self._refresh_ports()
        return bar
    
    def _create_left_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # test controls
        ctrl_grp = QGroupBox("Test Control")
        ctrl_layout = QVBoxLayout(ctrl_grp)
        
        self.test_label = QLabel("No active test")
        self.test_label.setAlignment(Qt.AlignCenter)
        self.test_label.setStyleSheet("color: #888;")
        ctrl_layout.addWidget(self.test_label)
        
        self.duration_label = QLabel("00:00:00")
        self.duration_label.setAlignment(Qt.AlignCenter)
        self.duration_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #00d9ff;")
        ctrl_layout.addWidget(self.duration_label)
        
        btns1 = QHBoxLayout()
        self.start_btn = QPushButton("Start Test")
        self.start_btn.clicked.connect(self._start_test)
        btns1.addWidget(self.start_btn)
        
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self._stop_test)
        self.stop_btn.setEnabled(False)
        btns1.addWidget(self.stop_btn)
        ctrl_layout.addLayout(btns1)
        
        btns2 = QHBoxLayout()
        self.pass_btn = QPushButton("PASS")
        self.pass_btn.setStyleSheet("background-color: #00a86b;")
        self.pass_btn.clicked.connect(lambda: self._finish_test("PASS"))
        self.pass_btn.setEnabled(False)
        btns2.addWidget(self.pass_btn)
        
        self.fail_btn = QPushButton("FAIL")
        self.fail_btn.setStyleSheet("background-color: #e94560;")
        self.fail_btn.clicked.connect(lambda: self._finish_test("FAIL"))
        self.fail_btn.setEnabled(False)
        btns2.addWidget(self.fail_btn)
        ctrl_layout.addLayout(btns2)
        
        layout.addWidget(ctrl_grp)
        
        # meters
        meters = QGridLayout()
        
        self.v_meter = Meter("VOLTAGE", "V", "#ff6b6b")
        meters.addWidget(self.v_meter, 0, 0)
        
        self.i_meter = Meter("CURRENT", "A", "#4ecdc4")
        meters.addWidget(self.i_meter, 0, 1)
        
        self.p_meter = Meter("POWER", "W", "#ffe66d")
        meters.addWidget(self.p_meter, 1, 0)
        
        self.r_meter = Meter("RESISTANCE", "Ω", "#95e1d3")
        meters.addWidget(self.r_meter, 1, 1)
        
        self.f_meter = Meter("FREQUENCY", "Hz", "#a388ee")
        meters.addWidget(self.f_meter, 2, 0)
        
        self.wl_meter = Meter("WAVELENGTH", "m", "#f9b4ab")
        meters.addWidget(self.wl_meter, 2, 1)
        
        self.vrms_meter = Meter("V RMS", "V", "#ff6b6b")
        meters.addWidget(self.vrms_meter, 3, 0)
        
        self.vpp_meter = Meter("V P-P", "V", "#ff6b6b")
        meters.addWidget(self.vpp_meter, 3, 1)
        
        layout.addLayout(meters)
        layout.addStretch()
        
        return panel
    
    def _create_right_panel(self):
        tabs = QTabWidget()
        tabs.addTab(self._create_graphs_tab(), "Graphs")
        tabs.addTab(self._create_records_tab(), "Records")
        tabs.addTab(self._create_thresholds_tab(), "Thresholds")
        tabs.addTab(self._create_stats_tab(), "Statistics")
        return tabs
    
    def _create_graphs_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        
        pg.setConfigOptions(antialias=True)
        
        # V/I plot
        self.vi_plot = PlotWidget()
        self.vi_plot.setBackground('#1a1a2e')
        self.vi_plot.showGrid(x=True, y=True, alpha=0.3)
        self.vi_plot.addLegend()
        self.v_curve = self.vi_plot.plot(pen=pg.mkPen('#ff6b6b', width=2), name='Voltage')
        self.i_curve = self.vi_plot.plot(pen=pg.mkPen('#4ecdc4', width=2), name='Current')
        layout.addWidget(QLabel("Voltage & Current"))
        layout.addWidget(self.vi_plot)
        
        # power plot
        self.p_plot = PlotWidget()
        self.p_plot.setBackground('#1a1a2e')
        self.p_plot.showGrid(x=True, y=True, alpha=0.3)
        self.p_curve = self.p_plot.plot(pen=pg.mkPen('#ffe66d', width=2))
        layout.addWidget(QLabel("Power"))
        layout.addWidget(self.p_plot)
        
        # freq plot
        self.f_plot = PlotWidget()
        self.f_plot.setBackground('#1a1a2e')
        self.f_plot.showGrid(x=True, y=True, alpha=0.3)
        self.f_curve = self.f_plot.plot(pen=pg.mkPen('#a388ee', width=2))
        layout.addWidget(QLabel("Frequency"))
        layout.addWidget(self.f_plot)
        
        return w
    
    def _create_records_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        
        # controls
        ctrls = QHBoxLayout()
        
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search...")
        self.search_edit.textChanged.connect(self._search_records)
        ctrls.addWidget(self.search_edit)
        
        self.filter_cb = QComboBox()
        self.filter_cb.addItems(["All", "PASS", "FAIL", "ABORTED"])
        self.filter_cb.currentTextChanged.connect(self._filter_records)
        ctrls.addWidget(self.filter_cb)
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._load_records)
        ctrls.addWidget(refresh_btn)
        
        tmpl_btn = QPushButton("Templates")
        tmpl_btn.clicked.connect(self._show_templates)
        ctrls.addWidget(tmpl_btn)
        
        export_btn = QPushButton("Export All")
        export_btn.clicked.connect(self._export_all)
        ctrls.addWidget(export_btn)
        
        layout.addLayout(ctrls)
        
        # stats row
        stats_row = QHBoxLayout()
        self.total_lbl = QLabel("Total: 0")
        self.total_lbl.setStyleSheet("font-weight: bold;")
        stats_row.addWidget(self.total_lbl)
        
        self.pass_lbl = QLabel("Passed: 0")
        self.pass_lbl.setStyleSheet("color: #00a86b; font-weight: bold;")
        stats_row.addWidget(self.pass_lbl)
        
        self.fail_lbl = QLabel("Failed: 0")
        self.fail_lbl.setStyleSheet("color: #e94560; font-weight: bold;")
        stats_row.addWidget(self.fail_lbl)
        
        self.rate_lbl = QLabel("Pass Rate: 0%")
        self.rate_lbl.setStyleSheet("color: #00d9ff; font-weight: bold;")
        stats_row.addWidget(self.rate_lbl)
        
        stats_row.addStretch()
        layout.addLayout(stats_row)
        
        # records list
        self.records_scroll = QScrollArea()
        self.records_scroll.setWidgetResizable(True)
        self.records_widget = QWidget()
        self.records_layout = QVBoxLayout(self.records_widget)
        self.records_layout.setAlignment(Qt.AlignTop)
        self.records_scroll.setWidget(self.records_widget)
        layout.addWidget(self.records_scroll)
        
        self._load_records()
        
        return w
    
    def _create_thresholds_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        
        v_grp = QGroupBox("Voltage")
        v_layout = QHBoxLayout(v_grp)
        v_layout.addWidget(QLabel("Min:"))
        self.th_v_min = QDoubleSpinBox()
        self.th_v_min.setRange(0, 100)
        self.th_v_min.valueChanged.connect(self._apply_thresholds)
        v_layout.addWidget(self.th_v_min)
        v_layout.addWidget(QLabel("Max:"))
        self.th_v_max = QDoubleSpinBox()
        self.th_v_max.setRange(0, 100)
        self.th_v_max.setValue(50)
        self.th_v_max.valueChanged.connect(self._apply_thresholds)
        v_layout.addWidget(self.th_v_max)
        layout.addWidget(v_grp)
        
        i_grp = QGroupBox("Current")
        i_layout = QHBoxLayout(i_grp)
        i_layout.addWidget(QLabel("Min:"))
        self.th_i_min = QDoubleSpinBox()
        self.th_i_min.setRange(0, 50)
        self.th_i_min.setDecimals(3)
        self.th_i_min.valueChanged.connect(self._apply_thresholds)
        i_layout.addWidget(self.th_i_min)
        i_layout.addWidget(QLabel("Max:"))
        self.th_i_max = QDoubleSpinBox()
        self.th_i_max.setRange(0, 50)
        self.th_i_max.setDecimals(3)
        self.th_i_max.setValue(5)
        self.th_i_max.valueChanged.connect(self._apply_thresholds)
        i_layout.addWidget(self.th_i_max)
        layout.addWidget(i_grp)
        
        f_grp = QGroupBox("Frequency")
        f_layout = QHBoxLayout(f_grp)
        f_layout.addWidget(QLabel("Min:"))
        self.th_f_min = QDoubleSpinBox()
        self.th_f_min.setRange(0, 1e6)
        self.th_f_min.valueChanged.connect(self._apply_thresholds)
        f_layout.addWidget(self.th_f_min)
        f_layout.addWidget(QLabel("Max:"))
        self.th_f_max = QDoubleSpinBox()
        self.th_f_max.setRange(0, 1e6)
        self.th_f_max.setValue(100000)
        self.th_f_max.valueChanged.connect(self._apply_thresholds)
        f_layout.addWidget(self.th_f_max)
        layout.addWidget(f_grp)
        
        layout.addStretch()
        
        self._apply_thresholds()
        
        return w
    
    def _create_stats_tab(self):
        w = QWidget()
        layout = QVBoxLayout(w)
        
        grp = QGroupBox("Session Statistics")
        g_layout = QGridLayout(grp)
        
        headers = ["", "Min", "Max", "Avg", "StdDev"]
        for c, h in enumerate(headers):
            lbl = QLabel(h)
            lbl.setStyleSheet("font-weight: bold;")
            g_layout.addWidget(lbl, 0, c)
        
        self.stat_labels = {}
        for r, name in enumerate(["Voltage", "Current", "Power", "Frequency"], 1):
            g_layout.addWidget(QLabel(name), r, 0)
            self.stat_labels[name] = {}
            for c, key in enumerate(["min", "max", "avg", "std"], 1):
                lbl = QLabel("---")
                self.stat_labels[name][key] = lbl
                g_layout.addWidget(lbl, r, c)
        
        layout.addWidget(grp)
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._calc_stats)
        layout.addWidget(refresh_btn)
        
        layout.addStretch()
        
        return w
    
    # --- handlers ---
    
    def _refresh_ports(self):
        self.port_cb.clear()
        for p in serial.tools.list_ports.comports():
            self.port_cb.addItem(f"{p.device} - {p.description}", p.device)
    
    def _toggle_connection(self):
        if self.serial.running:
            self.serial.disconnect()
        else:
            port = self.port_cb.currentData()
            baud = int(self.baud_cb.currentText())
            if port:
                self.serial.connect_to(port, baud)
            else:
                QMessageBox.warning(self, "Error", "Select a port first")
    
    def _on_status(self, connected, msg):
        self.status_light.set_connected(connected, msg)
        self.conn_btn.setText("Disconnect" if connected else "Connect")
        self.status.showMessage(msg)
        if connected:
            self.t0 = time.time()
    
    def _on_error(self, msg):
        self.status.showMessage(f"Error: {msg}")
    
    def _on_data(self, d):
        t = time.time() - self.t0
        
        v = d.get('V', 0)
        i = d.get('I', 0)
        p = d.get('P', 0)
        r = d.get('R', 0)
        f = d.get('F', 0)
        wl = d.get('WL', 0)
        vrms = d.get('Vrms', 0)
        vpp = d.get('Vpp', 0)
        
        # update meters
        self.v_meter.set_value(v)
        self.i_meter.set_value(i, 4)
        self.p_meter.set_value(p)
        self.r_meter.set_value(r, 1)
        self.f_meter.set_value(f, 1)
        self.wl_meter.set_value(wl, 2)
        self.vrms_meter.set_value(vrms)
        self.vpp_meter.set_value(vpp)
        
        # store
        self.time_buf.append(t)
        self.v_buf.append(v)
        self.i_buf.append(i)
        self.p_buf.append(p)
        self.f_buf.append(f)
        
        # violations
        if self.v_meter.violation:
            self.v_viols += 1
        if self.i_meter.violation:
            self.i_viols += 1
        if self.f_meter.violation:
            self.f_viols += 1
        
        # record if testing
        if self.testing:
            self.test_data.append({
                'time': t, 'voltage': v, 'current': i,
                'power': p, 'resistance': r, 'frequency': f, 'wavelength': wl
            })
    
    def _update_plots(self):
        if len(self.time_buf) > 0:
            t = np.array(self.time_buf)
            self.v_curve.setData(t, np.array(self.v_buf))
            self.i_curve.setData(t, np.array(self.i_buf))
            self.p_curve.setData(t, np.array(self.p_buf))
            self.f_curve.setData(t, np.array(self.f_buf))
    
    def _apply_thresholds(self):
        self.v_meter.set_thresholds(self.th_v_min.value(), self.th_v_max.value())
        self.i_meter.set_thresholds(self.th_i_min.value(), self.th_i_max.value())
        self.f_meter.set_thresholds(self.th_f_min.value(), self.th_f_max.value())
    
    def _start_test(self):
        templates = self.db.get_templates()
        dlg = NewTestDialog(self, templates)
        
        if dlg.exec_() != QDialog.Accepted:
            return
        
        data = dlg.get_data()
        
        # set thresholds
        self.th_v_min.setValue(data['v_min'])
        self.th_v_max.setValue(data['v_max'])
        self.th_i_min.setValue(data['i_min'])
        self.th_i_max.setValue(data['i_max'])
        self.th_f_min.setValue(data['f_min'])
        self.th_f_max.setValue(data['f_max'])
        self._apply_thresholds()
        
        # save info
        self.test_info = data
        
        # reset
        self.test_data = []
        self.v_viols = 0
        self.i_viols = 0
        self.f_viols = 0
        self.test_start = datetime.now()
        
        # clear buffers
        self.time_buf.clear()
        self.v_buf.clear()
        self.i_buf.clear()
        self.p_buf.clear()
        self.f_buf.clear()
        self.t0 = time.time()
        
        # ui
        self.testing = True
        self.test_label.setText(f"Testing: {data['name']}")
        self.test_label.setStyleSheet("color: #00d9ff; font-weight: bold;")
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.pass_btn.setEnabled(True)
        self.fail_btn.setEnabled(True)
        
        self.test_timer.start(1000)
        self.status.showMessage(f"Test started: {data['name']}")
    
    def _stop_test(self):
        self._finish_test("ABORTED")
    
    def _finish_test(self, status):
        if not self.testing:
            return
        
        self.testing = False
        self.test_timer.stop()
        
        end = datetime.now()
        duration = (end - self.test_start).total_seconds()
        
        # calc stats
        if self.test_data:
            v_arr = np.array([d['voltage'] for d in self.test_data])
            i_arr = np.array([d['current'] for d in self.test_data])
            p_arr = np.array([d['power'] for d in self.test_data])
            f_arr = np.array([d['frequency'] for d in self.test_data])
        else:
            v_arr = i_arr = p_arr = f_arr = np.array([0])
        
        # save
        record = {
            'name': self.test_info.get('name', ''),
            'board': self.test_info.get('board', ''),
            'serial_num': self.test_info.get('serial', ''),
            'operator': self.test_info.get('operator', ''),
            'start_time': self.test_start.isoformat(),
            'end_time': end.isoformat(),
            'duration': duration,
            'status': status,
            'v_min': float(np.min(v_arr)),
            'v_max': float(np.max(v_arr)),
            'v_avg': float(np.mean(v_arr)),
            'i_min': float(np.min(i_arr)),
            'i_max': float(np.max(i_arr)),
            'i_avg': float(np.mean(i_arr)),
            'p_min': float(np.min(p_arr)),
            'p_max': float(np.max(p_arr)),
            'p_avg': float(np.mean(p_arr)),
            'f_min': float(np.min(f_arr)),
            'f_max': float(np.max(f_arr)),
            'f_avg': float(np.mean(f_arr)),
            'v_violations': self.v_viols,
            'i_violations': self.i_viols,
            'f_violations': self.f_viols,
            'notes': self.test_info.get('notes', ''),
            'raw_data': json.dumps(self.test_data[-1000:])
        }
        
        test_id = self.db.save_test(record)
        
        # reset ui
        self.test_label.setText("No active test")
        self.test_label.setStyleSheet("color: #888;")
        self.duration_label.setText("00:00:00")
        
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.pass_btn.setEnabled(False)
        self.fail_btn.setEnabled(False)
        
        self._load_records()
        
        result_txt = {'PASS': 'PASSED', 'FAIL': 'FAILED', 'ABORTED': 'ABORTED'}
        QMessageBox.information(
            self, "Test Complete",
            f"Test #{test_id}\nResult: {result_txt.get(status)}\nDuration: {duration:.1f}s\nSamples: {len(self.test_data)}"
        )
        
        self.test_data = []
        self.test_info = {}
    
    def _update_duration(self):
        if self.test_start:
            elapsed = datetime.now() - self.test_start
            secs = int(elapsed.total_seconds())
            h, rem = divmod(secs, 3600)
            m, s = divmod(rem, 60)
            self.duration_label.setText(f"{h:02d}:{m:02d}:{s:02d}")
    
    def _load_records(self):
        # clear
        while self.records_layout.count():
            item = self.records_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        records = self.db.get_all_tests()
        
        for r in records[:50]:
            card = TestCard(r)
            card.clicked.connect(self._show_record)
            self.records_layout.addWidget(card)
        
        stats = self.db.get_stats()
        self.total_lbl.setText(f"Total: {stats['total']}")
        self.pass_lbl.setText(f"Passed: {stats['passed']}")
        self.fail_lbl.setText(f"Failed: {stats['failed']}")
        self.rate_lbl.setText(f"Pass Rate: {stats['pass_rate']:.1f}%")
    
    def _search_records(self, query):
        while self.records_layout.count():
            item = self.records_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        records = self.db.search(query) if query else self.db.get_all_tests()
        
        for r in records[:50]:
            card = TestCard(r)
            card.clicked.connect(self._show_record)
            self.records_layout.addWidget(card)
    
    def _filter_records(self, status):
        while self.records_layout.count():
            item = self.records_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        records = self.db.filter_by_status(status) if status != "All" else self.db.get_all_tests()
        
        for r in records[:50]:
            card = TestCard(r)
            card.clicked.connect(self._show_record)
            self.records_layout.addWidget(card)
    
    def _show_record(self, rid):
        r = self.db.get_test(rid)
        if r:
            dlg = TestDetailsDialog(r, self)
            dlg.exec_()
    
    def _show_templates(self):
        dlg = TemplatesDialog(self.db, self)
        dlg.exec_()
    
    def _export_all(self):
        fname, _ = QFileDialog.getSaveFileName(
            self, "Export",
            f"tests_{datetime.now().strftime('%Y%m%d')}.csv",
            "CSV (*.csv)"
        )
        
        if not fname:
            return
        
        records = self.db.get_all_tests()
        if not records:
            QMessageBox.information(self, "Info", "No records to export")
            return
        
        # get keys except raw_data
        keys = [k for k in records[0].keys() if k != 'raw_data']
        
        with open(fname, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=keys)
            writer.writeheader()
            for r in records:
                row = {k: r[k] for k in keys}
                writer.writerow(row)
        
        QMessageBox.information(self, "Done", f"Exported {len(records)} records")
    
    def _calc_stats(self):
        data = {
            'Voltage': np.array(self.v_buf),
            'Current': np.array(self.i_buf),
            'Power': np.array(self.p_buf),
            'Frequency': np.array(self.f_buf)
        }
        
        for name, arr in data.items():
            if len(arr) > 0:
                self.stat_labels[name]['min'].setText(f"{np.min(arr):.4f}")
                self.stat_labels[name]['max'].setText(f"{np.max(arr):.4f}")
                self.stat_labels[name]['avg'].setText(f"{np.mean(arr):.4f}")
                self.stat_labels[name]['std'].setText(f"{np.std(arr):.4f}")
    
    def closeEvent(self, event):
        if self.testing:
            if QMessageBox.question(self, "Exit", "Test in progress. Stop and exit?") == QMessageBox.No:
                event.ignore()
                return
            self._finish_test("ABORTED")
        
        self.serial.disconnect()
        event.accept()


# --------------------------------
# Run
# --------------------------------
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())