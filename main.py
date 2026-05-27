import ctypes
import json
import logging
import os
import sqlite3
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QAction, QCloseEvent, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSystemTrayIcon,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_NAME = "TimeIncomeTracker"


def app_data_dir() -> Path:
    base = Path(os.environ.get("APPDATA", str(Path.home())))
    path = base / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    (path / "logs").mkdir(exist_ok=True)
    return path


def setup_logging(base_path: Path):
    log_file = base_path / "logs" / "app.log"
    logging.basicConfig(level=logging.INFO, filename=log_file, format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")


class DB:
    def __init__(self, db_path: Path):
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.init()

    def init(self):
        cur = self.conn.cursor()
        cur.executescript(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL,
                hourly_rate INTEGER NULL,
                has_hourly_rate INTEGER DEFAULT 0,
                is_active INTEGER DEFAULT 1,
                created_at TEXT,
                updated_at TEXT
            );
            CREATE TABLE IF NOT EXISTS time_entries (
                id INTEGER PRIMARY KEY,
                project_id INTEGER,
                project_name_snapshot TEXT,
                task_description TEXT,
                start_time TEXT,
                end_time TEXT,
                duration_seconds INTEGER,
                hourly_rate_snapshot INTEGER NULL,
                amount INTEGER NULL,
                status TEXT,
                created_at TEXT,
                updated_at TEXT
            );
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );
            CREATE TABLE IF NOT EXISTS active_timer_state (
                id INTEGER PRIMARY KEY,
                project_id INTEGER,
                project_name_snapshot TEXT,
                task_description TEXT,
                start_time TEXT,
                accumulated_seconds INTEGER,
                is_paused INTEGER,
                pause_started_at TEXT,
                auto_paused INTEGER,
                hourly_rate_snapshot INTEGER NULL,
                updated_at TEXT
            );
            """
        )
        self.conn.commit()

    def set_setting(self, key, value):
        self.conn.execute("INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))
        self.conn.commit()

    def get_setting(self, key, default=None):
        row = self.conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row[0] if row else default


class IdleMonitor:
    class LASTINPUTINFO(ctypes.Structure):
        _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

    @staticmethod
    def idle_seconds() -> int:
        if os.name != "nt":
            return 0
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32
        lii = IdleMonitor.LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(lii)
        if not user32.GetLastInputInfo(ctypes.byref(lii)):
            return 0
        tick = kernel32.GetTickCount()
        return int((tick - lii.dwTime) / 1000)


@dataclass
class TimerState:
    project_id: int
    project_name: str
    task_description: str
    start_time: datetime
    accumulated_seconds: int = 0
    is_paused: bool = False
    pause_started_at: Optional[datetime] = None
    auto_paused: bool = False
    hourly_rate_snapshot: Optional[int] = None


class SettingsDialog(QDialog):
    def __init__(self, db: DB, app_dir: Path, parent=None):
        super().__init__(parent)
        self.db = db
        self.app_dir = app_dir
        self.setWindowTitle("Settings")
        self.resize(700, 500)
        self.setModal(True)
        tabs = QTabWidget()

        # General
        general = QWidget()
        gform = QFormLayout(general)
        self.hotkey = QLineEdit(self.db.get_setting("hotkey", "<shift>+q"))
        self.notify_enabled = QCheckBox()
        self.notify_enabled.setChecked(self.db.get_setting("notify_enabled", "1") == "1")
        self.notify_min = QSpinBox(); self.notify_min.setRange(5, 240); self.notify_min.setValue(int(self.db.get_setting("notify_minutes", "30")))
        self.idle_enabled = QCheckBox(); self.idle_enabled.setChecked(self.db.get_setting("idle_enabled", "1") == "1")
        self.idle_min = QSpinBox(); self.idle_min.setRange(1, 60); self.idle_min.setValue(int(self.db.get_setting("idle_minutes", "2")))
        self.default_rate = QLineEdit(self.db.get_setting("default_rate", ""))
        self.auto_start = QCheckBox()
        self.auto_start.setChecked(self.db.get_setting("auto_start", "0") == "1")
        self.show_on_startup = QCheckBox()
        self.show_on_startup.setChecked(self.db.get_setting("show_main_window_on_startup", "0") == "1")
        self.excel_dir = QLineEdit(self.db.get_setting("excel_dir", str(app_dir)))
        browse = QPushButton("Browse")
        browse.clicked.connect(self.pick_dir)
        h = QHBoxLayout(); h.addWidget(self.excel_dir); h.addWidget(browse)
        wrap = QWidget(); wrap.setLayout(h)
        gform.addRow("Global Hotkey", self.hotkey)
        gform.addRow("Enable Notifications", self.notify_enabled)
        gform.addRow("Notify every (minutes)", self.notify_min)
        gform.addRow("Auto-pause on idle", self.idle_enabled)
        gform.addRow("Idle minutes", self.idle_min)
        gform.addRow("Default hourly rate (Toman)", self.default_rate)
        gform.addRow("Auto-start with Windows", self.auto_start)
        gform.addRow("Show main window on startup", self.show_on_startup)
        gform.addRow("Excel output directory", wrap)

        # Projects
        projects = QWidget()
        pv = QVBoxLayout(projects)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Rate", "Active"])
        pv.addWidget(self.table)
        hp = QHBoxLayout()
        add_btn, save_btn = QPushButton("Add"), QPushButton("Save Projects")
        hp.addWidget(add_btn); hp.addWidget(save_btn); hp.addStretch()
        pv.addLayout(hp)
        add_btn.clicked.connect(self.add_project_row)
        save_btn.clicked.connect(self.save_projects)

        tabs.addTab(general, "General")
        tabs.addTab(projects, "Projects")
        self.load_projects()

        btn = QPushButton("Save Settings")
        btn.setObjectName("primaryButton")
        btn.clicked.connect(self.save_settings)
        layout = QVBoxLayout(self); layout.addWidget(tabs); layout.addWidget(btn)

    def pick_dir(self):
        path = QFileDialog.getExistingDirectory(self, "Select folder", self.excel_dir.text())
        if path:
            self.excel_dir.setText(path)

    def load_projects(self):
        rows = self.db.conn.execute("SELECT * FROM projects ORDER BY id DESC").fetchall()
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(str(r["id"])))
            self.table.setItem(i, 1, QTableWidgetItem(r["name"]))
            self.table.setItem(i, 2, QTableWidgetItem("" if r["hourly_rate"] is None else str(r["hourly_rate"])))
            self.table.setItem(i, 3, QTableWidgetItem("1" if r["is_active"] else "0"))

    def add_project_row(self):
        i = self.table.rowCount(); self.table.insertRow(i)
        self.table.setItem(i, 0, QTableWidgetItem(""))
        self.table.setItem(i, 1, QTableWidgetItem(""))
        self.table.setItem(i, 2, QTableWidgetItem(""))
        self.table.setItem(i, 3, QTableWidgetItem("1"))

    def save_projects(self):
        for r in range(self.table.rowCount()):
            pid = (self.table.item(r, 0).text() if self.table.item(r, 0) else "").strip()
            name = (self.table.item(r, 1).text() if self.table.item(r, 1) else "").strip()
            if not name:
                continue
            rate_txt = (self.table.item(r, 2).text() if self.table.item(r, 2) else "").replace(",", "").strip()
            rate = int(rate_txt) if rate_txt else None
            has_rate = 1 if rate is not None else 0
            active = 1 if (self.table.item(r, 3).text() if self.table.item(r, 3) else "1").strip() != "0" else 0
            now = datetime.now().isoformat()
            if pid:
                self.db.conn.execute("UPDATE projects SET name=?,hourly_rate=?,has_hourly_rate=?,is_active=?,updated_at=? WHERE id=?",
                                     (name, rate, has_rate, active, now, int(pid)))
            else:
                self.db.conn.execute("INSERT INTO projects(name,hourly_rate,has_hourly_rate,is_active,created_at,updated_at) VALUES(?,?,?,?,?,?)",
                                     (name, rate, has_rate, active, now, now))
        self.db.conn.commit()
        self.load_projects()

    def save_settings(self):
        self.save_projects()
        self.db.set_setting("hotkey", self.hotkey.text().strip() or "<shift>+q")
        self.db.set_setting("notify_enabled", int(self.notify_enabled.isChecked()))
        self.db.set_setting("notify_minutes", self.notify_min.value())
        self.db.set_setting("idle_enabled", int(self.idle_enabled.isChecked()))
        self.db.set_setting("idle_minutes", self.idle_min.value())
        self.db.set_setting("default_rate", self.default_rate.text().strip())
        self.db.set_setting("auto_start", int(self.auto_start.isChecked()))
        self.db.set_setting("show_main_window_on_startup", int(self.show_on_startup.isChecked()))
        self.db.set_setting("excel_dir", self.excel_dir.text().strip())

        try:
            MainWindow.sync_startup_shortcut(self.auto_start.isChecked())
        except Exception as e:
            logging.exception("auto-start sync failed")
            QMessageBox.warning(self, "Auto-start", f"Auto-start update failed: {e}")
        self.accept()


class MainWindow(QMainWindow):
    @staticmethod
    def startup_shortcut_path() -> Path:
        startup_dir = Path(os.environ.get("APPDATA", str(Path.home()))) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
        return startup_dir / f"{APP_NAME}.lnk"

    @staticmethod
    def current_app_path() -> Path:
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve()
        return Path(__file__).resolve()

    @staticmethod
    def sync_startup_shortcut(enabled: bool):
        if os.name != "nt":
            return
        shortcut = MainWindow.startup_shortcut_path()
        shortcut.parent.mkdir(parents=True, exist_ok=True)

        if not enabled:
            if shortcut.exists():
                shortcut.unlink()
            return

        target = MainWindow.current_app_path()
        workdir = target.parent
        icon_location = str(target)
        escaped_shortcut = str(shortcut).replace("'", "''")
        escaped_target = str(target).replace("'", "''")
        escaped_workdir = str(workdir).replace("'", "''")
        escaped_icon = icon_location.replace("'", "''")
        ps_script = (
            f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{escaped_shortcut}');"
            f"$s.TargetPath='{escaped_target}';"
            f"$s.WorkingDirectory='{escaped_workdir}';"
            f"$s.IconLocation='{escaped_icon}';"
            "$s.Save();"
        )
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            check=True,
            capture_output=True,
            text=True,
        )
    def __init__(self, db, app_dir):
        super().__init__()
        self.db, self.app_dir = db, app_dir
        self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.Tool)
        self.setWindowTitle("Time Income Tracker")
        self.resize(980, 640)
        self.timer_state: Optional[TimerState] = None
        self.last_notification_at: Optional[datetime] = None
        self.hotkey_listener = None
        self.tray: Optional[QSystemTrayIcon] = None
        self.is_quitting = False
        self.tray_menu = None
        self.open_act = None
        self.hide_act = None
        self.start_resume_act = None
        self.pause_act = None
        self.stop_act = None

        self.init_ui()
        self.apply_dark()
        self.load_projects()
        self.restore_active_state()
        self.init_timers()
        self.init_tray()
        self.setup_hotkey()

    def init_ui(self):
        w = QWidget(); self.setCentralWidget(w)
        v = QVBoxLayout(w)
        v.setContentsMargins(22, 22, 22, 22)
        v.setSpacing(16)

        self.project_combo = QComboBox()
        self.desc = QTextEdit(); self.desc.setPlaceholderText("شرح فعالیت")
        self.timer_lbl = QLabel("00:00:00"); self.timer_lbl.setObjectName("timerLabel")
        self.amount_lbl = QLabel("0 تومان")
        self.amount_lbl.setObjectName("amountLabel")
        self.status_lbl = QLabel("Status: idle")
        self.status_lbl.setObjectName("statusLabel")
        self.banner = QLabel("")
        self.banner.setObjectName("bannerLabel")

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        form.setVerticalSpacing(12)
        form.addRow("Project", self.project_combo)
        form.addRow("Description", self.desc)

        btns = QHBoxLayout()
        button_map = [
            ("Start", self.start_timer, "primaryButton"),
            ("Pause", self.pause_timer, ""),
            ("Resume", self.resume_timer, ""),
            ("Stop & Save", self.stop_save, "successButton"),
            ("Cancel", self.cancel_timer, "dangerButton"),
            ("Settings", self.open_settings, ""),
            ("Export Excel", self.export_excel, ""),
        ]
        for t, fn, cls in button_map:
            b = QPushButton(t)
            if cls:
                b.setObjectName(cls)
            b.clicked.connect(fn)
            btns.addWidget(b)

        v.addLayout(form); v.addWidget(self.timer_lbl); v.addWidget(self.amount_lbl); v.addWidget(self.status_lbl); v.addWidget(self.banner); v.addLayout(btns)

    def apply_dark(self):
        self.setStyleSheet("""
            QWidget { background:#0f172a; color:#e2e8f0; font-size:14px; }
            QMainWindow { background:#0f172a; }
            QLineEdit, QTextEdit, QComboBox, QSpinBox, QDateEdit, QTableWidget, QTabWidget::pane {
                background:#1e293b; color:#f8fafc; border:1px solid #334155; border-radius:10px; padding:7px;
            }
            QComboBox::drop-down { border:none; }
            QPushButton {
                background:#334155; color:#f8fafc; border:1px solid #475569; border-radius:10px; padding:10px 14px; font-weight:600;
            }
            QPushButton:hover { background:#475569; }
            QPushButton#primaryButton { background:#2563eb; border:1px solid #3b82f6; }
            QPushButton#primaryButton:hover { background:#1d4ed8; }
            QPushButton#successButton { background:#059669; border:1px solid #10b981; }
            QPushButton#successButton:hover { background:#047857; }
            QPushButton#dangerButton { background:#dc2626; border:1px solid #ef4444; }
            QPushButton#dangerButton:hover { background:#b91c1c; }
            QLabel#timerLabel { font-size:52px; font-weight:800; color:#f8fafc; padding-top:8px; }
            QLabel#amountLabel { font-size:26px; font-weight:700; color:#93c5fd; }
            QLabel#statusLabel { color:#94a3b8; font-size:13px; }
            QLabel#bannerLabel { color:#facc15; font-weight:600; min-height:22px; }
            QHeaderView::section { background:#1e293b; color:#cbd5e1; border:1px solid #334155; padding:6px; font-weight:700; }
            QTabBar::tab { background:#1e293b; color:#cbd5e1; padding:8px 14px; border-top-left-radius:8px; border-top-right-radius:8px; margin-right:4px; }
            QTabBar::tab:selected { background:#2563eb; color:#ffffff; }
        """)

    def load_projects(self):
        self.project_combo.clear()
        rows = self.db.conn.execute("SELECT * FROM projects WHERE is_active=1 ORDER BY name").fetchall()
        for r in rows:
            self.project_combo.addItem(r["name"], dict(r))

    def current_duration(self):
        if not self.timer_state:
            return 0
        d = self.timer_state.accumulated_seconds
        if not self.timer_state.is_paused:
            d += int((datetime.now() - self.timer_state.start_time).total_seconds())
        return max(0, d)

    def start_timer(self):
        data = self.project_combo.currentData()
        if not data:
            QMessageBox.warning(self, "Warning", "ابتدا یک پروژه فعال بسازید.")
            return
        self.timer_state = TimerState(
            project_id=data["id"], project_name=data["name"], task_description=self.desc.toPlainText().strip(),
            start_time=datetime.now(), hourly_rate_snapshot=data["hourly_rate"] if data["has_hourly_rate"] else None
        )
        self.status_lbl.setText("Status: running")
        self.save_active_state()

    def pause_timer(self, auto=False):
        if not self.timer_state or self.timer_state.is_paused:
            return
        self.timer_state.accumulated_seconds = self.current_duration()
        self.timer_state.is_paused = True
        self.timer_state.pause_started_at = datetime.now()
        self.timer_state.auto_paused = auto
        self.status_lbl.setText("Status: paused")
        self.save_active_state()

    def resume_timer(self):
        if not self.timer_state or not self.timer_state.is_paused:
            return
        self.timer_state.start_time = datetime.now()
        self.timer_state.is_paused = False
        self.timer_state.pause_started_at = None
        self.timer_state.auto_paused = False
        self.banner.setText("")
        self.status_lbl.setText("Status: running")
        self.save_active_state()

    def stop_save(self):
        if not self.timer_state:
            return
        end = datetime.now()
        dur = self.current_duration()
        rate = self.timer_state.hourly_rate_snapshot
        amount = int(dur / 3600 * rate) if rate is not None else None
        now = datetime.now().isoformat()
        self.db.conn.execute("""INSERT INTO time_entries(project_id,project_name_snapshot,task_description,start_time,end_time,duration_seconds,hourly_rate_snapshot,amount,status,created_at,updated_at)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                           (self.timer_state.project_id, self.timer_state.project_name, self.timer_state.task_description,
                            self.timer_state.start_time.isoformat(), end.isoformat(), dur, rate, amount, "saved", now, now))
        self.db.conn.execute("DELETE FROM active_timer_state")
        self.db.conn.commit()
        self.timer_state = None
        self.status_lbl.setText("Status: saved")

    def cancel_timer(self):
        self.timer_state = None
        self.db.conn.execute("DELETE FROM active_timer_state")
        self.db.conn.commit()
        self.status_lbl.setText("Status: cancelled")

    def save_active_state(self):
        if not self.timer_state:
            return
        self.db.conn.execute("DELETE FROM active_timer_state")
        self.db.conn.execute("""INSERT INTO active_timer_state(id,project_id,project_name_snapshot,task_description,start_time,accumulated_seconds,is_paused,pause_started_at,auto_paused,hourly_rate_snapshot,updated_at)
                           VALUES(1,?,?,?,?,?,?,?,?,?,?)""",
                           (self.timer_state.project_id, self.timer_state.project_name, self.timer_state.task_description,
                            self.timer_state.start_time.isoformat(), self.timer_state.accumulated_seconds, int(self.timer_state.is_paused),
                            self.timer_state.pause_started_at.isoformat() if self.timer_state.pause_started_at else None,
                            int(self.timer_state.auto_paused), self.timer_state.hourly_rate_snapshot, datetime.now().isoformat()))
        self.db.conn.commit()

    def restore_active_state(self):
        row = self.db.conn.execute("SELECT * FROM active_timer_state WHERE id=1").fetchone()
        if not row:
            return
        self.timer_state = TimerState(
            project_id=row["project_id"], project_name=row["project_name_snapshot"], task_description=row["task_description"] or "",
            start_time=datetime.fromisoformat(row["start_time"]), accumulated_seconds=row["accumulated_seconds"] or 0,
            is_paused=bool(row["is_paused"]), pause_started_at=datetime.fromisoformat(row["pause_started_at"]) if row["pause_started_at"] else None,
            auto_paused=bool(row["auto_paused"]), hourly_rate_snapshot=row["hourly_rate_snapshot"]
        )
        self.desc.setPlainText(self.timer_state.task_description)
        self.status_lbl.setText("Status: paused" if self.timer_state.is_paused else "Status: running")

    def init_timers(self):
        self.ui_timer = QTimer(self); self.ui_timer.setInterval(1000); self.ui_timer.timeout.connect(self.tick); self.ui_timer.start()
        self.idle_timer = QTimer(self); self.idle_timer.setInterval(5000); self.idle_timer.timeout.connect(self.check_idle); self.idle_timer.start()

    def tick(self):
        d = self.current_duration()
        self.timer_lbl.setText(f"{d//3600:02}:{(d%3600)//60:02}:{d%60:02}")
        if self.timer_state and self.timer_state.hourly_rate_snapshot:
            amount = int(d / 3600 * self.timer_state.hourly_rate_snapshot)
            self.amount_lbl.setText(f"{amount:,} تومان")
        else:
            self.amount_lbl.setText("0 تومان")
        self.notify_if_needed()

    def notify_if_needed(self):
        if not self.timer_state or self.timer_state.is_paused:
            return
        if self.db.get_setting("notify_enabled", "1") != "1":
            return
        mins = int(self.db.get_setting("notify_minutes", "30"))
        now = datetime.now()
        if self.last_notification_at and (now - self.last_notification_at).total_seconds() < mins * 60:
            return
        self.last_notification_at = now
        if self.tray:
            self.tray.showMessage("Time Income Tracker", f"هنوز مشغول «{self.timer_state.task_description or self.timer_state.project_name}» هستی؟", QSystemTrayIcon.MessageIcon.Information, 6000)

    def check_idle(self):
        if not self.timer_state or self.timer_state.is_paused:
            return
        if self.db.get_setting("idle_enabled", "1") != "1":
            return
        mins = int(self.db.get_setting("idle_minutes", "2"))
        if IdleMonitor.idle_seconds() >= mins * 60:
            self.pause_timer(auto=True)
            self.banner.setText("تایمر به دلیل عدم فعالیت متوقف شد. Resume / Keep Paused / Stop & Save")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.windowIcon() or QIcon())
        self.tray_menu = QMenu()
        self.open_act = QAction("Open / Show", self); self.open_act.triggered.connect(self.show_from_tray)
        self.hide_act = QAction("Hide", self); self.hide_act.triggered.connect(self.hide)
        self.start_resume_act = QAction("Start Timer", self); self.start_resume_act.triggered.connect(self.start_or_resume_timer)
        self.pause_act = QAction("Pause Timer", self); self.pause_act.triggered.connect(self.pause_timer)
        self.stop_act = QAction("Stop & Save", self); self.stop_act.triggered.connect(self.stop_save)
        settings_act = QAction("Settings", self); settings_act.triggered.connect(self.open_settings)
        exit_act = QAction("Exit", self); exit_act.triggered.connect(self.quit_app)
        for a in [self.open_act, self.hide_act, self.start_resume_act, self.pause_act, self.stop_act, settings_act, exit_act]:
            self.tray_menu.addAction(a)
        self.tray_menu.aboutToShow.connect(self.update_tray_menu)
        self.tray.setContextMenu(self.tray_menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()
        self.update_tray_menu()

    def setup_hotkey(self):
        try:
            from pynput import keyboard
            hotkey = self.db.get_setting("hotkey", "<shift>+q")
            self.hotkey_listener = keyboard.GlobalHotKeys({hotkey: self.toggle_window})
            self.hotkey_listener.start()
        except Exception as e:
            logging.exception("hotkey failed")
            self.status_lbl.setText(f"Status: Hotkey error ({e})")

    def toggle_window(self):
        if self.isHidden():
            self.show_from_tray()
        else:
            self.hide()

    def show_from_tray(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def on_tray_activated(self, reason):
        if reason in (QSystemTrayIcon.ActivationReason.Trigger, QSystemTrayIcon.ActivationReason.DoubleClick):
            self.toggle_window()

    def start_or_resume_timer(self):
        if self.timer_state and self.timer_state.is_paused:
            self.resume_timer()
        elif not self.timer_state:
            self.start_timer()

    def update_tray_menu(self):
        if self.timer_state:
            if self.timer_state.is_paused:
                self.start_resume_act.setText("Resume Timer")
                self.start_resume_act.setEnabled(True)
                self.pause_act.setEnabled(False)
                self.stop_act.setEnabled(True)
            else:
                self.start_resume_act.setText("Timer Running")
                self.start_resume_act.setEnabled(False)
                self.pause_act.setEnabled(True)
                self.stop_act.setEnabled(True)
        else:
            self.start_resume_act.setText("Start Timer")
            self.start_resume_act.setEnabled(True)
            self.pause_act.setEnabled(False)
            self.stop_act.setEnabled(False)

    def open_settings(self):
        dlg = SettingsDialog(self.db, self.app_dir, self)
        if dlg.exec():
            self.load_projects()

    def export_excel(self):
        dlg = QDialog(self); dlg.setWindowTitle("Excel Report")
        layout = QFormLayout(dlg)
        period = QComboBox(); period.addItems(["Today", "This Week", "This Month", "Last Month", "Custom"])
        project = QComboBox(); project.addItem("All", None)
        for r in self.db.conn.execute("SELECT id,name FROM projects ORDER BY name"): project.addItem(r["name"], r["id"])
        start = QDateEdit(); start.setCalendarPopup(True); start.setDate(datetime.now().date())
        end = QDateEdit(); end.setCalendarPopup(True); end.setDate(datetime.now().date())
        go = QPushButton("Export")
        layout.addRow("Period", period); layout.addRow("Project", project); layout.addRow("Start", start); layout.addRow("End", end); layout.addRow(go)

        def do_export():
            sdt, edt = self.resolve_period(period.currentText(), start.date().toPyDate(), end.date().toPyDate())
            pid = project.currentData()
            q = "SELECT * FROM time_entries WHERE status='saved' AND start_time>=? AND start_time<?"
            args = [sdt.isoformat(), (edt + timedelta(days=1)).isoformat()]
            if pid:
                q += " AND project_id=?"; args.append(pid)
            rows = self.db.conn.execute(q, args).fetchall()
            wb = Workbook(); ws = wb.active; ws.title = "Report"
            headers = ["ردیف", "تاریخ", "نام پروژه", "شرح فعالیت", "ساعت شروع", "ساعت پایان", "مدت", "نرخ ساعتی", "مبلغ"]
            ws.append(headers)
            for c in ws[1]:
                c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="333333")
            total_sec = total_amount = 0
            for idx, r in enumerate(rows, 1):
                st = datetime.fromisoformat(r["start_time"]); en = datetime.fromisoformat(r["end_time"])
                dur = r["duration_seconds"] or 0
                total_sec += dur; total_amount += r["amount"] or 0
                ws.append([idx, st.date().isoformat(), r["project_name_snapshot"], r["task_description"], st.strftime("%H:%M"), en.strftime("%H:%M"), f"{dur//3600:02}:{(dur%3600)//60:02}:{dur%60:02}", r["hourly_rate_snapshot"], r["amount"]])
            ws.append([]); ws.append(["", "", "", "", "", "جمع کل زمان", f"{total_sec//3600:02}:{(total_sec%3600)//60:02}:{total_sec%60:02}", "جمع کل مبلغ", total_amount])
            for i in range(1, 10): ws.column_dimensions[get_column_letter(i)].width = 18
            out_dir = Path(self.db.get_setting("excel_dir", str(self.app_dir))); out_dir.mkdir(parents=True, exist_ok=True)
            file = out_dir / f"TimeIncomeReport-{datetime.now().date().isoformat()}.xlsx"
            wb.save(file)
            QMessageBox.information(self, "Done", f"Saved: {file}")
            dlg.accept()

        go.clicked.connect(do_export)
        dlg.exec()

    def resolve_period(self, p, custom_s, custom_e):
        now = datetime.now().date()
        if p == "Today": return now, now
        if p == "This Week": return now - timedelta(days=now.weekday()), now
        if p == "This Month": return now.replace(day=1), now
        if p == "Last Month":
            first = now.replace(day=1)
            last_prev = first - timedelta(days=1)
            return last_prev.replace(day=1), last_prev
        return custom_s, custom_e

    def quit_app(self):
        if self.timer_state and not self.timer_state.is_paused:
            r = QMessageBox.question(
                self,
                "Exit",
                "یک تایمر فعال وجود دارد. آیا از خروج مطمئن هستید؟",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        self.is_quitting = True
        if self.tray:
            self.tray.hide()
        QApplication.quit()

    def closeEvent(self, event: QCloseEvent):
        if self.is_quitting:
            event.accept()
            return
        event.ignore()
        self.hide()


def main():
    d = app_data_dir()
    setup_logging(d)
    db = DB(d / "time_income.db")
    app = QApplication(sys.argv)
    win = MainWindow(db, d)
    if db.get_setting("show_main_window_on_startup", "0") == "1":
        win.show_from_tray()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
