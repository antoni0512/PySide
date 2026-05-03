"""
Playwright Automation UI — PySide6
Upload Excel → Run Automation → Write Results → Export Updated Excel
"""

import sys
import os
import asyncio
import threading
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QTextEdit, QProgressBar, QFrame, QSplitter, QHeaderView,
    QStatusBar, QMessageBox, QComboBox, QLineEdit, QGroupBox,
    QGridLayout, QScrollArea, QSizePolicy
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QObject, QTimer, QPropertyAnimation,
    QEasingCurve, QSize, QPoint, QRectF
)
from PySide6.QtGui import (
    QColor, QPalette, QFont, QIcon, QPixmap, QPainter,
    QLinearGradient, QBrush, QPen, QFontDatabase, QMovie
)

# ── Optional dependencies (graceful fallback) ──────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Font as XLFont, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from playwright.async_api import async_playwright
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False

# ── Colour palette ─────────────────────────────────────────────────────────
DARK_BG      = "#0D1117"
PANEL_BG     = "#161B22"
SURFACE      = "#21262D"
BORDER       = "#30363D"
ACCENT_BLUE  = "#58A6FF"
ACCENT_GREEN = "#3FB950"
ACCENT_RED   = "#F85149"
ACCENT_AMBER = "#D29922"
TEXT_PRIMARY = "#E6EDF3"
TEXT_MUTED   = "#8B949E"
HIGHLIGHT    = "#1F6FEB"

# ══════════════════════════════════════════════════════════════════════════════
#  Sample Excel generator (demo only)
# ══════════════════════════════════════════════════════════════════════════════

def create_sample_excel(path: str) -> None:
    """Create a demo Excel file with sample automation tasks."""
    if not HAS_OPENPYXL:
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Automation Tasks"

    header_fill = PatternFill(start_color="1F6FEB", end_color="1F6FEB", fill_type="solid")
    header_font = XLFont(bold=True, color="FFFFFF", size=11)

    headers = ["ID", "URL", "Action", "Selector", "Input Value", "Status", "Result", "Timestamp"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    sample_rows = [
        [1, "https://example.com",        "navigate",    "",               "",             "Pending", "", ""],
        [2, "https://example.com",        "screenshot",  "",               "",             "Pending", "", ""],
        [3, "https://quotes.toscrape.com","scrape_title","h1",             "",             "Pending", "", ""],
        [4, "https://httpbin.org/get",    "get_content", "body",           "",             "Pending", "", ""],
        [5, "https://example.com",        "get_links",   "a",              "",             "Pending", "", ""],
    ]

    for row_data in sample_rows:
        ws.append(row_data)

    # Column widths
    for col, width in zip("ABCDEFGH", [5, 35, 15, 20, 20, 10, 30, 20]):
        ws.column_dimensions[chr(64 + col)].width = width

    wb.save(path)


# ══════════════════════════════════════════════════════════════════════════════
#  Worker — runs in a QThread
# ══════════════════════════════════════════════════════════════════════════════

class AutomationWorker(QObject):
    log        = Signal(str, str)   # (message, level)
    row_update = Signal(int, str, str)   # (row_idx, status, result)
    progress   = Signal(int)
    finished   = Signal(bool, str)  # (success, message)

    def __init__(self, tasks: list[dict], headless: bool = True):
        super().__init__()
        self.tasks    = tasks
        self.headless = headless
        self._stop    = False

    def stop(self):
        self._stop = True

    def run(self):
        asyncio.run(self._run_async())

    async def _run_async(self):
        total = len(self.tasks)
        if total == 0:
            self.finished.emit(False, "No tasks to run.")
            return

        if not HAS_PLAYWRIGHT:
            # Simulate without playwright
            self.log.emit("⚠  Playwright not installed — running in simulation mode.", "warn")
            for i, task in enumerate(self.tasks):
                if self._stop:
                    break
                await asyncio.sleep(0.6)
                status  = "Done"
                result  = f"[Simulated] {task.get('action','?')} on {task.get('url','?')}"
                self.row_update.emit(i, status, result)
                self.progress.emit(int((i + 1) / total * 100))
                self.log.emit(f"  ✔  Row {i+1}: {result}", "info")
            self.finished.emit(True, "Simulation complete.")
            return

        self.log.emit("🚀  Launching browser…", "info")
        try:
            async with async_playwright() as pw:
                browser = await pw.chromium.launch(headless=self.headless)
                page    = await browser.new_page()

                for i, task in enumerate(self.tasks):
                    if self._stop:
                        self.log.emit("⛔  Stopped by user.", "warn")
                        break

                    url      = task.get("url", "")
                    action   = task.get("action", "").lower()
                    selector = task.get("selector", "")
                    value    = task.get("input_value", "")

                    self.log.emit(f"▶  Row {i+1}: {action} → {url}", "info")
                    status = "Done"
                    result = ""

                    try:
                        if action == "navigate":
                            await page.goto(url, timeout=15000)
                            result = page.url

                        elif action == "screenshot":
                            await page.goto(url, timeout=15000)
                            fname = f"screenshot_{i+1}_{datetime.now():%H%M%S}.png"
                            await page.screenshot(path=fname)
                            result = f"Saved: {fname}"

                        elif action == "scrape_title":
                            await page.goto(url, timeout=15000)
                            if selector:
                                el  = page.locator(selector).first
                                result = await el.inner_text()
                            else:
                                result = await page.title()

                        elif action == "get_content":
                            await page.goto(url, timeout=15000)
                            el     = page.locator(selector or "body").first
                            result = (await el.inner_text())[:120].replace("\n", " ")

                        elif action == "get_links":
                            await page.goto(url, timeout=15000)
                            links  = await page.eval_on_selector_all(
                                selector or "a", "els => els.map(e => e.href)"
                            )
                            result = f"{len(links)} links found"

                        elif action == "fill":
                            await page.goto(url, timeout=15000)
                            await page.fill(selector, value)
                            result = "Filled"

                        else:
                            result = f"Unknown action: {action}"
                            status = "Skipped"

                    except Exception as ex:
                        status = "Error"
                        result = str(ex)[:100]
                        self.log.emit(f"  ✖  {ex}", "error")

                    self.row_update.emit(i, status, result)
                    self.progress.emit(int((i + 1) / total * 100))
                    self.log.emit(f"  ✔  {status}: {result}", "info" if status == "Done" else "error")

                await browser.close()
            self.finished.emit(True, "Automation complete.")
        except Exception as ex:
            self.finished.emit(False, str(ex))


# ══════════════════════════════════════════════════════════════════════════════
#  Styled widgets
# ══════════════════════════════════════════════════════════════════════════════

class StyledButton(QPushButton):
    def __init__(self, text, accent=ACCENT_BLUE, parent=None):
        super().__init__(text, parent)
        self._accent = accent
        self.setMinimumHeight(38)
        self.setCursor(Qt.PointingHandCursor)
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet(f"""
            QPushButton {{
                background: {self._accent};
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 0 18px;
                font-size: 13px;
                font-weight: 600;
                letter-spacing: 0.3px;
            }}
            QPushButton:hover {{
                background: {self._accent}CC;
            }}
            QPushButton:pressed {{
                background: {self._accent}99;
            }}
            QPushButton:disabled {{
                background: {SURFACE};
                color: {TEXT_MUTED};
            }}
        """)


class GhostButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(38)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 0 16px;
                font-size: 13px;
            }}
            QPushButton:hover {{
                background: {SURFACE};
                border-color: {ACCENT_BLUE};
                color: {ACCENT_BLUE};
            }}
            QPushButton:pressed {{ background: {BORDER}; }}
            QPushButton:disabled {{ color: {TEXT_MUTED}; }}
        """)


class StatusBadge(QLabel):
    COLOURS = {
        "Pending": (ACCENT_AMBER, "#3D2F07"),
        "Done":    (ACCENT_GREEN, "#0A2416"),
        "Error":   (ACCENT_RED,   "#2E0A09"),
        "Skipped": (TEXT_MUTED,   SURFACE),
        "Running": (ACCENT_BLUE,  "#071B3D"),
    }

    def __init__(self, status="Pending", parent=None):
        super().__init__(status, parent)
        self.set_status(status)

    def set_status(self, status):
        fg, bg = self.COLOURS.get(status, (TEXT_MUTED, SURFACE))
        self.setText(status)
        self.setStyleSheet(f"""
            QLabel {{
                background: {bg};
                color: {fg};
                border: 1px solid {fg}55;
                border-radius: 4px;
                padding: 1px 8px;
                font-size: 11px;
                font-weight: 600;
            }}
        """)
        self.setAlignment(Qt.AlignCenter)


# ══════════════════════════════════════════════════════════════════════════════
#  Main Window
# ══════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Playwright Automation Studio")
        self.resize(1280, 820)
        self.setMinimumSize(900, 600)

        self._excel_path:  str | None = None
        self._tasks:       list[dict] = []
        self._worker:      AutomationWorker | None = None
        self._thread:      QThread | None = None
        self._results:     list[dict] = []   # mirrors table rows

        self._apply_global_style()
        self._build_ui()
        self._log("Welcome to Playwright Automation Studio.", "info")
        self._log("1. Load an Excel file  →  2. Run automation  →  3. Export results.", "muted")

    # ── Style ────────────────────────────────────────────────────────────────

    def _apply_global_style(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background: {DARK_BG};
                color: {TEXT_PRIMARY};
                font-family: 'Segoe UI', 'SF Pro Text', system-ui, sans-serif;
                font-size: 13px;
            }}
            QSplitter::handle {{ background: {BORDER}; }}
            QScrollBar:vertical {{
                background: {PANEL_BG}; width: 8px; border: none;
            }}
            QScrollBar::handle:vertical {{
                background: {BORDER}; border-radius: 4px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
            QScrollBar:horizontal {{
                background: {PANEL_BG}; height: 8px; border: none;
            }}
            QScrollBar::handle:horizontal {{
                background: {BORDER}; border-radius: 4px;
            }}
            QHeaderView::section {{
                background: {SURFACE};
                color: {TEXT_MUTED};
                border: none;
                border-bottom: 1px solid {BORDER};
                padding: 6px 10px;
                font-size: 11px;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            QTableWidget {{
                background: {PANEL_BG};
                border: 1px solid {BORDER};
                border-radius: 8px;
                gridline-color: {BORDER};
                selection-background-color: {HIGHLIGHT}44;
            }}
            QTableWidget::item {{
                padding: 6px 10px;
                border-bottom: 1px solid {BORDER}55;
            }}
            QTableWidget::item:selected {{
                background: {HIGHLIGHT}44;
                color: {TEXT_PRIMARY};
            }}
            QComboBox {{
                background: {SURFACE};
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 6px 12px;
                min-height: 34px;
            }}
            QComboBox::drop-down {{ border: none; width: 20px; }}
            QComboBox QAbstractItemView {{
                background: {SURFACE};
                color: {TEXT_PRIMARY};
                selection-background-color: {HIGHLIGHT};
                border: 1px solid {BORDER};
            }}
            QLineEdit {{
                background: {SURFACE};
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 6px 10px;
                min-height: 34px;
            }}
            QLineEdit:focus {{ border-color: {ACCENT_BLUE}; }}
            QGroupBox {{
                border: 1px solid {BORDER};
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 8px;
                font-weight: 600;
                color: {TEXT_MUTED};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }}
            QProgressBar {{
                background: {SURFACE};
                border: 1px solid {BORDER};
                border-radius: 4px;
                height: 6px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ACCENT_BLUE}, stop:1 {ACCENT_GREEN});
                border-radius: 4px;
            }}
            QStatusBar {{
                background: {PANEL_BG};
                border-top: 1px solid {BORDER};
                color: {TEXT_MUTED};
                padding: 4px 12px;
                font-size: 12px;
            }}
        """)

    # ── UI construction ──────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Top bar ──────────────────────────────────────────────────────────
        root.addWidget(self._build_topbar())

        # ── Body splitter ────────────────────────────────────────────────────
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setSizes([820, 440])
        root.addWidget(splitter, 1)

        # ── Bottom bar ───────────────────────────────────────────────────────
        root.addWidget(self._build_bottom_bar())

        # Status bar
        self.statusBar().showMessage("Ready")

    def _build_topbar(self) -> QWidget:
        bar = QFrame()
        bar.setFixedHeight(56)
        bar.setStyleSheet(f"""
            QFrame {{
                background: {PANEL_BG};
                border-bottom: 1px solid {BORDER};
            }}
        """)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(20, 0, 20, 0)

        # Logo / title
        ico = QLabel("⚡")
        ico.setStyleSheet("font-size: 22px;")
        title = QLabel("Playwright Automation Studio")
        title.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {TEXT_PRIMARY};")
        subtitle = QLabel("v1.0")
        subtitle.setStyleSheet(f"font-size: 11px; color: {TEXT_MUTED}; margin-left: 6px; margin-top: 3px;")

        lay.addWidget(ico)
        lay.addSpacing(8)
        lay.addWidget(title)
        lay.addWidget(subtitle)
        lay.addStretch()

        # Playwright / openpyxl badges
        for lib, ok in [("Playwright", HAS_PLAYWRIGHT), ("openpyxl", HAS_OPENPYXL)]:
            badge = QLabel(f"{'✔' if ok else '✖'}  {lib}")
            badge.setStyleSheet(f"""
                background: {'#0A2416' if ok else '#2E0A09'};
                color: {ACCENT_GREEN if ok else ACCENT_RED};
                border: 1px solid {ACCENT_GREEN if ok else ACCENT_RED}55;
                border-radius: 4px;
                padding: 3px 10px;
                font-size: 11px;
                font-weight: 600;
            """)
            lay.addWidget(badge)
            lay.addSpacing(6)

        return bar

    def _build_left_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet(f"background: {DARK_BG};")
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(16, 16, 8, 16)
        lay.setSpacing(12)

        # ── File section ─────────────────────────────────────────────────────
        file_group = QGroupBox("Excel File")
        fg_lay = QVBoxLayout(file_group)
        fg_lay.setSpacing(8)

        file_row = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet(f"""
            background: {SURFACE};
            border: 1px dashed {BORDER};
            border-radius: 6px;
            padding: 8px 12px;
            color: {TEXT_MUTED};
        """)
        self.file_label.setMinimumHeight(36)

        btn_browse = GhostButton("Browse…")
        btn_browse.setFixedWidth(90)
        btn_browse.clicked.connect(self._browse_excel)

        btn_sample = GhostButton("Sample")
        btn_sample.setFixedWidth(80)
        btn_sample.clicked.connect(self._create_sample)
        btn_sample.setToolTip("Generate a sample Excel file")

        file_row.addWidget(self.file_label, 1)
        file_row.addWidget(btn_browse)
        file_row.addWidget(btn_sample)
        fg_lay.addLayout(file_row)

        lay.addWidget(file_group)

        # ── Options ──────────────────────────────────────────────────────────
        opt_group = QGroupBox("Options")
        og_lay = QGridLayout(opt_group)
        og_lay.setSpacing(8)

        og_lay.addWidget(QLabel("Browser mode:"), 0, 0)
        self.browser_mode = QComboBox()
        self.browser_mode.addItems(["Headless (background)", "Headed (visible)"])
        og_lay.addWidget(self.browser_mode, 0, 1)

        og_lay.addWidget(QLabel("Output path:"), 1, 0)
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Auto (same folder as input)")
        og_lay.addWidget(self.output_path, 1, 1)

        lay.addWidget(opt_group)

        # ── Table ────────────────────────────────────────────────────────────
        table_label = QLabel("Tasks")
        table_label.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {TEXT_MUTED};")
        lay.addWidget(table_label)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["#", "URL", "Action", "Selector", "Status", "Result"])
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet(self.table.styleSheet() + f"""
            QTableWidget {{ alternate-background-color: {SURFACE}22; }}
        """)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        hh.setSectionResizeMode(5, QHeaderView.Stretch)
        for c, w in [(0, 36), (2, 110), (3, 110), (4, 80)]:
            self.table.setColumnWidth(c, w)
        self.table.setMinimumHeight(200)
        lay.addWidget(self.table, 1)

        return panel

    def _build_right_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet(f"background: {DARK_BG};")
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(8, 16, 16, 16)
        lay.setSpacing(12)

        # ── Log ──────────────────────────────────────────────────────────────
        log_label = QLabel("Console")
        log_label.setStyleSheet(f"font-size: 12px; font-weight: 600; color: {TEXT_MUTED};")
        lay.addWidget(log_label)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet(f"""
            QTextEdit {{
                background: {PANEL_BG};
                border: 1px solid {BORDER};
                border-radius: 8px;
                padding: 10px;
                font-family: 'JetBrains Mono', 'Cascadia Code', 'Consolas', monospace;
                font-size: 12px;
                color: {TEXT_PRIMARY};
            }}
        """)
        lay.addWidget(self.log_box, 1)

        # ── Stats ────────────────────────────────────────────────────────────
        stats_group = QGroupBox("Run Summary")
        sg = QGridLayout(stats_group)
        sg.setSpacing(6)

        self.stat_total   = self._stat_label("0")
        self.stat_done    = self._stat_label("0", ACCENT_GREEN)
        self.stat_error   = self._stat_label("0", ACCENT_RED)
        self.stat_pending = self._stat_label("0", ACCENT_AMBER)

        for col, (lbl, val) in enumerate([
            ("Total", self.stat_total),
            ("Done",  self.stat_done),
            ("Error", self.stat_error),
            ("Pending", self.stat_pending),
        ]):
            sg.addWidget(QLabel(lbl), 0, col, alignment=Qt.AlignCenter)
            sg.addWidget(val,         1, col, alignment=Qt.AlignCenter)
            sg.itemAtPosition(0, col).widget().setStyleSheet(f"font-size:11px; color:{TEXT_MUTED};")

        lay.addWidget(stats_group)

        return panel

    def _stat_label(self, text, colour=TEXT_PRIMARY) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"font-size: 22px; font-weight: 700; color: {colour};")
        lbl.setAlignment(Qt.AlignCenter)
        return lbl

    def _build_bottom_bar(self) -> QWidget:
        bar = QFrame()
        bar.setFixedHeight(64)
        bar.setStyleSheet(f"""
            QFrame {{
                background: {PANEL_BG};
                border-top: 1px solid {BORDER};
            }}
        """)
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(16, 0, 16, 0)
        lay.setSpacing(10)

        self.progress = QProgressBar()
        self.progress.setFixedHeight(6)
        self.progress.setTextVisible(False)
        self.progress.setValue(0)

        self.btn_run = StyledButton("▶  Run Automation", ACCENT_BLUE)
        self.btn_run.setFixedWidth(180)
        self.btn_run.clicked.connect(self._run_automation)

        self.btn_stop = StyledButton("⛔  Stop", ACCENT_RED)
        self.btn_stop.setFixedWidth(110)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop_automation)

        self.btn_export = StyledButton("⬇  Export Excel", ACCENT_GREEN)
        self.btn_export.setFixedWidth(150)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_excel)

        btn_clear = GhostButton("Clear Log")
        btn_clear.setFixedWidth(100)
        btn_clear.clicked.connect(self.log_box.clear)

        inner = QVBoxLayout()
        inner.setSpacing(4)
        inner.addWidget(QLabel("Progress"), alignment=Qt.AlignLeft)
        inner.addWidget(self.progress)
        inner.itemAt(0).widget().setStyleSheet(f"font-size:11px; color:{TEXT_MUTED};")

        lay.addLayout(inner, 1)
        lay.addWidget(btn_clear)
        lay.addWidget(self.btn_stop)
        lay.addWidget(self.btn_run)
        lay.addWidget(self.btn_export)

        return bar

    # ── Excel I/O ────────────────────────────────────────────────────────────

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self._load_excel(path)

    def _create_sample(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Sample Excel", "sample_tasks.xlsx", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        if not HAS_OPENPYXL:
            QMessageBox.warning(self, "Missing Library",
                "openpyxl is not installed.\n\npip install openpyxl")
            return
        create_sample_excel(path)
        self._log(f"📄 Sample Excel created: {path}", "info")
        self._load_excel(path)

    def _load_excel(self, path: str):
        if not HAS_OPENPYXL:
            QMessageBox.warning(self, "Missing Library",
                "openpyxl is not installed.\n\npip install openpyxl")
            return

        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                self._log("⚠  Excel file is empty.", "warn")
                return

            headers = [str(h).strip().lower().replace(" ", "_") if h else "" for h in rows[0]]
            self._tasks   = []
            self._results = []

            self.table.setRowCount(0)

            for ri, row in enumerate(rows[1:], 1):
                task = dict(zip(headers, row))
                task["_row_index"] = ri - 1
                self._tasks.append(task)
                self._results.append({"status": "Pending", "result": ""})

                r = self.table.rowCount()
                self.table.insertRow(r)
                self.table.setRowHeight(r, 36)

                self.table.setItem(r, 0, self._cell(str(ri)))
                self.table.setItem(r, 1, self._cell(str(task.get("url", ""))))
                self.table.setItem(r, 2, self._cell(str(task.get("action", ""))))
                self.table.setItem(r, 3, self._cell(str(task.get("selector", ""))))

                badge = StatusBadge("Pending")
                self.table.setCellWidget(r, 4, badge)
                self.table.setItem(r, 5, self._cell(""))

            self._excel_path = path
            fname = Path(path).name
            self.file_label.setText(f"📄  {fname}")
            self.file_label.setStyleSheet(f"""
                background: {SURFACE};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 8px 12px;
                color: {TEXT_PRIMARY};
            """)
            self._log(f"✔  Loaded {len(self._tasks)} tasks from '{fname}'.", "info")
            self._update_stats()
            self.btn_run.setEnabled(True)

        except Exception as ex:
            self._log(f"✖  Failed to load Excel: {ex}", "error")
            QMessageBox.critical(self, "Load Error", str(ex))

    def _cell(self, text: str) -> QTableWidgetItem:
        item = QTableWidgetItem(text)
        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        item.setForeground(QColor(TEXT_PRIMARY))
        return item

    # ── Automation control ───────────────────────────────────────────────────

    def _run_automation(self):
        if not self._tasks:
            QMessageBox.information(self, "No Tasks", "Load an Excel file first.")
            return

        # Reset statuses
        for i in range(self.table.rowCount()):
            w = self.table.cellWidget(i, 4)
            if isinstance(w, StatusBadge):
                w.set_status("Pending")
            self.table.setItem(i, 5, self._cell(""))
        for r in self._results:
            r["status"] = "Pending"
            r["result"] = ""

        self.progress.setValue(0)
        self.btn_run.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_export.setEnabled(False)
        self.statusBar().showMessage("Running automation…")
        self._log("─" * 50, "muted")
        self._log(f"▶  Starting {len(self._tasks)} task(s)…", "info")

        headless = self.browser_mode.currentIndex() == 0
        self._worker = AutomationWorker(self._tasks, headless)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.log.connect(self._log)
        self._worker.row_update.connect(self._on_row_update)
        self._worker.progress.connect(self.progress.setValue)
        self._worker.finished.connect(self._on_finished)

        self._thread.start()

    def _stop_automation(self):
        if self._worker:
            self._worker.stop()
        self.btn_stop.setEnabled(False)
        self.statusBar().showMessage("Stopping…")

    def _on_row_update(self, row_idx: int, status: str, result: str):
        if row_idx < self.table.rowCount():
            w = self.table.cellWidget(row_idx, 4)
            if isinstance(w, StatusBadge):
                w.set_status(status)
            self.table.setItem(row_idx, 5, self._cell(result))
        if row_idx < len(self._results):
            self._results[row_idx]["status"] = status
            self._results[row_idx]["result"] = result
        self._update_stats()

    def _on_finished(self, success: bool, msg: str):
        self.btn_run.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.btn_export.setEnabled(True)
        level = "info" if success else "error"
        self._log(f"{'✔' if success else '✖'}  {msg}", level)
        self._log("─" * 50, "muted")
        self.statusBar().showMessage(msg)
        self._update_stats()

        if self._thread:
            self._thread.quit()
            self._thread.wait()

    # ── Export ───────────────────────────────────────────────────────────────

    def _export_excel(self):
        if not HAS_OPENPYXL:
            QMessageBox.warning(self, "Missing Library",
                "openpyxl is not installed.\n\npip install openpyxl")
            return

        default_path = ""
        if self._excel_path:
            p = Path(self._excel_path)
            default_path = str(p.parent / f"{p.stem}_results{p.suffix}")

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Export Results", default_path or "results.xlsx",
            "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        try:
            # Try to base on original file
            if self._excel_path and Path(self._excel_path).exists():
                wb = openpyxl.load_workbook(self._excel_path)
            else:
                wb = openpyxl.Workbook()

            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            headers = [str(h).strip().lower().replace(" ", "_") if h else "" for h in rows[0]] if rows else []

            # Find / create status + result columns
            def _col_for(name):
                for i, h in enumerate(headers):
                    if h == name:
                        return i + 1
                ws.cell(row=1, column=len(headers) + 1, value=name.capitalize())
                headers.append(name)
                return len(headers)

            status_col = _col_for("status")
            result_col = _col_for("result")
            ts_col     = _col_for("timestamp")

            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            done_fill  = PatternFill("solid", fgColor="0A2416")
            err_fill   = PatternFill("solid", fgColor="2E0A09")
            pend_fill  = PatternFill("solid", fgColor="3D2F07")

            for i, r in enumerate(self._results):
                excel_row = i + 2
                status = r.get("status", "Pending")
                result = r.get("result", "")
                ws.cell(row=excel_row, column=status_col, value=status)
                ws.cell(row=excel_row, column=result_col, value=result)
                ws.cell(row=excel_row, column=ts_col, value=ts)

                fill = done_fill if status == "Done" else (err_fill if status == "Error" else pend_fill)
                for c in [status_col, result_col, ts_col]:
                    ws.cell(row=excel_row, column=c).fill = fill

            wb.save(save_path)
            self._log(f"✔  Exported to: {save_path}", "info")
            self.statusBar().showMessage(f"Exported → {save_path}")
            QMessageBox.information(self, "Export Complete",
                f"Results saved to:\n{save_path}")

        except Exception as ex:
            self._log(f"✖  Export failed: {ex}", "error")
            QMessageBox.critical(self, "Export Error", str(ex))

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _log(self, msg: str, level: str = "info"):
        colours = {
            "info":  TEXT_PRIMARY,
            "warn":  ACCENT_AMBER,
            "error": ACCENT_RED,
            "muted": TEXT_MUTED,
        }
        colour = colours.get(level, TEXT_PRIMARY)
        ts = datetime.now().strftime("%H:%M:%S")
        html = (
            f'<span style="color:{TEXT_MUTED};">[{ts}]</span> '
            f'<span style="color:{colour};">{msg}</span>'
        )
        self.log_box.append(html)
        self.log_box.verticalScrollBar().setValue(
            self.log_box.verticalScrollBar().maximum()
        )

    def _update_stats(self):
        total   = len(self._results)
        done    = sum(1 for r in self._results if r["status"] == "Done")
        errors  = sum(1 for r in self._results if r["status"] == "Error")
        pending = sum(1 for r in self._results if r["status"] == "Pending")

        self.stat_total.setText(str(total))
        self.stat_done.setText(str(done))
        self.stat_error.setText(str(errors))
        self.stat_pending.setText(str(pending))


# ══════════════════════════════════════════════════════════════════════════════
#  Entry point
# ══════════════════════════════════════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    palette = QPalette()
    palette.setColor(QPalette.Window,          QColor(DARK_BG))
    palette.setColor(QPalette.WindowText,      QColor(TEXT_PRIMARY))
    palette.setColor(QPalette.Base,            QColor(PANEL_BG))
    palette.setColor(QPalette.AlternateBase,   QColor(SURFACE))
    palette.setColor(QPalette.Text,            QColor(TEXT_PRIMARY))
    palette.setColor(QPalette.Button,          QColor(SURFACE))
    palette.setColor(QPalette.ButtonText,      QColor(TEXT_PRIMARY))
    palette.setColor(QPalette.Highlight,       QColor(HIGHLIGHT))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()