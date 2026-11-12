"""K-SACCR desktop preview: a mock-up of the v1.7.0 desktop application (#694).

Demo data only. The trades are the five CRE99 worked examples of
tests/fixtures and the figures are the values CRE99 prints, from
tests/expected; the counterparties are invented. Nothing is calculated, no
workbook is opened and no file is written.

    python -m pip install PySide6
    python desktop/preview/ksaccr_desktop_preview.py
"""
import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QDate, QRectF, QSortFilterProxyModel, Qt, QTimer
from PySide6.QtGui import QColor, QKeySequence, QPainter, QShortcut, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QComboBox, QDateEdit, QFormLayout, QFrame,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QMainWindow,
                               QMessageBox, QProgressBar, QPushButton, QStackedWidget, QStatusBar, QTableView,
                               QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget)

ROOT = Path(__file__).resolve().parents[2]
TRADE_FIELDS = ("id", "asset_class", "sub_class", "risk_factor", "instrument", "direction", "option_type",
                "notional_amount", "market_value_amount", "maturity_date", "description")


def load_demo():
    """The five CRE99 examples: trades from the fixtures, figures as published."""
    data = {"netting_sets": [], "trades": [], "results": []}
    for n in range(1, 6):
        fixture = json.loads((ROOT / f"tests/fixtures/cre99-example-{n}.json").read_text(encoding="utf-8"))
        expected = json.loads((ROOT / f"tests/expected/cre99-example-{n}.bcbs.json").read_text(encoding="utf-8"))
        ns = fixture["netting_set"]
        data["netting_sets"].append({
            "id": ns["id"], "example": n, "margined": ns["margined"],
            "vm": ns["variation_margin_net_amount"], "ic": ns["independent_collateral_net_amount"],
            "th": ns["threshold_amount"], "mta": ns["minimum_transfer_amount"],
            "mpor": ns.get("remargin_period_business_days"), "description": fixture["description"]})
        for trade in fixture["trades"]:
            data["trades"].append({key: trade[key] for key in TRADE_FIELDS} | {"netting_set": ns["id"]})
        for output in expected["outputs"]:
            data["results"].append({
                "netting_set": ns["id"], "quantity": output["quantity"], "rule": output["rule"],
                "trade": output.get("trade"), "value": output["value"], "unit": output["unit"],
                "locator": output["reference"]["locator"]})
    return data


DATA = load_demo()

# ---------------------------------------------------------------------------
# Demo counterparties (synthetic) and helpers
# ---------------------------------------------------------------------------
COUNTERPARTIES = [
    {"id": "CP-001", "name": "Demo Bank A", "sector": "CI", "netting_sets": ["EX1", "EX4"]},
    {"id": "CP-002", "name": "Demo Insurer B", "sector": "OFC", "netting_sets": ["EX2"]},
    {"id": "CP-003", "name": "Demo Utility C", "sector": "NFC", "netting_sets": ["EX3"]},
    {"id": "CP-004", "name": "Demo Fund D", "sector": "OFC", "netting_sets": ["EX5"]},
]
SECTOR_NAMES = {"CB": "Central banks", "GG": "General governments", "CI": "Credit institutions",
                "OFC": "Other financial corporations", "NFC": "Non-financial corporations",
                "HH": "Households"}
QUANTITY_LABELS = {
    "replacement_cost": "Replacement cost (RC)", "aggregate_add_on": "Aggregate add-on",
    "add_on.interest_rate": "Add-on, interest rate", "multiplier": "PFE multiplier",
    "exposure_value": "Exposure at default (EAD)", "margin_period_of_risk": "Margin period of risk",
    "adjusted_notional": "Adjusted notional", "supervisory_delta": "Supervisory delta",
}
ORDER = ["replacement_cost", "add_on.interest_rate", "aggregate_add_on", "multiplier",
         "margin_period_of_risk", "adjusted_notional", "supervisory_delta", "exposure_value"]


def counterparty_of(ns_id):
    for cp in COUNTERPARTIES:
        if ns_id in cp["netting_sets"]:
            return cp
    return None


def ead_of(ns_id):
    for r in DATA["results"]:
        if r["netting_set"] == ns_id and r["quantity"] == "exposure_value":
            return r["value"]
    return None


def fmt_amount(value):
    return f"{value:,.0f}"


def fmt_value(r):
    if r["unit"] == "USD":
        return fmt_amount(r["value"])
    if r["quantity"] == "margin_period_of_risk":
        return f"{r['value']:g} business days"
    return f"{r['value']:g}"


# ---------------------------------------------------------------------------
# Theme
# ---------------------------------------------------------------------------
NAVY, BLUE, BG, CARD, BORDER, TEXT, MUTED = ("#0B2545", "#1F6FEB", "#F4F6FA", "#FFFFFF",
                                             "#DCE3EE", "#1B2430", "#5B6B82")
STYLE = f"""
QMainWindow, QWidget#page {{ background: {BG}; }}
QWidget {{ font-family: 'Segoe UI', 'Carlito', 'DejaVu Sans'; font-size: 10pt; color: {TEXT}; }}
QListWidget#nav {{ background: {NAVY}; border: none; padding-top: 8px; outline: 0; }}
QListWidget#nav::item {{ color: #C9D6EA; padding: 10px 18px; border-left: 3px solid transparent; }}
QListWidget#nav::item:selected {{ background: #13345F; color: white; border-left: 3px solid {BLUE}; }}
QListWidget#nav::item:hover {{ background: #102E55; }}
QLabel#brand {{ background: {NAVY}; color: white; font-size: 15pt; font-weight: 600; padding: 18px 18px 6px 18px; }}
QLabel#brandsub {{ background: {NAVY}; color: #8FA6C8; font-size: 8.5pt; padding: 0 18px 12px 18px; }}
QLabel#title {{ font-size: 16pt; font-weight: 600; }}
QLabel#subtitle {{ color: {MUTED}; }}
QFrame#card {{ background: {CARD}; border: 1px solid {BORDER}; border-radius: 8px; }}
QLabel#kpi {{ font-size: 18pt; font-weight: 600; color: {NAVY}; }}
QLabel#kpilabel {{ color: {MUTED}; font-size: 9pt; }}
QLabel#badge {{ background: #E7F0FF; color: {BLUE}; border-radius: 9px; padding: 2px 10px; font-size: 8.5pt; }}
QLabel#warn {{ background: #FFF4E5; color: #8A5300; border: 1px solid #F3D19E; border-radius: 6px; padding: 8px 10px; }}
QTableView, QTreeWidget, QListWidget#log {{ background: {CARD}; border: 1px solid {BORDER}; border-radius: 6px;
    gridline-color: #EEF2F7; selection-background-color: #DCE9FF; selection-color: {TEXT}; }}
QTableView::item {{ padding: 4px 10px; }}
QHeaderView::section {{ background: #F0F3F8; border: none; border-bottom: 1px solid {BORDER};
    padding: 6px; font-weight: 600; color: {MUTED}; }}
QPushButton {{ background: {CARD}; border: 1px solid {BORDER}; border-radius: 6px; padding: 7px 14px; }}
QPushButton:hover {{ border-color: {BLUE}; }}
QPushButton#primary {{ background: {BLUE}; color: white; border: none; font-weight: 600; }}
QPushButton#primary:disabled {{ background: #9DB9E8; }}
QLineEdit, QComboBox, QDateEdit {{ background: {CARD}; border: 1px solid {BORDER}; border-radius: 6px; padding: 5px 8px; }}
QProgressBar {{ background: #E6EBF2; border: none; border-radius: 4px; height: 8px; text-align: center; }}
QProgressBar::chunk {{ background: {BLUE}; border-radius: 4px; }}
QStatusBar {{ background: {CARD}; border-top: 1px solid {BORDER}; color: {MUTED}; }}
"""


# ---------------------------------------------------------------------------
# Widgets
# ---------------------------------------------------------------------------
def card(layout_cls=QVBoxLayout):
    frame = QFrame()
    frame.setObjectName("card")
    layout = layout_cls(frame)
    layout.setContentsMargins(16, 14, 16, 14)
    return frame, layout


def page(title, subtitle):
    widget = QWidget()
    widget.setObjectName("page")
    layout = QVBoxLayout(widget)
    layout.setContentsMargins(28, 22, 28, 22)
    layout.setSpacing(14)
    head = QLabel(title)
    head.setObjectName("title")
    sub = QLabel(subtitle)
    sub.setObjectName("subtitle")
    sub.setWordWrap(True)
    layout.addWidget(head)
    layout.addWidget(sub)
    return widget, layout


def table(headers, rows, numeric=()):
    model = QStandardItemModel(0, len(headers))
    model.setHorizontalHeaderLabels(headers)
    for row in rows:
        items = []
        for col, value in enumerate(row):
            item = QStandardItem("" if value is None else str(value))
            item.setEditable(False)
            if col in numeric:
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            items.append(item)
        model.appendRow(items)
    view = QTableView()
    view.setModel(model)
    view.verticalHeader().setVisible(False)
    view.setSelectionBehavior(QAbstractItemView.SelectRows)
    view.setSelectionMode(QAbstractItemView.SingleSelection)
    view.setAlternatingRowColors(False)
    view.setShowGrid(False)
    view.horizontalHeader().setStretchLastSection(True)
    view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    view.horizontalHeader().setSectionResizeMode(len(headers) - 1, QHeaderView.Stretch)
    return view, model


class BarChart(QWidget):
    """EAD per netting set, drawn with QPainter (no chart dependency)."""

    def __init__(self, values):
        super().__init__()
        self.values = values
        self.setMinimumHeight(200)

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        width, height = self.width(), self.height()
        top, label_w, right = 6, 70, 110
        bar_h = min((height - top) / max(len(self.values), 1), 44)
        peak = max(v for _, v in self.values) or 1
        for i, (label, value) in enumerate(self.values):
            y = top + i * bar_h
            painter.setPen(QColor(MUTED))
            painter.drawText(QRectF(0, y, label_w - 8, bar_h), Qt.AlignRight | Qt.AlignVCenter, label)
            length = (width - label_w - right) * value / peak
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor(BLUE))
            painter.drawRoundedRect(QRectF(label_w, y + bar_h * 0.22, max(length, 2), bar_h * 0.56), 4, 4)
            painter.setPen(QColor(TEXT))
            painter.drawText(QRectF(label_w + length + 8, y, right, bar_h), Qt.AlignLeft | Qt.AlignVCenter,
                             fmt_amount(value))


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
SECTIONS = ["Dashboard", "Counterparties", "Trades", "Netting Sets", "Calculation", "Results",
            "Reports", "Settings"]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("K-SACCR Desktop - preview")
        self.resize(1280, 800)
        self.last_run = None

        root = QWidget()
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        side = QWidget()
        side.setFixedWidth(220)
        side.setStyleSheet(f"background: {NAVY};")
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.setSpacing(0)
        brand = QLabel("K-SACCR")
        brand.setObjectName("brand")
        brand_sub = QLabel("SA-CCR exposure desktop\npreview, demo data")
        brand_sub.setObjectName("brandsub")
        self.nav = QListWidget()
        self.nav.setObjectName("nav")
        self.nav.addItems(SECTIONS)
        side_layout.addWidget(brand)
        side_layout.addWidget(brand_sub)
        side_layout.addWidget(self.nav, 1)

        self.stack = QStackedWidget()
        for build in (self.dashboard, self.counterparties, self.trades, self.netting_sets,
                      self.calculation, self.results, self.reports, self.settings):
            self.stack.addWidget(build())
        self.nav.currentRowChanged.connect(self.stack.setCurrentIndex)
        self.nav.setCurrentRow(0)

        outer.addWidget(side)
        outer.addWidget(self.stack, 1)
        self.setCentralWidget(root)

        status = QStatusBar()
        status.showMessage("Demo mode  |  no workbook opened  |  engine: demo (published CRE99 values)  |  "
                           "regime: Basel  |  preview build, not a release")
        self.setStatusBar(status)
        for i, name in enumerate(SECTIONS):
            QShortcut(QKeySequence(f"Ctrl+{i + 1}"), self, activated=lambda i=i: self.nav.setCurrentRow(i))

    # -- pages -----------------------------------------------------------------
    def dashboard(self):
        widget, layout = page("Dashboard", "Demo portfolio: the five worked examples of CRE99, one netting "
                              "set each. Figures are the published values, in USD.")
        row = QHBoxLayout()
        total = sum(ead_of(n["id"]) for n in DATA["netting_sets"])
        for value, label in ((fmt_amount(total), "Total EAD (USD)"),
                             (str(len(DATA["netting_sets"])), "Netting sets"),
                             (str(len(DATA["trades"])), "Trades"),
                             ("0", "Validation errors")):
            frame, inner = card()
            number = QLabel(value)
            number.setObjectName("kpi")
            caption = QLabel(label)
            caption.setObjectName("kpilabel")
            inner.addWidget(number)
            inner.addWidget(caption)
            row.addWidget(frame)
        layout.addLayout(row)

        lower = QHBoxLayout()
        frame, inner = card()
        heading = QLabel("EAD by netting set")
        heading.setStyleSheet("font-weight: 600;")
        inner.addWidget(heading)
        inner.addWidget(BarChart([(n["id"], ead_of(n["id"])) for n in DATA["netting_sets"]]), 1)
        lower.addWidget(frame, 3)

        frame, inner = card()
        heading = QLabel("Latest run")
        heading.setStyleSheet("font-weight: 600;")
        inner.addWidget(heading)
        form = QFormLayout()
        self.dash_run = QLabel("not run yet: open Calculation")
        form.addRow("Run:", self.dash_run)
        form.addRow("Engine:", QLabel("demo (published values)"))
        form.addRow("Regime:", QLabel("Basel (CRE52)"))
        form.addRow("Parameter set:", QLabel("REGULATORY"))
        form.addRow("Input snapshot:", QLabel("demo-cre99 (21 trades)"))
        inner.addLayout(form)
        inner.addStretch(1)
        lower.addWidget(frame, 2)
        layout.addLayout(lower, 1)
        return widget

    def counterparties(self):
        widget, layout = page("Counterparties", "Who each netting set faces. The sector is the FINREP "
                              "counterparty sector, for reporting only (#674).")
        rows = [(c["id"], c["name"], c["sector"], SECTOR_NAMES[c["sector"]], ", ".join(c["netting_sets"]))
                for c in COUNTERPARTIES]
        view, _ = table(["ID", "Name", "Sector", "Sector name", "Netting sets"], rows)
        buttons = QHBoxLayout()
        for text in ("New counterparty", "Edit", "Import CSV"):
            buttons.addWidget(QPushButton(text))
        buttons.addStretch(1)
        layout.addLayout(buttons)
        layout.addWidget(view, 1)
        return widget

    def trades(self):
        widget, layout = page("Trades", "Trade inputs, validated on entry against the CSV contract. "
                              "Select a trade to see its fields.")
        bar = QHBoxLayout()
        search = QLineEdit()
        search.setPlaceholderText("Search trades (ID, description, risk factor)")
        asset = QComboBox()
        asset.addItems(["All asset classes", "interest_rate", "credit", "commodity"])
        bar.addWidget(search, 3)
        bar.addWidget(asset, 1)
        bar.addWidget(QPushButton("New trade"))
        bar.addWidget(QPushButton("Import CSV"))
        layout.addLayout(bar)

        rows = [(t["id"], t["netting_set"], t["asset_class"], t["instrument"], t["direction"],
                 fmt_amount(t["notional_amount"]), fmt_amount(t["market_value_amount"]), t["maturity_date"],
                 t["description"]) for t in DATA["trades"]]
        view, model = table(["Trade", "Netting set", "Asset class", "Instrument", "Direction", "Notional",
                             "MtM", "Maturity", "Description"], rows, numeric=(5, 6))
        proxy = QSortFilterProxyModel()
        proxy.setSourceModel(model)
        proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        proxy.setFilterKeyColumn(-1)
        view.setModel(proxy)
        view.setSortingEnabled(True)
        view.sortByColumn(0, Qt.AscendingOrder)
        search.textChanged.connect(proxy.setFilterFixedString)
        asset.currentTextChanged.connect(
            lambda text: proxy.setFilterFixedString("" if text.startswith("All") else text))

        split = QHBoxLayout()
        split.addWidget(view, 3)
        frame, inner = card()
        heading = QLabel("Trade details")
        heading.setStyleSheet("font-weight: 600;")
        inner.addWidget(heading)
        self.trade_form = QFormLayout()
        inner.addLayout(self.trade_form)
        state = QLabel("Validation: VALID, no diagnostics")
        state.setObjectName("badge")
        inner.addWidget(state)
        inner.addStretch(1)
        frame.setMinimumWidth(300)
        split.addWidget(frame, 1)
        layout.addLayout(split, 1)

        def show(current, _previous=None):
            while self.trade_form.rowCount():
                self.trade_form.removeRow(0)
            if not current.isValid():
                return
            trade_id = proxy.index(current.row(), 0).data()
            trade = next(t for t in DATA["trades"] if t["id"] == trade_id)
            for key in ("id", "netting_set", "asset_class", "sub_class", "risk_factor", "instrument",
                        "direction", "option_type", "notional_amount", "market_value_amount", "maturity_date"):
                value = trade.get(key)
                if isinstance(value, (int, float)):
                    value = fmt_amount(value)
                self.trade_form.addRow(key.replace("_", " ").capitalize() + ":",
                                       QLabel("-" if value in (None, "") else str(value)))

        view.selectionModel().currentRowChanged.connect(show)
        view.selectRow(0)
        return widget

    def netting_sets(self):
        widget, layout = page("Netting Sets", "Margining, collateral and clearing terms of each netting "
                              "set. Collateral after haircuts, in the reporting currency.")
        rows = []
        for n in DATA["netting_sets"]:
            cp = counterparty_of(n["id"])
            count = sum(1 for t in DATA["trades"] if t["netting_set"] == n["id"])
            rows.append((n["id"], cp["name"] if cp else "", "Yes" if n["margined"] else "No",
                         fmt_amount(n["vm"]), fmt_amount(n["ic"]), fmt_amount(n["th"]), fmt_amount(n["mta"]),
                         n["mpor"] or "", count, n["description"]))
        view, _ = table(["Netting set", "Counterparty", "Margined", "VM", "IC", "Threshold", "MTA",
                         "Remargin days", "Trades", "Description"], rows, numeric=(3, 4, 5, 6, 7, 8))
        layout.addWidget(view, 1)
        return widget

    def calculation(self):
        widget, layout = page("Calculation", "Validate the inputs and run SA-CCR with the selected engine. "
                              "The window stays responsive; a run can be cancelled.")
        frame, inner = card()
        form = QFormLayout()
        self.engine = QComboBox()
        self.engine.addItems(["Demo: published CRE99 values", "Excel/VBA engine (hidden workbook): not found",
                              "Python engine: planned (#450)"])
        regime = QComboBox()
        regime.addItems(["Basel (CRE52)", "CRR (Articles 274-280f)"])
        date = QDateEdit(QDate(2026, 9, 30))
        date.setCalendarPopup(True)
        date.setDisplayFormat("yyyy-MM-dd")
        params = QComboBox()
        params.addItems(["REGULATORY"])
        for field in (self.engine, regime, date, params):
            field.setMaximumWidth(460)
        form.addRow("Engine:", self.engine)
        form.addRow("Regime:", regime)
        form.addRow("Valuation date:", date)
        form.addRow("Parameter set:", params)
        inner.addLayout(form)
        buttons = QHBoxLayout()
        validate = QPushButton("Validate inputs")
        self.run_button = QPushButton("Run SA-CCR")
        self.run_button.setObjectName("primary")
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setEnabled(False)
        buttons.addWidget(validate)
        buttons.addWidget(self.run_button)
        buttons.addWidget(self.cancel_button)
        buttons.addStretch(1)
        inner.addLayout(buttons)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        inner.addWidget(self.progress)
        layout.addWidget(frame)

        self.log = QListWidget()
        self.log.setObjectName("log")
        layout.addWidget(QLabel("Run log"))
        layout.addWidget(self.log, 1)

        self.steps = []
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.step)
        self.run_button.clicked.connect(self.start_run)
        self.cancel_button.clicked.connect(self.cancel_run)
        validate.clicked.connect(lambda: self.log.addItem(
            f"{datetime.now():%H:%M:%S}  Validation: 5 netting sets, 21 trades, 0 errors, 0 warnings"))
        self.engine.currentIndexChanged.connect(
            lambda i: self.run_button.setEnabled(i == 0))
        return widget

    def start_run(self):
        self.log.clear()
        self.steps = [
            (10, "Snapshot of the inputs taken: 5 netting sets, 21 trades"),
            (25, "Validation: 0 errors, 0 warnings"),
            (45, "Engine: demo, returning the published CRE99 values"),
            (70, "Netting sets EX1 to EX5 calculated"),
            (90, "Results bound to the input snapshot"),
            (100, "Run complete: total EAD " + fmt_amount(sum(ead_of(n['id']) for n in DATA['netting_sets'])) + " USD"),
        ]
        self.progress.setValue(0)
        self.run_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.timer.start(350)

    def step(self):
        if not self.steps:
            self.timer.stop()
            self.cancel_button.setEnabled(False)
            self.run_button.setEnabled(True)
            self.last_run = datetime.now()
            text = f"{self.last_run:%Y-%m-%d %H:%M:%S}, completed"
            self.dash_run.setText(text)
            self.results_meta.setText(f"Run {text}  |  engine: demo  |  regime: Basel  |  parameters: "
                                      "REGULATORY  |  inputs: demo-cre99")
            return
        value, text = self.steps.pop(0)
        self.progress.setValue(value)
        self.log.addItem(f"{datetime.now():%H:%M:%S}  {text}")

    def cancel_run(self):
        self.timer.stop()
        self.steps = []
        self.log.addItem(f"{datetime.now():%H:%M:%S}  Run cancelled: no results written")
        self.cancel_button.setEnabled(False)
        self.run_button.setEnabled(True)

    def results(self):
        widget, layout = page("Results", "What the engine returned, by netting set, with the rule and the "
                              "source of each figure. Nothing is recomputed in the interface.")
        self.results_meta = QLabel("No run yet: the figures below are the demo engine's.")
        self.results_meta.setObjectName("badge")
        layout.addWidget(self.results_meta, 0, Qt.AlignLeft)
        tree = QTreeWidget()
        tree.setHeaderLabels(["Netting set / quantity", "Value", "Rule", "Source"])
        for n in DATA["netting_sets"]:
            cp = counterparty_of(n["id"])
            ead = next(r for r in DATA["results"]
                       if r["netting_set"] == n["id"] and r["quantity"] == "exposure_value")
            top = QTreeWidgetItem([f"{n['id']}  ({cp['name'] if cp else ''}): EAD", fmt_amount(ead["value"]),
                                   ead["rule"], ead["locator"]])
            font = top.font(0)
            font.setBold(True)
            for c in range(4):
                top.setFont(c, font)
            top.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
            items = [r for r in DATA["results"]
                     if r["netting_set"] == n["id"] and r["quantity"] != "exposure_value"]
            items.sort(key=lambda r: ORDER.index(r["quantity"]))
            for r in items:
                label = QUANTITY_LABELS.get(r["quantity"], r["quantity"])
                if r["trade"]:
                    label += f", trade {r['trade']}"
                child = QTreeWidgetItem([label, fmt_value(r), r["rule"], r["locator"]])
                child.setTextAlignment(1, Qt.AlignRight | Qt.AlignVCenter)
                top.addChild(child)
            tree.addTopLevelItem(top)
        tree.expandAll()
        tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(tree, 1)
        note = QLabel("Only the quantities CRE99 prints are shown; the Excel/VBA and Python engines will "
                      "return every intermediate (RC, add-ons by asset class, multiplier, PFE).")
        note.setObjectName("subtitle")
        note.setWordWrap(True)
        layout.addWidget(note)
        return widget

    def reports(self):
        widget, layout = page("Reports", "Exports keep the run's identity: input snapshot, engine and "
                              "version, parameter set and timestamp.")
        for title, text in (("Results (CSV)", "One row per netting set and quantity, machine-readable."),
                            ("Evidence (CSV)", "The evidence record of the run, as the workbook exports it."),
                            ("Run report (PDF)", "Dashboard, results and diagnostics for review.")):
            frame, inner = card(QHBoxLayout)
            label = QLabel(f"<b>{title}</b><br><span style='color:{MUTED}'>{text}</span>")
            inner.addWidget(label, 1)
            button = QPushButton("Export...")
            button.clicked.connect(lambda _=False, t=title: QMessageBox.information(
                self, "Preview", f"{t}: export is not part of this preview."))
            inner.addWidget(button)
            layout.addWidget(frame)
        layout.addStretch(1)
        return widget

    def settings(self):
        widget, layout = page("Settings", "Local settings of this installation.")
        frame, inner = card()
        form = QFormLayout()
        theme = QComboBox()
        theme.addItems(["Light", "Dark", "Follow Windows"])
        units = QComboBox()
        units.addItems(["Units", "Thousands", "Millions"])
        folder = QLineEdit(r"%USERPROFILE%\Documents\K-SACCR")
        engine = QLineEdit("Excel/VBA engine: K-SACCR_1.7.0.xlsm (not installed in this preview)")
        engine.setReadOnly(True)
        form.addRow("Theme:", theme)
        form.addRow("Amounts shown in:", units)
        form.addRow("Data folder:", folder)
        form.addRow("Engine workbook:", engine)
        inner.addLayout(form)
        layout.addWidget(frame)
        warn = QLabel("Preview: no data leaves this computer, nothing is saved, and the figures are the "
                      "published CRE99 values, not a calculation.")
        warn.setObjectName("warn")
        warn.setWordWrap(True)
        layout.addWidget(warn)
        layout.addStretch(1)
        return widget


def main():
    parser = argparse.ArgumentParser(description="K-SACCR desktop preview (demo data only)")
    parser.add_argument("--screenshots", metavar="FOLDER", help="save one PNG per section and exit")
    args = parser.parse_args()
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE)
    window = MainWindow()
    window.show()
    if args.screenshots:
        folder = Path(args.screenshots)
        folder.mkdir(parents=True, exist_ok=True)

        def shoot():
            window.start_run()
            while window.timer.isActive() or window.steps:
                app.processEvents()
                time.sleep(0.05)
            app.processEvents()
            for i, name in enumerate(SECTIONS):
                window.nav.setCurrentRow(i)
                app.processEvents()
                window.grab().save(str(folder / f"{i + 1:02d}-{name.lower().replace(' ', '-')}.png"))
            app.quit()

        QTimer.singleShot(300, shoot)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
