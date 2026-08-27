"""Shared helpers and base classes used by the quotation/delivery-note/doc modules:
.ui loading, PDF/print export, Excel list logging, window positioning/sizing, and the
compact entry-form base class."""

import json
import os
import re

import openpyxl
import win32com.client
from PyQt5 import uic
from PyQt5.QtCore import QDate, QSize, Qt
from PyQt5.QtGui import QFontMetricsF, QPainter
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
from PyQt5.QtWidgets import QApplication, QMainWindow, QScrollArea, QTextBrowser, QTextEdit

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UI_DIR = os.path.join(BASE_DIR, "UI")
CLICK_SOUND = os.path.join(BASE_DIR, "sounds", "click.wav")

XL_TYPE_PDF = 0

# Rounded, bordered look for the compact forms' input fields (Quotation.ui / DeliveryNote.ui /
# Doc.ui), replacing the default flat QTextBrowser frame. Paired with turning the vertical
# scrollbar off wherever this is applied, so a stray scrollbar thumb never appears either.
COMPACT_INPUT_STYLESHEET = (
    "QTextBrowser {"
    " background: white;"
    " border: 1px solid #c3c3c3;"
    " border-radius: 8px;"
    " padding: 0 10px;"
    "}"
)

# --- Quotation.ui / DeliveryNote.ui (compact entry forms): identical item-row grid and
# widget names in both files, only differing in the button cluster that slides down
# when rows don't fit, and in which full-page window/number-prefix Go hands off to.
COMPACT_ITEM_ROW_COLUMNS = (
    ("item", 80, 201),
    ("description", 290, 431),
    ("quantity", 730, 91),
    ("unit_price", 830, 201),
)
COMPACT_ITEM_ROW_TOP = 400
COMPACT_ITEM_ROW_HEIGHT = 55
COMPACT_ITEM_ROW_SPACING = 55
COMPACT_ITEM_ROW_RIGHT_ALIGNED_COLUMNS = ("quantity", "unit_price")
COMPACT_ITEM_ROW_GAP = 20
COMPACT_BOTTOM_CLUSTER_WIDGET_NAMES = (
    "Plus_pushButton",
    "Minus_pushButton",
    "Go_pushButton",
    "Back_pushButton",
    "pushButton_3",
    "label_9",
    "label_11",
    "label_12",
)

# Newer Qt Designer saves enums in Qt6's scoped form (e.g. "Qt::Orientation::Horizontal"),
# which PyQt5's uic parser (Qt5) does not understand. Normalize to the old flat form
# before parsing so re-saving in Designer can't break loading.
_SCOPED_ENUM_RE = re.compile(r"\b(\w+)::\w+::(\w+)\b")


def load_ui(ui_path, baseinstance):
    with open(ui_path, "r", encoding="utf-8") as f:
        content = f.read()
    normalized = _SCOPED_ENUM_RE.sub(r"\1::\2", content)
    if normalized == content:
        uic.loadUi(ui_path, baseinstance)
        return

    tmp_path = ui_path + ".normalized.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(normalized)
    try:
        uic.loadUi(tmp_path, baseinstance)
    finally:
        os.remove(tmp_path)


def center_text_vertically(text_edit):
    """Vertically center text_edit's content, and keep it centered as the text
    changes (typing, prefill, live amount recalculation, 문서참조 loads, etc.)."""

    doc = text_edit.document()
    # The document's own top/bottom margin sits inside its reported size, so splitting
    # that size in half as a single top-only viewport margin double-counts it on top
    # and leaves it (plus any leftover viewport space) to silently pad the bottom
    # instead - text ends up sitting visibly above center. Zero it out and center
    # against the frame explicitly, in equal top/bottom viewport margins, instead.
    doc.setDocumentMargin(0)

    def recenter():
        no_wrap = text_edit.lineWrapMode() == QTextEdit.NoWrap
        doc.setTextWidth(-1 if no_wrap else text_edit.width())
        if no_wrap:
            # QTextDocument's own line-box height bakes in each font's line-spacing/leading,
            # which for some (esp. CJK) fonts runs far taller than the glyphs actually drawn -
            # centering against that box clamps the margin to 0 and pins single-line text to
            # the top with a large empty gap below. QFontMetrics' ascent+descent tracks the
            # visible ink instead, so single-line fields center the way a QLineEdit would.
            content_height = QFontMetricsF(text_edit.currentFont()).height()
        else:
            content_height = doc.size().height()
        # contentsRect() (not frameWidth()) accounts for the vertical inset correctly even
        # when a stylesheet gives the box asymmetric horizontal/vertical border+padding -
        # frameWidth() collapses that into one number and gets the vertical space wrong.
        available = text_edit.contentsRect().height()
        margin = int(max(0, (available - content_height) / 2))
        text_edit.setViewportMargins(0, margin, 0, margin)

    recenter()
    if not text_edit.property("_center_text_live"):
        text_edit.setProperty("_center_text_live", True)
        text_edit.textChanged.connect(recenter)


def next_document_number(prefix, seq_file):
    date_str = QDate.currentDate().toString("yyyyMMdd")
    seq = 0
    if os.path.exists(seq_file):
        with open(seq_file, "r", encoding="utf-8") as f:
            state = json.load(f)
        if state.get("date") == date_str:
            seq = (state.get("seq", -1) + 1) % 1000

    with open(seq_file, "w", encoding="utf-8") as f:
        json.dump({"date": date_str, "seq": seq}, f)

    return f"{prefix}-{date_str}-{seq:03d}"


def _advance_on_enter(widget, advance):
    original_key_press = widget.keyPressEvent

    def key_press(event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            advance()
            event.accept()
            return
        original_key_press(event)

    widget.keyPressEvent = key_press


def _as_number(text):
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return text


def export_pdf(xlsx_path, pdf_path):
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(xlsx_path)
        try:
            workbook.ExportAsFixedFormat(XL_TYPE_PDF, pdf_path)
        finally:
            workbook.Close(False)
    finally:
        excel.Quit()


def _paint_widget_onto_printer(widget, printer):
    pixmap = widget.grab()
    painter = QPainter(printer)
    try:
        page_rect = printer.pageRect()
        scaled = pixmap.scaled(page_rect.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        x = page_rect.x() + (page_rect.width() - scaled.width()) // 2
        y = page_rect.y() + (page_rect.height() - scaled.height()) // 2
        painter.drawPixmap(x, y, scaled)
    finally:
        painter.end()


def export_widget_pdf(widget, pdf_path):
    """Render exactly what's on screen (the document page as filled in) to a PDF,
    independent of the Excel-generated one."""
    printer = QPrinter(QPrinter.HighResolution)
    printer.setOutputFormat(QPrinter.PdfFormat)
    printer.setOutputFileName(pdf_path)
    printer.setPageSize(QPrinter.A4)
    printer.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
    _paint_widget_onto_printer(widget, printer)


def print_widget(widget, parent):
    """Open the OS print dialog and, if accepted, send exactly what's on screen (the
    same rendering export_widget_pdf() captures) to the chosen printer."""
    printer = QPrinter(QPrinter.HighResolution)
    printer.setPageSize(QPrinter.A4)
    dialog = QPrintDialog(printer, parent)
    if dialog.exec_() != QPrintDialog.Accepted:
        return
    _paint_widget_onto_printer(widget, printer)


def export_document_files(workbook, doc_no, save_dir, source_widget):
    """Save `workbook` as {doc_no}.xlsx into save_dir, export it to {doc_no}.pdf via
    Excel, and grab a {doc_no}_screen.pdf of source_widget. Returns the xlsx path."""
    os.makedirs(save_dir, exist_ok=True)

    xlsx_path = os.path.join(save_dir, f"{doc_no}.xlsx")
    workbook.save(xlsx_path)

    pdf_path = os.path.join(save_dir, f"{doc_no}.pdf")
    export_pdf(xlsx_path, pdf_path)

    screen_pdf_path = os.path.join(save_dir, f"{doc_no}_screen.pdf")
    export_widget_pdf(source_widget, screen_pdf_path)

    return xlsx_path


def append_to_list_xlsx(list_xlsx_path, date_text, name, item, description, quantity, unit_price):
    workbook = openpyxl.load_workbook(list_xlsx_path)
    sheet = workbook.active

    row = 3
    while sheet[f"A{row}"].value is not None:
        row += 1

    sheet[f"A{row}"] = row - 2
    sheet[f"B{row}"] = date_text
    sheet[f"C{row}"] = name
    sheet[f"D{row}"] = item
    sheet[f"E{row}"] = description
    sheet[f"J{row}"] = quantity
    sheet[f"K{row}"] = unit_price
    sheet[f"L{row}"] = f"=J{row}*K{row}"

    workbook.save(list_xlsx_path)


def move_centered_on(window, reference):
    """Center `window` on `reference`'s current frame, then clamp it to stay fully
    within the available area of whichever screen it landed on. Without this, on a
    multi-monitor setup a newly opened window can end up positioned partly or wholly
    off every visible screen (e.g. following a parent window that defaulted onto a
    secondary monitor) with no way for the user to reach or even notice it."""
    center = reference.frameGeometry().center()
    frame = window.frameGeometry()
    frame.moveCenter(center)

    screen = QApplication.screenAt(center) or window.screen() or QApplication.primaryScreen()
    available = screen.availableGeometry()
    x = min(max(frame.left(), available.left()), max(available.left(), available.right() - frame.width() + 1))
    y = min(max(frame.top(), available.top()), max(available.top(), available.bottom() - frame.height() + 1))
    window.move(x, y)


def make_window_scrollable(window):
    """Wrap window's centralwidget in a QScrollArea and cap the window to the screen's
    available size. The full-page templates (Doc2.ui / Quotation2.ui / DeliveryNote2.ui)
    are taller than most displays (up to ~1519px), so opened at their natural fixed size
    the bottom controls (Save/Print/...) land off-screen with no way to reach them."""
    # Right after load_ui, the window already has its .ui-declared size, but the
    # centralwidget hasn't been laid out to fill it yet (that happens lazily) - so read
    # the size from the window, not the not-yet-sized centralwidget, and stamp it on
    # explicitly since setWidgetResizable(False) below needs a real size to work with.
    # Some templates (e.g. Doc2.ui's approval-stage row/underline) place child widgets
    # past the .ui file's own declared window width, so also grow to childrenRect() -
    # otherwise that overflow is silently clipped with no way to reach it.
    full_size = window.size()
    content = window.centralWidget()
    children_rect = content.childrenRect()
    full_size = full_size.expandedTo(
        QSize(children_rect.x() + children_rect.width(), children_rect.y() + children_rect.height())
    )
    content.resize(full_size)

    scroll_area = QScrollArea()
    scroll_area.setWidget(content)
    scroll_area.setWidgetResizable(False)
    scroll_area.setFrameShape(QScrollArea.NoFrame)
    window.setCentralWidget(scroll_area)

    available = QApplication.primaryScreen().availableGeometry()
    v_scrollbar_allowance = scroll_area.verticalScrollBar().sizeHint().width()
    h_scrollbar_allowance = scroll_area.horizontalScrollBar().sizeHint().height()
    width = min(full_size.width() + v_scrollbar_allowance, available.width())
    height = min(full_size.height() + h_scrollbar_allowance, available.height() - 60)
    window.setFixedSize(width, height)


# --- Shared logic for the "full-page" document windows (Quotation2.ui / DeliveryNote2.ui):
# fixed item rows, live QTY*UnitPrice -> Amount, and a Save that fills the Excel
# template, exports it (+ a screen-capture PDF), and logs each item row to the running list.

def style_item_form(window, item_rows, extra_single_line_widgets):
    for text_edit in window.centralwidget.findChildren(QTextEdit):
        center_text_vertically(text_edit)

    single_line_widgets = list(extra_single_line_widgets)
    for row in item_rows:
        single_line_widgets.extend(getattr(window, name) for name in row)
    for widget in single_line_widgets:
        widget.setLineWrapMode(QTextEdit.NoWrap)
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

    for *_rest, amount_name in item_rows:
        getattr(window, amount_name).setReadOnly(True)


def connect_amount_calculations(window, item_rows):
    for _item_name, _desc_name, qty_name, price_name, amount_name in item_rows:
        qty_widget = getattr(window, qty_name)
        price_widget = getattr(window, price_name)
        amount_widget = getattr(window, amount_name)

        def recalc(_checked=False, qty_widget=qty_widget, price_widget=price_widget, amount_widget=amount_widget):
            qty = _as_number(qty_widget.toPlainText().strip())
            price = _as_number(price_widget.toPlainText().strip())
            if isinstance(qty, (int, float)) and isinstance(price, (int, float)):
                amount_widget.setPlainText(f"{qty * price:,.0f}")
            else:
                amount_widget.clear()

        qty_widget.textChanged.connect(recalc)
        price_widget.textChanged.connect(recalc)


def build_entry_chain(window, item_rows, leading_fields, submit_button):
    entry_fields = list(leading_fields)
    for item_name, desc_name, qty_name, price_name, _amount_name in item_rows:
        entry_fields.extend(getattr(window, n) for n in (item_name, desc_name, qty_name, price_name))

    for field in entry_fields:
        field.setTabChangesFocus(True)
    for current, following in zip(entry_fields, entry_fields[1:]):
        window.setTabOrder(current, following)
        _advance_on_enter(current, following.setFocus)
    window.setTabOrder(entry_fields[-1], submit_button)
    _advance_on_enter(entry_fields[-1], submit_button.click)
    return entry_fields


def save_filled_document(
    window,
    *,
    item_rows,
    header_cell_map,
    remarks_cell,
    remarks_widget_name,
    xlsx_template,
    save_dir,
    list_xlsx_path,
    name_widget_name,
    date_widget_name,
    number_widget_name,
):
    """Fill `xlsx_template` from `window`'s fields, export it + a screen-capture PDF into
    `save_dir`, and log each non-blank item row into `list_xlsx_path`. Returns the
    document number used."""
    rows_data = [
        (
            getattr(window, item_name).toPlainText().strip(),
            getattr(window, desc_name).toPlainText().strip(),
            getattr(window, qty_name).toPlainText().strip(),
            getattr(window, price_name).toPlainText().strip(),
        )
        for item_name, desc_name, qty_name, price_name, _amount_name in item_rows
    ]

    workbook = openpyxl.load_workbook(xlsx_template)
    sheet = workbook.active

    for cell, widget_name in header_cell_map:
        sheet[cell] = getattr(window, widget_name).toPlainText().strip()
    sheet[remarks_cell] = getattr(window, remarks_widget_name).toPlainText().strip()

    for offset, (item, desc, qty, price) in enumerate(rows_data):
        row = 16 + offset
        sheet[f"A{row}"] = item
        sheet[f"B{row}"] = desc
        sheet[f"G{row}"] = _as_number(qty) if qty else None
        sheet[f"H{row}"] = _as_number(price) if price else None

    doc_no = getattr(window, number_widget_name).toPlainText().strip()
    export_document_files(workbook, doc_no, save_dir, window.centralwidget)

    list_date = QDate.fromString(getattr(window, date_widget_name).toPlainText().strip(), "yyyy-MM-dd").toString("yyMMdd")
    name = getattr(window, name_widget_name).toPlainText().strip()
    for item, desc, qty, price in rows_data:
        if not item:
            continue
        append_to_list_xlsx(
            list_xlsx_path,
            list_date,
            name,
            item,
            desc,
            _as_number(qty) if qty else None,
            _as_number(price) if price else None,
        )

    return doc_no


class CompactItemFormWindow(QMainWindow):
    """Base for the compact entry forms (Quotation.ui / DeliveryNote.ui): a few header
    fields plus up to `max_item_rows` dynamically add/remove-able item rows. Go hands
    everything off to the paired full-page window (set via `full_page_window_class`)
    for review/save."""

    ui_path = None
    window_title = ""
    number_prefix = None
    seq_file = None
    bottom_cluster_names = ()
    full_page_window_class = None
    max_item_rows = 8

    def __init__(self):
        super().__init__()
        load_ui(self.ui_path, self)
        self.setFixedSize(self.size())
        self.setWindowTitle(self.window_title)

        # Date/number are auto-filled below and left read-only; customer name/address are
        # the only header fields the user actually types into.
        auto_filled = ("Date_textBrowser", "Number_textBrowse")
        for name in ("Date_textBrowser", "Number_textBrowse", "Name_textBrowse", "Address_textBrowse"):
            widget = getattr(self, name)
            widget.setReadOnly(name in auto_filled)
            widget.setLineWrapMode(QTextEdit.NoWrap)
            widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            widget.setStyleSheet(COMPACT_INPUT_STYLESHEET)
            font = widget.font()
            font.setPointSize(12)
            widget.setFont(font)
            center_text_vertically(widget)

        self.Date_textBrowser.setText(QDate.currentDate().toString("yyyy-MM-dd"))
        self.Number_textBrowse.setText(next_document_number(self.number_prefix, self.seq_file))

        self._base_window_size = self.size()
        self._bottom_cluster_origin = {
            getattr(self, name): getattr(self, name).pos() for name in self.bottom_cluster_names
        }
        self._bottom_cluster_top = min(pos.y() for pos in self._bottom_cluster_origin.values())

        self.item_rows = [[
            self.Item_textBrowse,
            self.Description_textBrowse,
            self.Quantity_textBrowse,
            self.UnitPrice_textBrowse,
        ]]
        for widget, (key, _x, _width) in zip(self.item_rows[0], COMPACT_ITEM_ROW_COLUMNS):
            self._style_row_widget(widget, right_align=key in COMPACT_ITEM_ROW_RIGHT_ALIGNED_COLUMNS)

        self._rebuild_entry_chain()
        self._update_bottom_layout()
        self.Name_textBrowse.setFocus()

        self._next_window = None
        self.Go_pushButton.clicked.connect(self.open_full_page)
        self.Back_pushButton.clicked.connect(self.close)
        self.Plus_pushButton.clicked.connect(self.add_item_row)
        self.Minus_pushButton.clicked.connect(self.remove_item_row)

    @staticmethod
    def _style_row_widget(widget, right_align=False):
        widget.setReadOnly(False)
        font = widget.font()
        font.setPointSize(12)
        widget.setFont(font)
        widget.setLineWrapMode(QTextEdit.NoWrap)
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setStyleSheet(COMPACT_INPUT_STYLESHEET)
        if right_align:
            # setAlignment() only affects the current paragraph, and a fresh
            # setPlainText() (prefill, live recalculation, etc.) resets it - reapply on
            # every text change so the field stays right-aligned no matter how its text
            # is set. setAlignment() itself re-emits textChanged, so guard against
            # re-entering when it's already right-aligned or this recurses forever.
            def apply_right_alignment():
                if widget.alignment() != Qt.AlignRight:
                    widget.setAlignment(Qt.AlignRight)

            apply_right_alignment()
            widget.textChanged.connect(apply_right_alignment)
        center_text_vertically(widget)

    def _make_item_row(self):
        row_index = len(self.item_rows)
        y = COMPACT_ITEM_ROW_TOP + row_index * COMPACT_ITEM_ROW_SPACING
        widgets = []
        for key, x, width in COMPACT_ITEM_ROW_COLUMNS:
            widget = QTextBrowser(self.centralwidget)
            widget.setGeometry(x, y, width, COMPACT_ITEM_ROW_HEIGHT)
            self._style_row_widget(widget, right_align=key in COMPACT_ITEM_ROW_RIGHT_ALIGNED_COLUMNS)
            widget.show()
            widgets.append(widget)
        return widgets

    def _rebuild_entry_chain(self):
        fields = [self.Name_textBrowse, self.Address_textBrowse]
        for row in self.item_rows:
            fields.extend(row)

        for field in fields:
            field.setTabChangesFocus(True)
        for current, following in zip(fields, fields[1:]):
            self.setTabOrder(current, following)
            _advance_on_enter(current, following.setFocus)
        self.setTabOrder(fields[-1], self.Go_pushButton)
        _advance_on_enter(fields[-1], self.Go_pushButton.click)

    def _update_bottom_layout(self):
        row_count = len(self.item_rows)
        last_row_bottom = COMPACT_ITEM_ROW_TOP + (row_count - 1) * COMPACT_ITEM_ROW_SPACING + COMPACT_ITEM_ROW_HEIGHT
        needed_top = last_row_bottom + COMPACT_ITEM_ROW_GAP
        delta = max(0, needed_top - self._bottom_cluster_top)

        for widget, origin in self._bottom_cluster_origin.items():
            widget.move(origin.x(), origin.y() + delta)
        self.setFixedSize(self._base_window_size.width(), self._base_window_size.height() + delta)

        self.Plus_pushButton.setEnabled(row_count < self.max_item_rows)
        self.Minus_pushButton.setEnabled(row_count > 1)

    def add_item_row(self):
        if len(self.item_rows) >= self.max_item_rows:
            return
        widgets = self._make_item_row()
        self.item_rows.append(widgets)
        self._rebuild_entry_chain()
        self._update_bottom_layout()
        widgets[0].setFocus()

    def remove_item_row(self):
        if len(self.item_rows) <= 1:
            return
        widgets = self.item_rows.pop()
        for widget in widgets:
            widget.deleteLater()
        self._rebuild_entry_chain()
        self._update_bottom_layout()

    def open_full_page(self):
        rows = [
            (
                item_w.toPlainText().strip(),
                desc_w.toPlainText().strip(),
                qty_w.toPlainText().strip(),
                price_w.toPlainText().strip(),
            )
            for item_w, desc_w, qty_w, price_w in self.item_rows
        ]
        prefill = {
            "date": self.Date_textBrowser.toPlainText().strip(),
            "number": self.Number_textBrowse.toPlainText().strip(),
            "name": self.Name_textBrowse.toPlainText().strip(),
            "address": self.Address_textBrowse.toPlainText().strip(),
            "rows": rows,
        }
        self._next_window = self.full_page_window_class(prefill=prefill)
        move_centered_on(self._next_window, self)
        self._next_window.show()
        self.close()
