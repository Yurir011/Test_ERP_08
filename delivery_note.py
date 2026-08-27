"""거래명세서 (Delivery Note) windows: the compact entry form (DeliveryNote.ui), the
full-page review/save form (DeliveryNote2.ui), and the "문서 참조" (load a previously
saved xlsx/pdf back into the compact form) parsers."""

import datetime
import os
import re

import openpyxl
import winsound
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QMessageBox

from common import (
    BASE_DIR,
    CLICK_SOUND,
    COMPACT_BOTTOM_CLUSTER_WIDGET_NAMES,
    UI_DIR,
    CompactItemFormWindow,
    build_entry_chain,
    connect_amount_calculations,
    load_ui,
    make_window_scrollable,
    next_document_number,
    print_widget,
    save_filled_document,
    style_item_form,
)

DELIVERY_NOTE_UI = os.path.join(UI_DIR, "DeliveryNote.ui")
DELIVERY_NOTE2_UI = os.path.join(UI_DIR, "DeliveryNote2.ui")
DELIVERY_NOTE_XLSX = os.path.join(BASE_DIR, "BN_거래명세서_양식.xlsx")
DELIVERY_NOTE_SAVE_DIR = os.path.join(BASE_DIR, "List", "DeliveryNott_List")
DELIVERY_NOTE_LIST_XLSX = os.path.join(DELIVERY_NOTE_SAVE_DIR, "List", "거래명세서List.xlsx")
DELIVERY_NOTE_SEQ_FILE = os.path.join(BASE_DIR, "delivery_note_seq.json")

# Same item-row layout as Quotation2.ui, just under its own (slightly differently
# named) widgets on DeliveryNote2.ui.
DELIVERY_ITEM_ROWS = (
    ("Item1_textEdit", "Description1_textEdit", "Qty1_textEdit", "UnitPrice1_textEdit", "Amount1_textEdit"),
    ("Item2_textEdit", "Description2_textEdit", "Qty2_textEdit", "UnitPrice2_textEdit", "Amount2_textEdit"),
    ("Item3_textEdit", "Description3_textEdit", "Qty3_textEdit", "UnitPrice3_textEdit", "Amount3_textEdit"),
    ("Item4_textEdit", "Description4_textEdit", "Qty4_textEdit", "UnitPrice4_textEdit", "Amount4_textEdit"),
    ("Item5_textEdit", "Description5_textEdit", "Qty5_textEdit", "UnitPrice5_textEdit", "Amount5_textEdit"),
    ("Item6_textEdit", "Description6_textEdit", "Qty6_textEdit", "UnitPrice6_textEdit", "Amount6_textEdit"),
)
DELIVERY_HEADER_CELL_MAP = (
    ("G2", "Date_textEdit"),  # date
    ("G3", "textEdit_7"),  # document number
    ("A11", "ClientNametextEdit"),  # customer name
    ("A12", "ClientAdd_textEdit"),  # customer address
)
DELIVERY_REMARKS_CELL = "F10"  # textEdit_2

DELIVERY_NOTE_BOTTOM_CLUSTER_WIDGET_NAMES = COMPACT_BOTTOM_CLUSTER_WIDGET_NAMES + ("Go_pushButton_2", "label_13")


# --- "문서 참조" (load a previously saved xlsx/pdf back into the compact entry form).
# Both parsers read the same cell/text layout save_filled_document() writes, so a
# hand-edited or differently formatted file can throw them off - this is a best-effort
# convenience, not a guaranteed-correct import.

def parse_saved_xlsx(xlsx_path):
    workbook = openpyxl.load_workbook(xlsx_path)
    sheet = workbook.active

    def cell_text(coord):
        value = sheet[coord].value
        if value is None:
            return ""
        if isinstance(value, (datetime.date, datetime.datetime)):
            return value.strftime("%Y-%m-%d")
        return str(value).strip()

    rows = []
    for row_num in range(16, 24):
        item = cell_text(f"A{row_num}")
        desc = cell_text(f"B{row_num}")
        qty = cell_text(f"G{row_num}")
        price = cell_text(f"H{row_num}")
        if item or desc or qty or price:
            rows.append((item, desc, qty, price))

    return {
        "date": cell_text("G2"),
        "number": cell_text("G3"),
        "name": cell_text("A11"),
        "address": cell_text("A12"),
        "rows": rows,
    }


def _cluster_pdf_words(words, tolerance=2.0):
    """Group words (from pdfplumber's extract_words()) into lines by their 'top'
    coordinate: top-to-bottom, each line sorted left-to-right by x0."""
    lines = []
    for word in sorted(words, key=lambda w: w["top"]):
        for line in lines:
            if abs(line[0]["top"] - word["top"]) <= tolerance:
                line.append(word)
                break
        else:
            lines.append([word])
    for line in lines:
        line.sort(key=lambda w: w["x0"])
    lines.sort(key=lambda line: line[0]["top"])
    return lines


def _cluster_pdf_words_into_lines(words, tolerance=2.0):
    return [" ".join(w["text"] for w in line) for line in _cluster_pdf_words(words, tolerance)]


def parse_saved_pdf(pdf_path):
    import pdfplumber  # imported lazily so a missing dependency only breaks this feature

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        table = page.extract_table() or []

    full_text = " ".join(w["text"] for w in words)
    date_match = re.search(r"\d{4}-\d{2}-\d{2}", full_text)
    number_match = re.search(r"BN-[A-Z]+-\d{8}-\d{3}", full_text)

    # Customer name/address sit between the "Email : ..." line (end of our fixed
    # company info block) and the "아래와 같이 ~합니다." line (start of the item
    # table section) on both templates - the first line in between is a one-word
    # section label ("고객"/"납품") to skip, then name, then address.
    email_words = [w for w in words if w["text"] == "Email"]
    start_words = [w for w in words if w["text"].startswith("아래와")]
    name, address = "", ""
    if email_words and start_words:
        top_bound = email_words[0]["bottom"]
        bottom_bound = start_words[0]["top"]
        between = [w for w in words if top_bound < w["top"] < bottom_bound]
        content_lines = _cluster_pdf_words_into_lines(between)[1:]
        if content_lines:
            name = content_lines[0]
        if len(content_lines) > 1:
            address = content_lines[1]

    rows = []
    header_idx = next((i for i, row in enumerate(table) if row and row[0] == "ITEM #"), None)
    if header_idx is not None:
        for row in table[header_idx + 1:]:
            if not row or row[0] == "TOTAL":
                break
            item = (row[0] or "").strip()
            desc_candidates = [c for c in row[1:6] if c]
            desc = desc_candidates[0].strip() if desc_candidates else ""
            qty = (row[6] or "").strip()
            price = (row[7] or "").strip()
            if item or desc or qty or price:
                rows.append((item, desc, qty, price))
            if len(rows) >= len(DELIVERY_ITEM_ROWS):
                break

    return {
        "date": date_match.group(0) if date_match else "",
        "number": number_match.group(0) if number_match else "",
        "name": name,
        "address": address,
        "rows": rows,
    }


class DeliveryNote2Window(QMainWindow):
    def __init__(self, prefill=None):
        super().__init__()
        load_ui(DELIVERY_NOTE2_UI, self)
        self.setWindowTitle("거래명세서")
        make_window_scrollable(self)

        self.Date_textEdit.setText(
            prefill["date"] if prefill and prefill.get("date") else QDate.currentDate().toString("yyyy-MM-dd")
        )
        self.textEdit_7.setText(
            prefill["number"] if prefill and prefill.get("number") else next_document_number("BN-INV", DELIVERY_NOTE_SEQ_FILE)
        )

        style_item_form(
            self,
            DELIVERY_ITEM_ROWS,
            [
                self.Date_textEdit,
                self.textEdit_7,
                self.ClientNametextEdit,
                self.ClientAdd_textEdit,
                # Static letterhead fields - see the matching comment in Quotation2Window.
                self.textEdit,
                self.textEdit_3,
                self.textEdit_5,
                self.Tel_textEdit,
                self.mail_textEdit,
                self.textEdit_35,
            ],
        )
        connect_amount_calculations(self, DELIVERY_ITEM_ROWS)
        entry_fields = build_entry_chain(
            self, DELIVERY_ITEM_ROWS, [self.ClientNametextEdit, self.ClientAdd_textEdit], self.Save_pushButton
        )
        self.textEdit_2.setTabChangesFocus(True)

        if prefill:
            self._apply_prefill(prefill)

        entry_fields[0].setFocus()

        self._saving = False
        self.Save_pushButton.clicked.connect(self.save_delivery_note)
        self.Print_pushButton.clicked.connect(lambda: print_widget(self.centralwidget, self))

    def _apply_prefill(self, data):
        self.ClientNametextEdit.setPlainText(data.get("name", ""))
        self.ClientAdd_textEdit.setPlainText(data.get("address", ""))
        for row_names, (item, desc, qty, price) in zip(DELIVERY_ITEM_ROWS, data.get("rows", [])):
            item_name, desc_name, qty_name, price_name, _amount_name = row_names
            getattr(self, item_name).setPlainText(item)
            getattr(self, desc_name).setPlainText(desc)
            getattr(self, qty_name).setPlainText(qty)
            getattr(self, price_name).setPlainText(price)

    def save_delivery_note(self):
        if self._saving:
            return
        self._saving = True
        self.Save_pushButton.setEnabled(False)
        winsound.PlaySound(CLICK_SOUND, winsound.SND_FILENAME | winsound.SND_ASYNC)

        try:
            doc_no = save_filled_document(
                self,
                item_rows=DELIVERY_ITEM_ROWS,
                header_cell_map=DELIVERY_HEADER_CELL_MAP,
                remarks_cell=DELIVERY_REMARKS_CELL,
                remarks_widget_name="textEdit_2",
                xlsx_template=DELIVERY_NOTE_XLSX,
                save_dir=DELIVERY_NOTE_SAVE_DIR,
                list_xlsx_path=DELIVERY_NOTE_LIST_XLSX,
                name_widget_name="ClientNametextEdit",
                date_widget_name="Date_textEdit",
                number_widget_name="textEdit_7",
            )
            QMessageBox.information(self, "저장 완료", f"{doc_no} 거래명세서가 저장되었습니다.")
        except Exception as exc:
            QMessageBox.critical(self, "거래명세서 저장 실패", f"거래명세서를 저장하는 중 오류가 발생했습니다.\n\n{exc}")
        finally:
            self._saving = False
            self.Save_pushButton.setEnabled(True)


class DeliveryNoteWindow(CompactItemFormWindow):
    ui_path = DELIVERY_NOTE_UI
    window_title = "거래명세서"
    number_prefix = "BN-INV"
    seq_file = DELIVERY_NOTE_SEQ_FILE
    bottom_cluster_names = DELIVERY_NOTE_BOTTOM_CLUSTER_WIDGET_NAMES
    full_page_window_class = DeliveryNote2Window
    max_item_rows = len(DELIVERY_ITEM_ROWS)

    def __init__(self):
        super().__init__()
        self.Go_pushButton_2.clicked.connect(self.load_reference_document)

    def load_reference_document(self):
        path, _filter = QFileDialog.getOpenFileName(
            self,
            "문서 참조 - 불러오기",
            DELIVERY_NOTE_SAVE_DIR,
            "Excel/PDF 파일 (*.xlsx *.pdf)",
        )
        if not path:
            return

        try:
            data = parse_saved_xlsx(path) if path.lower().endswith(".xlsx") else parse_saved_pdf(path)
        except Exception as exc:
            QMessageBox.critical(self, "불러오기 실패", f"문서를 불러오는 중 오류가 발생했습니다.\n\n{exc}")
            return

        self.Name_textBrowse.setPlainText(data["name"])
        self.Address_textBrowse.setPlainText(data["address"])

        rows = data["rows"] or [("", "", "", "")]
        target_count = min(len(rows), self.max_item_rows)
        while len(self.item_rows) < target_count:
            self.add_item_row()
        while len(self.item_rows) > max(target_count, 1):
            self.remove_item_row()

        for row_widgets, (item, desc, qty, price) in zip(self.item_rows, rows):
            item_w, desc_w, qty_w, price_w = row_widgets
            item_w.setPlainText(item)
            desc_w.setPlainText(desc)
            qty_w.setPlainText(qty)
            price_w.setPlainText(price)
