"""견적서 (Quotation) windows: the compact entry form (Quotation.ui) and the full-page
review/save form (Quotation2.ui)."""

import os

import winsound
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QMainWindow, QMessageBox

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

QUOTATION_UI = os.path.join(UI_DIR, "Quotation.ui")
QUOTATION2_UI = os.path.join(UI_DIR, "Quotation2.ui")
QUOTATION_XLSX = os.path.join(BASE_DIR, "BN_견적서_양식.xlsx")
QUOTATION_SAVE_DIR = os.path.join(BASE_DIR, "List", "Quotation_List")
QUOTATION_LIST_XLSX = os.path.join(QUOTATION_SAVE_DIR, "List", "견적서List.xlsx")
QUOTATION_SEQ_FILE = os.path.join(BASE_DIR, "quotation_seq.json")

# Each item row on Quotation2.ui as (item, description, quantity, unit_price, amount_display)
# widget names. Rows correspond, in order, to the quotation template's item rows 16-21.
ITEM_ROWS = (
    ("Item1_textEdit", "Description1_textEdit", "Qty1_textEdit", "UnitPrice1_textEdit", "Amount1_textEdit"),
    ("Item2_textEdit", "Description2_textEdit", "Qty2_textEdit", "UnitPrice2_textEdit", "Amount2_textEdit"),
    ("Item3_textEdit", "Description3_textEdit", "Qty3_textEdit", "UnitPrice3_textEdit", "Amount3_textEdit"),
    ("Item4_textEdit", "Description4_textEdit", "Qty4_textEdit", "UnitPrice4_textEdit", "Amount4_textEdit"),
    ("Item5_textEdit", "Description5_textEdit", "Qty5_textEdit", "UnitPrice5_textEdit", "Amount5_textEdit"),
    ("Item6_textEdit", "Description6_textEdit", "Qty6_textEdit", "UnitPrice6_textEdit", "Amount6_textEdit"),
)

# (excel cell, Quotation2.ui widget name) for the single-value header/customer fields.
HEADER_CELL_MAP = (
    ("G2", "textEdit_6"),  # date
    ("G3", "textEdit_7"),  # quotation number
    ("A11", "ClientName_textEdit"),  # customer name
    ("A12", "ClientAddress_textEdit"),  # customer address
)
REMARKS_CELL = "F10"  # textEdit_2

QUOTATION_BOTTOM_CLUSTER_WIDGET_NAMES = COMPACT_BOTTOM_CLUSTER_WIDGET_NAMES


class Quotation2Window(QMainWindow):
    def __init__(self, prefill=None):
        super().__init__()
        load_ui(QUOTATION2_UI, self)
        self.setWindowTitle("견적서")
        make_window_scrollable(self)

        self.textEdit_6.setText(
            prefill["date"] if prefill and prefill.get("date") else QDate.currentDate().toString("yyyy-MM-dd")
        )
        self.textEdit_7.setText(
            prefill["number"] if prefill and prefill.get("number") else next_document_number("BN-QUT", QUOTATION_SEQ_FILE)
        )

        style_item_form(
            self,
            ITEM_ROWS,
            [
                self.textEdit_6,
                self.textEdit_7,
                self.ClientName_textEdit,
                self.ClientAddress_textEdit,
                # Static letterhead fields (company name/reg. no./line of business/phone/
                # email/rep. name) - single-line but never listed here before, so they
                # kept the default word-wrap mode. center_text_vertically() then measured
                # their height from the wrapped document layout instead of font metrics,
                # which for CJK text overstates the line height and leaves them looking
                # off-center (see center_text_vertically()'s no_wrap branch comment).
                self.textEdit,
                self.textEdit_3,
                self.textEdit_5,
                self.Tel_textEdit,
                self.mail_textEdit,
                self.textEdit_35,
            ],
        )
        connect_amount_calculations(self, ITEM_ROWS)
        entry_fields = build_entry_chain(
            self, ITEM_ROWS, [self.ClientName_textEdit, self.ClientAddress_textEdit], self.Save_pushButton
        )
        self.textEdit_2.setTabChangesFocus(True)

        if prefill:
            self._apply_prefill(prefill)

        entry_fields[0].setFocus()

        self._saving = False
        self.Save_pushButton.clicked.connect(self.save_quotation)
        self.Print_pushButton.clicked.connect(lambda: print_widget(self.centralwidget, self))

    def _apply_prefill(self, data):
        self.ClientName_textEdit.setPlainText(data.get("name", ""))
        self.ClientAddress_textEdit.setPlainText(data.get("address", ""))
        for row_names, (item, desc, qty, price) in zip(ITEM_ROWS, data.get("rows", [])):
            item_name, desc_name, qty_name, price_name, _amount_name = row_names
            getattr(self, item_name).setPlainText(item)
            getattr(self, desc_name).setPlainText(desc)
            getattr(self, qty_name).setPlainText(qty)
            getattr(self, price_name).setPlainText(price)

    def save_quotation(self):
        if self._saving:
            return
        self._saving = True
        self.Save_pushButton.setEnabled(False)
        winsound.PlaySound(CLICK_SOUND, winsound.SND_FILENAME | winsound.SND_ASYNC)

        try:
            doc_no = save_filled_document(
                self,
                item_rows=ITEM_ROWS,
                header_cell_map=HEADER_CELL_MAP,
                remarks_cell=REMARKS_CELL,
                remarks_widget_name="textEdit_2",
                xlsx_template=QUOTATION_XLSX,
                save_dir=QUOTATION_SAVE_DIR,
                list_xlsx_path=QUOTATION_LIST_XLSX,
                name_widget_name="ClientName_textEdit",
                date_widget_name="textEdit_6",
                number_widget_name="textEdit_7",
            )
            QMessageBox.information(self, "저장 완료", f"{doc_no} 견적서가 저장되었습니다.")
        except Exception as exc:
            QMessageBox.critical(self, "견적서 저장 실패", f"견적서를 저장하는 중 오류가 발생했습니다.\n\n{exc}")
        finally:
            self._saving = False
            self.Save_pushButton.setEnabled(True)


class QuotationWindow(CompactItemFormWindow):
    ui_path = QUOTATION_UI
    window_title = "견적서"
    number_prefix = "BN-QUT"
    seq_file = QUOTATION_SEQ_FILE
    bottom_cluster_names = QUOTATION_BOTTOM_CLUSTER_WIDGET_NAMES
    full_page_window_class = Quotation2Window
    max_item_rows = len(ITEM_ROWS)
