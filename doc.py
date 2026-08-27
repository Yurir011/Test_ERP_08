"""공문서 (Official Document) windows: the compact entry form (Doc.ui) and the
full-page body/approval/save form (Doc2.ui)."""

import os

import openpyxl
import winsound
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QMessageBox, QTextEdit

from common import (
    BASE_DIR,
    CLICK_SOUND,
    COMPACT_INPUT_STYLESHEET,
    UI_DIR,
    _advance_on_enter,
    center_text_vertically,
    export_document_files,
    load_ui,
    make_window_scrollable,
    move_centered_on,
    next_document_number,
    print_widget,
)

DOC_UI = os.path.join(UI_DIR, "Doc.ui")
DOC2_UI = os.path.join(UI_DIR, "Doc2.ui")
DOC_XLSX = os.path.join(BASE_DIR, "BN_공문서_양식.xlsx")
DOC_SAVE_DIR = os.path.join(BASE_DIR, "List", "OfficialDoc_List")
DOC_LIST_XLSX = os.path.join(DOC_SAVE_DIR, "List", "공문서List.xlsx")
DOC_SEQ_FILE = os.path.join(BASE_DIR, "doc_seq.json")
DOC_APPROVE_ICON = os.path.join(BASE_DIR, "ICON", "icon-approve.png")
DOC_REJECT_ICON = os.path.join(BASE_DIR, "ICON", "icon-reject.png")

# --- Doc2.ui (공문서 full-page document): single-value header fields, a free-form
# body split one line per template row, and a 4-stage approval strip.
DOC_HEADER_CELL_MAP = (
    ("C6", "Number_textEdit"),  # 문서번호
    ("C7", "Date_textEdit"),  # 시행일자
    ("C8", "textEdit_3"),  # 경유
    ("C9", "textEdit_4"),  # 수신
    ("C10", "Title_textEdit"),  # 제목
)
DOC_PERSON_CELL = "A3"  # letterhead contact person - cell already reads "담당자     <name>"
DOC_PERSON_LABEL = "담당자"
DOC_FOOTER_PERSON_CELL = "A40"
DOC_FOOTER_COOP_CELL = "A41"
DOC_BODY_ROWS = tuple(r for r in range(15, 39) if r != 36)  # one MainBody_textEdit line
# per row, A15..A38 - skipping row 36, which holds the fixed company signature text.

# (approval button widget name, excel cell for that stage's stamp/status, role label)
DOC_APPROVAL_STAGES = (
    ("pushButton_4", "H7", "담당자"),
    ("pushButton", "I7", "관리자"),
    ("pushButton_2", "J7", "부서장"),
    ("pushButton_3", "K7", "대표이사"),
)


def append_to_doc_list(list_xlsx_path, date_text, manager, number, routing, recipient, title):
    workbook = openpyxl.load_workbook(list_xlsx_path)
    sheet = workbook.active

    row = 3
    while sheet[f"B{row}"].value is not None:
        row += 1

    sheet[f"B{row}"] = date_text
    sheet[f"C{row}"] = manager
    sheet[f"D{row}"] = number
    sheet[f"E{row}"] = routing
    sheet[f"F{row}"] = recipient
    sheet[f"G{row}"] = title

    workbook.save(list_xlsx_path)


class DocWindow(QMainWindow):
    """Compact entry form (Doc.ui): fill the header fields here, then Go hands
    everything off to Doc2Window for the full-page body/approval/save."""

    FIELD_NAMES = (
        "Person_textBrowse",  # 담당자
        "_textBrowser",  # 문서번호
        "Number_textBrowse",  # 시행일자
        "Name_textBrowse",  # 경유
        "Address_textBrowse",  # 수신
        "Item_textBrowse",  # 제목
    )

    def __init__(self):
        super().__init__()
        load_ui(DOC_UI, self)
        self.setFixedSize(self.size())
        self.setWindowTitle("공문서")

        # 문서번호/시행일자 are auto-filled below and left read-only; the rest are typed in.
        auto_filled = ("_textBrowser", "Number_textBrowse")
        for name in self.FIELD_NAMES:
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

        self._textBrowser.setText(next_document_number("BN-DLM", DOC_SEQ_FILE))
        self.Number_textBrowse.setText(QDate.currentDate().toString("yyyy-MM-dd"))

        entry_fields = [getattr(self, name) for name in self.FIELD_NAMES]
        for field in entry_fields:
            field.setTabChangesFocus(True)
        for current, following in zip(entry_fields, entry_fields[1:]):
            self.setTabOrder(current, following)
            _advance_on_enter(current, following.setFocus)
        self.setTabOrder(entry_fields[-1], self.Go_pushButton)
        _advance_on_enter(entry_fields[-1], self.Go_pushButton.click)
        entry_fields[0].setFocus()

        self._next_window = None
        self.Go_pushButton.clicked.connect(self.open_doc2)
        self.Back_pushButton.clicked.connect(self.close)

    def open_doc2(self):
        prefill = {
            "number": self._textBrowser.toPlainText().strip(),
            "date": self.Number_textBrowse.toPlainText().strip(),
            "person": self.Person_textBrowse.toPlainText().strip(),
            "routing": self.Name_textBrowse.toPlainText().strip(),
            "recipient": self.Address_textBrowse.toPlainText().strip(),
            "title": self.Item_textBrowse.toPlainText().strip(),
        }
        self._next_window = Doc2Window(prefill=prefill)
        move_centered_on(self._next_window, self)
        self._next_window.show()
        self.close()


class Doc2Window(QMainWindow):
    def __init__(self, prefill=None):
        super().__init__()
        load_ui(DOC2_UI, self)
        self.setWindowTitle("공문서")
        make_window_scrollable(self)

        for text_edit in self.centralwidget.findChildren(QTextEdit):
            center_text_vertically(text_edit)

        single_line_widgets = [
            self.Person_textEdit,
            self.Number_textEdit,
            self.Date_textEdit,
            self.textEdit_3,
            self.textEdit_4,
            self.Title_textEdit,
            self.textEdit_6,
            self.textEdit_7,
            # Static letterhead fields - see the matching comment in Quotation2Window.
            self.Tel_textEdit,
            self.mail_textEdit,
        ]
        for widget in single_line_widgets:
            widget.setLineWrapMode(QTextEdit.NoWrap)
            widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.Number_textEdit.setText(
            prefill["number"] if prefill and prefill.get("number") else next_document_number("BN-DLM", DOC_SEQ_FILE)
        )
        self.Date_textEdit.setText(
            prefill["date"] if prefill and prefill.get("date") else QDate.currentDate().toString("yyyy-MM-dd")
        )

        # The footer "담당자" line is the same person as the letterhead one - shown
        # read-only here and kept in sync, so it's only ever typed once.
        self.textEdit_6.setReadOnly(True)
        self.Person_textEdit.textChanged.connect(
            lambda: self.textEdit_6.setPlainText(self.Person_textEdit.toPlainText())
        )

        entry_fields = [
            self.Person_textEdit,
            self.textEdit_3,
            self.textEdit_4,
            self.Title_textEdit,
            self.MainBody_textEdit,
            self.textEdit_7,
        ]
        for field in entry_fields:
            field.setTabChangesFocus(True)
        for current, following in zip(entry_fields, entry_fields[1:]):
            self.setTabOrder(current, following)
            if current is not self.MainBody_textEdit:
                _advance_on_enter(current, following.setFocus)
        self.setTabOrder(entry_fields[-1], self.Save_pushButton)
        _advance_on_enter(entry_fields[-1], self.Save_pushButton.click)
        # MainBody_textEdit is free-form multi-line text, so Enter there is left alone
        # to insert a newline instead of jumping focus.

        if prefill:
            self.Person_textEdit.setPlainText(prefill.get("person", ""))
            self.textEdit_3.setPlainText(prefill.get("routing", ""))
            self.textEdit_4.setPlainText(prefill.get("recipient", ""))
            self.Title_textEdit.setPlainText(prefill.get("title", ""))
            if prefill.get("body"):
                self.MainBody_textEdit.setPlainText(prefill["body"])

        for button_name, _cell, _role in DOC_APPROVAL_STAGES:
            button = getattr(self, button_name)
            button.clicked.connect(lambda _checked=False, b=button: self._cycle_approval(b))

        entry_fields[0].setFocus()

        self._saving = False
        self.Save_pushButton.clicked.connect(self.save_doc)
        self.Print_pushButton.clicked.connect(lambda: print_widget(self.centralwidget, self))

    def _cycle_approval(self, button):
        """결재대기 -> 승인 -> 반려 -> 결재대기 ... one click per state change."""
        state = button.property("approval_state") or "pending"
        next_state = {"pending": "approved", "approved": "rejected", "rejected": "pending"}[state]
        button.setProperty("approval_state", next_state)

        today = QDate.currentDate().toString("yyyy-MM-dd")
        if next_state == "approved":
            button.setIcon(QIcon(DOC_APPROVE_ICON))
            button.setText(f"승인\n{today}")
            button.setStyleSheet("background-color: rgb(198, 239, 206); color: rgb(30, 90, 30);")
        elif next_state == "rejected":
            button.setIcon(QIcon(DOC_REJECT_ICON))
            button.setText(f"반려\n{today}")
            button.setStyleSheet("background-color: rgb(245, 205, 203); color: rgb(120, 30, 30);")
        else:
            button.setIcon(QIcon())
            button.setText("결재대기")
            button.setStyleSheet("color: rgb(209, 209, 209);")

    def save_doc(self):
        if self._saving:
            return
        self._saving = True
        self.Save_pushButton.setEnabled(False)
        winsound.PlaySound(CLICK_SOUND, winsound.SND_FILENAME | winsound.SND_ASYNC)

        try:
            workbook = openpyxl.load_workbook(DOC_XLSX)
            sheet = workbook.active

            for cell, widget_name in DOC_HEADER_CELL_MAP:
                sheet[cell] = getattr(self, widget_name).toPlainText().strip()

            person = self.Person_textEdit.toPlainText().strip()
            sheet[DOC_PERSON_CELL] = f"{DOC_PERSON_LABEL}     {person}" if person else DOC_PERSON_LABEL
            sheet[DOC_FOOTER_PERSON_CELL] = f"담당자   :   {person}"
            coop = self.textEdit_7.toPlainText().strip()
            sheet[DOC_FOOTER_COOP_CELL] = f"협조자   :   {coop}"

            body_lines = self.MainBody_textEdit.toPlainText().split("\n")
            for row, line in zip(DOC_BODY_ROWS, body_lines):
                sheet[f"A{row}"] = line

            for button_name, cell, _role in DOC_APPROVAL_STAGES:
                button = getattr(self, button_name)
                if (button.property("approval_state") or "pending") != "pending":
                    sheet[cell] = button.text()

            doc_no = self.Number_textEdit.toPlainText().strip()
            export_document_files(workbook, doc_no, DOC_SAVE_DIR, self.centralwidget)

            list_date = QDate.fromString(self.Date_textEdit.toPlainText().strip(), "yyyy-MM-dd").toString("yyMMdd")
            append_to_doc_list(
                DOC_LIST_XLSX,
                list_date,
                person,
                doc_no,
                self.textEdit_3.toPlainText().strip(),
                self.textEdit_4.toPlainText().strip(),
                self.Title_textEdit.toPlainText().strip(),
            )

            QMessageBox.information(self, "저장 완료", f"{doc_no} 공문서가 저장되었습니다.")
        except Exception as exc:
            QMessageBox.critical(self, "공문서 저장 실패", f"공문서를 저장하는 중 오류가 발생했습니다.\n\n{exc}")
        finally:
            self._saving = False
            self.Save_pushButton.setEnabled(True)
