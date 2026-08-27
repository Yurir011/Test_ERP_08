import os
import sys
import traceback

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QLabel, QMainWindow, QMessageBox

from common import UI_DIR, load_ui, move_centered_on
from doc import DOC_LIST_XLSX, DocWindow
from delivery_note import DELIVERY_NOTE_LIST_XLSX, DeliveryNoteWindow
from quotation import QUOTATION_LIST_XLSX, QuotationWindow

HOME_UI = os.path.join(UI_DIR, "Home.ui")


class HomeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        load_ui(HOME_UI, self)
        self.setFixedSize(self.size())
        self.setWindowTitle("Benchsoft")
        self._children = []

        # The caption labels are laid out on top of the icon buttons, so by default they
        # swallow every click that lands on the text and the button never fires.
        for label in self.findChildren(QLabel):
            label.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        self.pushButton.clicked.connect(self.open_doc)
        self.pushButton_4.clicked.connect(self.open_doc_list)
        self.pushButton_2.clicked.connect(self.open_quotation)
        self.pushButton_5.clicked.connect(self.open_quotation_list)
        self.pushButton_3.clicked.connect(self.open_delivery_note)
        self.pushButton_6.clicked.connect(self.open_delivery_note_list)

    def _open_centered(self, child):
        move_centered_on(child, self)
        # Keep a reference so the window is not garbage collected, and drop the ones the
        # user has already closed instead of holding them for the whole session.
        self._children = [c for c in self._children if c.isVisible()]
        self._children.append(child)
        child.show()
        child.raise_()
        child.activateWindow()

    def open_quotation(self):
        self._open_centered(QuotationWindow())

    def open_quotation_list(self):
        os.startfile(QUOTATION_LIST_XLSX)

    def open_delivery_note(self):
        self._open_centered(DeliveryNoteWindow())

    def open_delivery_note_list(self):
        os.startfile(DELIVERY_NOTE_LIST_XLSX)

    def open_doc(self):
        self._open_centered(DocWindow())

    def open_doc_list(self):
        os.startfile(DOC_LIST_XLSX)


def _show_uncaught_exception(exc_type, exc_value, exc_tb):
    """PyQt swallows exceptions raised inside a slot (a button click, etc.): by default
    it only prints them to the console and the UI just sits there, which looks exactly
    like the click did nothing. That's the most common way editing a .ui file in Qt
    Designer breaks this app - a renamed/removed widget the code still refers to by its
    old name - so surface the error in a dialog instead of leaving it silent."""
    traceback.print_exception(exc_type, exc_value, exc_tb)
    QMessageBox.critical(
        None,
        "예상치 못한 오류",
        "작업 중 오류가 발생했습니다:\n\n"
        f"{exc_type.__name__}: {exc_value}\n\n"
        "최근 Qt Designer에서 .ui 파일을 수정했다면, 이름이 바뀌었거나 삭제된 위젯을 "
        "코드가 아직 그 이름으로 찾고 있는 것일 수 있습니다.",
    )


def main():
    sys.excepthook = _show_uncaught_exception
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    window = HomeWindow()
    frame = window.frameGeometry()
    frame.moveCenter(QApplication.primaryScreen().availableGeometry().center())
    window.move(frame.topLeft())
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
