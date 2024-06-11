import os
import sys

import xlwings as xw
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon, QKeyEvent
from PySide6.QtWidgets import (QApplication, QComboBox, QFileDialog,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMainWindow, QPushButton,
                               QVBoxLayout, QWidget)

from excel.lookup_excel import DirectoryExcel, SearchResult


def open_file(item: SearchResult):
    wb = xw.Book(os.path.abspath(item.file))

    ws = wb.sheets[item.sheet]

    ws.activate()
    ws.range(item.cell).select()
    wb.app.activate(steal_focus=True)


class QDirComboBox(QComboBox):
    editingFinished = Signal(str)

    def __init__(self):
        self.folder_list = set()
        super().__init__()
        self.setInsertPolicy(QComboBox.InsertPolicy.InsertAtBottom)

    def add_dir(self, dir_path):
        if dir_path in self.folder_list:
            return
        else:
            self.folder_list.add(dir_path)
            self.addItem(dir_path)

    def keyPressEvent(self, e: QKeyEvent) -> None:
        if e.key() in [Qt.Key_Return, Qt.Key_Enter]:
            self.editingFinished.emit(self.currentText())
        super().keyPressEvent(e)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Lookup App")
        self.key_to_search: str = ""
        self.dir_context = DirectoryExcel()

        layout = QVBoxLayout()
        dir_selector_layout = QHBoxLayout()

        label_select_dir = QLabel("Select or type directory here:")
        layout.addWidget(label_select_dir)

        self.combobox_select_dir = QDirComboBox()
        self.combobox_select_dir.setEditable(True)
        self.combobox_select_dir.editingFinished.connect(self.enter_dir)
        dir_selector_layout.addWidget(self.combobox_select_dir)

        button = QPushButton()
        button.setIcon(QIcon("./public/dir_icon.png"))
        button.setMaximumSize(32, 32)
        button.clicked.connect(self.button_click_handler)
        dir_selector_layout.addWidget(button)

        layout.addLayout(dir_selector_layout)

        label_input_key = QLabel("Input search here:")
        layout.addWidget(label_input_key)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Enter search keys")
        self.search_box.editingFinished.connect(self.text_edited)
        layout.addWidget(self.search_box)

        # Add a QListWidget to display search results
        self.result_list = QListWidget()
        self.result_list.itemDoubleClicked.connect(self.item_clicked_handler)
        layout.addWidget(self.result_list)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def enter_dir(self, dir):
        self.dir_context.root_dir = dir
        self.combobox_select_dir.add_dir(dir)
        self.search_and_update()

    def search_and_update(self):
        self.result_list.clear()
        for result in self.dir_context.search_keyword(self.key_to_search):
            print(result)
            item = QListWidgetItem(str(result))
            item.setData(Qt.UserRole, result)
            self.result_list.addItem(item)

    def text_edited(self):
        self.key_to_search = self.search_box.text()
        print(self.key_to_search)
        self.search_and_update()

    def dir_handler(self, dir):
        self.selected_dir = dir
        self.combobox_select_dir.add_dir(dir)
        self.dir_context.root_dir = dir
        self.combobox_select_dir.setCurrentText(dir)
        self.search_and_update()

    def button_click_handler(self, checked):
        self.dir_selector = QFileDialog(self)
        self.dir_selector.setFileMode(QFileDialog.FileMode.Directory)
        self.dir_selector.fileSelected.connect(self.dir_handler)
        self.dir_selector.exec()

    def item_clicked_handler(self, item):
        open_file(item.data(Qt.UserRole))


app = QApplication(sys.argv)

window = MainWindow()
window.show()
if __name__ == "__main__":
    app.exec()
