import sys
from random import choice

from PySide6.QtCore import QFileSelector, QSize
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (QApplication, QComboBox, QFileDialog, QLabel,
                               QLineEdit, QMainWindow, QPushButton,
                               QVBoxLayout, QWidget)


class QDirComboBox(QComboBox):
    def __init__(self):
        self.folder_list = set()
        super().__init__()
        self.setInsertPolicy(QComboBox.InsertPolicy.InsertAtTop)

    def add_dir(self, dir_path):
        if dir_path in self.folder_list:
            return
        else:
            self.folder_list.add(dir_path)
            self.addItem(dir_path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(QSize(400, 300))
        self.setWindowTitle("Excel Lookup App")
        self.selected_dir: str = ""
        self.searched_key: str = ""

        layout = QVBoxLayout()

        label_select_dir = QLabel("Select or type directory here:")
        layout.addWidget(label_select_dir)

        button = QPushButton("Click to select folder")
        button.clicked.connect(self.button_click_handler)
        layout.addWidget(button)

        self.combobox_select_dir = QDirComboBox()
        self.combobox_select_dir.setEditable(True)
        layout.addWidget(self.combobox_select_dir)

        label_input_key = QLabel("Input search here:")
        layout.addWidget(label_input_key)

        self.search_box = QLineEdit()
        layout.addWidget(self.search_box)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def dir_handler(self, dir):
        self.selected_dir = dir
        self.combobox_select_dir.add_dir(dir)

        self.combobox_select_dir.setCurrentText(dir)
        print(dir)

    def button_click_handler(self, checked):
        self.dir_selector = QFileDialog(self)
        self.dir_selector.setFileMode(QFileDialog.FileMode.Directory)
        self.dir_selector.fileSelected.connect(self.dir_handler)
        self.dir_selector.exec()


app = QApplication(sys.argv)

window = MainWindow()
window.show()
if __name__ == "__main__":
    app.exec()
