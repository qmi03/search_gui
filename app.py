import sys
from random import choice

from PySide6.QtCore import QFileSelector, QSize
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (QApplication, QComboBox, QFileDialog,
                               QHBoxLayout, QLabel, QLineEdit, QMainWindow,
                               QPushButton, QVBoxLayout, QWidget)


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
        self.setWindowTitle("Excel Lookup App")
        self.searched_key: str = ""

        layout = QVBoxLayout()
        dir_selector_layout = QHBoxLayout()

        label_select_dir = QLabel("Select or type directory here:")
        layout.addWidget(label_select_dir)

        self.combobox_select_dir = QDirComboBox()
        self.combobox_select_dir.setEditable(True)
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
        self.search_box.textEdited.connect(self.text_edited)
        layout.addWidget(self.search_box)

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)

    def text_edited(self, text):

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

    def search_handler(self, dir):
        pass


app = QApplication(sys.argv)

window = MainWindow()
window.show()
if __name__ == "__main__":
    app.exec()
