import sys
from random import choice

from PySide6.QtCore import QFileSelector, QSize
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (QApplication, QComboBox, QFileDialog, QLabel,
                               QMainWindow, QPushButton, QVBoxLayout, QWidget)


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
        button = QPushButton("Click to select folder")
        button.clicked.connect(self.button_click_handler)
        self.selected_dir = ""

        layout = QVBoxLayout()
        label_select_dir = QLabel("Select dir here:")
        self.combobox_select_dir = QDirComboBox()
        self.combobox_select_dir.setEditable(True)

        layout.addWidget(label_select_dir)
        layout.addWidget(button)
        layout.addWidget(self.combobox_select_dir)

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
