import sys
from random import choice

from PySide6.QtCore import QFileSelector, QSize
from PySide6.QtWidgets import (QApplication, QFileDialog, QMainWindow,
                               QPushButton)

window_titles = [
    "My App",
    "My App",
    "Still My App",
    "Still My App",
    "What on earth",
    "What on earth",
    "This is surprising",
    "This is surprising",
    "Something went wrong",
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Lookup App")
        self.n_times_clicked = 0
        self.selected_folder = None
        self.button = QPushButton("Click to select folder")
        self.button.clicked.connect(self.button_click_handler)
        self.setCentralWidget(self.button)
        self.set

        self.setMinimumSize(QSize(400, 300))

    def directory_handler(self, directory):
        print(directory)

    def button_click_handler(self, checked):
        self.directory_selector = QFileDialog(self)
        self.directory_selector.setFileMode(QFileDialog.FileMode.Directory)
        self.directory_selector.fileSelected.connect(self.directory_handler)
        self.directory_selector.exec()


app = QApplication(sys.argv)

window = MainWindow()
window.show()
if __name__ == "__main__":
    app.exec()
