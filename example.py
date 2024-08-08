import sys
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QLabel

class FileDialogDemo(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("No file selected", self)
        layout.addWidget(self.label)

        self.button = QPushButton("Open File Dialog", self)
        self.button.clicked.connect(self.showFileDialog)
        layout.addWidget(self.button)

        self.setLayout(layout)
        self.setWindowTitle('File Dialog Example')
        self.setGeometry(300, 300, 350, 150)

    def showFileDialog(self):
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]
            self.label.setText(f"Selected file: {file_path}")

def main():
    app = QApplication(sys.argv)
    demo = FileDialogDemo()
    demo.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
