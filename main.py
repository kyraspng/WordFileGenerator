from PyQt6.QtWidgets import (QApplication, QGridLayout, QLabel, QFileDialog, QPushButton, QVBoxLayout, QLineEdit,
                             QWidget, QMessageBox)
import sys
import os
import functions
from pathlib import Path

class WordFileGenerator(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()
        self.plans_file_paths = []  # Instance variable to store plans file paths
        self.trucks_file_paths = []  # Instance variable to store trucks file paths

    def initUI(self):
        # layout = QVBoxLayout()
        grid = QGridLayout()
        self.setWindowTitle("文档生成器")
        self.setLayout(grid)
        self.setGeometry(300, 300, 550, 250)
        # The first two numbers in setGeometry represent the position of the top-left corner of the window (x, y).
        # The last two numbers represent the size of the window (width, height).

        # Set the fixed width of line edit
        filepath_line_edit_fixed_width = 350
        self.filepath_line_edit_maximum_width = 400

        # Create Widget
        # Line that displays the name of the selected 《计划汇集报送》文件
        self.plans_files_line_edit = QLineEdit()
        self.plans_files_line_edit.setFixedWidth(filepath_line_edit_fixed_width)

        # 选择选择《计划汇集报送》文件button
        self.plans_files_selected_button = QPushButton("选择计划汇集报送excel", self)
        self.plans_files_selected_button.clicked.connect(lambda: self.showFileDialog(self.plans_files_line_edit,
                                                                                     'plans'))
        self.plans_files_selected_button.setFixedWidth(160)

        # Line that displays the name of the selected 《车辆信息》文件
        self.trucks_files_line_edit = QLineEdit()
        self.trucks_files_line_edit.setFixedWidth(filepath_line_edit_fixed_width)

        # 选择《车辆信息汇集》文件button
        self.trucks_files_selected_button = QPushButton("选择车辆信息excel", self)
        self.trucks_files_selected_button.clicked.connect(lambda: self.showFileDialog(self.trucks_files_line_edit,
                                                                                      'trucks'))
        self.trucks_files_selected_button.setFixedWidth(160)

        # Execution button
        self.execution_button = QPushButton("点击运行")
        self.execution_button.clicked.connect(self.generatorAll)

        # Add widgets to grid
        grid.addWidget(self.plans_files_line_edit, 0, 0)
        grid.addWidget(self.plans_files_selected_button, 0, 1)
        grid.addWidget(self.trucks_files_line_edit, 1, 0)
        grid.addWidget(self.trucks_files_selected_button, 1, 1)
        grid.addWidget(self.execution_button, 2, 0, 1, 2)

    def showFileDialog(self, line_edit, file_type):
        try:
            file_dialog = QFileDialog(self)
            file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles) # The users can select multiple existing files
            if file_dialog.exec(): # Executes the dialog. If the user selects a file and clicks "Open," it returns True.
                file_paths = file_dialog.selectedFiles() # The selectedFiles() method returns a list of selected files.

                # Store file paths in the appropriate instance variable
                if file_type == "plans":
                    self.plans_file_paths = file_paths
                elif file_type == "trucks":
                    self.trucks_file_paths = file_paths

                # Extract file names only
                file_names = [os.path.basename(file_path) for file_path in file_paths]
                display_text = "; ".join(file_names)
                line_edit.setText(display_text)
                # self.adjustLineEditSize(line_edit, display_text)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"表格错误，请重新选择")

    def adjustLineEditSize(self, line_edit, text):
        """不调整长度了"""
        # Use font metrics to determine the width of the text
        font_metrics = line_edit.fontMetrics()
        text_width = font_metrics.horizontalAdvance(text)
        # line_edit.setFixedWidth(max(text_width + 10, self.filepath_line_edit_minimum_width))
        line_edit.setFixedWidth(min(text_width + 10, self.filepath_line_edit_maximum_width))

    def plansSumWordFileGenerator(self):
        """计划汇集报送.xlsx word generator"""
        # Get the desktop path
        desktop_path = Path.home() / 'Desktop'
        if self.plans_file_paths:
            for filepath in self.plans_file_paths:
                filename_with_extension = os.path.basename(filepath)
                file_name, file_extension = os.path.splitext(filename_with_extension)
                generator = functions.PlansSumWordFileGenerate(filepath)

                generator.save_docx(desktop_path / file_name)

    def trucksSumWordFileGenerator(self):
        """车辆信息.xlsx word generator"""
        desktop_path = Path.home() / 'Desktop'
        if self.trucks_file_paths:
            for filepath in self.trucks_file_paths:
                filename_with_extension = os.path.basename(filepath)
                file_name, file_extension = os.path.splitext(filename_with_extension)
                generator = functions.TrucksWordFileGenerate(filepath)

                generator.save_docx(desktop_path / file_name)

    def generatorAll(self):
        plans_success = False
        trucks_success = False
        try:
            self.plansSumWordFileGenerator()
            plans_success = True
        except Exception as e:
            QMessageBox.critical(self, "Error", f"表格错误，请重新选择")

        try:
            self.trucksSumWordFileGenerator()
            trucks_success = True
        except Exception as e:
            QMessageBox.critical(self, "Error", f"表格错误，请重新选择")

        if plans_success and trucks_success:
            QMessageBox.information(self, "Success", "所有文档已生成并保存在桌面。")


def main():
    app = QApplication(sys.argv)
    demo = WordFileGenerator()
    demo.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()


