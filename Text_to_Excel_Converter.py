import sys
import os
import subprocess
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTextEdit, QPushButton, 
                             QVBoxLayout, QWidget, QFileDialog, QProgressBar, QLabel)
from PyQt5.QtCore import Qt
from openpyxl import Workbook
from openpyxl.styles import Font

class Document:
    def __init__(self, text):
        self.text = text
        self.sections = []
        self.parse_text()

    def parse_text(self):
        sections = self.text.split('### ')
        for section_text in sections[1:]:
            self.sections.append(Section(section_text.strip()))

    def to_excel(self, file_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Magnetic Bearings"
        row = 1

        for section in self.sections:
            row = section.write_to_excel(ws, row)

        wb.save(file_path)

class Section:
    def __init__(self, text):
        self.title = ""
        self.subsections = []
        self.parse_text(text)

    def parse_text(self, text):
        lines = text.split('\n')
        self.title = lines[0].strip()
        subsection_text = ""
        for line in lines[1:]:
            if line.startswith('#### '):
                if subsection_text:
                    self.subsections.append(SubSection(subsection_text.strip()))
                    subsection_text = ""
                subsection_text += line + '\n'
            else:
                subsection_text += line + '\n'
        if subsection_text:
            self.subsections.append(SubSection(subsection_text.strip()))

    def write_to_excel(self, ws, row):
        ws.append([self.title])
        ws[row][0].font = Font(bold=True, size=14)
        row += 1
        for subsection in self.subsections:
            row = subsection.write_to_excel(ws, row)
        return row

class SubSection:
    def __init__(self, text):
        self.title = ""
        self.content = []
        self.parse_text(text)

    def parse_text(self, text):
        lines = text.split('\n')
        self.title = lines[0].strip()
        content = ""
        for line in lines[1:]:
            if line.startswith('1. ') or line.startswith('- '):
                if content:
                    self.content.append(content.strip())
                    content = ""
            content += line + '\n'
        if content:
            self.content.append(content.strip())

    def write_to_excel(self, ws, row):
        ws.append([self.title])
        ws[row][0].font = Font(bold=True, size=12)
        row += 1
        for content in self.content:
            ws.append([content])
            row += 1
        return row

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Website to Excel Downloader")
        self.setGeometry(100, 100, 800, 600)
        
        self.text_edit = QTextEdit(self)
        
        self.reset_button = QPushButton("Reset", self)
        self.reset_button.setEnabled(False)
        self.reset_button.clicked.connect(self.reset)

        self.exit_button = QPushButton("Exit", self)
        self.exit_button.setEnabled(True)
        self.exit_button.clicked.connect(self.exit_application)
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setAlignment(Qt.AlignCenter)

        self.label = QLabel("", self)

        self.start_button = QPushButton("Start", self)
        self.start_button.setEnabled(True)
        self.start_button.clicked.connect(self.run_application)

        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.label)
        layout.addWidget(self.start_button)
        layout.addWidget(self.reset_button)
        layout.addWidget(self.exit_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.file_path = None
        self.input_text = None

    def run_application(self):
        self.input_text = self.text_edit.toPlainText()
        if not self.input_text.strip():
            self.label.setText("No input text provided.")
            return

        self.show_progress("Processing")
        self.file_path, _ = QFileDialog.getSaveFileName(self, "Save File", os.path.expanduser("~"), "Excel Files (*.xlsx);;All Files (*)")
        if self.file_path:
            self.process_document()
            self.show_progress("Finalizing")
            self.clear_memory()
            self.reset_button.setEnabled(True)
            self.exit_button.setEnabled(True)
        else:
            self.label.setText("Save operation cancelled.")
            self.reset_button.setEnabled(True)
            self.exit_button.setEnabled(True)

    def show_progress(self, label_text):
        self.label.setText(label_text)
        for i in range(10):
            time.sleep(0.1)  # Simulate processing
            self.progress_bar.setValue((i + 1) * 10)

    def process_document(self):
        document = Document(self.input_text)
        document.to_excel(self.file_path)
        self.open_directory()

    def open_directory(self):
        path = os.path.dirname(self.file_path)
        if os.name == 'nt':
            os.startfile(path)
        elif os.name == 'posix':
            subprocess.Popen(['xdg-open', path])

    def clear_memory(self):
        self.input_text = None
        self.file_path = None
        self.progress_bar.reset()
        self.label.setText("")

    def reset(self):
        self.reset_button.setEnabled(False)
        self.exit_button.setEnabled(False)
        self.start_button.setEnabled(True)
        self.text_edit.clear()
        self.progress_bar.reset()
        self.label.setText("")

    def exit_application(self):
        self.close()
        sys.exit()

def main():
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
