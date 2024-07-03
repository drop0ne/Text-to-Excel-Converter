
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, scrolledtext
import os
import subprocess
import sys
import time
import openpyxl
from openpyxl.styles import Font

class ProgressBar:
    def __init__(self, total_steps, label):
        self.total_steps = total_steps
        self.current_step = 0
        self.label = label

    def start(self):
        print(f"{self.label}...")
        self.update()

    def update(self):
        self.current_step += 1
        progress = int((self.current_step / self.total_steps) * 100)
        bar = f"[{'#' * (progress // 2)}{'.' * (50 - progress // 2)}] {progress}%"
        print(f"\r{bar}", end='', flush=True)

    def finish(self):
        self.current_step = self.total_steps
        self.update()
        print()  # Move to the next line

class TextInputDialog:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("UTF-8 Text Input")
        self.text_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=100, height=40)
        self.text_area.pack(padx=10, pady=10)
        self.submit_button = tk.Button(self.root, text="Submit", command=self.on_submit)
        self.submit_button.pack(pady=5)
        self.input_text = None

    def on_submit(self):
        self.input_text = self.text_area.get("1.0", tk.END)
        self.root.quit()
        self.root.destroy()

    def get_input_text(self):
        self.root.mainloop()
        return self.input_text

class FileDialog:
    @staticmethod
    def get_save_path():
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(initialdir=r"D:\Users\Main Profile\Documents\Developer Files\My Documents", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        return file_path

class ExplorerOpener:
    @staticmethod
    def open_directory():
        path = r"D:\Users\Main Profile\Documents\Developer Files\My Documents"
        subprocess.run(f'explorer {path}', shell=True)

class ConsoleCloser:
    @staticmethod
    def close():
        sys.exit()

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
        wb = openpyxl.Workbook()
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

def main():
    print("Website to Excel Downloader")

    text_input_dialog = TextInputDialog()
    input_text = text_input_dialog.get_input_text()

    if not input_text.strip():
        print("No text provided.")
        ConsoleCloser.close()

    loading_bar = ProgressBar(total_steps=10, label="Processing")
    loading_bar.start()
    for _ in range(10):
        time.sleep(0.1)  # Simulate processing
        loading_bar.update()
    loading_bar.finish()

    file_path = FileDialog.get_save_path()
    if file_path:
        document = Document(input_text)
        document.to_excel(file_path)

        ExplorerOpener.open_directory()

        finalizing_bar = ProgressBar(total_steps=10, label="Finalizing")
        finalizing_bar.start()
        for _ in range(10):
            time.sleep(0.1)  # Simulate finalizing
            finalizing_bar.update()
        finalizing_bar.finish()

        ConsoleCloser.close()
    else:
        print("Save operation cancelled.")
        ConsoleCloser.close()

if __name__ == "__main__":
    main()
