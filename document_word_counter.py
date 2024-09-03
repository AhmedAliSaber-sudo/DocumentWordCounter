import os
import docx2txt
from openpyxl import load_workbook, Workbook
import re
from zipfile import BadZipFile
import win32com.client
from pptx import Presentation
import pythoncom
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


def count_words_in_file(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()

    try:
        if file_extension == '.docx':
            text = docx2txt.process(file_path)
        elif file_extension == '.doc':
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            try:
                doc = word.Documents.Open(file_path)
                text = doc.Content.Text
                doc.Close()
            finally:
                word.Quit()
                pythoncom.CoUninitialize()
        elif file_extension == '.pptx':
            prs = Presentation(file_path)
            text = " ".join(shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, 'text'))
        elif file_extension == '.xlsx':
            wb = load_workbook(file_path, read_only=True, data_only=True)
            text = " ".join(
                str(cell.value) for sheet in wb.worksheets for row in sheet.rows for cell in row if cell.value)
            wb.close()
        else:
            print(f"Unsupported file type: {file_path}")
            return None

        words = re.findall(r'\S+', text)
        return len(words)
    except BadZipFile:
        print(f"Error: {file_path} is not a valid Office file. Skipping...")
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}. Skipping...")

    return None

def list_docs_and_word_counts(folder_path, output_path, status_label):
    wb = Workbook()
    ws = wb.active
    ws.title = "Document Names and Word Counts"

    ws.append(["Document Name", "Word Count", "Date Modified"])

    file_extensions = ('.docx', '.doc', '.pptx', '.xlsx')
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(file_extensions)]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)))

    total_files = len(files)
    for index, filename in enumerate(files, 1):
        file_path = os.path.join(folder_path, filename)
        word_count = count_words_in_file(file_path)
        if word_count is not None:
            date_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
            ws.append([filename, word_count, date_modified])

        status_label.config(text=f"Processing: {index}/{total_files}")
        status_label.update()

    wb.save(output_path)
    return output_path


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Document Word Counter")
        self.geometry("500x300")

        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()

        tk.Label(self, text="Select Input Folder:").pack(pady=(10, 0))
        tk.Entry(self, textvariable=self.folder_path, width=50).pack()
        tk.Button(self, text="Browse Input", command=self.browse_input_folder).pack(pady=(0, 10))

        tk.Label(self, text="Select Output File:").pack(pady=(10, 0))
        tk.Entry(self, textvariable=self.output_path, width=50).pack()
        tk.Button(self, text="Browse Output", command=self.browse_output_file).pack(pady=(0, 10))

        self.status_label = tk.Label(self, text="")
        self.status_label.pack(pady=10)

        tk.Button(self, text="Process Files", command=self.process_files).pack(pady=10)

    def browse_input_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def browse_output_file(self):
        file_selected = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")])
        self.output_path.set(file_selected)

    def process_files(self):
        input_folder = self.folder_path.get()
        output_file = self.output_path.get()

        if not input_folder:
            messagebox.showerror("Error", "Please select an input folder")
            return
        if not output_file:
            messagebox.showerror("Error", "Please select an output file")
            return

        self.status_label.config(text="Processing...")
        self.update()

        try:
            output_path = list_docs_and_word_counts(input_folder, output_file, self.status_label)
            self.status_label.config(text="Processing complete!")
            messagebox.showinfo("Success", f"Excel file created:\n{output_path}")
        except Exception as e:
            self.status_label.config(text="An error occurred")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = Application()
    app.mainloop()