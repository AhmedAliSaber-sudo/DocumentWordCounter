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
def list_docs_and_word_counts(folder_path, status_label):
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

    output_path = os.path.join(folder_path, "Document_Names_and_Word_Counts.xlsx")
    wb.save(output_path)
    return output_path


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Document Word Counter")
        self.geometry("400x200")

        self.folder_path = tk.StringVar()

        tk.Label(self, text="Select Folder:").pack(pady=10)
        tk.Entry(self, textvariable=self.folder_path, width=50).pack()
        tk.Button(self, text="Browse", command=self.browse_folder).pack(pady=5)

        self.status_label = tk.Label(self, text="")
        self.status_label.pack(pady=10)

        tk.Button(self, text="Process Files", command=self.process_files).pack(pady=10)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def process_files(self):
        folder_path = self.folder_path.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder")
            return

        self.status_label.config(text="Processing...")
        self.update()

        try:
            output_path = list_docs_and_word_counts(folder_path, self.status_label)
            self.status_label.config(text="Processing complete!")
            messagebox.showinfo("Success", f"Excel file created:\n{output_path}")
        except Exception as e:
            self.status_label.config(text="An error occurred")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = Application()
    app.mainloop()