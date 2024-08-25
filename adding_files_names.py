import os
import docx2txt
from openpyxl import Workbook
import re


def count_words_in_docx(file_path):
    # Extract text from the docx file
    text = docx2txt.process(file_path)

    # Remove extra whitespace and split into words
    words = re.findall(r'\S+', text)

    # Return the count of words
    return len(words)


def list_docs_and_word_counts(folder_path):
    # Create a workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Document Names and Word Counts"

    # Write the header
    ws.append(["Document Name", "Word Count", "Date Modified"])

    # List all files in the folder and sort by date modified
    files = [f for f in os.listdir(folder_path) if f.endswith(".docx")]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)))

    # Loop through each file and get word count
    for filename in files:
        file_path = os.path.join(folder_path, filename)
        word_count = count_words_in_docx(file_path)
        date_modified = os.path.getmtime(file_path)
        ws.append([filename, word_count, date_modified])

    # Save the workbook
    wb.save("Document_Names_and_Word_Counts.xlsx")


# Specify the folder path
folder_path = r"G:\documents\work\translation\2024_8"
list_docs_and_word_counts(folder_path)