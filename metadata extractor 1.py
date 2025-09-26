from PyPDF2 import PdfReader
from docx import Document
import tkinter as tk
from tkinter import filedialog

# Open file dialog to select file
root = tk.Tk()
root.withdraw()  # Hide main window
file_path = filedialog.askopenfilename(
    title="Select PDF or DOCX File",
    filetypes=[("PDF files", "*.pdf"), ("Word files", "*.docx")]
)

if not file_path:
    print("No file selected. Exiting.")
    exit()

output_file = "metadata_output.txt"

with open(output_file, "w") as f:
    f.write(f"Metadata for: {file_path}\n\n")
    ext = file_path.split('.')[-1].lower()

    if ext == "pdf":
        try:
            reader = PdfReader(file_path)
            metadata = reader.metadata
            if metadata:
                for key, value in metadata.items():
                    f.write(f"{key}: {value}\n")
            else:
                f.write("No metadata found.\n")
        except Exception as e:
            f.write(f"Error reading PDF: {e}\n")

    elif ext == "docx":
        try:
            doc = Document(file_path)
            core = doc.core_properties
            f.write(f"Author: {core.author}\n")
            f.write(f"Title: {core.title}\n")
            f.write(f"Created: {core.created}\n")
            f.write(f"Last modified by: {core.last_modified_by}\n")
            f.write(f"Modified: {core.modified}\n")
        except Exception as e:
            f.write(f"Error reading DOCX: {e}\n")
    else:
        f.write("Unsupported file type.\n")

print(f"Metadata saved to {output_file}")