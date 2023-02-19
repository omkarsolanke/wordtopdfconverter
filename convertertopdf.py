import docx2pdf
import os
import tkinter as tk
from tkinter import filedialog

# Create Tkinter root window
root = tk.Tk()
root.withdraw()

# Get input from user for word file path
word_file_path = filedialog.askopenfilename(title="Select Word file to convert", filetypes=[("Word Documents", "*.docx")])

# Check if the user canceled the file dialog box
if not word_file_path:
    print("File selection canceled.")
    exit()

# Get input from user for output folder path
output_folder_path = filedialog.askdirectory(title="Select output folder for PDF file")

# Check if the user canceled the file dialog box
if not output_folder_path:
    print("Folder selection canceled.")
    exit()

# Convert the Word document to PDF
try:
    pdf_file_path = os.path.join(output_folder_path, os.path.splitext(os.path.basename(word_file_path))[0] + ".pdf")
    docx2pdf.convert(word_file_path, pdf_file_path)
    print("File converted successfully!")
except Exception as e:
    print("An error occurred while converting the file:", e)
