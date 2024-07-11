import sys
print(sys.executable)
print(sys.path)
import tkinter as tk
from tkinter import filedialog
from pdf2docx.converter import Converter
import os

class PDFtoWordConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF to Word Converter")

        self.label = tk.Label(master, text="Select a PDF file to convert:")
        self.label.pack()

        self.select_button = tk.Button(master, text="Select PDF", command=self.select_pdf)
        self.select_button.pack()

        self.convert_button = tk.Button(master, text="Convert", command=self.convert_pdf, state=tk.DISABLED)
        self.convert_button.pack()

        self.status_label = tk.Label(master, text="")
        self.status_label.pack()

    def select_pdf(self):
        self.pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_file:
            self.convert_button['state'] = tk.NORMAL
            self.status_label.config(text=f"Selected: {os.path.basename(self.pdf_file)}")

    def convert_pdf(self):
        if hasattr(self, 'pdf_file'):
            word_file = os.path.splitext(self.pdf_file)[0] + ".docx"
            cv = Converter(self.pdf_file)
            cv.convert(word_file)
            cv.close()
            self.status_label.config(text=f"Converted: {os.path.basename(word_file)}")
        else:
            self.status_label.config(text="Please select a PDF file first.")

root = tk.Tk()
app = PDFtoWordConverter(root)
root.mainloop()