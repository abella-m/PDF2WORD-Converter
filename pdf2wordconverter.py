import sys
import tkinter as tk
from tkinter import filedialog, ttk
from pdf2docx.converter import Converter
import os

class PDFtoWordConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF to Word Converter")
        master.geometry("400x300")
        master.configure(bg="#f0f0f0")

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.frame = ttk.Frame(master, padding="20")
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(self.frame, text="Select a PDF file to convert:", font=("Arial", 12))
        self.label.grid(row=0, column=0, columnspan=2, pady=(0, 10))

        self.select_button = ttk.Button(self.frame, text="Select PDF", command=self.select_pdf)
        self.select_button.grid(row=1, column=0, padx=(0, 5))

        self.convert_button = ttk.Button(self.frame, text="Convert", command=self.convert_pdf, state=tk.DISABLED)
        self.convert_button.grid(row=1, column=1, padx=(5, 0))

        self.status_label = ttk.Label(self.frame, text="", font=("Arial", 10), wraplength=350)
        self.status_label.grid(row=2, column=0, columnspan=2, pady=(20, 0))

        self.progress_bar = ttk.Progressbar(self.frame, orient=tk.HORIZONTAL, length=350, mode='indeterminate')
        self.progress_bar.grid(row=3, column=0, columnspan=2, pady=(20, 0))

    def select_pdf(self):
        self.pdf_file = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_file:
            self.convert_button['state'] = tk.NORMAL
            self.status_label.config(text=f"Selected: {os.path.basename(self.pdf_file)}")

    def convert_pdf(self):
        if hasattr(self, 'pdf_file'):
            save_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
            if save_file:
                self.progress_bar.start()
                self.status_label.config(text="Converting...")
                self.master.update()

                cv = Converter(self.pdf_file)
                cv.convert(save_file)
                cv.close()

                self.progress_bar.stop()
                self.status_label.config(text=f"Converted: {os.path.basename(save_file)}")
        else:
            self.status_label.config(text="Please select a PDF file first.")

root = tk.Tk()
app = PDFtoWordConverter(root)
root.mainloop()