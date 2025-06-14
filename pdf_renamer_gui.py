import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

class PDFRenamerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Renamer - Drag and Drop")
        self.geometry("500x300")
        self.configure(bg="#f0f0f0")

        self.excel_path = None
        self.pdf_folder = None

        self.excel_btn = tk.Button(self, text="Select Excel Mapping File", command=self.select_excel)
        self.excel_btn.pack(pady=10)

        self.excel_label = tk.Label(self, text="No Excel file selected", bg="#f0f0f0")
        self.excel_label.pack()

        self.drop_label = tk.Label(self, text="Drag and Drop PDF Folder Here", relief="groove", bg="#e0e0e0", height=5)
        self.drop_label.pack(fill="both", padx=20, pady=20)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.drop_pdf_folder)

        self.rename_btn = tk.Button(self, text="Rename PDFs", command=self.rename_pdfs)
        self.rename_btn.pack(pady=10)

    def select_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        self.excel_label.config(text=os.path.basename(self.excel_path) if self.excel_path else "No Excel file selected")

    def drop_pdf_folder(self, event):
        path = event.data.strip("{").strip("}")
        if os.path.isdir(path):
            self.pdf_folder = path
            self.drop_label.config(text=f"Folder selected:\n{path}")
        else:
            messagebox.showerror("Invalid", "Please drop a valid folder.")

    def rename_pdfs(self):
        if not self.excel_path or not self.pdf_folder:
            messagebox.showerror("Missing Info", "Please select both Excel file and PDF folder.")
            return

        try:
            df = pd.read_excel(self.excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel: {e}")
            return

        if df.shape[1] < 2:
            messagebox.showerror("Error", "Excel must have at least 2 columns (Original and New names).")
            return

        original_col, new_col = df.columns[:2]
        renamed = 0
        missing = 0

        for _, row in df.iterrows():
            orig = str(row[original_col]).strip()
            new = str(row[new_col]).strip()
            if not orig.lower().endswith(".pdf"):
                orig += ".pdf"
            if not new.lower().endswith(".pdf"):
                new += ".pdf"

            orig_path = os.path.join(self.pdf_folder, orig)
            new_path = os.path.join(self.pdf_folder, new)

            if os.path.exists(orig_path):
                try:
                    os.rename(orig_path, new_path)
                    renamed += 1
                except Exception as e:
                    print(f"Failed to rename {orig}: {e}")
            else:
                missing += 1

        msg = f"✅ Renamed: {renamed} files\n❌ Missing: {missing} files"
        messagebox.showinfo("Done", msg)

if __name__ == "__main__":
    app = PDFRenamerApp()
    app.mainloop()
