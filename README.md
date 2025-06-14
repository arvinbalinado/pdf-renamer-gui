# PDF Renamer GUI (Drag-and-Drop)

A simple Python GUI tool to rename hundreds of PDF files using an Excel mapping sheet.

## Features

✅ Drag and drop PDF folder  
✅ Choose Excel mapping file  
✅ Automatically renames PDFs based on original and new filename  
✅ Built with `tkinter`, `pandas`, and `tkinterDnD2`

## How to Run

1. Install dependencies:

pip install -r requirements.txt


2. Run the app:
python pdf_renamer_gui.py


## Excel Format

The Excel file must have:
- **First column**: Current file names (e.g., `old.pdf`)
- **Second column**: New file names (e.g., `new.pdf`)

## Screenshot

![PDF Renamer GUI](screenshot.png)


## Author

Arvin Balinado
