import os
import shutil
import sys

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

# have to import PIL but not use it any where for pictures to work but maybe now I don't anymore
# from PIL import Image

# pip install pywin32
import win32com.client as win32

# Prep file dialog
root = tk.Tk()
root.withdraw()

# noinspection PyBroadException
try:
    file_path = filedialog.askopenfilename(title="Please select excel (.xlsx) file:")
except Exception:
    print(Exception)
    messagebox.showerror(title="Greška", message="Selektovanje excel file-a neuspešno")
    sys.exit(1)
else:
    if not file_path:
        sys.exit(1)
    folder = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)

# noinspection PyBroadException
try:
    prefix_end_index = file_name.rindex('-') + 2
except Exception:
    messagebox.showerror(title="Greška", message="Ime excel-a u formatu: blablabla - bla.xlsx (mora '-' da bude)")
    sys.exit(1)
else:
    prefix = file_name[0:prefix_end_index]

# noinspection PyBroadException
try:
    pdf_file_path = filedialog.askopenfilename(title="Please select pdf file:")
except Exception:
    print(Exception)
    messagebox.showerror(title="Greška", message="Selektovanje pdf file-a neuspešno")
    sys.exit(1)
else:
    if not pdf_file_path:
        sys.exit(1)
    pdf_file_name = os.path.basename(pdf_file_path)

# noinspection PyBroadException
try:
    pdf_prefix_end_index = pdf_file_name.rindex('-') + 2
except Exception:
    messagebox.showerror(title="Greška", message="Ime pdf-a u formatu: blablabla - bla.pdf (mora '-' da bude)")
    sys.exit(1)
else:
    pdf_prefix = pdf_file_name[0:pdf_prefix_end_index]

files_to_make = simpledialog.askinteger("Koliko?", "Koliko?")

app = win32.DispatchEx("Excel.Application")
app.Visible = False

for i in range(2, files_to_make + 1):

    xlsx_path = os.path.join(folder, f"{prefix}{i}.xlsx")
    pdf_path = os.path.join(folder, f"{pdf_prefix}{i}.pdf").replace("/", "\\")

    if not os.path.isfile(xlsx_path):
        shutil.copyfile(file_path, xlsx_path)

    if os.path.isfile(pdf_path):
        continue

    wb = app.Workbooks.open(xlsx_path)
    ws = wb.Worksheets(1)
    ws.Range('I8').Value = i

    print(f'Start conversion to PDF - {i}')

    app.PrintCommunication = False

    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesTall = 3

    side_margin_val = 0.393700787401575
    footer_header_val = 0.0

    ws.PageSetup.LeftMargin = app.InchesToPoints(side_margin_val)
    ws.PageSetup.RightMargin = app.InchesToPoints(side_margin_val)
    ws.PageSetup.TopMargin = app.InchesToPoints(0.0)
    ws.PageSetup.BottomMargin = app.InchesToPoints(0.0)
    ws.PageSetup.HeaderMargin = app.InchesToPoints(footer_header_val)
    ws.PageSetup.FooterMargin = app.InchesToPoints(footer_header_val)

    app.PrintCommunication = True

    wb.WorkSheets([1]).Select()
    wb.ExportAsFixedFormat(0, pdf_path)

    wb.Close(True)

app.Quit()
