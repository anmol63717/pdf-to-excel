import pdfplumber
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_table_from_pdf(pdf_path):
    """Extracts tables from a PDF and returns structured data."""
    table_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if words:
                rows = {}
                for word in words:
                    top = round(word['top'])
                    text = word['text']
                    try:
                        text = text.encode('latin1').decode('utf-8')  # Decode text properly
                    except (UnicodeEncodeError, UnicodeDecodeError):
                        pass  # If decoding fails, keep the original text
                    if top not in rows:
                        rows[top] = []
                    rows[top].append(text)
                
                sorted_rows = [rows[key] for key in sorted(rows.keys())]
                table_data.extend(sorted_rows)
                print(f"Extracted words from page {page_number}: {sorted_rows}")  # Debug print
            else:
                print(f"No words found on page {page_number}")  # Debug print
    return table_data

def write_table_to_excel(table_data, excel_path):
    """Writes the extracted table data to an Excel file."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    if not table_data:
        print("No table data to write to Excel.")  # Debug print
        return

    for row_index, row in enumerate(table_data, start=1):
        for col_index, cell_value in enumerate(row, start=1):
            sheet.cell(row=row_index, column=col_index, value=cell_value)

    workbook.save(excel_path)
    print(f"Table data written to {excel_path}")  # Debug print

def select_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, pdf_path)

def select_excel():
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if excel_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, excel_path)

def extract_and_save():
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()
    if not pdf_path or not excel_path:
        messagebox.showerror("Error", "Please select both PDF and Excel file paths.")
        return
    table = extract_table_from_pdf(pdf_path)
    write_table_to_excel(table, excel_path)
    messagebox.showinfo("Success", f"Table extraction complete. Check '{excel_path}'")

# Create the main window
root = tk.Tk()
root.title("PDF Table Extractor")

# Create and place the widgets
tk.Label(root, text="Select PDF file:").grid(row=0, column=0, padx=10, pady=10)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_pdf).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Save as Excel file:").grid(row=1, column=0, padx=10, pady=10)
excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_excel).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Extract and Save", command=extract_and_save).grid(row=2, column=0, columnspan=3, padx=10, pady=20)

# Run the main loop
root.mainloop()