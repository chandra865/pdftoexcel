import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader

def normalize_header(header):
    if not header:
        return []
    return [col.strip().lower().replace("\n", " ") if col else '' for col in header]

def ensure_unique_columns(columns):
    seen = {}
    for i, col in enumerate(columns):
        if col in seen:
            seen[col] += 1
            columns[i] = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
    return columns


def clean_and_align_dataframe(df, combined_df):
    df.columns = ensure_unique_columns(list(df.columns))
    for col in combined_df.columns:
        if col not in df.columns:
            df[col] = ''
    return df[combined_df.columns]


def remove_blank_rows(df):
    df = df.dropna(how='all')
    df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0] != '')]
    return df

def convert_pdf_to_excel(pdf_path, excel_path):
    print(f"Processing PDF: {pdf_path}")

    if not os.access(os.path.dirname(excel_path), os.W_OK):
        messagebox.showerror("Permission Error", f"Cannot write to {os.path.dirname(excel_path)}")
        return

    reader = PdfReader(pdf_path)
    tables_by_header = {}

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                try:
                    tables = page.extract_tables()
                    if not tables:
                        print(f"No tables found on page {page_num + 1}")
                        continue

                    for table in tables:
                        if not table or len(table) < 2 or not table[0]:
                            print(f"Skipping empty table on page {page_num + 1}")
                            continue

                        header = normalize_header(table[0])
                        if not any(header):
                            print(f"Skipping table on page {page_num + 1} due to invalid header.")
                            continue

                        header_tuple = tuple(header)
                        df = pd.DataFrame(table[1:], columns=header)
                        df.columns = ensure_unique_columns(list(df.columns))

                        print(f"Page {page_num + 1} headers: {header}")

                        if header_tuple in tables_by_header:
                            df = clean_and_align_dataframe(df, tables_by_header[header_tuple])
                            tables_by_header[header_tuple] = pd.concat([tables_by_header[header_tuple], df], ignore_index=True)
                        else:
                            tables_by_header[header_tuple] = df

                except Exception as e:
                    print(f"Error processing page {page_num + 1}: {e}")

        if tables_by_header:
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                for i, (header_tuple, combined_table) in enumerate(tables_by_header.items()):
                    combined_table.columns = ensure_unique_columns(list(combined_table.columns))
                    combined_table = remove_blank_rows(combined_table)
                    sheet_name = f"Table_{i+1}"
                    combined_table.to_excel(writer, sheet_name=sheet_name, index=False)

            messagebox.showinfo("Success", f"Data successfully saved to {excel_path}")
        else:
            messagebox.showinfo("No Tables Found", "No tables were extracted from the PDF.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def select_pdf():
    file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)

def select_excel_save_location():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Save Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, file_path)

def start_conversion():
    pdf_path = pdf_entry.get()
    excel_path = excel_entry.get()

    if not pdf_path or not excel_path:
        messagebox.showwarning("Input Required", "Please select both the PDF file and the destination to save the Excel file.")
        return

    convert_pdf_to_excel(pdf_path, excel_path)


root = tk.Tk()
root.title("PDF to Excel Converter")

tk.Label(root, text="Select PDF File:").grid(row=0, column=0, padx=10, pady=10)
pdf_entry = tk.Entry(root, width=50)
pdf_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_pdf).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Save Excel File As:").grid(row=1, column=0, padx=10, pady=10)
excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_excel_save_location).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Convert", command=start_conversion, bg="green", fg="white").grid(row=2, columnspan=3, pady=20)

root.mainloop()
