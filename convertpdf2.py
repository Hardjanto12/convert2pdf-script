import pdfplumber
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

def pdf_to_excel_with_borders(pdf_path, excel_path):
    """Extract tables from PDF and save them into one Excel sheet with borders."""
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []

        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                # Convert each table to DataFrame and append to list
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append(df)

        # Combine all tables into one DataFrame
        combined_df = pd.concat(all_tables, ignore_index=True)

        # Save the combined DataFrame to Excel
        combined_df.to_excel(excel_path, index=False)

    # Load the saved Excel file with openpyxl to add borders
    wb = load_workbook(excel_path)
    ws = wb.active

    # Define the border style
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Apply borders to all cells in the sheet
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border_style

    # Save the modified Excel file with borders
    wb.save(excel_path)
    print(f"Converted {pdf_path} to {excel_path} with borders.")

def convert_all_pdfs_in_folder(folder_path):
    """Converts all PDF files in the specified folder to one Excel file each with borders."""
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            excel_path = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}.xlsx")
            pdf_to_excel_with_borders(pdf_path, excel_path)

# Replace 'path/to/your/folder' with the path to the folder containing your PDFs
folder_path = 'X:\Python\Bot\input'
convert_all_pdfs_in_folder(folder_path)
