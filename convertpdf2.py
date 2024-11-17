import pdfplumber
import pandas as pd
import os

def pdf_to_excel(pdf_path, excel_path):
    """Extract tables from PDF and save them to an Excel file."""
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                # Convert to DataFrame and append to list
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append(df)

        # Write all tables to an Excel file
        with pd.ExcelWriter(excel_path) as writer:
            for i, df in enumerate(all_tables):
                df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
        print(f"Converted {pdf_path} to {excel_path}")

def convert_all_pdfs_in_folder(folder_path):
    """Converts all PDF files in the specified folder to Excel files."""
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            excel_path = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}.xlsx")
            pdf_to_excel(pdf_path, excel_path)

# Replace 'path/to/your/folder' with the path to the folder containing your PDFs
folder_path = 'X:\Python\input'
convert_all_pdfs_in_folder(folder_path)
