import pdfplumber
import pandas as pd
import os

def pdf_to_excel(pdf_path, excel_path):
    """Extract tables from PDF and save them into one Excel sheet."""
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
        print(f"Converted {pdf_path} to {excel_path}")

def convert_all_pdfs_in_folder(folder_path):
    """Converts all PDF files in the specified folder to one Excel file each."""
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            excel_path = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}.xlsx")
            pdf_to_excel(pdf_path, excel_path)

# Replace 'path/to/your/folder' with the path to the folder containing your PDFs
folder_path = 'X:\Python\input'
convert_all_pdfs_in_folder(folder_path)
