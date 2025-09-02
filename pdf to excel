import pdfplumber
import pandas as pd
import os

# ask user for PDF file path
pdf_file = input("Enter the PDF file name (e.g., messy_table.pdf): ")

if not os.path.exists(pdf_file):
    print(
        f" File '{pdf_file}' not found. Make sure it’s in the same folder as this script.")
else:
    all_tables = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:  # avoid empty
                    # first row as header
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)

    if all_tables:
        final = pd.concat(all_tables, ignore_index=True)
        output_file = "extracted.xlsx"
        final.to_excel(output_file, index=False)
        print(f" Saved extracted data to {output_file}")
    else:
        print(" No tables found in this PDF.")
