import pandas as pd
import glob
import os

# ask user for folder path
folder_path = input("Enter the folder path that contains CSV files: ")

# find all CSV files in the folder
all_files = glob.glob(os.path.join(folder_path, "*.csv"))

if not all_files:
    print(" No CSV files found in that folder.")
else:
    df_list = []
    for file in all_files:
        df = pd.read_csv(file)
        df_list.append(df)

    combined = pd.concat(df_list, ignore_index=True)

    # cleaning
    combined.drop_duplicates(inplace=True)
    combined.dropna(how="all", inplace=True)

    # optional: strip spaces in column names
    combined.columns = [col.strip() for col in combined.columns]

    output_file = "cleaned_output.xlsx"
    combined.to_excel(output_file, index=False)
    print(f" Saved: {output_file}")
