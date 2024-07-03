import PyPDF2
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
import re

def extract_data_from_pdf(pdf_path):
    # Create a PDF reader object
    pdf_reader = PyPDF2.PdfReader(open(pdf_path, "rb"))
    all_text = []
    requisition_numbers = []

    # Iterate through the pages and extract text
    for page in pdf_reader.pages:
        text = page.extract_text()
        if text:
            lines = text.split("\n")
            for line in lines:
                # Extract requisition number if present
                match = re.search(r"rekvisisjonsnr\.\s*:\s*(\d+)", text)
                if match:
                    requisition_number = match.group(1)
                    requisition_numbers.append(requisition_number)
                else:
                    requisition_numbers.append(None)
                
                if "stk." in line:
                    line = line.replace("stk.", "").strip()
                    first_part, _, second_part = line.partition(" ")
                    if "15.00" in second_part:
                        second_part = second_part.replace("15.00", " ")
                    second_part = " ".join(second_part.split())  # Remove double spaces
                    parts = second_part.rsplit(" ", 4)
                    if len(parts) == 5:
                        product_details = parts[0]
                        rest_details = " ".join(parts[1:])
                    else:
                        product_details = second_part
                        rest_details = ""
                    all_text.append([first_part, product_details, rest_details])

    # Convert combined text to DataFrame
    df = pd.DataFrame(all_text, columns=["Part1", "ProductDetails", "RestDetails"])
    
    # Split the "RestDetails" column by spaces into multiple columns
    rest_details_split = df["RestDetails"].str.split(expand=True)
    df = pd.concat([df.drop(columns=["RestDetails"]), rest_details_split], axis=1)
    
    # Swap the columns "Part1" and "ProductDetails"
    columns = df.columns.tolist()
    columns[0], columns[1] = columns[1], columns[0]
    df = df[columns]

    # Rename the columns to the specified names
    df.columns = ["Navn", "Antall", "Enh.pris", "Rabatt%", "Nto.Enh.pris", "Beløp"]
    
    # Convert specified columns from text to numbers
    numeric_columns = ["Antall", "Enh.pris", "Rabatt%", "Nto.Enh.pris", "Beløp"]
    for col in numeric_columns:
        if col == "Rabatt%":
            # Remove percentage sign and convert to float
            df[col] = df[col].str.replace('%', '').astype(float) / 100
        else:
            # Remove thousand separators and convert to float
            df[col] = df[col].str.replace(',', '').astype(float)

    # Filter out rows where "Beløp" is 0
    df = df[df["Beløp"] != 0]
    
    # Add a new column "Rabatt"
    df["Rabatt"] = df["Antall"] * df["Enh.pris"] - df["Beløp"]
    
    # Add the requisition number as a new column
    df["Rekvisisjonsnr"] = requisition_numbers[:len(df)]

    return df

def select_pdf_files():
    root = tk.Tk()
    root.withdraw()
    pdf_paths = filedialog.askopenfilenames(title="Velg PDF-filer", filetypes=[("PDF files", "*.pdf")])
    root.update_idletasks()
    root.destroy()  # Ensure tkinter window is closed properly
    return pdf_paths

def save_excel_file(dfs):
    root = tk.Tk()
    root.withdraw()
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if excel_path:  # Check if a file path is selected
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for i, df in enumerate(dfs):
                sheet_name = str(i + 1)  # Use numeric names for sheets
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        root.update_idletasks()
    root.destroy()  # Ensure tkinter window is closed properly

def main():
    pdf_paths = select_pdf_files()
    if pdf_paths:
        dfs = []
        for pdf_path in pdf_paths:
            df = extract_data_from_pdf(pdf_path)
            dfs.append(df)
        save_excel_file(dfs)

if __name__ == "__main__":
    main()
