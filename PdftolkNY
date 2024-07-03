import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tabula import read_pdf
from openpyxl import load_workbook, Workbook

def extract_data_from_pdf(pdf_path):
    # Les PDF og trekk ut tabeller
    tables = read_pdf(pdf_path, pages="all", multiple_tables=True)
    
    # Kombiner alle tabellene til en enkelt DataFrame
    df = pd.concat(tables, ignore_index=True)
    
    # Gi nytt navn til kolonnene hvis nødvendig
    df.columns = ["Ordre", "Dato", "Referanse/Detaljer", "Enh.pris", "Rabatt%", "Nto.Enh.pris", "MVA", "Beløp"]
    
    # Splitt "Referanse/Detaljer" kolonnen og legg til "Antall" kolonne
    df[['Referanse/Detaljer', 'Antall']] = df['Referanse/Detaljer'].apply(split_reference_details)
    
    # Filtrer ut rader hvor "Nto.Enh.pris" er tom, null eller 0
    df['Nto.Enh.pris'].replace('', pd.NA, inplace=True)  # Erstatter tomme strenger med NaN
    df = df.dropna(subset=['Nto.Enh.pris'])  # Fjerner rader hvor "Nto.Enh.pris" er NaN
    df = df[df['Nto.Enh.pris'].astype(float) != 0]  # Fjerner rader hvor "Nto.Enh.pris" er 0
    
    # Erstatt komma med punktum i "Enh.pris" og "Nto.Enh.pris"
    df['Enh.pris'] = df['Enh.pris'].astype(str).str.replace(',', '.').astype(float)
    df['Nto.Enh.pris'] = df['Nto.Enh.pris'].astype(str).str.replace(',', '.').astype(float)
    
    # Fjern tusenskille (komma) i "Beløp"
    df['Beløp'] = df['Beløp'].astype(str).str.replace(',', '').astype(float)
    
    # Konverter "Antall" til heltall
    df['Antall'] = df['Antall'].fillna(0).astype(float).astype(int)
    
    # Konverter "Rabatt%" til prosentverdier
    df['Rabatt%'] = df['Rabatt%'].astype(str).str.replace(',', '.').str.replace('%', '').astype(float) / 100
    
    # Beregn "Rabattbeløp"
    df['Rabattbeløp'] = (df['Antall'] * df['Enh.pris']) - df['Beløp']
    df.loc[df['Rabatt%'] == 0, 'Rabattbeløp'] = pd.NA  # Sett "Rabattbeløp" til blank hvis det ikke er rabatt
    
    # Fjern de to første kolonnene (Ordre og Dato)
    df = df.drop(columns=["Ordre", "Dato"])
    
    return df

def split_reference_details(text):
    # Hvis teksten inneholder et mønster som matcher antall og produktnavn
    import re
    match = re.match(r'(\d+(\.\d+)?)\s*stk\.\s*(.*)', text)
    if match:
        return pd.Series([match.group(3), match.group(1)])
    else:
        return pd.Series([text, ''])

def select_pdf_files():
    root = tk.Tk()
    root.withdraw()
    pdf_paths = filedialog.askopenfilenames(title="Velg PDF-filer", filetypes=[("PDF files", "*.pdf")])
    root.update_idletasks()
    root.destroy()  # Sørg for at tkinter-vinduet lukkes korrekt
    return pdf_paths

def save_excel_file(dfs):
    root = tk.Tk()
    root.withdraw()
    excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if excel_path:  # Sjekk om en filbane er valgt
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for i, df in enumerate(dfs):
                sheet_name = f"Sheet{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Åpne filen med openpyxl for å justere kolonnebredden
        workbook = load_workbook(excel_path)
        for sheet in workbook.sheetnames:
            worksheet = workbook[sheet]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter  # Get the column name
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        workbook.save(excel_path)
        
    root.update_idletasks()
    root.destroy()  # Sørg for at tkinter-vinduet lukkes korrekt

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
