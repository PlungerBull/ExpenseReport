import pandas as pd
import os
import ctypes

def is_hidden_windows(filepath):
    """
    Checks if a file is hidden on Windows using ctypes to access file attributes.
    """
    if os.name == 'nt':
        try:
            FILE_ATTRIBUTE_HIDDEN = 0x02
            attrs = ctypes.windll.kernel32.GetFileAttributesW(str(filepath))
            return (attrs != -1) and (attrs & FILE_ATTRIBUTE_HIDDEN)
        except Exception as e:
            print(f"Warning: Could not check hidden attribute for '{filepath}': {e}")
            return False
    return False

# --- Helper Function to simulate Power Query's "Transform File" ---
def transform_excel_file(file_path):
    """
    Reads an Excel file (assuming first sheet) and returns a DataFrame.
    This simulates the Power Query 'Transform File' custom function.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file '{file_path}': {e}")
        return pd.DataFrame()

### Main Data Processing Function
def process_sales_reports(input_folder_path, output_folder_path, period):
    """
    Consolidates and transforms data from Excel files to generate a sales report
    and saves the result to an Excel file.

    Args:
        input_folder_path (str): The path to the folder containing source Excel files.
        output_folder_path (str): The path to the folder where the output Excel file will be saved.
        period (str): The period string (e.g., '2023-12') used for naming the output file.
    """

    output_filename = f"{period}. SalesReport.xlsx"
    output_full_path = os.path.join(output_folder_path, output_filename)

    all_data_frames = []

    print(f"Scanning for Excel files in: {input_folder_path}")
    for entry in os.scandir(input_folder_path):
        if entry.is_file() and entry.name.endswith('.xlsx') and not is_hidden_windows(entry.path):
            file_path = entry.path
            file_name = entry.name
            print(f"Processing file: {file_name}")

            transformed_df = transform_excel_file(file_path)

            if not transformed_df.empty:
                # Extract company name from file name
                name_parts = file_name.replace('.xlsx', '').split('-')
                company_name = name_parts[-1] if len(name_parts) > 2 else "UNKNOWN"

                transformed_df['company'] = company_name
                all_data_frames.append(transformed_df)

    if not all_data_frames:
        print("No valid Excel files found or processed in the input directory.")
        return

    df_combined = pd.concat(all_data_frames, ignore_index=True)

    # Filter out 'FLN' company
    # df_combined = df_combined[df_combined['company'] != "FLN"].copy()

    # Change column types
    type_mapping = {
        "Fecha": 'datetime64[ns]',
        "Débito Local": float,
        "Débito Dólar": float,
        "Crédito Local": float,
        "Crédito Dolar": float
    }
    for col, dtype in type_mapping.items():
        if col in df_combined.columns:
            df_combined[col] = pd.to_numeric(df_combined[col], errors='coerce')
            if dtype == 'datetime64[ns]':
                df_combined[col] = pd.to_datetime(df_combined[col], errors='coerce')
            else:
                 df_combined[col] = df_combined[col].astype(dtype)


    # Filter rows where 'Cuenta Contable' starts with "70"
    if 'Cuenta Contable' in df_combined.columns:
        df_combined = df_combined[df_combined['Cuenta Contable'].astype(str).str.startswith("70", na=False)].copy()

    # Add 'saldoPEN' column
    df_combined['Crédito Local'] = df_combined['Crédito Local'].fillna(0)
    df_combined['Débito Local'] = df_combined['Débito Local'].fillna(0)
    df_combined['saldoPEN'] = df_combined['Crédito Local'] - df_combined['Débito Local']

    # Clean 'Fuente' column
    if 'Fuente' in df_combined.columns:
        df_combined['Fuente'] = df_combined['Fuente'].astype(str)
        df_combined['Fuente'] = (
            df_combined['Fuente']
            .str.replace("INTERFASCE ODOO ", "", regex=False)
            .str.replace("INTERFACE ODOO ", "", regex=False)
            .str.replace("N/C", "", regex=False)
            .str.replace("FAC", "", regex=False)
            .str.replace("B/V", "", regex=False)
            .str.replace("#", "", regex=False)
            .str.replace(" ", "", regex=False)
        )

    # Remove specified columns
    columns_to_remove = [
        "Tipo De Documento", "Documento", "Débito Local", "Débito Dólar",
        "Crédito Local", "Crédito Dolar", "Centro Costo", "Consecutivo",
        "Tipo De Asiento", "Glosa", "Descripción Tipo De Asiento", "Módulo",
        "Usuario Impresión", "Flujo Efectivo", "Patrimonio Neto", "Proyecto",
        "Descripción de Proyecto", "Fase", "Descripción de Fase"
    ]
    df_final = df_combined.drop(columns=[col for col in columns_to_remove if col in df_combined.columns], errors='ignore')

    # Save to Excel
    try:
        df_final.to_excel(output_full_path, index=False, sheet_name='rawData')
        print(f"\nSuccessfully processed data and saved to: {output_full_path}")

        os.startfile(output_full_path)  # This line will open the Excel file

    except Exception as e:
        print(f"Error saving or opening the Excel file: {e}")