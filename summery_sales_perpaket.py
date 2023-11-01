import gspread
import pandas as pd
from datetime import datetime

# Authenticate and open "spreadsheet generik sample"
sample_sheet_id = '4.1. Generik - Sample'
finance_sheet_id = '4.2. Generik - Finance'
sample_worksheet_name = 'Data Sampel'
finance_worksheet_name = 'Variasi Special Pricing_Paket'

try:
    # Authenticate and access "spreadsheet generik sample"
    sheet_service = gspread.service_account(filename='key.json')
    sample_worksheet = sheet_service.open(sample_sheet_id)
    sheet_ds = sample_worksheet.worksheet(sample_worksheet_name)

    # Access data from "spreadsheet generik sample"
    sample_data = pd.DataFrame(sheet_ds.get_all_records())

    # Convert 'Tanggal Pickup' to datetime
    sample_data['Tanggal Pickup'] = pd.to_datetime(sample_data['Tanggal Pickup'], format='%d/%m/%Y')

    # Filter for 'Special Price' being TRUE and not null 'Tanggal Pickup' and 'RS/Klinik'
    special_price_true = sample_data[(sample_data['Special Price'] == 'TRUE') & 
                                     sample_data['Tanggal Pickup'].notna() & 
                                     sample_data['RS/Klinik'].notna()]

    # Authenticate and access "spreadsheet generik finance"
    finance_worksheet = sheet_service.open(finance_sheet_id)
    sheet_variasi = finance_worksheet.worksheet(finance_worksheet_name)
    finance_data = pd.DataFrame(sheet_variasi.get_all_records())

    # Check for matching 'rs/klinik' in "variasi special pricing paket"
    matching_rs_klinik = finance_data[finance_data['Nama Instansi'].isin(special_price_true['RS/Klinik'])]

    # Check for 'tanggal pickup' in between 'tanggal mulai' and 'tanggal selesai'
    summary = []
    for index, row in special_price_true.iterrows():
        pickup_date = row['Tanggal Pickup']
        matching_rs = matching_rs_klinik[matching_rs_klinik['Nama Instansi'] == row['RS/Klinik']]
        for _, match_row in matching_rs.iterrows():
            start_date = pd.to_datetime(match_row['Tanggal Mulai'], format='%d/%m/%Y')
            end_date = pd.to_datetime(match_row['Tanggal Selesai'], format='%d/%m/%Y')
            if not pd.isnull(start_date) and not pd.isnull(end_date) and start_date <= pickup_date <= end_date:
                summary.append({
                    'ID Sample': row['ID Sample'],
                    'Tanggal Pickup': pickup_date,
                    'RS/Klinik': row['RS/Klinik'],
                    'Tanggal Mulai': start_date,
                    'Tabel Harga Special': match_row['Tabel Harga Special']
                })

    # Create a DataFrame from the summary data
    summary_df = pd.DataFrame(summary)

    # Filter for specific months (July, August, September, October)
    selected_months = [7, 8, 9, 10]
    summary_df = summary_df[summary_df['Tanggal Pickup'].dt.month.isin(selected_months)]

    # Calculate the counts of 'rs/klinik' and 'tabel harga special'
    id_sampel_count = summary_df['ID Sample'].nunique()
    tabel_harga_special_count = summary_df['Tabel Harga Special'].nunique()

    # Calculate the final result
    result = id_sampel_count * tabel_harga_special_count

    # Export to Excel
    result_df = pd.DataFrame({'Data Sample Paket': [id_sampel_count], 'Tabel Harga Special': [tabel_harga_special_count], 'Result': [result]})
    result_df.to_excel('summary_sales_perpaket.xlsx', index=False)

    print("Summary data successfully exported to summary_salel_perpaket.xlsx")
except Exception as e:
    print(f"An error occurred: {e}")
