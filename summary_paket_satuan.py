import gspread
import pandas as pd
from datetime import datetime

# Authenticate and open "spreadsheet generik sample"
sample_sheet_id = '4.1. Generik - Sample'
finance_sheet_id = '4.2. Generik - Finance'
sample_worksheet_name = 'Data Sampel'
layanan_worksheet_name = 'Data Layanan'
finance_worksheet_name = 'Variasi Pricing_Satuan'

try:
    # Authenticate and access "spreadsheet generik sample"
    sheet_service = gspread.service_account(filename='key.json')
    sample_worksheet = sheet_service.open(sample_sheet_id)
    sheet_ds = sample_worksheet.worksheet(sample_worksheet_name)

    # Access data from "spreadsheet generik sample"
    sample_data = pd.DataFrame(sheet_ds.get_all_records())

    # Convert 'Tanggal Pickup' to datetime
    sample_data['Tanggal Pickup'] = pd.to_datetime(sample_data['Tanggal Pickup'], format='%d/%m/%Y')

    # Filter for 'Special Price' being FALSE
    special_price_true = sample_data[sample_data['Special Price'] == 'FALSE']

    # Authenticate and access "spreadsheet generik sample" (DATA LAYANAN)
    layanan_worksheet = sample_worksheet.worksheet(layanan_worksheet_name)
    layanan_data = pd.DataFrame(layanan_worksheet.get_all_records())

    # Get 'ID Sample' with 'layanan'
    id_samples_with_layanan = layanan_data[layanan_data['Layanan'].notna()]['ID Sample']

    # Authenticate and access "spreadsheet generik finance"
    finance_worksheet = sheet_service.open(finance_sheet_id)
    sheet_variasi= finance_worksheet.worksheet(finance_worksheet_name)
    finance_data = pd.DataFrame(sheet_variasi.get_all_records())

    # Check for matching 'rs/klinik' in "variasi special pricing paket"
    sample_rs_klinik = special_price_true['RS/Klinik'].unique()
    matching_rs_klinik = finance_data[finance_data['Nama Instansi'].isin(sample_rs_klinik)]

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
                    'Tanggal Mulai': match_row['Tanggal Mulai'],
                    'Nama Instansi':match_row['Nama Instansi'],
                    'Tabel Harga': match_row['Tabel Harga']
                })

    # Create a DataFrame from the summary data
    summary_df = pd.DataFrame(summary)

    # Filter for specific months (July, August, September, October)
    selected_months = [7, 8, 9, 10]
    summary_df = summary_df[summary_df['Tanggal Pickup'].dt.month.isin(selected_months)]

    # Group and aggregate data for the summary
    summary_grouped = summary_df.groupby(['ID Sample', 'RS/Klinik', 'Tabel Harga', 'Nama Instansi','Tanggal Mulai']).size().reset_index(name='Count')

    # Export to Excel
    summary_grouped.to_excel('summary_Paket_satuan.xlsx', index=False)

    print("Summary data successfully exported to summary_Paket_satuan.xlsx")
except Exception as e:
    print(f"An error occurred: {e}")
