import gspread
import pandas as pd
from datetime import datetime

# Authentication and open the Google Sheets document
sheet_id = '4.1. Generik - Sample'
worksheet_name = 'Data Sampel'

try:
    # Authenticate and access the Google Sheets document
    sheet_service = gspread.service_account(filename='key.json')
    worksheet = sheet_service.open(sheet_id).worksheet(worksheet_name)

    # Retrieve data from the "Data Sampel" worksheet
    data = pd.DataFrame(worksheet.get_all_records())

    # Convert 'Tanggal Pickup' to datetime
    data['Tanggal Pickup'] = pd.to_datetime(data['Tanggal Pickup'], format='%d/%m/%Y')

    # Filter for 'Special Price' being True
    special_price_true = data[(data['Special Price'] == 'TRUE') & data['RS/Klinik'].notnull()]

    # Filter for specific months (July, August, September, October)
    selected_months = [7, 8, 9, 10]
    data_filtered = special_price_true[special_price_true['Tanggal Pickup'].dt.month.isin(selected_months)]

    # Group and aggregate data for the summary
    summary = special_price_true.groupby('RS/Klinik')['ID Sample'].count().reset_index()
    summary.columns = ['RS/Klinik', 'Jumlah Pembelian']

    # Export to Excel
    summary.to_excel('summary_jumlah_paket.xlsx', index=False)

    print("Summary data successfully exported to summary_jumlah_paket.xlsx")
except Exception as e:
    print(f"An error occurred: {e}")