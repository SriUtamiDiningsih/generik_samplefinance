import gspread
import pandas as pd

# Fungsi untuk menghitung jumlah 'rs/klinik' dengan 'special price' TRUE pada bulan-bulan tertentu
def count_rs_klinik_true(df, months):
    return df[(df['Special Price'] == 'TRUE') &
               df['Tanggal Pickup'].notna() &
               df['Tanggal Pickup'].dt.month.isin(months)]['RS/Klinik'].nunique()

# Fungsi untuk menghitung jumlah data 'tabel harga special' pada bulan-bulan tertentu
def count_tabel_harga_special(df, months):
    df['Tanggal Mulai'] = pd.to_datetime(df['Tanggal Mulai'], format='%d/%m/%Y')
    return df[df['Nama Instansi'].notna()&
              df['Tanggal Mulai'].dt.month.isin(months)]['Tabel Harga Special'].nunique()

try:
    # Autentikasi dan akses "spreadsheet generik sample"
    sheet_service = gspread.service_account(filename='key.json')
    sample_worksheet = sheet_service.open('4.1. Generik - Sample')
    sheet_ds = sample_worksheet.worksheet('Data Sampel')
    sample_data = pd.DataFrame(sheet_ds.get_all_records())
    sample_data['Tanggal Pickup'] = pd.to_datetime(sample_data['Tanggal Pickup'], format='%d/%m/%Y')

    # Autentikasi dan akses "spreadsheet generik finance"
    finance_worksheet = sheet_service.open('4.2. Generik - Finance')
    sheet_variasi = finance_worksheet.worksheet('Variasi Special Pricing_Paket')
    finance_data = pd.DataFrame(sheet_variasi.get_all_records())

    # Tentukan bulan-bulan yang diinginkan
    selected_months = [7, 8, 9, 10]

    # Hitung jumlah 'rs/klinik' dengan 'special price' TRUE dan jumlah 'tabel harga special' untuk bulan-bulan tertentu
    rs_klinik_count = count_rs_klinik_true(sample_data, selected_months)
    tabel_harga_special_count = count_tabel_harga_special(finance_data, selected_months)

    # Hitung hasil akhir
    result = rs_klinik_count * tabel_harga_special_count

    # Export ke Excel
    result_df = pd.DataFrame({
        'RS/Klinik': [rs_klinik_count],
        'Tabel Harga Special': [tabel_harga_special_count],
        'Result': [result]
    })
    result_df.to_excel('summary_sales_perRS.xlsx', index=False)

    print("Summary data successfully exported to summary_sales_perRS.xlsx")
except Exception as e:
    print(f"An error occurred: {e}")
