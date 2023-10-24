# sheet_id_sample = '4.1. Generik - Sample'
# sheet_name_sample = 'Data Sampel'
# sheet_name_layanan = 'Data Layanan'

import gspread
import pandas as pd

# Fungsi untuk mengecek apakah nilai "True" atau bukan
def is_true(val):
    return val.lower() == 'true'

# Nama spreadsheet dan lembar Google
sheet_id = '4.1. Generik - Sample'
sheet_name_sample = 'Data Sampel'
sheet_name_service = 'Data Layanan'

# Mendapatkan autentikasi ke Google Sheets
sheet_service = gspread.service_account(filename='key.json')
df_sample = sheet_service.open(sheet_id).worksheet(sheet_name_sample)
df_service = sheet_service.open(sheet_id).worksheet(sheet_name_service)

# Mendapatkan data dari lembar Google
data_sample = df_sample.get_all_records()
data_service = df_service.get_all_records()

# Membuat DataFrames
df_sample = pd.DataFrame(data_sample)
df_service = pd.DataFrame(data_service)

# Menggabungkan DataFrames berdasarkan kolom "id_sample"
merged_df = df_sample.merge(df_service, on='ID Sample', how='inner')


# Filter data berdasarkan tanggal pickup untuk bulan Juli hingga Oktober
merged_df['Tanggal Pickup'] = pd.to_datetime(merged_df['Tanggal Pickup'], format='%d/%m/%Y')
df_layanan_data = merged_df[(merged_df['Tanggal Pickup'].dt.month >= 7) & (merged_df['Tanggal Pickup'].dt.month <= 10)]

# Menerapkan filter untuk harga khusus (special price) yang benar (True)
filtered_df = df_layanan_data[df_layanan_data['Special Price'].apply(is_true)]

# Membuat ringkasan berdasarkan RS/Klinik
summary_df = filtered_df.groupby(['RS/Klinik','ID Sample', 'Layanan']).size().reset_index(name='count')

# Export ke dalam format Excel
summary_df.to_excel('summary_paket.xlsx', index=False)

