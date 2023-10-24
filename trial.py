import gspread
import pandas as pd

# Nama spreadsheet dan lembar Google
sheet_id = '4.2. Generik - Finance'
sheet_name = 'Variasi Special Pricing_Paket'

# Mendapatkan autentikasi ke Google Sheets
sheet_service = gspread.service_account(filename='key.json')
df_paket = sheet_service.open(sheet_id).worksheet(sheet_name)

# Mendapatkan data dari lembar Google
data_paket = df_paket.get_all_records()

# Membuat DataFrame
df_paket = pd.DataFrame(data_paket)

# Menerapkan filter untuk bulan Juli, Agustus, September, Oktober
df_paket['Tanggal Mulai'] = pd.to_datetime(df_paket['Tanggal Mulai'], format='%d/%m/%Y')
filtered_df = df_paket[(df_paket['Tanggal Mulai'].dt.month >= 7) & (df_paket['Tanggal Mulai'].dt.month <= 10)]

# Mengelompokkan berdasarkan nama instansi dan menghitung jumlah tabel harga special
summary_df = filtered_df.groupby('Nama Instansi')['Tabel Harga Special'].count().reset_index()

# Export ke dalam format Excel
summary_df.to_excel('summary_jumlahpaket.xlsx', index=False)
