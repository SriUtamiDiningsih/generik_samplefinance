from datetime import datetime
import gspread
import pandas as pd

class SampleSummary:
    def __init__(self, keyfile, spreadsheet_name, worksheet_name):
        self.sheet_service = gspread.service_account(filename=keyfile)
        self.df_ds = self.sheet_service.open(spreadsheet_name).worksheet(worksheet_name)

    def get_data(self):
        try:
            data = self.df_ds.get_all_records()
            return data
        except Exception as e:
            print(f"Error while fetching data: {str(e)}")
            return None

    def generate_summary(self, data):
        if data is None:
            return None

        try:
           # Membuat DataFrame dari data
            df_dp = pd.DataFrame(data)

            # Mengonversi kolom tanggal menjadi tipe datetime
            df_dp['Tanggal Pickup'] = pd.to_datetime(df_dp['Tanggal Pickup'],  format='%d/%m/%Y')

            # Buat kolom bulan dari tanggal
            df_dp['Bulan'] = df_dp['Tanggal Pickup'].dt.strftime("%B")

            # Filter data berdasarkan bulan Juli, Agustus, September, dan Oktober
            start_date = datetime(2023, 7, 1)
            end_date = datetime(2023, 10, 31)
            filtered_data = df_dp[(df_dp['Tanggal Pickup'] >= start_date) & (df_dp['Tanggal Pickup'] <= end_date)]

            # Membuat ringkasan berdasarkan RS/Klinik
            summary = filtered_data.groupby(['RS/Klinik', 'Bulan']).size().reset_index(name='Summary Pembelian')
            return summary
        except Exception as e:
            print(f"Error while generating summary: {str(e)}")
            return None

    def export_to_excel(self, summary, output_filename):
        if summary is None:
            return

        try:
            summary.to_excel(output_filename, index=False)
            print(f"Summary berhasil disimpan dalam file Excel: {output_filename}")
        except Exception as e:
            print(f"Error while exporting to Excel: {str(e)}")

# Menggunakan kelas untuk menghasilkan summary dan ekspor ke Excel
keyfile = 'key.json'
spreadsheet_name = '4.1. Generik - Sample'
worksheet_name = 'Data Sampel'
output_filename = 'summary_RS.xlsx'

sample_summary = SampleSummary(keyfile, spreadsheet_name, worksheet_name)
data = sample_summary.get_data()
summary = sample_summary.generate_summary(data)
sample_summary.export_to_excel(summary, output_filename)
