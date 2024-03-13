import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from gspread import authorize

# Sidebar
st.sidebar.header('Upload Data Baru')
uploaded_file = st.sidebar.file_uploader('Pilih file Excel', type='xlsx')
sheet_name = st.sidebar.selectbox('Pilih Sheet', ('MARET', 'Sheet2', 'Sheet3'))

# Proses data jika file diupload
if uploaded_file is not None:
    # Baca data dari file Excel
    data_baru = pd.read_excel(uploaded_file, sheet_name=sheet_name, usecols=['Bulan','Tanggal Pesanan','Customer','Kontak','Alamat','QUANTITY','CRM','Kode Trans'])

    # Google Sheets authentication
    credentials = service_account.Credentials.from_service_account_file('up-media-415604-102e9d61851c.json')
    gc = authorize(credentials)
    worksheet = gc.open('data_bersih_allCRM').worksheet(sheet_name)

    # Mencocokkan data dan menambahkan ke spreadsheet
    for index, row in data_baru.iterrows():
        # Mencari baris yang cocok di spreadsheet
        kontak = row['Kontak']
        baris_cocok = None
        for j in range(1, worksheet.row_count + 1):
            if worksheet.cell(j, 4).value == kontak:  # Kolom ke-4 (D) adalah kolom Kontak
                baris_cocok = j
                break
    
        # Jika kontak ditemukan, masukkan Qty ke kolom RO
        if baris_cocok:
            kolom_ro = 9  # Kolom RO pertama (I)
            while worksheet.cell(baris_cocok, kolom_ro).value:
                kolom_ro += 1
            worksheet.update_cell(baris_cocok, kolom_ro, row['QUANTITY'])
            worksheet.update_cell(baris_cocok, 6, row['CRM'])  # Kolom ke-6 (F) adalah kolom PIC CRM
    
        # Jika kontak tidak ditemukan, tambahkan baris baru
        else:
            data_baru_dict = {
                'Bulan': row['Bulan'],
                'Tanggal': row['Tanggal Pesanan'],
                'Nama': row['Customer'],
                'Kontak': row['Kontak'],
                'Alamat': row['Alamat'],
                'PIC CRM': row['CRM'],
                'QUANTITY': row['QUANTITY'],
            }
            worksheet.append_row(list(data_baru_dict.values()))

    # Hitung total botol
    for i in range(1, worksheet.row_count + 1):
        total_botol = 0
        for j in range(9, 25):  # Kolom RO (I - Y)
            if worksheet.cell(i, j).value:
                total_botol += worksheet.cell(i, j).value
        worksheet.update_cell(i, 25, total_botol)  # Kolom ke-25 (AA) adalah kolom TOTAL BOTOL

    # Simpan perubahan
    worksheet.save()

    # Tampilkan pesan
    st.success('Data telah berhasil diupload dan ditambahkan!')
else:
    st.info('Silakan upload file Excel data baru.')