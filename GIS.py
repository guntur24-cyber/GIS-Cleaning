import pandas as pd
import glob
import numpy as np
import time
import datetime as dt
import re
import streamlit as st
from io import BytesIO
from xlsxwriter import Workbook

st.title('GIS')
selected_option = st.selectbox("Pilih salah satu:", ['32.07','32.15','32.23'])
uploaded_file = st.file_uploader("Upload File", type="xlsx")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Mengakses workbook dan worksheet untuk format header
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Menambahkan format khusus untuk header
        header_format = workbook.add_format({'border': 0, 'bold': False, 'font_size': 12})
        
        # Menulis header manual dengan format khusus
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    processed_data = output.getvalue()
    return processed_data
            
if uploaded_file is not None:
    st.write('File berhasil diupload')
    # Baca konten zip file

    if st.button('Process'):
      if selected_option=='32.07':
        df  = pd.read_excel(uploaded_file,header=1).fillna('')
        
        # Find the indices of start and end markers
        start_indices = df[df.apply(lambda row: 'Cabang :' in str(row.values), axis=1)].index
        end_indices = df[df.apply(lambda row: 'ACCURATE Accounting System Report' in str(row.values), axis=1)].index
        
        # Concatenate sections into a single DataFrame
        concatenated_df = pd.concat([df.iloc[start_idx+1:end_idx] for start_idx, end_idx in zip(start_indices, end_indices)], ignore_index=True)
        
        # Remove existing header
        concatenated_df.columns = range(concatenated_df.shape[1])
        
        # Set the second row as the new header
        new_header = concatenated_df.iloc[0]
        concatenated_df = concatenated_df[1:]
        concatenated_df.columns = new_header
        
        # Delete the blank Column
        concatenated_df = concatenated_df.loc[:,['Nomor # PR','Tanggal # PR','Nomor # PO','Tanggal # PO','Pemasok','Kode #','Nama Barang','Kuantitas','@Harga','Total Harga','Rasio Satuan','Nama Satuan','Tgl/Jam Pembuatan PO#','Tgl/Jam Pembuatan PR#']]
        
        # Drop Unnecessary Column
        concatenated_df = concatenated_df[concatenated_df['Nomor # PO']     !=      'Nomor # PO']
        concatenated_df = concatenated_df[concatenated_df['Nomor # PO']     !=      '']
        
        
        # Reset the index
        concatenated_df.reset_index(drop=True, inplace=True)
        
        concatenated_df['Tanggal # PR']           =   pd.to_datetime(concatenated_df['Tanggal # PR'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y')
        concatenated_df['Tanggal # PO']           =   pd.to_datetime(concatenated_df['Tanggal # PO'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y')
        concatenated_df['Tgl/Jam Pembuatan PO#']  =   pd.to_datetime(concatenated_df['Tgl/Jam Pembuatan PO#'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')
        concatenated_df['Tgl/Jam Pembuatan PR#']  =   pd.to_datetime(concatenated_df['Tgl/Jam Pembuatan PR#'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')
        
        
        excel_data = to_excel(concatenated_df)
        st.download_button(
            label="Download Excel",
            data=excel_data,
            file_name='32.07.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

      if selected_option=='32.15':
        concatenated_df=[]
        for file in uploaded_file:
            # Read the Excel file, skipping blank lines and using the second row as the header
            df = pd.read_excel(file, header=1).fillna('')
        
            # Initialize lists to store extracted dataframes
            extracted_dfs = []
        
            # Find the indices of start and end markers
            start_indices = df[df.apply(lambda row: 'Permintaan Barang' in str(row.values), axis=1)].index
            end_indices = df[df.apply(lambda row: 'ACCURATE Accounting System Report' in str(row.values), axis=1)].index
        
            # Loop through each pair of start and end indices
            for start_idx, end_idx in zip(start_indices, end_indices):
                # Extract the desired range of rows
                selected_rows = df.loc[start_idx:end_idx-1]
        
                # Remove existing header
                selected_rows.columns = range(selected_rows.shape[1])
        
                # Set the first row as the new header
                new_header = selected_rows.iloc[0]
                selected_rows = selected_rows[1:]
                selected_rows.columns = new_header
        
                # Delete the blank Column
                selected_rows = selected_rows.loc[:,['Permintaan Barang', 'Pesanan Pembelian', 'Penerimaan Barang', 'Uang Muka Pembelian', 'Faktur Pembelian', 'Retur Pembelian', 'Pembayaran Pembelian']]
        
                # Drop Unnecessary Column
                selected_rows = selected_rows[selected_rows['Permintaan Barang'] != 'Permintaan Barang']
                selected_rows = selected_rows[selected_rows['Permintaan Barang'] != '']
        
                # Reset the index
                selected_rows.reset_index(drop=True, inplace=True)
        
                # Append the selected DataFrame to the list
                extracted_dfs.append(selected_rows)
        
            # Concatenate all extracted dataframes
            concatenated_df.append(pd.concat(extracted_dfs, ignore_index=True))
            
        concatenated_df = pd.concat(concatenated_df, ignore_index=True)
        excel_data = to_excel(concatenated_df)
        st.download_button(
            label="Download Excel",
            data=excel_data,
            file_name='32.07.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )          
