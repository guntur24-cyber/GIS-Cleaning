import pandas as pd
import glob
import numpy as np
import time
import datetime as dt
import re
import streamlit as st

st.markdown('GIS')
selected_option = st.selectbox("Pilih salah satu:", ['32.17','32.15','32.23'])
uploaded_file = st.file_uploader("Upload File", type="xlsx")

if uploaded_file is not None:
    st.write('File berhasil diupload')
    # Baca konten zip file

    if st.button('Process'):
      if selected_option=='32.17':
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
        

        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.save()
            processed_data = output.getvalue()
            return processed_data
        
        # Konversi DataFrame ke file Excel
        excel_data = to_excel(concatenated_df)
        
        # Buat tombol unduhan untuk file Excel
        st.download_button(
            label="Unduh file Excel",
            data=excel_data,
            file_name='32.07.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

