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
selected_option = st.selectbox("Pilih salah satu:", ['32.07','32.15','32.23', '42.05','42.08','42.15','42.17'])
uploaded_file = st.file_uploader("Upload File", type="xlsx", accept_multiple_files=True)

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
        with st.spinner('Data sedang diproses...'):
            if selected_option=='22.05':
                df_2205 = pd.read_excel(uploaded_file[0], skiprows=4).fillna('')
                df_2205 = df_2205.iloc[:-5]
                df_2205 = df_2205.loc[:, ~df_2205.columns.str.startswith('Unnamed:')]
                
                excel_data = to_excel(df_2205)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='22.05.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='22.19':
                df_2219 = pd.read_excel(uploaded_file[0]).fillna('')

                # Drop the first four columns
                df_2219 = df_2219.iloc[:, 2:]
                # Drop the first three rows
                df_2219 = df_2219.iloc[3:, :]
                # Reset the index (optional, if you want a clean index)
                df_2219.reset_index(drop=True, inplace=True)
                
                # Set the first row as the header
                df_2219.columns = df_2219.iloc[0]  # Set the first row as the column headers
                df_2219 = df_2219.drop(df_2219.index[0])  # Drop the first row now that it's the header
                
                # Reset the index again (optional, if you want a clean index)
                df_2219.reset_index(drop=True, inplace=True)
                # Fill the blank "Pelanggan" cells with the preceding value
                df_2219['Nama Cabang'] = df_2219['Nama Cabang'].replace('', None).ffill()
                
                # Fill the blank "Pelanggan" cells with the preceding value
                df_2219['Pelanggan'] = df_2219['Pelanggan'].replace('', None).ffill()
                
                # Convert "Tgl. SI #" column to datetime format
                df_2219['Tgl. SI #'] = pd.to_datetime(df_2219['Tgl. SI #'], format='%d/%m/%Y')
                
                # Format "Total" column as numbers (assuming they are stored as strings)
                df_2219['Total'] = pd.to_numeric(df_2219['Total'])
                
                df_2219 =       df_2219[df_2219['Nama Cabang']      !=      "Total Nama Cabang"]

                excel_data = to_excel(df_2219)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='22.19.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            
            if selected_option=='32.07':
                df  = pd.read_excel(uploaded_file[0],header=1).fillna('')
                
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
                    file_name='32.15.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )          

            if selected_option=='32.23':
                df_3223 = pd.read_excel(uploaded_file[0], header=4).fillna('')
                df_3223 = df_3223.iloc[:-5]
                dfDPB       =       df_3223.loc[:,["Nama Cabang",
                                     "Nomor #",
                                     "Tanggal",
                                     "Tgl/Jam Pembuatan",
                                     "Pemasok",
                                     "Pengiriman"]].rename(columns={'Nomor #':'Nomor'}).fillna("")
                dfDPB['Tanggal']                = pd.to_datetime(dfDPB['Tanggal'], format='%Y-%m-%d')
                dfDPB['Tgl/Jam Pembuatan']      = pd.to_datetime(dfDPB['Tgl/Jam Pembuatan'], format='%Y-%m-%d %H:%M:%S')
                
                excel_data = to_excel(dfDPB)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='32.23.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='42.05':
                df_4205 = pd.read_excel(uploaded_file[0], header=4).fillna('')
                df_4205 = df_4205.iloc[:-5]
                df_4205 = df_4205.drop(columns=['Unnamed: 0'])
                
                # Rename columns with names like "Unnamed: 1", "Unnamed: 2", etc. to empty strings
                df_4205.rename(columns=lambda x: '' if 'Unnamed' in x else x, inplace=True)
                df_4205['Tanggal #Kirim']           =   pd.to_datetime(df_4205['Tanggal #Kirim'], format='%d-%b-%y').dt.strftime('%d %b %Y')
                df_4205['Tanggal #Terima']          =   pd.to_datetime(df_4205['Tanggal #Terima'], format='%d-%b-%y').dt.strftime('%d %b %Y')
                df_4205['#Tgl/Jam Pembuatan RI']    =   pd.to_datetime(df_4205['#Tgl/Jam Pembuatan RI'], format='%d-%b-%Y %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')

                excel_data = to_excel(df_4205)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='42.05.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='42.08':
                concatenated_df = []
                for file in uploaded_file:
                    df_4208     =   pd.read_excel(file, header=4).fillna('')
                
                    df_4208 = df_4208.drop(df_4208.columns[[0, 1]], axis=1)
                
                    df_4208     =   df_4208.rename(columns={'Kode Barang':'Nama Barang','Unnamed: 3':'Cabang','Unnamed: 4':'Nomor #',df_4208.columns[4]:'Barang','Unnamed: 9':'Tanggal','Unnamed: 11':'Deskripsi','Unnamed: 14':'Satuan','Unnamed: 16':'Masuk','Unnamed: 18':'Keluar','Unnamed: 20':'Saldo'})
                
                    # Drop columns that start with 'Unnamed'
                    df_4208 = df_4208.loc[:, ~df_4208.columns.str.startswith('Unnamed')].drop(columns=(':'))
                
                    for i in range(len(df_4208)):
                        if df_4208['Deskripsi'][i].startswith("Saldo Barang"):
                            if ((i + 1) < len(df_4208)) and (df_4208['Tanggal'][i+1] != ""):
                                df_4208.at[i, 'Tanggal'] = 'BOTTOM'
                            else:
                                df_4208.at[i, 'Tanggal'] = 'TOP'
                
                    # Forward fill 'BOTTOM' and backward fill 'TOP'
                    df_4208['Tanggal'] = df_4208['Tanggal'].replace('BOTTOM', method='bfill').replace('TOP', method='ffill')
                
                    # Check for consecutive blank rows
                    def is_blank_row(row):
                        return all(cell == '' for cell in row)
                
                    # Track indices of rows to delete
                    rows_to_delete = []
                    consecutive_blanks = 0
                
                    for i, row in df_4208.iterrows():
                        if is_blank_row(row):
                            consecutive_blanks += 1
                            if consecutive_blanks == 9:
                                # Mark the 9 rows for deletion
                                rows_to_delete.extend(range(i - 8, i + 1))
                                consecutive_blanks = 0  # Reset the counter after marking
                        else:
                            consecutive_blanks = 0  # Reset the counter if a row is not blank
                
                    # Drop the rows
                    df_4208 = df_4208.drop(rows_to_delete)
                
                    # Reset the index of the DataFrame
                    df_4208.reset_index(drop=True, inplace=True)
                
                    # Forward fill the 'Barang' column
                    df_4208['Barang'] = df_4208['Barang'].replace('', pd.NA).ffill()
                    # Forward fill the 'Cabang' column
                    df_4208['Cabang'] = df_4208['Cabang'].replace('', pd.NA).ffill()
                
                    df_4208     =   df_4208[df_4208['Nomor #']      !=      "Nomor #"]
                    df_4208     =   df_4208[df_4208['Nomor #']      !=      ""]
                
                    df_4208['Nama Barang']     =   df_4208['Barang']
                
                    # Drop the 'Barang' column
                    df_4208 = df_4208.drop(columns='Barang')
                
                    df_4208['Tanggal']      =   pd.to_datetime(df_4208['Tanggal'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d/%m/%Y')
                    concatenated_df.append(df_4208)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                
                excel_data = to_excel(df_4208)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='42.08.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='42.15':
                df_4215 = pd.read_excel(uploaded_file[0], header=4).fillna('')
                df_4215 = df_4215.iloc[:-5]
                df_4215 = df_4215.drop(columns=['Unnamed: 0','Nomor # Permintaan Barang'])
                df_4215.rename(columns=lambda x: '' if 'Unnamed' in x else x, inplace=True)
                df_4215['Tanggal']              =   pd.to_datetime(df_4215['Tanggal'], format='%d-%b-%y').dt.strftime('%d %b %Y')
                df_4215['Tgl/Jam Pembuatan']    =   pd.to_datetime(df_4215['Tgl/Jam Pembuatan'], format='%d-%b-%Y %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')

                excel_data = to_excel(df_4215)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='42.15.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='42.17':
                df_4217     =   pd.read_excel(uploaded_file[0], header=4).fillna('')
                df_4217 = df_4217.drop(columns=[x for x in df_4217.reset_index().T[(df_4217.reset_index().T[1]=='')].index if 'Unnamed' in x])
                df_4217.columns = df_4217.T.reset_index()['index'].apply(lambda x: np.nan if 'Unnamed' in x else x).ffill().values
                df_4217 = df_4217.iloc[1:,:-3]
                
                df_melted =pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang'], 
                    value_vars=df_4217.columns[6:].values,
                    var_name='Nama Cabang', value_name='Total Stok').reset_index(drop=True)

                df_melted2 = pd.melt(pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Satuan #1','Satuan #2','Satuan #3'], 
                    value_vars=df_4217.columns[6:].values,
                    var_name='Nama Cabang', value_name='Total Stok').drop_duplicates(),
                    id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Nama Cabang','Total Stok'],
                    var_name='Variabel', value_name='Satuan')

                df_melted2 = df_melted2[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Satuan','Variabel']].drop_duplicates().reset_index(drop=True)

                df_melted = df_melted.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                df_melted2 = df_melted2.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                
                df_4217_final = pd.concat([df_melted2.drop(columns='Variabel'), df_melted[['Total Stok']]], axis=1)
                df_4217_final['Kode Barang'] = df_4217_final['Kode Barang'].astype('int')

                excel_data = to_excel(df_4217_final)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name='42.17.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
