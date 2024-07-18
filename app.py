import streamlit as st
import pandas as pd
import re
from io import BytesIO

def clean_kolompok(kelompok):
    return re.sub(r'\D', '', str(kelompok)).rstrip('0')

def convert_df_to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()  # Use close() instead of save()
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("Filter kelompok lebih dari 8")
    uploaded_file = st.file_uploader("Upload File Excelnya Disini", type=["xlsx"])
    if uploaded_file is not None:
        df1 = pd.read_excel(uploaded_file, skiprows=2)

        if 'Kelompok' in df1.columns:
            df1['Kelompok'] = df1['Kelompok'].apply(clean_kolompok)

            grouped = df1.groupby(['Center', 'Kelompok']).size().reset_index(name='JumlahAnggota')
            filtered_group = grouped[grouped['JumlahAnggota'] > 8]
            result = pd.merge(filtered_group, df1, on=['Center', 'Kelompok'], how='inner')
            result = result.loc[:, ['No', 'Cabang', 'ID Anggota', 'Nama Anggota', 'Center', 'Kelompok']]

            st.write("Filtered Results:")
            st.dataframe(result)

            excel_data = convert_df_to_excel(result)
            
            st.markdown("""
                ## Catatan:
                1. Ambil data dari modul Detail Nasabah SRSS
                2. Format nama pada file ini harus "Detail Nasabah.xlsx"
            """)
            st.download_button(
                label="Download File tersebut disini",
                data=excel_data,
                file_name='Kelompok_lebih_dari_8.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("Kolom 'Kelompok' tidak ditemukan dalam dataframe.")
    else:
        st.info("1.Sumber data dari Detail Nasabah SRSS 2. Ganti nama jadi Detail Nasabah.xlsx 3. Jika hasilnya blank berarti tidak ada anggota lebih dari 8 per kelompoknya.")

if __name__ == "__main__":
    main()
