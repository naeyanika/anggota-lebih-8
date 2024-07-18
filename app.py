import streamlit as st
import pandas as pd
import re


def clean_kolompok(kelompok):
    return re.sub(r'\D', '', str(kelompok))


def main():
    st.title("Kelompok Analysis")


    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    
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

            @st.cache_data
            def convert_df_to_excel(df):
                return df.to_excel(index=False, engine='openpyxl')

            excel_data = convert_df_to_excel(result)

            st.title('Data Pengolahan THC Simpanan')
st.markdown("""
            ## Catatan:
            Format nama pada file ini harus "Detail Nasabah.xlsx"
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
        st.info("Please upload an Excel file to proceed.")

if __name__ == "__main__":
    main()
