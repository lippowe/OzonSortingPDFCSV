import streamlit as st
import base64
import PyPDF2
import pandas as pd
import io
from datetime import datetime

##Стикеры из pdf файла
def get_data_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    ans = []

    for page in pdf_reader.pages:
        text = page.extract_text()
        text = text.split()
        for i in text:
            if i.isdigit():
                pair = (i, page)
                ans.append(pair)
    return ans

##Доступ к хедеру excel фалйа
def get_header_xlsx(xlsx_file):
    df = pd.read_excel(xlsx_file)
    first_two_rows = df.head(4).copy()
    selected_columns = first_two_rows.iloc[:, [0, 4]]
    title = selected_columns.columns.values[0]
    data = selected_columns.values[0][0]
    type = selected_columns.values[2][0]
    quantity = selected_columns.values[2][1]
    return (title, data, type, quantity)

def get_tables(xlsx_file):
    df = pd.read_excel(xlsx_file, skiprows = 4)
    columns_to_drop = ['Фото', 'Размер', 'Цвет']
    df = df.drop(columns=columns_to_drop)
        
    value_counts = df['Артикул продавца'].value_counts() ## общее количество
    df_repeats = df[df['Артикул продавца'].isin(value_counts[value_counts > 1].index)] ## С повторениями
    df_unique = df[df['Артикул продавца'].isin(value_counts[value_counts == 1].index)] ## Без повторений
    df_repeats_counts = df_repeats['Артикул продавца'].value_counts()
    df_unique_counts = df_unique['Артикул продавца'].value_counts()

    ##Готовые таблицы
    sorted_df = df.loc[df['Артикул продавца'].isin(value_counts.index)].sort_values(by=['Артикул продавца', 'Бренд'], key=lambda x: x.map(value_counts), ascending=[False, True])
    df_repeats_sorted = df.loc[df['Артикул продавца'].isin(df_repeats_counts.index)].sort_values(by=['Артикул продавца', 'Бренд'], key=lambda x: x.map(value_counts), ascending=[False, True]) ## C повторениями
    df_unique_sorted = df.loc[df['Артикул продавца'].isin(df_unique_counts.index)].sort_values(by=['Артикул продавца', 'Бренд'], key=lambda x: x.map(value_counts), ascending=[False, True]) ## Без повторений

    return (df_repeats_sorted, df_unique_sorted, sorted_df, df)

def create_xlsx_file(xlsx_file, sorted_df, df):
    output_buffer_xlsx = io.BytesIO()
    writer = pd.ExcelWriter(output_buffer_xlsx, engine='xlsxwriter')
    sorted_df.to_excel(writer, sheet_name='Лист подбора', index=False, startrow=4)

    workbook = writer.book
    worksheet = writer.sheets['Лист подбора']

    for idx, col in enumerate(df):
        max_len = max(df[col].astype(str).str.len().max(), len(col))
        worksheet.set_column(idx, idx, max_len + 2)

    for row_num, value in enumerate(sorted_df['Стикер'], start=0):
        worksheet.write_rich_string(row_num+5, 4, value[:-4],  workbook.add_format({'bold': True}), value[-4:]+" ", workbook.add_format({'bold': False}))

    title, data, type, quantity = get_header_xlsx(xlsx_file)
    worksheet.merge_range('A1:E1', title, workbook.add_format({'bold': False, 'font_size': 11}))
    worksheet.merge_range('A2:I2', data, workbook.add_format({'bold': True, 'font_size': 14,'bg_color': '#0000FF', 'font_color': '#FFFFFF'}))
    worksheet.merge_range('A4:D4', type)
    # worksheet.merge_range('A2:D2', quantity)
    workbook.close()
    output_buffer_xlsx.seek(0)
    return output_buffer_xlsx

def create_pdf_file(sorted_df, pdf_file):
    ans = get_data_pdf(pdf_file)
    column_data = sorted_df['Стикер']
    stickers = []

    for i in column_data:
        tmp = i.split(" ")
        tmp = "".join(tmp)
        stickers.append(tmp)

    filltered_ans = [(item[0], item[1]) for item in ans if item[0] in stickers] 
    index_map = {item: index for index, item in enumerate(stickers)}
    sorted_list = sorted(filltered_ans, key=lambda x: index_map[x[0]])

    output_pdf = PyPDF2.PdfWriter()
    for i in sorted_list:
        output_pdf.add_page(i[1])

    output_buffer_pdf = io.BytesIO()
    output_pdf.write(output_buffer_pdf)
    
    output_buffer_pdf.seek(0)
    return output_buffer_pdf

def main():
    st.title("Листы подбора и стикеры для WB")

    with st.sidebar:
        pdf_file = st.file_uploader("Загрузите PDF файл cо стикерами", type=['pdf'])
        xlsx_file = st.file_uploader("Загрузите XLSX файл с информацией о товарах", type=['xlsx'])

    if (pdf_file is not None) and (xlsx_file is not None):

        df_repeats_sorted, df_unique_sorted, sorted_df, df = get_tables(xlsx_file)

        xlsx_repeats = create_xlsx_file(xlsx_file, df_repeats_sorted, df)
        pdf_repeats = create_pdf_file(df_repeats_sorted, pdf_file)

        xlsx_unique = create_xlsx_file(xlsx_file, df_unique_sorted, df)
        pdf_unique = create_pdf_file(df_unique_sorted, pdf_file)

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("C повторением артикула продавца")
            st.download_button(label = "PDF", data = pdf_repeats.read(), file_name = f"Repeats-{datetime.now().strftime('%H-%M-%S')}.pdf")
            st.download_button(label = "XLSX", data = xlsx_repeats.read(), file_name = f"Repeats-{datetime.now().strftime('%H-%M-%S')}.xlsx")
        
        with col2:
            st.subheader("Без повторения артикула продавца")
            st.download_button(label = "PDF", data = pdf_unique.read(), file_name = f"Unique-{datetime.now().strftime('%H-%M-%S')}.pdf")
            st.download_button(label = "XLSX", data = xlsx_unique.read(), file_name = f"Unique-{datetime.now().strftime('%H-%M-%S')}.xlsx")

if __name__ == "__main__":
    main()