import streamlit as st
import pandas as pd
import re
import io
from pypdf import PdfReader, PdfWriter
from datetime import datetime


def extract_order_number_prefix(order_string):

    if not isinstance(order_string, str):
        order_string = str(order_string)

    match = re.search(r'^(\d+)-', order_string)
    if match:
        return match.group(1)
    else:
        return None

def extract_sticker_from_order(order_number):

    if not isinstance(order_number, str):
        order_number = str(order_number)


    match = re.search(r'(\d{4})-', order_number)
    if match:
        return match.group(1)
    else:
        return None


def sort_dataframe(df):
    required_cols = ['Артикул', 'Количество', 'Наименование товара']
    for col in required_cols:
        if col not in df.columns:
            df[col] = ''

    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)

    original_article_case = df['Артикул'].astype(str)

    df['Артикул_lower'] = df['Артикул'].astype(str).str.lower()
    df['Наименование товара_lower'] = df['Наименование товара'].astype(str).str.lower()

    def get_article_core(article):

        match = re.search(r'([a-z]\d+)$', article)
        if match:
            end_of_core = match.start()
            return article[:end_of_core].strip()
        else:
            return article.strip()

    df['article_core'] = df['Артикул_lower'].apply(get_article_core)


    core_counts = df['article_core'].value_counts()
    df['core_repeat_count'] = df['article_core'].map(core_counts)
    sticker_counts = df['Артикул_lower'].value_counts()
    df['full_sticker_repeat_count'] = df['Артикул_lower'].map(sticker_counts)
    df['has_k_prefix_num'] = df['Артикул_lower'].str.contains(r'.*[k][2-5]\d*.*', na=False)
    df['k_num_suffix'] = 0
    k_match = df['Артикул_lower'].str.extract(r'.*[k]([2-5]\d*)$', expand=False)
    df['k_num_suffix'] = pd.to_numeric(k_match, errors='coerce').fillna(0)
    has_qty_greater_than_1 = df['Количество'] > 1
    is_full_duplicate = df['full_sticker_repeat_count'] > 1
    df['sort_level'] = 4.0
    priority1_mask = (df['core_repeat_count'] > 1) & df['has_k_prefix_num']
    df.loc[priority1_mask, 'sort_level'] = 1.0
    priority2_mask = (is_full_duplicate & has_qty_greater_than_1) & (df['sort_level'] == 4.0)
    df.loc[priority2_mask, 'sort_level'] = 2.0
    priority3_mask = is_full_duplicate & (df['sort_level'] == 4.0)
    df.loc[priority3_mask, 'sort_level'] = 3.0

    # --- Cортировка ---
    # Порядок:
    # 1. sort_level (1.0, 2.0, 3.0, 4.0)
    # 2. article_core (для группировки похожих ядер) - ТОЛЬКО если sort_level одинаковый
    # 3. k_num_suffix (убывание) - ТОЛЬКО для sort_level 1.0
    # 4. core_repeat_count (убывание) - для sort_level 1.0 и 3.0
    # 5. Количество (убывание) - для sort_level 2.0 и 4.0
    # 6. Наименование товара (А-Я) - для sort_level 4.0
    # 7. Артикул (А-Я) - для стабильности

    df = df.sort_values(
        by=[
            'sort_level',
            'article_core',

            'k_num_suffix',
            'core_repeat_count',

            'Количество',

            'core_repeat_count',

            'Наименование товара_lower',

            'Артикул_lower'
        ],
        ascending=[
            True,                   # sort_level (1.0 выше 4.0)
            True,                   # article_core (А-Я)
            False,                  # k_num_suffix (убывание)
            False,                  # core_repeat_count (убывание)
            False,                  # Количество (убывание)
            False,                  # core_repeat_count (убывание) - для дубликатов
            True,                   # Наименование товара (А-Я)
            True                    # Артикул (А-Я)
        ]
    )

    df['Артикул'] = original_article_case

    df = df.drop([
        'Артикул_lower', 'Наименование товара_lower', 'article_core', 'core_repeat_count',
        'full_sticker_repeat_count', 'has_k_prefix_num', 'k_num_suffix',
        'sort_level'
    ], axis=1)

    return df

# --- Функции для работы с PDF ---

def extract_sticker_data_from_pdf(pdf_file):

    sticker_data = {}
    try:
        reader = PdfReader(pdf_file)
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                # Ищем число после "FBS: 204514"
                match = re.search(r"FBS:\s*204514\s*(\d+)", text)
                if match:
                    sticker_number = match.group(1)
                    sticker_data[page_num + 1] = sticker_number
                else:
                    pass
            else:
                pass
    except Exception as e:
        st.error(f"Ошибка при обработке PDF файла: {e}")
    return sticker_data

def reorder_pdf_pages(pdf_file, page_order_mapping):
    try:
        reader = PdfReader(pdf_file)
        writer = PdfWriter()
        pages_dict = {i + 1: page for i, page in enumerate(reader.pages)}
        for original_page_num, _ in page_order_mapping:
            if original_page_num not in pages_dict:
                st.error(f"Страница {original_page_num} из PDF не найдена. Проверьте соответствие стикеров.")
                return None
        for original_page_num, _ in page_order_mapping:
            page_to_add = pages_dict[original_page_num]
            writer.add_page(page_to_add)

        return writer

    except Exception as e:
        st.error(f"Ошибка при переупорядочивании страниц PDF: {e}")
        return None

# --- Основная логика Streamlit приложения ---
def main():
    st.set_page_config(layout="wide")
    st.title("Обработка заказов: PDF и CSV")

    st.header("1. Загрузка файлов")
    uploaded_csv_file = st.file_uploader("Загрузите CSV файл с заказами", type=["csv", "txt"])
    uploaded_pdf_file = st.file_uploader("Загрузите PDF файл со стикерами", type="pdf")

    if uploaded_csv_file and uploaded_pdf_file:
        st.success("Файлы успешно загружены!")

        st.header("2. Обработка CSV")
        try:
            try:
                df_original = pd.read_csv(uploaded_csv_file, sep=';')
            except Exception:
                try:
                    df_original = pd.read_csv(uploaded_csv_file, sep=',')
                except Exception:
                    try:
                        df_original = pd.read_csv(uploaded_csv_file, sep='\t')
                    except Exception:
                        uploaded_csv_file.seek(0)
                        df_original = pd.read_csv(io.StringIO(uploaded_csv_file.read().decode('cp1251')))

            st.write("Исходные данные CSV:")
            st.dataframe(df_original)

            df_original['Стикер'] = df_original['Номер заказа'].apply(extract_order_number_prefix)

            df_with_order_prefix = df_original.dropna(subset=['Стикер']).copy()

            if df_with_order_prefix.empty:
                st.warning(
                    "Не найдено ни одного номера заказа в формате 'число-' в колонке 'Номер заказа' CSV файла. Проверьте формат номеров заказов.")
            else:
                st.write("Данные CSV с извлеченными префиксами номеров заказов:")
                st.dataframe(df_with_order_prefix)
                df_sorted = sort_dataframe(df_with_order_prefix)
                st.write("Отсортированные данные CSV:")
                st.dataframe(df_sorted)

                st.header("3. Обработка PDF и сопоставление")

                pdf_sticker_data = extract_sticker_data_from_pdf(uploaded_pdf_file)

                if not pdf_sticker_data:
                    st.warning(
                        "Не удалось извлечь ни одного стикера из PDF файла. Проверьте, соответствует ли формат стикера шаблону 'FBS: 204514 XXXXX'.")
                else:
                    st.write("Извлеченные стикеры из PDF (страница: стикер):")
                    st.write(pdf_sticker_data)

                    pdf_pages_in_csv_order = []
                    missing_pdf_pages = []

                    for index, row in df_sorted.iterrows():

                        csv_identifier = row['Стикер']
                        found_page = None
                        for page_num, pdf_sticker_value in pdf_sticker_data.items():
                            if pdf_sticker_value == csv_identifier:
                                found_page = (page_num, pdf_sticker_value)
                                del pdf_sticker_data[page_num]
                                break

                        if found_page:
                            pdf_pages_in_csv_order.append(found_page)
                        else:

                            missing_pdf_pages.append(
                                csv_identifier)

                    if missing_pdf_pages:
                        st.warning(
                            f"Следующие идентификаторы (префиксы заказов) из отсортированного CSV не были найдены в PDF: {', '.join(missing_pdf_pages)}. Страницы с соответствующими стикерами не будут включены в новый PDF.")
                    if pdf_sticker_data:
                        st.info(
                            f"В PDF файле остались стикеры, которые не были найдены в CSV: {', '.join(pdf_sticker_data.values())}. Эти страницы не будут использованы.")

                    if not pdf_pages_in_csv_order:
                        st.error(
                            "Не удалось найти соответствие между идентификаторами из CSV и стикерами из PDF. Переупорядочивание PDF невозможно.")
                    else:
                        st.write("Порядок страниц PDF для нового файла (исходная_страница_PDF, стикер_из_PDF):")
                        st.write(pdf_pages_in_csv_order)

                        reordered_pdf_writer = reorder_pdf_pages(uploaded_pdf_file, pdf_pages_in_csv_order)

                        if reordered_pdf_writer:
                            st.success("Страницы PDF успешно переупорядочены!")

                            # --- НОВЫЙ БЛОК: Подготовка и скачивание отсортированного CSV ---
                            st.header("4. Результат")

                            columns_to_display_base = ['Номер отправления', 'Наименование товара', 'Артикул', 'Количество',
                                                       'Стикер']

                            display_data_values = {}
                            for col in columns_to_display_base:
                                if col in df_sorted.columns:
                                    display_data_values[col] = df_sorted[col]
                                elif col == 'Номер отправления':
                                    st.warning(
                                        "Колонка 'Номер отправления' не найдена. Для отображения будет использоваться 'Артикул'.")
                                    display_data_values['Номер отправления'] = df_sorted.get('Артикул',
                                                                                             pd.Series(dtype='str'))
                                else:
                                    st.warning(f"Колонка '{col}' не найдена в данных CSV.")
                                    display_data_values[col] = pd.Series(dtype='str')

                            num_rows = len(df_sorted)
                            display_data_values['Код'] = pd.Series(range(1, num_rows + 1), index=df_sorted.index)

                            df_display = pd.DataFrame(display_data_values)

                            desired_column_order = ['Код']
                            for col in df_display.columns:
                                if col != 'Код':
                                    desired_column_order.append(col)

                            df_display = df_display[desired_column_order]

                            st.write("Отсортированные данные (выбранные колонки):")
                            st.dataframe(df_display)

                            # --- БЛОК для скачивания Excel ---
                            st.header("- Лист подбора(Excel) -")

                            def get_last_4_digits(value):
                                if pd.isna(value):
                                    return ""

                                value_str = str(value)

                                match = re.search(r'(\d{4})$', value_str)

                                if match:
                                    return match.group(0)
                                else:

                                    digits_only = "".join(filter(str.isdigit, value_str))
                                    if len(digits_only) >= 4:
                                        return digits_only[-4:]
                                    else:
                                        return ""

                            df_for_excel = df_display.copy()
                            df_for_excel['Стикер'] = df_for_excel['Стикер'].apply(get_last_4_digits)

                            excel_output_buffer = io.BytesIO()

                            df_for_excel.to_excel(excel_output_buffer, index=False,
                                                  sheet_name='Последние 4 цифры стикера')

                            excel_output_buffer.seek(0)

                            st.download_button(
                                label="Скачать отсортированный Excel",
                                data=excel_output_buffer,
                                file_name = f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            # --- КОНЕЦ БЛОКА Excel ---
                            # --- Блок для скачивания Csv ---
                            st.header("- CSV файл -")
                            csv_output_buffer = io.StringIO()
                            df_display.to_csv(csv_output_buffer, index=False, sep=';', encoding='utf-8-sig')
                            csv_output_buffer.seek(0)

                            st.download_button(
                                label="Скачать отсортированный CSV",
                                data=csv_output_buffer.getvalue(),
                                file_name = f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.csv",
                                mime="text/csv"
                            )
                            # --- КОНЕЦ БЛОКА Csv ---

                            # --- Блок для скачивания PDF ---
                            pdf_output_buffer = io.BytesIO()
                            reordered_pdf_writer.write(pdf_output_buffer)
                            pdf_output_buffer.seek(0)
                            st.header("- Стикеры(PDF файл) -")
                            st.write("Ваш новый PDF файл с переупорядоченными страницами:")
                            st.download_button(
                                label="Скачать Стикеры",
                                data=pdf_output_buffer,
                                file_name = f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.pdf",
                                mime="application/pdf"
                            )
                            # --- КОНЕЦ БЛОКА PDF ---
        except Exception as e:
                    st.error(f"Произошла ошибка при обработке файлов: {e}")
                    st.exception(e)


if __name__ == "__main__":
 main()

