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

    # 1. Сортировка по 'Количество'
    if 'Количество' in df.columns:

        df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)
        # Сортируем по количеству по убыванию
        df = df.sort_values(by='Количество', ascending=False)
    else:
        st.warning("Колонка 'Количество' не найдена. Сортировка по этому полю будет пропущена.")

    # 2. Определение приоритета из 'Артикул'
    def get_priority(row):

        article = str(row.get('Артикул', '')).lower() # Берем артикул, приводим к строке и нижнему регистру
        match = re.search(r'k(\d+)', article) # Ищем 'k' с последующими цифрами
        if match:
            return int(match.group(1)) # Возвращаем число после 'k'
        else:
            return 0 # Если 'k' не найдено, присваиваем минимальный приоритет

    df['Приоритет_Сортировки'] = df.apply(get_priority, axis=1) # Создаем временную колонку для приоритета

    # 3. Сортировка по приоритету, Наименованию товара и Артикулу

    df['Наименование товара'] = df.get('Наименование товара', pd.Series(dtype='str')).astype(str)
    df['Артикул'] = df.get('Артикул', pd.Series(dtype='str')).astype(str)

    # Сортируем:
    # - Приоритет (убывание)
    # - Наименование товара (возрастание)
    # - Артикул (возрастание)

    df = df.sort_values(by=['Приоритет_Сортировки', 'Наименование товара', 'Артикул'], ascending=[False, True, True])


    df = df.drop('Приоритет_Сортировки', axis=1)

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

