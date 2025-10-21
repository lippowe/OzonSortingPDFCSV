
import streamlit as st
import pandas as pd
import re
import csv
import io
from pypdf import PdfReader, PdfWriter # Используем pypdf

# --- Ваши существующие функции (extract_sticker, sort_dataframe) ---

def extract_sticker(order_number):
    """Извлекает последние 4 цифры перед первым дефисом."""
    if not isinstance(order_number, str):
        order_number = str(order_number)

    match = re.search(r'(\d{4})-', order_number)
    if match:
        return match.group(1)
    else:
        return ""

def sort_dataframe(df):
    """Сортирует DataFrame."""
    # Убедимся, что 'Количество' является числом для сортировки
    if 'Количество' in df.columns:
        df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0) # Заменяем нечисловые на 0
        df = df.sort_values(by='Количество', ascending=False)
    else:
        st.warning("Колонка 'Количество' не найдена для сортировки.")

    def get_priority(row):
        # Преобразуем Артикул в строку для поиска, на случай если он None или число
        article = str(row['Артикул']).lower()
        match = re.search(r'k(\d+)', article)  # Ищем "k" с одной или более цифрами после него и запоминаем число
        if match:
            return int(match.group(1))  # Возвращаем число после "k" как приоритет
        else:
            return 0  # Наименьший приоритет для тех, у кого нет "k"

    df['Приоритет'] = df.apply(get_priority, axis=1)

    # Убедимся, что 'Наименование товара' и 'Артикул' существуют и являются строками для сортировки
    if 'Наименование товара' in df.columns:
        df['Наименование товара'] = df['Наименование товара'].astype(str)
    if 'Артикул' in df.columns:
        df['Артикул'] = df['Артикул'].astype(str)

    # Сортируем сначала по приоритету (убыванию), затем по наименованию товара (возрастанию), и в конце по артикулу (возрастанию)
    df = df.sort_values(by=['Приоритет', 'Наименование товара', 'Артикул'], ascending=[False, True, True])

    df = df.drop('Приоритет', axis=1)

    return df

# --- Обновленная функция для извлечения номера стикера из текстового слоя PDF-страницы ---
def extract_sticker_number_from_pdf_page_text(page_object):
    """
    Извлекает номер стикера из текстового содержимого PDF-страницы.
    Ищет номер стикера после фразы "FBS: 204514".
    :param page_object: Объект страницы pypdf.PageObject
    """
    try:
        text = page_object.extract_text()
        if text:
            # Ищем номер стикера после "FBS: 204514"
            match = re.search(r"FBS:\s*204514\s*(\d+)", text)
            if match:
                return match.group(1)  # Возвращаем найденный номер
            else:
                return None  # Если не нашли, возвращаем None
        else:
            return None  # Если текст не извлечен, возвращаем None
    except Exception as e:
        print(f"Ошибка при извлечении текста со страницы: {e}")
        return None

# --- Основная логика Streamlit приложения ---
def main():
    st.set_page_config(layout="wide")
    st.title("Сортировка страниц PDF по данным CSV (извлечение текста)")

    st.header("1. Загрузка файлов")
    uploaded_csv_file = st.file_uploader("Загрузите CSV файл", type="csv")
    uploaded_pdf_file = st.file_uploader("Загрузите PDF файл со стикерами", type="pdf")

    # --- Обработка CSV ---
    sticker_order = [] # Инициализация списка
    if uploaded_csv_file is not None:
        st.subheader("Обработка CSV файла")
        try:
            uploaded_csv_file.seek(0)
            # Пытаемся прочитать с разделителем ';' сначала, затем ',', для большей гибкости
            try:
                df = pd.read_csv(uploaded_csv_file, encoding='utf-8', sep=';')
            except pd.errors.ParserError:
                uploaded_csv_file.seek(0)
                df = pd.read_csv(uploaded_csv_file, encoding='utf-8', sep=',')
            st.info("CSV файл успешно прочитан.")
        except Exception as e:
            st.error(f"Ошибка при чтении CSV файла: {e}")
            return

        if df is not None:
            try:
                selected_columns = ['Номер отправления', 'Наименование товара', 'Артикул', 'Количество']
                missing_cols = [col for col in selected_columns if col not in df.columns]
                if missing_cols:
                    st.error(f"В CSV файле отсутствуют следующие колонки: {', '.join(missing_cols)}. Убедитесь, что названия колонок совпадают с ожидаемыми.")
                    return
                df = df[selected_columns]
                df['Стикер'] = df['Номер отправления'].apply(extract_sticker)
                df['Артикул'] = df['Артикул'].fillna('')
                sorted_df = sort_dataframe(df.copy()) # Сортируем DataFrame

                # Извлекаем список Номеров Заказов (или стикеров, если это то же самое) из отсортированного CSV
                sticker_order = sorted_df['Номер отправления'].astype(str).tolist()
                st.info("Список номеров заказов из отсортированного CSV успешно извлечен.")

                st.subheader("Обработанные данные CSV (отсортировано для получения порядка стикеров)")
                st.dataframe(sorted_df)
                csv_data = sorted_df.to_csv(index=False, encoding='utf-8', sep=';')
                st.download_button(
                    label="Скачать обработанный CSV файл",
                    data=csv_data,
                    file_name='processed_data.csv',
                    mime='text/csv'
                )

            except Exception as e:
                st.error(f"Произошла ошибка при обработке данных CSV: {e}")
                st.exception(e)
                return

    # --- Обработка PDF и сортировка ---
    if uploaded_pdf_file is not None and sticker_order:
        st.header("2. Сортировка страниц PDF по данным CSV")

        try:
            pdf_bytes = uploaded_pdf_file.read()
            pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
            total_pdf_pages = len(pdf_reader.pages)
            st.write(f"PDF файл содержит {total_pdf_pages} страниц.")

            # Шаг 1: Создание карты "Номер стикера (из PDF-текста) -> Объект страницы"
            st.write("Извлечение номеров стикеров из страниц PDF (из текстового слоя)...")
            sticker_to_page = {}
            progress_bar = st.progress(0)

            for i, page_object in enumerate(pdf_reader.pages):
                sticker_number_from_pdf = extract_sticker_number_from_pdf_page_text(page_object)
                if sticker_number_from_pdf:
                    # Убедимся, что ключ - это строка
                    sticker_to_page[str(sticker_number_from_pdf)] = page_object
                progress_bar.progress((i + 1) / total_pdf_pages)

            st.write(f"Найдено {len(sticker_to_page)} уникальных номеров стикеров в PDF.")


            # Шаг 2: Сортировка страниц PDF в соответствии с CSV
            st.write("Сортировка страниц PDF согласно данным CSV...")
            pdf_writer = PdfWriter()
            pages_added_count = 1
            missing_order_numbers = []
            progress_bar = st.progress(0)

            for i, order_number in enumerate(sticker_order):
                # Ищем Номер отправления из CSV в карте стикеров, извлеченных из PDF
                if str(order_number) in sticker_to_page:
                    pdf_writer.add_page(sticker_to_page[str(order_number)])
                    pages_added_count += 1
                else:
                    missing_order_numbers.append(order_number)
                progress_bar.progress((i + 1) / len(sticker_order))

            if missing_order_numbers:
                st.warning(
                    f"Следующие номера заказов из CSV не найдены в PDF: "
                    f"{', '.join(map(str, missing_order_numbers[:10]))}{'...' if len(missing_order_numbers) > 10 else ''}. "
                    f"Всего не найдено уникальных номеров заказов: {len(set(missing_order_numbers))}."
                )

            if pages_added_count == 0:
                st.error("Не удалось добавить ни одной страницы в новый PDF. Проверьте номера заказов в CSV и содержимое PDF.")
                return

            # Шаг 3: Сохранение и предоставление для скачивания нового PDF
            output_pdf_bytes = io.BytesIO()
            pdf_writer.write(output_pdf_bytes)
            output_pdf_bytes.seek(0)

            st.success(f"Новый PDF файл успешно создан с {pages_added_count} страницами.")
            st.download_button(
                label="Скачать отсортированный PDF",
                data=output_pdf_bytes,
                file_name="sorted_stickers.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"Произошла ошибка при обработке PDF файла: {e}")
            st.exception(e)
            return

    elif uploaded_pdf_file is not None and not sticker_order:
        st.info("Пожалуйста, загрузите и обработайте CSV файл, чтобы использовать его для сортировки PDF.")
    elif uploaded_csv_file is None and uploaded_pdf_file is not None:
         st.info("Пожалуйста, загрузите CSV файл для сортировки страниц PDF.")

if __name__ == "__main__":
    main()
