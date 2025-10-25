import streamlit as st
import pandas as pd
import re
import io
from pypdf import PdfReader, PdfWriter
from datetime import datetime
from openpyxl.styles import Font, Alignment, Border, Side


def extract_order_number_prefix(order_string):
    """Извлекает префикс номера заказа."""
    if not isinstance(order_string, str):
        order_string = str(order_string)
    match = re.search(r'^(\d+)-', order_string)
    if match:
        return match.group(1)
    else:
        return None


def extract_sticker_from_order(order_number):
    """Извлекает стикер из номера заказа."""
    if not isinstance(order_number, str):
        order_number = str(order_number)
    match = re.search(r'(\d{4})-', order_number)
    if match:
        return match.group(1)
    else:
        return None

def sort_dataframe(df):
    """Сортирует DataFrame в соответствии с заданными приоритетами."""
    required_cols = ['Артикул', 'Количество', 'Наименование товара', 'Номер отправления', 'Стикер']
    for col in required_cols:
        if col not in df.columns:
            df[col] = ''

    df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)
    original_article_case = df['Артикул'].astype(str)  # Сохраняем оригинальный регистр
    df['Артикул_lower'] = df['Артикул'].astype(str).str.lower()
    df['Наименование товара_lower'] = df['Наименование товара'].astype(str).str.lower()

    def get_article_core(article):
        """Извлекает основную часть артикула, убирая суффиксы."""
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

    # Создаем новые признаки для приоритетов

    # Считаем повторения "Номер отправления" и "Стикер"
    df['shipment_sticker_key'] = df['Номер отправления'].astype(str) + '_' + df['Стикер'].astype(str)
    shipment_sticker_counts = df['shipment_sticker_key'].value_counts()
    df['shipment_sticker_repeated'] = df['shipment_sticker_key'].map(shipment_sticker_counts)
    df['shipment_sticker_repeated_flag'] = df['shipment_sticker_repeated'] > 1  # Flag для сортировки

    df['has_k_prefix_num'] = df['Артикул_lower'].str.contains(r'.*[k][2-5]\d*.*', na=False) #Товары с артикулом, содержащим k или K с числом от 2 до 5
    df['qty_greater_than_1'] = df['Количество'] > 1 #Товары, где количество больше 1.
    df['article_repeated'] = df['full_sticker_repeat_count'] > 1 #Товары, где повторяется артикул.


    # Считаем повторения "Наименование товара" и "Артикул"
    df['name_article_key'] = df['Наименование товара_lower'].astype(str) + '_' + df['Артикул_lower'].astype(str)
    name_article_counts = df['name_article_key'].value_counts()
    df['name_article_repeated'] = df['name_article_key'].map(name_article_counts)

    # Извлекаем число после k/K для сортировки
    df['k_num_suffix'] = 0
    k_match = df['Артикул_lower'].str.extract(r'.*[k]([2-5]\d*)$', expand=False)
    df['k_num_suffix'] = pd.to_numeric(k_match, errors='coerce').fillna(0)

    # Инициализируем столбец 'sort_level' значением по умолчанию
    df['sort_level'] = 4.0

    # Применяем маски для изменения 'sort_level'
    priority1_mask = (df['core_repeat_count'] > 1) & (df['has_k_prefix_num'])
    df.loc[priority1_mask, 'sort_level'] = 1.0
    priority2_mask = (df['full_sticker_repeat_count'] > 1) & (df['qty_greater_than_1']) & (df['sort_level'] == 4.0)
    df.loc[priority2_mask, 'sort_level'] = 2.0
    priority3_mask = (df['full_sticker_repeat_count'] > 1) & (df['sort_level'] == 4.0)
    df.loc[priority3_mask, 'sort_level'] = 3.0


    # Сортируем DataFrame в соответствии с приоритетами
    df = df.sort_values(
        by=[
            'shipment_sticker_repeated_flag',  # Приоритет 1: Повторение "Номер отправления" и "Стикер"
            'has_k_prefix_num',             # Приоритет 2: Наличие k/K с числом
            'k_num_suffix',               # Сортировка по номеру после k/K (убывание)
            'qty_greater_than_1',            # Приоритет 3: Количество > 1
            'article_repeated',              # Приоритет 4: Повторяющийся артикул
            'name_article_repeated',         # Приоритет 5: Повторение "Наименование товара" и "Артикул" (убывание)
            'sort_level',                    # Остальные критерии из вашего кода
            'article_core',
            'core_repeat_count',
            'Количество',
            'core_repeat_count',
            'Наименование товара_lower',
            'Артикул_lower'
        ],
        ascending=[
            False,                          # 'shipment_sticker_repeated_flag': Сначала True (повторяется)
            False,                          # 'has_k_prefix_num': Сначала True (есть k/K)
            False,                          # 'k_num_suffix':  Убывание (сначала больше)
            False,                          # 'qty_greater_than_1': Сначала True (Количество > 1)
            False,                          # 'article_repeated': Сначала True (повторяется)
            False,                          # 'name_article_repeated': По убыванию (сначала больше)
            True,                           # 'sort_level':  По возрастанию
            True,
            False,
            False,
            False,
            True,
            True
        ]
    )

    df['Артикул'] = original_article_case  # Восстанавливаем оригинальный регистр
    return df



def extract_sticker_data_from_pdf(pdf_file):
    """Извлекает данные стикеров из PDF."""
    sticker_data = {}
    try:
        reader = PdfReader(pdf_file)
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
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
    """Переупорядочивает страницы PDF."""
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


def get_last_4_digits(value):
    """Извлекает последние 4 цифры из значения."""
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


def customize_excel(df, combine_shipments=False):
    """Настраивает Excel файл."""
    try:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            sheet_name = 'Лист1'
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet = writer.sheets[sheet_name]

            # Auto-adjust column widths
            sheet['E1'].value = 'Кол-во'

            for column_cells in sheet.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                sheet.column_dimensions[column_cells[0].column_letter].width = length + 2

            for index, row in enumerate(df['Кол-во'], start=2):
                cell = sheet[f'E{index}']
                if isinstance(row, (int, float)) and row > 1:
                    cell.font = Font(bold=True)

            for index, row in enumerate(df['Стикер'], start=2):
                cell = sheet[f'F{index}']
                cell.font = Font(Font(bold=True))

            border_style = Border(left=Side(style='thin'),
                                  right=Side(style='thin'),
                                  top=Side(style='thin'),
                                  bottom=Side(style='thin'))

            for index, row in enumerate(df['Код'], start=2):
                cell = sheet[f'A{index}']
                cell.border = border_style

        excel_buffer.seek(0)
        return excel_buffer

    except Exception as e:
        st.error(f"Произошла ошибка при настройке Excel файла: {e}")
        return None


def main():
    """Основная логика приложения Streamlit."""
    st.set_page_config(layout="wide")
    st.title("Обработка заказов Озон: PDF и CSV")

    st.header("1. Загрузка файлов")
    uploaded_csv_file = st.file_uploader("Загрузите CSV файл с заказами", type=["csv", "txt"])
    uploaded_pdf_file = st.file_uploader("Загрузите PDF файл со стикерами", type="pdf")

    if uploaded_csv_file and uploaded_pdf_file:
        st.success("Файлы успешно загружены!")

        try:  # Обернули все в try except
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

            df_original['Стикер'] = df_original['Номер заказа'].apply(extract_order_number_prefix)
            df_with_order_prefix = df_original.dropna(subset=['Стикер']).copy()

            if df_with_order_prefix.empty:
                st.warning(
                    "Не найдено ни одного номера заказа в формате 'число-' в колонке 'Номер заказа' CSV файла. Проверьте формат номеров заказов.")
            else:
                # Шаг 1: Сортировка для группировки одинаковых номеров отправления и стикеров
                df_sorted_by_shipment_sticker = df_with_order_prefix.copy()

                # Шаг 2: Применяем основную сложную сортировку к уже сгруппированным данным
                df_sorted = sort_dataframe(df_sorted_by_shipment_sticker)

                # !!! КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: Сбрасываем индекс после всех сортировок
                df_sorted = df_sorted.reset_index(drop=True)

                # Создаем столбцы для значений, которые будут отображаться в объединенных ячейках
                df_sorted['Номер отправления для отображения'] = df_sorted['Номер отправления']
                df_sorted['Стикер для отображения'] = df_sorted['Стикер']

                num_rows = len(df_sorted)
                df_sorted['Код'] = pd.Series(range(1, num_rows + 1), index=df_sorted.index)

                df_sorted = df_sorted.rename(columns={'Количество': 'Кол-во'})

                desired_columns = ['Код', 'Номер отправления для отображения', 'Наименование товара', 'Артикул',
                                   'Кол-во', 'Стикер для отображения']
                df_for_excel = df_sorted[desired_columns].copy()

                # Переименовываем столбцы для Excel
                df_for_excel = df_for_excel.rename(columns={
                    'Номер отправления для отображения': 'Номер отправления',
                    'Стикер для отображения': 'Стикер'
                })

                pdf_sticker_data = extract_sticker_data_from_pdf(uploaded_pdf_file)

                if not pdf_sticker_data:
                    st.warning(
                        "Не удалось извлечь ни одного стикера из PDF файла. Проверьте, соответствует ли формат стикера шаблону 'FBS: 204514 XXXXX'.")
                else:
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
                            missing_pdf_pages.append(csv_identifier)

                    if missing_pdf_pages:
                        st.warning(
                            f"Следующие идентификаторы (префиксы заказов) из отсортированного CSV не были найдены в PDF: {', '.join(missing_pdf_pages)}. Страницы с соответствующими стикерами не будут включены в новый PDF.")
                    if pdf_sticker_data:
                        st.info(
                            f"Найдены заказы одному клиенту, их номер заказов: {', '.join(pdf_sticker_data.values())}. Эти страницы не будут использованы.")

                    if not pdf_pages_in_csv_order:
                        st.error(
                            "Не удалось найти соответствие между идентификаторами из CSV и стикерами из PDF. Переупорядочивание PDF невозможно.")
                    else:
                        reordered_pdf_writer = reorder_pdf_pages(uploaded_pdf_file, pdf_pages_in_csv_order)

                        if reordered_pdf_writer:
                            st.success("Страницы PDF успешно переупорядочены!")

                            # Подготовка и скачивание отсортированного Excel
                            st.header("- Лист подбора(Excel) -")

                            # Извлекаем последние 4 цифры стикера
                            df_for_excel['Стикер'] = df_for_excel['Стикер'].apply(get_last_4_digits)

                            # Указываем combine_shipments=True
                            excel_buffer = customize_excel(df_for_excel, combine_shipments=True)

                            if excel_buffer:
                                st.download_button(
                                    label="Скачать отсортированный Excel файл",
                                    data=excel_buffer.getvalue(),
                                    file_name=f"Repeats_Ozon_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )

                            # Блок для скачивания PDF
                            pdf_output_buffer = io.BytesIO()
                            reordered_pdf_writer.write(pdf_output_buffer)
                            pdf_output_buffer.seek(0)
                            st.header("- Стикеры(PDF файл) -")
                            st.write("Ваш новый PDF файл с переупорядоченными страницами:")
                            st.download_button(
                                label="Скачать Стикеры",
                                data=pdf_output_buffer,
                                file_name=f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.pdf",
                                mime="application/pdf"
                            )

        except Exception as e:  # Обработали все исключения
            st.error(f"Произошла ошибка при обработке файлов: {e}")
            st.exception(e)


if __name__ == "__main__":
    main()
