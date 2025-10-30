import streamlit as st
import pandas as pd
import re
import io
from pypdf import PdfReader, PdfWriter
from datetime import datetime
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

FBS_PREFIXES = {
    "Озон": "204514",
    "Рига": "2503733",
    "Плутон": "3021812"
}


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
    original_article_case = df['Артикул'].astype(str)
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

    df['shipment_sticker_key'] = df['Номер отправления'].astype(str) + '_' + df['Стикер'].astype(str)
    shipment_sticker_counts = df['shipment_sticker_key'].value_counts()
    df['shipment_sticker_repeated'] = df['shipment_sticker_key'].map(shipment_sticker_counts)
    df['shipment_sticker_repeated_flag'] = df['shipment_sticker_repeated'] > 1

    df['has_k_prefix_num'] = df['Артикул_lower'].str.contains(r'.*[k][2-5]\d*.*', na=False)
    df['qty_greater_than_1'] = df['Количество'] > 1
    df['article_repeated'] = df['full_sticker_repeat_count'] > 1

    df['name_article_key'] = df['Наименование товара_lower'].astype(str) + '_' + df['Артикул_lower'].astype(str)
    name_article_counts = df['name_article_key'].value_counts()
    df['name_article_repeated'] = df['name_article_key'].map(name_article_counts)

    df['k_num_suffix'] = 0
    k_match = df['Артикул_lower'].str.extract(r'.*[k]([2-6]\d*)$', expand=False)
    df['k_num_suffix'] = pd.to_numeric(k_match, errors='coerce').fillna(0)

    df['sort_level'] = 4.0
    priority1_mask = (df['core_repeat_count'] > 1) & (df['has_k_prefix_num'])
    df.loc[priority1_mask, 'sort_level'] = 1.0
    priority2_mask = (df['full_sticker_repeat_count'] > 1) & (df['qty_greater_than_1']) & (df['sort_level'] == 4.0)
    df.loc[priority2_mask, 'sort_level'] = 2.0
    priority3_mask = (df['full_sticker_repeat_count'] > 1) & (df['sort_level'] == 4.0)
    df.loc[priority3_mask, 'sort_level'] = 3.0

    df = df.sort_values(
        by=[
            'shipment_sticker_repeated_flag',  # Приоритет 1: Повторение "Номер отправления" и "Стикер"
            'has_k_prefix_num',  # Приоритет 2: Наличие k/K с числом
            'k_num_suffix',  # Сортировка по номеру после k/K (убывание)
            'qty_greater_than_1',  # Приоритет 3: Количество > 1
            'article_repeated',  # Приоритет 4: Повторяющийся артикул
            'name_article_repeated',  # Приоритет 5: Повторение "Наименование товара" и "Артикул" (убывание)
            'sort_level',  # Остальные критерии
            'article_core',
            'core_repeat_count',
            'Количество',
            'core_repeat_count',
            'Наименование товара_lower',
            'Артикул_lower'
        ],
        ascending=[
            False,  # 'shipment_sticker_repeated_flag': Сначала True (повторяется)
            False,  # 'has_k_prefix_num': Сначала True (есть k/K)
            False,  # 'k_num_suffix':  Убывание (сначала больше)
            False,  # 'qty_greater_than_1': Сначала True (Количество > 1)
            False,  # 'article_repeated': Сначала True (повторяется)
            False,  # 'name_article_repeated': По убыванию (сначала больше)
            True,  # 'sort_level':  По возрастанию
            True,
            False,
            False,
            False,
            True,
            True
        ]
    )

    df['Артикул'] = original_article_case
    return df


def extract_sticker_data_from_pdf(pdf_file, fbs_prefix):
    """Извлекает данные стикеров из PDF."""
    sticker_data = {}
    try:
        reader = PdfReader(pdf_file)
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                pattern = r"FBS:\s*" + re.escape(fbs_prefix) + r"\s*(\d+)"
                match = re.search(pattern, text)

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


def customize_excel(df, df_repeats, fbs_option, num_pdf_pages):
    """Настраивает Excel файл."""
    try:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            # === Лист 1: Основной ===
            sheet_name_main = 'Лист подбора'
            df.to_excel(writer, sheet_name=sheet_name_main, index=False, startrow=5)
            sheet_main = writer.sheets[sheet_name_main]

            # === Заголовки и инфо для основного листа ===
            sheet_main['B1'] = f'Лист подбора OZON'
            sheet_main['B1'].font = Font(bold=True, size=16)

            sheet_main['B2'] = f'Склад: {fbs_option}'
            sheet_main['B2'].font = Font(bold=True, size=13)

            sheet_main['B3'] = 'Дата: ' + datetime.now().strftime("%Y-%m-%d %H:%M")
            sheet_main['B3'].font = Font(bold=True, size=13)

            sheet_main['B4'] = f'Количество отправлений: {num_pdf_pages}'
            sheet_main['B4'].font = Font(bold=True, size=13)

            # === Настройки печати для основного листа ===
            sheet_main.page_setup.orientation = 'landscape'  # Альбомная ориентация
            sheet_main.page_setup.paperSize = 9  # Размер A4 (9 - индекс A4)
            sheet_main.page_margins.left = 0
            sheet_main.page_margins.right = 0
            sheet_main.page_margins.top = 0
            sheet_main.page_margins.bottom = 0
            sheet_main.page_margins.header = 0
            sheet_main.page_margins.footer = 0

            # Масштабирование: вписать все столбцы на одну страницу
            sheet_main.sheet_view.fitToPage = True
            sheet_main.sheet_view.zoomScale = 100  # Оставляем зум 100%, но fitToPage имеет приоритет
            sheet_main.sheet_view.zoomToFit = True  # Включаем подгонку по ширине

            # === Лист 2: Повторы ===
            if not df_repeats.empty:
                repeats_sheet_name = 'Повторы'
                # Очищаем повторяющиеся стикеры на листе "Повторы"
                df_repeats_processed = df_repeats.copy()
                seen_stickers = set()
                for index, row in df_repeats_processed.iterrows():
                    sticker = row['Стикер']
                    if sticker in seen_stickers:
                        row['Стикер'] = ''  # Очищаем стикер для повторяющихся записей
                    else:
                        seen_stickers.add(sticker)
                    df_repeats_processed.loc[index] = row

                df_repeats_processed.to_excel(writer, sheet_name=repeats_sheet_name, index=False, startrow=5)
                sheet_repeats = writer.sheets[repeats_sheet_name]

                # === Заголовки и инфо для листа повторов ===
                sheet_repeats['B1'] = f'Повторяющиеся заказы'
                sheet_repeats['B1'].font = Font(bold=True, size=16)

                sheet_repeats['B2'] = f'Склад: {fbs_option}'
                sheet_repeats['B2'].font = Font(bold=True, size=13)

                sheet_repeats['B3'] = 'Дата: ' + datetime.now().strftime("%Y-%m-%d %H:%M")
                sheet_repeats['B3'].font = Font(bold=True, size=13)

                # === Настройки печати для листа повторов ===
                sheet_repeats.page_setup.orientation = 'landscape'
                sheet_repeats.page_setup.paperSize = 9
                sheet_repeats.page_margins.left = 0
                sheet_repeats.page_margins.right = 0
                sheet_repeats.page_margins.top = 0
                sheet_repeats.page_margins.bottom = 0
                sheet_repeats.page_margins.header = 0
                sheet_repeats.page_margins.footer = 0

                sheet_repeats.sheet_view.fitToPage = True
                sheet_repeats.sheet_view.zoomScale = 100
                sheet_repeats.sheet_view.zoomToFit = True

                # === Стилизация листа повторов ===
                header_font = Font(bold=True)
                header_alignment = Alignment(horizontal='center')
                data_alignment = Alignment(horizontal='left')

                for col_num in range(1, df_repeats_processed.shape[1] + 1):
                    cell = sheet_repeats.cell(row=6, column=col_num)
                    cell.font = header_font
                    cell.alignment = header_alignment

                for row_num in range(7, sheet_repeats.max_row + 1):
                    for col_num in range(1, df_repeats_processed.shape[1] + 1):
                        cell = sheet_repeats.cell(row=row_num, column=col_num)
                        cell.alignment = data_alignment
                        if sheet_repeats.cell(row=6, column=col_num).value == 'Кол-во':
                            if isinstance(cell.value, (int, float)) and cell.value > 1:
                                cell.font = Font(bold=True)

                for col_num in range(1, df_repeats_processed.shape[1] + 1):
                    column_letter = get_column_letter(col_num)
                    max_length = 0
                    for row_num in range(6, sheet_repeats.max_row + 1):
                        cell = sheet_repeats[column_letter + str(row_num)]
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    header_cell = sheet_repeats[column_letter + '6']
                    if header_cell.value is not None:
                        max_length = max(max_length, len(str(header_cell.value)))
                    sheet_repeats.column_dimensions[column_letter].width = max_length + 2

                # === Заморозка области для листа повторов ===
                sheet_repeats.freeze_panes = 'A7'

            # === Стилизация основного листа ===
            header_font = Font(bold=True)
            header_alignment = Alignment(horizontal='center')
            data_alignment = Alignment(horizontal='left')

            for col_num in range(1, df.shape[1] + 1):
                cell = sheet_main.cell(row=6, column=col_num)
                cell.font = header_font
                cell.alignment = header_alignment

            for row_num in range(7, sheet_main.max_row + 1):
                for col_num in range(1, df.shape[1] + 1):
                    cell = sheet_main.cell(row=row_num, column=col_num)
                    cell.alignment = data_alignment

                    if sheet_main.cell(row=6, column=col_num).value == 'Кол-во':
                        if isinstance(cell.value, (int, float)) and cell.value > 1:
                            cell.font = Font(bold=True)

            for col_num in range(1, df.shape[1] + 1):
                column_letter = get_column_letter(col_num)
                max_length = 0
                for row_num in range(6, sheet_main.max_row + 1):
                    cell = sheet_main[column_letter + str(row_num)]
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                header_cell = sheet_main[column_letter + '6']
                if header_cell.value is not None:
                    max_length = max(max_length, len(str(header_cell.value)))

                sheet_main.column_dimensions[column_letter].width = max_length + 2

            # === Заморозка области для основного листа ===
            sheet_main.freeze_panes = 'A7'

        excel_buffer.seek(0)
        return excel_buffer

    except Exception as e:
        st.error(f"Произошла ошибка при настройке Excel файла: {e}")
        st.error(f"Тип ошибки: {type(e)}")
        st.error(f"Аргументы ошибки: {e.args}")
        st.exception(e)
        return None


def read_csv_with_encoding(uploaded_csv_file):
    """
    Пытается прочитать CSV файл с разными кодировками и определяет столбец 'Наименование товара'.
    """
    encodings_to_try = ['utf-8', 'cp1251', 'latin1']
    sep_options = [';', ',', '\t']
    possible_name_columns = ['Наименование товара', 'Название товара']

    for sep in sep_options:
        for encoding in encodings_to_try:
            try:
                df = pd.read_csv(uploaded_csv_file, sep=sep, encoding=encoding)
                # st.write(f"Файл успешно прочитан с кодировкой '{encoding}' и разделителем '{sep}'")

                name_column = None
                for col in possible_name_columns:
                    if col in df.columns:
                        name_column = col
                        break

                if name_column is None:
                    st.error(f"Не найден столбец с наименованием товара. Проверены: {possible_name_columns}")
                    # st.write("Список столбцов в DataFrame:")
                    st.write(df.columns.tolist())
                    return None

                if name_column != 'Наименование товара':
                    df = df.rename(columns={name_column: 'Наименование товара'})
                    # st.write(f"Столбец '{name_column}' переименован в 'Наименование товара'")

                # st.write("Список столбцов в DataFrame:")
                # st.write(df.columns.tolist())
                return df

            except UnicodeDecodeError:
                st.write(f"Ошибка UnicodeDecodeError при чтении с кодировкой '{encoding}' и разделителем '{sep}'")
                pass
            except Exception as e:
                st.write(f"Другая ошибка при чтении с разделителем '{sep}' и кодировкой '{encoding}': {e}")
                pass

    try:
        uploaded_csv_file.seek(0)
        df = pd.read_csv(io.StringIO(uploaded_csv_file.read().decode('cp1251')))

        name_column = None
        for col in possible_name_columns:
            if col in df.columns:
                name_column = col
                break

        if name_column is None:
            st.error(f"Не найден столбец с наименованием товара (StringIO). Проверены: {possible_name_columns}")
            st.write("Список столбцов в DataFrame (StringIO):")
            st.write(df.columns.tolist())
            return None

        if name_column != 'Наименование товара':
            df = df.rename(columns={name_column: 'Наименование товара'})
            st.write(f"Столбец '{name_column}' переименован в 'Наименование товара' (StringIO)")

        st.write("Файл успешно прочитан с использованием StringIO и кодировки 'cp1251'")
        st.write("Список столбцов в DataFrame (StringIO):")
        st.write(df.columns.tolist())
        return df

    except Exception as e:
        st.error(f"Не удалось прочитать CSV файл ни с одной из предложенных кодировок/разделителей. Ошибка: {e}")
        return None


def main():
    """Основная логика приложения Streamlit."""
    st.set_page_config(layout="wide")
    st.title("Обработка заказов Озон: PDF и CSV")

    fbs_option = st.selectbox("Выберите тип FBS", list(FBS_PREFIXES.keys()))
    fbs_prefix = FBS_PREFIXES[fbs_option]

    st.header("1. Загрузка файлов")
    uploaded_csv_file = st.file_uploader("Загрузите CSV файл с заказами", type=["csv", "txt"])
    uploaded_pdf_file = st.file_uploader("Загрузите PDF файл со стикерами", type="pdf")

    if uploaded_csv_file and uploaded_pdf_file:
        st.success("Файлы успешно загружены!")

        try:
            df_original = read_csv_with_encoding(uploaded_csv_file)

            if df_original is None:
                st.stop()

            # st.write(f"Тип данных столбца 'Наименование товара': {df_original['Наименование товара'].dtype}")
            # st.write(f"Количество NaN в столбце 'Наименование товара': {df_original['Наименование товара'].isnull().sum()}")
            df_original['Наименование товара'] = df_original['Наименование товара'].astype(str).fillna('')

            df_original['Стикер'] = df_original['Номер заказа'].apply(extract_order_number_prefix)
            df_with_order_prefix = df_original.dropna(subset=['Стикер']).copy()

            if df_with_order_prefix.empty:
                st.warning(
                    "Не найдено ни одного номера заказа в формате 'число-' в колонке 'Номер заказа' CSV файла. Проверьте формат номеров заказов.")
            else:
                df_sorted_by_shipment_sticker = df_with_order_prefix.copy()
                df_sorted = sort_dataframe(df_sorted_by_shipment_sticker)

                df_sorted = df_sorted.reset_index(drop=True)

                df_sorted['Номер отправления для отображения'] = df_sorted['Номер отправления']
                df_sorted['Стикер для отображения'] = df_sorted['Стикер']
                df_repeats = df_sorted[df_sorted['shipment_sticker_repeated_flag']].copy()
                df_repeats = df_repeats.sort_values(by=['Номер отправления'])
                df_sorted = df_sorted[~df_sorted['shipment_sticker_repeated_flag']].copy()

                num_rows = len(df_sorted)
                df_sorted['Код'] = pd.Series(range(1, num_rows + 1), index=df_sorted.index)

                start_num_repeats = df_sorted['Код'].max() + 1 if not df_sorted.empty else 1
                num_rows_repeats = len(df_repeats)
                df_repeats['Код'] = pd.Series(range(start_num_repeats, start_num_repeats + num_rows_repeats),
                                              index=df_repeats.index)

                df_sorted = df_sorted.rename(columns={'Количество': 'Кол-во'})
                df_repeats = df_repeats.rename(columns={'Количество': 'Кол-во'})

                desired_columns = ['Код', 'Номер отправления для отображения', 'Наименование товара', 'Артикул',
                                   'Кол-во', 'Стикер для отображения']

                df_for_excel = df_sorted[desired_columns].copy()
                df_repeats_for_excel = df_repeats[desired_columns].copy()
                df_for_excel = df_for_excel.rename(columns={
                    'Номер отправления для отображения': 'Номер отправления',
                    'Стикер для отображения': 'Стикер'
                })

                df_repeats_for_excel = df_repeats_for_excel.rename(columns={
                    'Номер отправления для отображения': 'Номер отправления',
                    'Стикер для отображения': 'Стикер'
                })

                # ==Отладочный вывод DataFrame перед Excel==
                # st.write("DataFrame основной перед функцией customize_excel:")
                # st.write(df_for_excel)

                # st.write("DataFrame повторов перед функцией customize_excel:")
                # st.write(df_repeats_for_excel)

                pdf_sticker_data = extract_sticker_data_from_pdf(uploaded_pdf_file, fbs_prefix)

                reader = PdfReader(uploaded_pdf_file)
                num_pdf_pages = len(reader.pages)

                if not pdf_sticker_data:
                    st.warning(
                        f"Не удалось извлечь ни одного стикера из PDF файла. Проверьте, соответствует ли формат стикера шаблону 'FBS: {fbs_prefix} XXXXX'.")
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

                    for index, row in df_repeats.iterrows():
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
                            f"Следующие заказы из подбора листа отмечены как соединёнными: {', '.join(missing_pdf_pages)}. Список будет выписан на новый лист (Повторы) в excel файле.")
                    if pdf_sticker_data:
                        st.info(
                            f"Найдены заказы одному клиенту, их номер заказов: {', '.join(pdf_sticker_data.values())}. Эти страницы не будут использованы.")

                    if not pdf_pages_in_csv_order:
                        st.error(
                            "Не удалось найти соответствие между идентификаторами из CSV и стикерами из PDF. Переупорядочивание PDF невозможно.")
                    else:
                        reordered_pdf_writer = reorder_pdf_pages(uploaded_pdf_file, pdf_pages_in_csv_order)

                        if reordered_pdf_writer:
                            st.success("Стикеры успешно переупорядочены!")

                            st.header("- Лист подбора(Excel) -")

                            df_for_excel['Стикер'] = df_for_excel['Стикер'].apply(get_last_4_digits)
                            df_repeats_for_excel['Стикер'] = df_repeats_for_excel['Стикер'].apply(get_last_4_digits)

                            excel_buffer = customize_excel(df_for_excel, df_repeats_for_excel, fbs_option,
                                                           num_pdf_pages)
                            # Блок для скачивания Excel
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

        except Exception as e:
            st.error(f"Произошла ошибка при обработке файлов: {e}")
            st.exception(e)


if __name__ == "__main__":
    main()
