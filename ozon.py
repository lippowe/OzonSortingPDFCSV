import streamlit as st
import pandas as pd
import re
import io
from pypdf import PdfReader, PdfWriter
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font


def normalize_old_id(value):
    """Нормализация старого номера отправления."""
    if value is None or pd.isna(value):
        return None

    value = str(value)
    value = re.sub(r'\.0$', '', value)
    value = re.sub(r'\s+', '', value)
    return value.strip().lower() or None


def normalize_label_code(value):
    """
    Нормализация нового кода этикетки Ozon:
    ii5009821 1493 -> ii50098211493
    ii50098211493 -> ii50098211493
    """
    if value is None or pd.isna(value):
        return None

    value = str(value).strip().lower()
    value = re.sub(r'\s+', '', value)

    if value.startswith('ii'):
        digits = re.sub(r'\D', '', value[2:])
        if digits:
            return 'ii' + digits

    return value or None


def robust_extract_id(text):
    """
    Извлекает из PDF оба формата стикеров.

    Новый:
        ii5009821 1493
        -> ii50098211493

    Старый:
        0208627373-0136-1
        или если PDF разбил строку:
        0208627373
        -0136-1
    """
    if not text:
        return None

    cleaned_text = text.replace('\n', ' ').replace('\r', ' ')

    # ------------------------------------------------------------
    # НОВЫЙ ФОРМАТ
    # ii5009821 1493
    # ii50098211493
    # ------------------------------------------------------------
    match = re.search(
        r'ii\s*((?:\d\s*){7,16})',
        cleaned_text,
        re.IGNORECASE
    )

    if match:
        digits = re.sub(r'\D', '', match.group(1))
        if 7 <= len(digits) <= 16:
            return 'ii' + digits

    # ------------------------------------------------------------
    # СТАРЫЙ ФОРМАТ
    # 0208627373-0136-1
    # Допускаем 4-5 цифр после первого дефиса.
    # ------------------------------------------------------------
    match = re.search(
        r'(\d{7,12}\s*-\s*\d{4,5}(?:\s*-\s*\d+)?)',
        cleaned_text
    )

    if match:
        return re.sub(r'\s+', '', match.group(1)).lower()

    # Резервный поиск старого номера без дефисов
    all_numbers = re.findall(r'\d{7,12}', cleaned_text)
    if all_numbers:
        return all_numbers[-1].lower()

    return None


def get_sticker_number(match_id):
    """
    Отображаемый номер стикера.

    Новый формат:
        ii50098211493 -> 1493

    Старый формат:
        0208627373-0136-1 -> 0136
    """
    if not match_id:
        return ''

    value = str(match_id)

    # Новый формат
    if value.lower().startswith('ii'):
        digits = re.sub(r'\D', '', value)
        return digits[-4:] if len(digits) >= 4 else digits

    # Старый формат
    match = re.search(r'-(\d{4,5})(?:-|$)', value)
    if match:
        return match.group(1)

    digits = re.sub(r'\D', '', value)
    return digits[-4:] if len(digits) >= 4 else digits


def get_column_by_variants(df, variants):
    for variant in variants:
        if variant in df.columns:
            return variant
    return None


def create_pdf_from_order(pdf_reader, pdf_map, ordered_match_ids):
    writer = PdfWriter()
    added_pages = 0
    current_map = pdf_map.copy()

    for m_id in ordered_match_ids:
        if not m_id:
            continue

        # Один и тот же заказ/этикетка может встречаться только один раз.
        for pg_num, pdf_id in list(current_map.items()):
            if m_id == pdf_id:
                writer.add_page(pdf_reader.pages[pg_num - 1])
                del current_map[pg_num]
                added_pages += 1
                break

    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue(), added_pages


def save_styled_excel(df, global_total_orders, fbs_name):
    output = io.BytesIO()
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    display_cols = [
        '№',
        'Наименование товара',
        'Артикул',
        'Кол-во',
        'Стикер'
    ]

    final_df = df[
        [c for c in display_cols if c in df.columns]
    ]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(
            writer,
            index=False,
            sheet_name='Лист подбора',
            startrow=5
        )

        worksheet = writer.sheets['Лист подбора']

        header_data = [
            (2, f"Лист подбора ({fbs_name})"),
            (3, f"Дата: {now_str}"),
            (4, f"Количество отправлений: {global_total_orders}")
        ]

        for row_num, text in header_data:
            worksheet.merge_cells(
                start_row=row_num,
                start_column=1,
                end_row=row_num,
                end_column=max(len(final_df.columns), 1)
            )

            cell = worksheet.cell(row=row_num, column=1)
            cell.value = text
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal='left')

        worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
        worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4

        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.25
        worksheet.page_margins.bottom = 0.25

        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0

        sticker_col_idx = -1
        name_col_idx = -1

        for idx, col in enumerate(final_df.columns, 1):
            if col == 'Стикер':
                sticker_col_idx = idx
            if col == 'Наименование товара':
                name_col_idx = idx

        for i, column in enumerate(final_df.columns, 1):
            col_letter = get_column_letter(i)

            if len(final_df) > 0:
                data_max_len = (
                    final_df[column]
                    .astype(str)
                    .map(len)
                    .max()
                )
            else:
                data_max_len = 0

            col_width = max(data_max_len, len(column)) + 3

            if i == name_col_idx:
                worksheet.column_dimensions[col_letter].width = min(
                    col_width, 80
                )
            else:
                worksheet.column_dimensions[col_letter].width = min(
                    col_width, 40
                )

        thin = Side(style='thin')

        header_border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        bold_font = Font(bold=True)

        align_center = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=True
        )

        align_left = Alignment(
            horizontal='left',
            vertical='center',
            wrap_text=True
        )

        if len(final_df.columns) > 0:
            for row in worksheet.iter_rows(
                min_row=6,
                max_row=len(final_df) + 6,
                min_col=1,
                max_col=len(final_df.columns)
            ):
                for cell in row:
                    if cell.column == name_col_idx:
                        cell.alignment = align_left
                    else:
                        cell.alignment = align_center

                    if cell.row == 6:
                        cell.border = header_border
                        cell.font = bold_font

                    if (
                        cell.column == sticker_col_idx
                        and cell.row > 6
                    ):
                        cell.font = bold_font

    return output.getvalue()


def main():
    st.set_page_config(
        layout="wide",
        page_title="Ozon Sorter Final"
    )

    st.title("📦 Ozon: Сортировщик")

    fbs_choice = st.selectbox(
        "Выберите склад (FBS):",
        ["Озон", "Рига", "Плутон"]
    )

    col_u1, col_u2 = st.columns(2)

    with col_u1:
        uploaded_csv = st.file_uploader(
            "1. Загрузите CSV",
            type=["csv", "txt"]
        )

    with col_u2:
        uploaded_pdf = st.file_uploader(
            "2. Загрузите PDF",
            type="pdf"
        )

    if uploaded_csv and uploaded_pdf:
        try:
            bytes_data = uploaded_csv.read()
            df = None

            # Читаем CSV
            for enc in ['utf-8', 'cp1251', 'latin1']:
                try:
                    test_df = pd.read_csv(
                        io.BytesIO(bytes_data),
                        sep=None,
                        engine='python',
                        encoding=enc
                    )

                    ship_col = get_column_by_variants(
                        test_df,
                        ['Номер отправления', 'Номер заказа']
                    )

                    if ship_col:
                        df = test_df.rename(
                            columns={
                                ship_col: 'Номер отправления'
                            }
                        )
                        break

                except Exception:
                    continue

            if df is None:
                st.error("❌ Ошибка чтения CSV")
                return

            # Наименование товара
            qty_col = (
                get_column_by_variants(
                    df,
                    ['Количество', 'Кол-во']
                ) or 'Кол-во'
            )

            art_col = (
                get_column_by_variants(
                    df,
                    ['Артикул', 'Артикул товара']
                ) or 'Артикул'
            )

            name_col = (
                get_column_by_variants(
                    df,
                    ['Наименование товара', 'Название товара']
                ) or 'Наименование товара'
            )

            rename_map = {
                qty_col: 'Кол-во',
                art_col: 'Артикул',
                name_col: 'Наименование товара'
            }

            df = df.rename(columns=rename_map)

            df['Кол-во'] = pd.to_numeric(
                df['Кол-во'],
                errors='coerce'
            ).fillna(1)

            # --------------------------------------------------------
            # ГЛАВНОЕ ИЗМЕНЕНИЕ
            #
            # В этом CSV есть два источника номера:
            #
            # 1. Старые стикеры:
            #    Номер отправления = 0208627373-0136-1
            #
            # 2. Новые стикеры:
            #    Код этикетки = ii50098211493
            #
            # Если "Код этикетки" заполнен, используем его.
            # Если пустой — используем старый "Номер отправления".
            # --------------------------------------------------------

            has_label_code = 'Код этикетки' in df.columns

            if has_label_code:
                df['label_id'] = df['Код этикетки'].apply(
                    normalize_label_code
                )
            else:
                df['label_id'] = None

            df['shipment_id'] = df[
                'Номер отправления'
            ].apply(normalize_old_id)

            # Универсальный ID для сопоставления с PDF
            df['match_id'] = df['label_id']

            df.loc[
                df['match_id'].isna()
                | (df['match_id'].astype(str).str.strip() == ''),
                'match_id'
            ] = df['shipment_id']

            # Номер стикера для Excel
            df['Стикер'] = df['match_id'].apply(
                get_sticker_number
            )

            global_total = df['match_id'].nunique()

            # --------------------------------------------------------
            # Группировка по артикулу
            # --------------------------------------------------------
            art_counts = df['Артикул'].value_counts()

            df_repeats_raw = df[
                df['Артикул'].isin(
                    art_counts[art_counts > 1].index
                )
            ].copy()

            df_main_raw = df[
                df['Артикул'].isin(
                    art_counts[art_counts == 1].index
                )
            ].copy()

            def get_prio(row):
                if row['Кол-во'] >= 2:
                    return 1

                if any(
                    s in str(row['Артикул'])
                    for s in ['K2', 'K3', 'K4', 'K5']
                ):
                    return 2

                return 3

            df_main_raw['priority'] = df_main_raw.apply(
                get_prio,
                axis=1
            )

            df_main = df_main_raw.sort_values(
                by=['priority', 'Наименование товара'],
                ascending=[True, True]
            )

            df_main['№'] = range(
                1,
                len(df_main) + 1
            )

            df_repeats_raw['counts'] = df_repeats_raw[
                'Артикул'
            ].map(art_counts)

            df_repeats = df_repeats_raw.sort_values(
                by=['counts', 'Наименование товара'],
                ascending=[False, True]
            )

            df_repeats['№'] = range(
                1,
                len(df_repeats) + 1
            )

            # --------------------------------------------------------
            # PDF
            # --------------------------------------------------------
            pdf_reader = PdfReader(uploaded_pdf)

            pdf_map = {}

            for i, page in enumerate(pdf_reader.pages):
                text = page.extract_text() or ''
                pdf_id = robust_extract_id(text)

                if pdf_id:
                    pdf_map[i + 1] = pdf_id

            # --------------------------------------------------------
            # Диагностика сопоставления
            # --------------------------------------------------------
            csv_match_ids = set(
                df['match_id'].dropna().astype(str)
            )

            pdf_match_ids = set(pdf_map.values())

            matched_count = len(
                csv_match_ids.intersection(pdf_match_ids)
            )

            st.info(
                f"🔎 Найдено стикеров в PDF: {len(pdf_map)} | "
                f"совпало с CSV: {matched_count}"
            )

            st.divider()

            # --------------------------------------------------------
            # Группа №1
            # --------------------------------------------------------
            st.subheader("📂 Группа №1: Основная")

            c1_1, c1_2 = st.columns(2)

            p_main, cnt_m = create_pdf_from_order(
                pdf_reader,
                pdf_map,
                df_main['match_id'].dropna().unique()
            )

            xlsx_main = save_styled_excel(
                df_main,
                global_total,
                fbs_choice
            )

            c1_1.download_button(
                f"📥 PDF Группа №1 ({cnt_m} ст.)",
                p_main,
                "1_main.pdf"
            )

            c1_2.download_button(
                "📥 Excel Группа №1",
                xlsx_main,
                "1_main.xlsx"
            )

            # --------------------------------------------------------
            # Группа №2
            # --------------------------------------------------------
            if not df_repeats.empty:
                st.divider()
                st.subheader("📂 Группа №2: Повторы")

                c2_1, c2_2 = st.columns(2)

                p_rep, cnt_r = create_pdf_from_order(
                    pdf_reader,
                    pdf_map,
                    df_repeats['match_id'].dropna().unique()
                )

                xlsx_rep = save_styled_excel(
                    df_repeats,
                    global_total,
                    fbs_choice
                )

                c2_1.download_button(
                    f"📥 PDF Группа №2 ({cnt_r} ст.)",
                    p_rep,
                    "2_repeats.pdf"
                )

                c2_2.download_button(
                    "📥 Excel Группа №2",
                    xlsx_rep,
                    "2_repeats.xlsx"
                )

        except Exception as e:
            st.error(f"Ошибка: {e}")


if __name__ == "__main__":
    main()
