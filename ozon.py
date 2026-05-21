import streamlit as st
import pandas as pd
import re
import io
from pypdf import PdfReader, PdfWriter
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font


def robust_extract_id(text):
    if not text: return None
    cleaned_text = text.replace('\n', ' ').replace('\r', ' ')
    match = re.search(r'(\d{8,12})\s*-', cleaned_text)
    if match: return match.group(1)
    all_numbers = re.findall(r'\d{7,12}', cleaned_text)
    return all_numbers[-1] if all_numbers else None


def get_column_by_variants(df, variants):
    for variant in variants:
        if variant in df.columns: return variant
    return None


def create_pdf_from_order(pdf_reader, pdf_map, ordered_match_ids):
    writer = PdfWriter()
    added_pages = 0
    current_map = pdf_map.copy()
    for m_id in ordered_match_ids:
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
    """Excel с выборочным выравниванием и границами только в шапке."""
    output = io.BytesIO()
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    display_cols = ['№', 'Наименование товара', 'Артикул', 'Кол-во', 'Стикер']
    final_df = df[[c for c in display_cols if c in df.columns]]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Лист подбора', startrow=5)

        workbook = writer.book
        worksheet = writer.sheets['Лист подбора']

        # 1. Заголовки (Строки 2, 3, 4)
        header_data = [
            (2, f"Лист подбора ({fbs_name})"),
            (3, f"Дата: {now_str}"),
            (4, f"Количество отправлений: {global_total_orders}")
        ]

        for row_num, text in header_data:
            worksheet.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=len(final_df.columns))
            cell = worksheet.cell(row=row_num, column=1)
            cell.value = text
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal='left')

        # 2. Настройки страницы
        worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
        worksheet.page_setup.paperSize = worksheet.PAPERSIZE_A4
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.25
        worksheet.page_margins.bottom = 0.25

        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0

        # 3. Ширина колонок и поиск индексов
        sticker_col_idx = -1
        name_col_idx = -1
        for idx, col in enumerate(final_df.columns, 1):
            if col == 'Стикер': sticker_col_idx = idx
            if col == 'Наименование товара': name_col_idx = idx

        for i, column in enumerate(final_df.columns, 1):
            col_letter = get_column_letter(i)
            data_max_len = final_df[column].astype(str).map(len).max()
            col_width = max(data_max_len, len(column)) + 3
            if i == name_col_idx:
                worksheet.column_dimensions[col_letter].width = min(col_width, 80)
            else:
                worksheet.column_dimensions[col_letter].width = min(col_width, 40)

        # 4. Оформление, Границы и Выравнивание
        thin = Side(style='thin')
        header_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        bold_font = Font(bold=True)

        # Стили выравнивания
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for row in worksheet.iter_rows(min_row=6, max_row=len(final_df) + 6, min_col=1, max_col=len(final_df.columns)):
            for cell in row:
                # ВЫРАВНИВАНИЕ: если это колонка Наименование, то влево, иначе по центру
                if cell.column == name_col_idx:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center

                # ШАПКА (Строка 6)
                if cell.row == 6:
                    cell.border = header_border
                    cell.font = bold_font

                # СТИКЕР (Жирный шрифт в данных)
                if cell.column == sticker_col_idx and cell.row > 6:
                    cell.font = bold_font

    return output.getvalue()


def main():
    st.set_page_config(layout="wide", page_title="Ozon Sorter Final")
    st.title("📦 Ozon: Сортировщик")

    fbs_choice = st.selectbox("Выберите склад (FBS):", ["Озон", "Рига", "Плутон"])

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        uploaded_csv = st.file_uploader("1. Загрузите CSV", type=["csv", "txt"])
    with col_u2:
        uploaded_pdf = st.file_uploader("2. Загрузите PDF", type="pdf")

    if uploaded_csv and uploaded_pdf:
        try:
            bytes_data = uploaded_csv.read()
            df = None
            for enc in ['utf-8', 'cp1251', 'latin1']:
                try:
                    df = pd.read_csv(io.BytesIO(bytes_data), sep=None, engine='python', encoding=enc)
                    ship_col = get_column_by_variants(df, ['Номер отправления', 'Номер заказа'])
                    if ship_col:
                        df = df.rename(columns={ship_col: 'Номер отправления'})
                        break
                except:
                    continue

            if df is None:
                st.error("❌ Ошибка чтения CSV");
                return

            qty_col = get_column_by_variants(df, ['Количество', 'Кол-во']) or 'Кол-во'
            art_col = get_column_by_variants(df, ['Артикул', 'Артикул товара']) or 'Артикул'
            name_col = get_column_by_variants(df, ['Наименование товара', 'Название товара']) or 'Наименование товара'

            df = df.rename(columns={qty_col: 'Кол-во', art_col: 'Артикул', name_col: 'Наименование товара'})
            df['Кол-во'] = pd.to_numeric(df['Кол-во'], errors='coerce').fillna(1)
            df['match_id'] = df['Номер отправления'].apply(lambda x: re.sub(r'-.*', '', str(x)).strip().lstrip('0'))
            df['Стикер'] = df['match_id'].apply(lambda x: x[-4:])

            global_total = df['match_id'].nunique()

            art_counts = df['Артикул'].value_counts()
            df_repeats_raw = df[df['Артикул'].isin(art_counts[art_counts > 1].index)].copy()
            df_main_raw = df[df['Артикул'].isin(art_counts[art_counts == 1].index)].copy()

            def get_prio(row):
                if row['Кол-во'] >= 2: return 1
                if any(s in str(row['Артикул']) for s in ['K2', 'K3', 'K4', 'K5']): return 2
                return 3

            df_main_raw['priority'] = df_main_raw.apply(get_prio, axis=1)
            df_main = df_main_raw.sort_values(by=['priority', 'Наименование товара'], ascending=[True, True])
            df_main['№'] = range(1, len(df_main) + 1)

            df_repeats_raw['counts'] = df_repeats_raw['Артикул'].map(art_counts)
            df_repeats = df_repeats_raw.sort_values(by=['counts', 'Наименование товара'], ascending=[False, True])
            df_repeats['№'] = range(1, len(df_repeats) + 1)

            pdf_reader = PdfReader(uploaded_pdf)
            pdf_map = {i + 1: robust_extract_id(p.extract_text()).lstrip('0')
                       for i, p in enumerate(pdf_reader.pages) if robust_extract_id(p.extract_text())}

            st.divider()

            # Вывод файлов
            st.subheader(f"📂 Группа №1: Основная")
            c1_1, c1_2 = st.columns(2)
            p_main, cnt_m = create_pdf_from_order(pdf_reader, pdf_map, df_main['match_id'].unique())
            xlsx_main = save_styled_excel(df_main, global_total, fbs_choice)
            c1_1.download_button(f"📥 PDF Группа №1 ({cnt_m} ст.)", p_main, "1_main.pdf")
            c1_2.download_button("📥 Excel Группа №1", xlsx_main, "1_main.xlsx")

            if not df_repeats.empty:
                st.divider()
                st.subheader(f"📂 Группа №2: Повторы")
                c2_1, c2_2 = st.columns(2)
                p_rep, cnt_r = create_pdf_from_order(pdf_reader, pdf_map, df_repeats['match_id'].unique())
                xlsx_rep = save_styled_excel(df_repeats, global_total, fbs_choice)
                c2_1.download_button(f"📥 PDF Группа №2 ({cnt_r} ст.)", p_rep, "2_repeats.pdf")
                c2_2.download_button("📥 Excel Группа №2", xlsx_rep, "2_repeats.xlsx")

        except Exception as e:
            st.error(f"Ошибка: {e}")


if __name__ == "__main__":
    main()
