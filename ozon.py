import streamlit as st
import pandas as pd
import re
import io
from pypdf import PdfReader, PdfWriter
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill


def normalize_old_id(value):
    if value is None or pd.isna(value):
        return None
    value = re.sub(r'\.0$', '', str(value))
    value = re.sub(r'\s+', '', value).strip().lower()
    return value or None


def normalize_label_code(value):
    if value is None or pd.isna(value):
        return None
    value = re.sub(r'\s+', '', str(value).strip().lower())
    if value.startswith('ii'):
        digits = re.sub(r'\D', '', value[2:])
        return 'ii' + digits if digits else None
    return value or None


def robust_extract_id(text):
    if not text:
        return None
    text = text.replace('\n', ' ').replace('\r', ' ')
    match = re.search(r'ii\s*((?:\d\s*){7,16})', text, re.IGNORECASE)
    if match:
        digits = re.sub(r'\D', '', match.group(1))
        if 7 <= len(digits) <= 16:
            return 'ii' + digits
    match = re.search(r'(\d{7,12}\s*-\s*\d{4,5}(?:\s*-\s*\d+)?)', text)
    if match:
        return re.sub(r'\s+', '', match.group(1)).lower()
    return None


def get_sticker_number(match_id):
    if not match_id:
        return ''
    value = str(match_id)
    if value.lower().startswith('ii'):
        digits = re.sub(r'\D', '', value)
        return digits[-4:] if len(digits) >= 4 else digits
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


def apply_priorities(df):
    df = df.copy()
    df['_article'] = df['Артикул'].fillna('').astype(str).str.strip().str.upper()
    df['_sticker'] = df['Стикер'].fillna('').astype(str).str.strip()
    article_counts = df['_article'].value_counts()
    sticker_counts = df.loc[df['_sticker'] != '', '_sticker'].value_counts()
    suffix_rank = {'K8': 0, 'K6': 1, 'K5': 2, 'K4': 3, 'K3': 4, 'K2': 5}

    def classify(row):
        # Повтор номера стикера всегда уходит в самый конец
        sticker = row['_sticker']
        if sticker and sticker_counts.get(sticker, 0) > 1:
            return 5, 0
        qty = pd.to_numeric(row['Кол-во'], errors='coerce')
        if pd.notna(qty) and qty >= 2:
            return 1, 0
        article = row['_article']
        suffix = next((s for s in suffix_rank if article.endswith(s)), None)
        if suffix:
            return 2, suffix_rank[suffix]
        count = article_counts.get(article, 0) if article else 0
        if article and count > 1:
            return 3, -count
        return 4, 0

    priorities = df.apply(classify, axis=1, result_type='expand')
    df['Приоритет'] = priorities[0].astype(int)
    df['_detail'] = priorities[1].astype(int)
    df['_name_sort'] = df['Наименование товара'].fillna('').astype(str).str.casefold()
    df['_article_sort'] = df['_article']
    df['_qty_sort'] = pd.to_numeric(df['Кол-во'], errors='coerce').fillna(0)

    # В приоритете 1 сначала товары с большим количеством; остальные приоритеты — по своим правилам.
    df = df.sort_values(
        by=['Приоритет', '_qty_sort', '_detail', '_name_sort', '_article_sort'],
        ascending=[True, False, True, True, True],
        kind='stable'
    )

    # В приоритете 3 одинаковые артикулы ставим рядом, начиная с наиболее частых.
    mask = df['Приоритет'] == 3
    if mask.any():
        repeated = df.loc[mask].sort_values(
            by=['_detail', '_article_sort', '_name_sort'],
            ascending=[True, True, True],
            kind='stable'
        )
        df.loc[mask, repeated.columns] = repeated

    df.drop(columns=['_article', '_sticker', '_detail', '_name_sort', '_article_sort', '_qty_sort'],
            inplace=True, errors='ignore')
    df['№'] = range(1, len(df) + 1)
    return df


def create_pdf_from_order(pdf_reader, pdf_map, ordered_match_ids):
    writer = PdfWriter()
    added_pages = 0
    current_map = pdf_map.copy()
    for m_id in ordered_match_ids:
        if not m_id:
            continue
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
    now_str = datetime.now().strftime('%d.%m.%Y %H:%M')
    cols = ['№', 'Приоритет', 'Наименование товара', 'Артикул', 'Кол-во', 'Стикер']
    final_df = df[[c for c in cols if c in df.columns]]
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Лист подбора', startrow=5)
        ws = writer.sheets['Лист подбора']
        for row_num, text in [
            (2, f'Лист подбора ({fbs_name})'),
            (3, f'Дата: {now_str}'),
            (4, f'Количество отправлений: {global_total_orders}')
        ]:
            ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num,
                           end_column=max(1, len(final_df.columns)))
            cell = ws.cell(row=row_num, column=1)
            cell.value = text
            cell.font = Font(bold=True, size=11)
            cell.alignment = Alignment(horizontal='left')
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_margins.left = ws.page_margins.right = 0.25
        ws.page_margins.top = ws.page_margins.bottom = 0.25
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.freeze_panes = 'A7'
        if len(final_df.columns):
            ws.auto_filter.ref = f'A6:{get_column_letter(len(final_df.columns))}{len(final_df) + 6}'
        for i, col in enumerate(final_df.columns, 1):
            letter = get_column_letter(i)
            values_len = final_df[col].fillna('').astype(str).map(len).tolist()
            width = max([len(str(col))] + values_len) + 3
            ws.column_dimensions[letter].width = min(width, 80 if col == 'Наименование товара' else 35)
        thin = Side(style='thin')
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in ws.iter_rows(min_row=6, max_row=len(final_df) + 6,
                                min_col=1, max_col=max(1, len(final_df.columns))):
            for cell in row:
                cell.alignment = Alignment(horizontal='left' if cell.column == 3 else 'center',
                                           vertical='center', wrap_text=True)
                if cell.row == 6:
                    cell.border = border
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill('solid', fgColor='D9EAF7')
                if cell.row > 6 and cell.column == 2:
                    cell.font = Font(bold=True)
    return output.getvalue()


def main():
    st.set_page_config(layout='wide', page_title='Ozon Sorter')
    st.title('📦 Ozon: Сортировщик')
    warehouse = st.selectbox('Выберите склад (FBS):', ['Озон', 'Рига', 'Плутон'])
    c1, c2 = st.columns(2)
    with c1:
        csv_file = st.file_uploader('1. Загрузите CSV', type=['csv', 'txt'])
    with c2:
        pdf_file = st.file_uploader('2. Загрузите PDF', type=['pdf'])
    if not (csv_file and pdf_file):
        return
    try:
        raw = csv_file.read()
        df = None
        for enc in ['utf-8-sig', 'utf-8', 'cp1251', 'latin1']:
            try:
                candidate = pd.read_csv(io.BytesIO(raw), sep=None, engine='python', encoding=enc)
                ship_col = get_column_by_variants(candidate, ['Номер отправления', 'Номер заказа'])
                if ship_col:
                    df = candidate.rename(columns={ship_col: 'Номер отправления'})
                    break
            except Exception:
                pass
        if df is None:
            st.error('Не удалось прочитать CSV или не найдена колонка «Номер отправления».')
            return

        qty_col = get_column_by_variants(df, ['Количество', 'Кол-во'])
        art_col = get_column_by_variants(df, ['Артикул', 'Артикул товара'])
        name_col = get_column_by_variants(df, ['Наименование товара', 'Название товара'])
        rename = {}
        if qty_col:
            rename[qty_col] = 'Кол-во'
        if art_col:
            rename[art_col] = 'Артикул'
        if name_col:
            rename[name_col] = 'Наименование товара'
        df = df.rename(columns=rename)
        for col, default in [('Кол-во', 1), ('Артикул', ''), ('Наименование товара', '')]:
            if col not in df.columns:
                df[col] = default
        df['Кол-во'] = pd.to_numeric(df['Кол-во'], errors='coerce').fillna(1)

        df['label_id'] = df['Код этикетки'].apply(normalize_label_code) if 'Код этикетки' in df.columns else None
        df['shipment_id'] = df['Номер отправления'].apply(normalize_old_id)
        df['match_id'] = df['label_id']
        df.loc[df['match_id'].isna() | (df['match_id'].astype(str).str.strip() == ''),
               'match_id'] = df['shipment_id']
        df['Стикер'] = df['match_id'].apply(get_sticker_number)

        sorted_df = apply_priorities(df)
        reader = PdfReader(pdf_file)
        pdf_map = {}
        for i, page in enumerate(reader.pages):
            pid = robust_extract_id(page.extract_text() or '')
            if pid:
                pdf_map[i + 1] = pid
        csv_ids = set(sorted_df['match_id'].dropna().astype(str))
        pdf_ids = set(pdf_map.values())
        st.info(f'Стикеров распознано в PDF: {len(pdf_map)}. Совпало с CSV: {len(csv_ids & pdf_ids)}.')

        total = sorted_df['match_id'].nunique()
        excel = save_styled_excel(sorted_df, total, warehouse)
        st.download_button('📥 Скачать общий Excel со всеми товарами',
                           data=excel, file_name='ozon_all_sorted.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        ordered_ids = sorted_df['match_id'].dropna().drop_duplicates().tolist()
        pdf_bytes, count = create_pdf_from_order(reader, pdf_map, ordered_ids)
        st.download_button(f'📥 Скачать PDF в порядке Excel ({count} страниц)',
                           data=pdf_bytes, file_name='ozon_all_sorted.pdf',
                           mime='application/pdf')
    except Exception as exc:
        st.error(f'Ошибка обработки файлов: {exc}')


if __name__ == '__main__':
    main()
