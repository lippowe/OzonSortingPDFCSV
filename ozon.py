

# --- Ваши существующие функции ---

def extract_order_number_prefix(order_string):
    """
    Извлекает часть номера заказа до первого тире.
    Например, из "12345-ABCDE" вернет "12345".
    """
    if not isinstance(order_string, str):
        order_string = str(order_string) # Убедимся, что это строка

    match = re.search(r'^(\d+)-', order_string) # Ищем цифры в начале строки до первого тире
    if match:
        return match.group(1) # Возвращаем найденные цифры
    else:
        return None # Если паттерн не найден (например, нет тире или не начинается с цифр)

def extract_sticker_from_order(order_number):
    """
    Извлекает последние 4 цифры перед первым дефисом из номера заказа.
    Предполагается, что это "стикер" для CSV.
    """
    if not isinstance(order_number, str):
        order_number = str(order_number)

    # Ищем паттерн: 4 цифры, за которыми следует дефис
    match = re.search(r'(\d{4})-', order_number)
    if match:
        return match.group(1) # Возвращаем найденные 4 цифры
    else:
        return None # Если паттерн не найден, возвращаем None

def sort_dataframe(df):
    """
    Сортирует DataFrame по 'Количество' (убывание), затем по
    приоритету, извлеченному из 'Артикул' (убывание),
    и далее по 'Наименование товара' и 'Артикул'.
    """
    # 1. Сортировка по 'Количество'
    if 'Количество' in df.columns:
        # Преобразуем в число, ошибки заменяем на 0 (или можно NaN)
        df['Количество'] = pd.to_numeric(df['Количество'], errors='coerce').fillna(0)
        # Сортируем по количеству по убыванию
        df = df.sort_values(by='Количество', ascending=False)
    else:
        st.warning("Колонка 'Количество' не найдена. Сортировка по этому полю будет пропущена.")

    # 2. Определение приоритета из 'Артикул'
    def get_priority(row):
        """Извлекает число из артикула после 'k'."""
        article = str(row.get('Артикул', '')).lower() # Берем артикул, приводим к строке и нижнему регистру
        match = re.search(r'k(\d+)', article) # Ищем 'k' с последующими цифрами
        if match:
            return int(match.group(1)) # Возвращаем число после 'k'
        else:
            return 0 # Если 'k' не найдено, присваиваем минимальный приоритет

    df['Приоритет_Сортировки'] = df.apply(get_priority, axis=1) # Создаем временную колонку для приоритета

    # 3. Сортировка по приоритету, Наименованию товара и Артикулу
    # Убедимся, что колонки существуют и являются строками
    df['Наименование товара'] = df.get('Наименование товара', pd.Series(dtype='str')).astype(str)
    df['Артикул'] = df.get('Артикул', pd.Series(dtype='str')).astype(str)

    # Сортируем:
    # - Приоритет (убывание)
    # - Наименование товара (возрастание)
    # - Артикул (возрастание)
    df = df.sort_values(by=['Приоритет_Сортировки', 'Наименование товара', 'Артикул'], ascending=[False, True, True])

    # Удаляем временную колонку приоритета
    df = df.drop('Приоритет_Сортировки', axis=1)

    return df

# --- Функции для работы с PDF ---

def extract_sticker_data_from_pdf(pdf_file):
    """
    Извлекает номера страниц и соответствующие им "стикеры" из PDF.
    Предполагается, что "стикер" — это число, следующее за "FBS: 204514".
    Возвращает словарь: {номер_страницы: стикер}
    """
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
                    # st.write(f"Найдено на странице {page_num + 1}: стикер {sticker_number}") # Для отладки
                else:
                    # Если паттерн не найден, можно попробовать извлечь номер стикера по другому правилу,
                    # или просто проигнорировать эту страницу, если она не содержит нужной информации.
                    pass # Пока пропускаем страницы без нужного паттерна
            else:
                # Можно добавить вывод предупреждения, если текст не удалось извлечь
                # st.warning(f"Не удалось извлечь текст со страницы {page_num + 1}.")
                pass
    except Exception as e:
        st.error(f"Ошибка при обработке PDF файла: {e}")
    return sticker_data

def reorder_pdf_pages(pdf_file, page_order_mapping):
    """
    Переупорядочивает страницы PDF согласно заданному списку.
    :param pdf_file: Объект загруженного PDF файла.
    :param page_order_mapping: Список кортежей (исходный_номер_страницы, стикер_из_PDF).
    :return: Объект PdfWriter с переупорядоченными страницами, или None при ошибке.
    """
    try:
        reader = PdfReader(pdf_file)
        writer = PdfWriter() # Используем PdfWriter из pypdf

        # Создаем словарь для быстрого поиска страниц по их исходному номеру
        # Нумерация в reader.pages начинается с 0
        pages_dict = {i + 1: page for i, page in enumerate(reader.pages)}

        # Проверяем, что все нужные страницы существуют
        for original_page_num, _ in page_order_mapping:
            if original_page_num not in pages_dict:
                st.error(f"Страница {original_page_num} из PDF не найдена. Проверьте соответствие стикеров.")
                return None

        # Добавляем страницы в новом порядке
        for original_page_num, _ in page_order_mapping:
            # Получаем страницу по ее исходному номеру (1-based)
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
    uploaded_csv_file = st.file_uploader("Загрузите CSV файл с заказами", type=["csv", "txt"]) # Добавим .txt для случая, если файл сохранили как .txt
    uploaded_pdf_file = st.file_uploader("Загрузите PDF файл со стикерами", type="pdf")

    if uploaded_csv_file and uploaded_pdf_file:
        st.success("Файлы успешно загружены!")

        st.header("2. Обработка CSV")
        try:
            # Пытаемся прочитать CSV, учитывая разные разделители и кодировки
            try:
                df_original = pd.read_csv(uploaded_csv_file, sep=';')
            except Exception:
                try:
                    df_original = pd.read_csv(uploaded_csv_file, sep=',')
                except Exception:
                    try:
                        df_original = pd.read_csv(uploaded_csv_file, sep='\t')
                    except Exception:
                        uploaded_csv_file.seek(0) # Возвращаем указатель в начало файла
                        df_original = pd.read_csv(io.StringIO(uploaded_csv_file.read().decode('cp1251'))) # Попытка с cp1251

            st.write("Исходные данные CSV:")
            st.dataframe(df_original)
            # Добавляем колонку "Номер_заказа_префикс" для сопоставления
            # Вместо extract_sticker_from_order используем новую функцию
            df_original['Стикер'] = df_original['Номер заказа'].apply(extract_order_number_prefix)

            # Удаляем строки, где префикс номера заказа не был найден
            df_with_order_prefix = df_original.dropna(subset=['Стикер']).copy()

            if df_with_order_prefix.empty:
                st.warning(
                    "Не найдено ни одного номера заказа в формате 'число-' в колонке 'Номер заказа' CSV файла. Проверьте формат номеров заказов.")
            else:
                st.write("Данные CSV с извлеченными префиксами номеров заказов:")
                st.dataframe(df_with_order_prefix)

                # Сортируем DataFrame (если сортировка по-прежнему нужна по старому принципу)
                # Если сортировка теперь должна быть по этому новому префиксу, логику sort_dataframe нужно будет скорректировать
                df_sorted = sort_dataframe(df_with_order_prefix)
                st.write("Отсортированные данные CSV:")
                st.dataframe(df_sorted)

                # --- Сопоставление и переупорядочивание PDF ---
                st.header("3. Обработка PDF и сопоставление")

                # Извлекаем данные о стикерах из PDF (эта функция остается той же, она ищет "FBS: 204514 XXXXX")
                pdf_sticker_data = extract_sticker_data_from_pdf(uploaded_pdf_file)

                if not pdf_sticker_data:
                    st.warning(
                        "Не удалось извлечь ни одного стикера из PDF файла. Проверьте, соответствует ли формат стикера шаблону 'FBS: 204514 XXXXX'.")
                else:
                    st.write("Извлеченные стикеры из PDF (страница: стикер):")
                    st.write(pdf_sticker_data)

                    # Создаем список для порядка страниц PDF
                    # Он будет содержать кортежи: (исходный_номер_страницы_PDF, стикер_из_PDF)
                    pdf_pages_in_csv_order = []
                    missing_pdf_pages = []

                    # Итерируемся по строкам отсортированного CSV
                    for index, row in df_sorted.iterrows():
                        # !!! ИЗМЕНЕНИЕ ЗДЕСЬ !!!
                        # Теперь мы используем извлеченный префикс номера заказа для поиска
                        csv_identifier = row['Стикер']  # Используем новую колонку

                        # Ищем этот идентификатор (префикс заказа) в данных из PDF
                        # !!!!! ВАЖНО !!!!!
                        # Ваша текущая функция extract_sticker_data_from_pdf ИЗВЛЕКАЕТ НОМЕР СТИКЕРА (XXXXX).
                        # Если вы хотите сопоставить ПРЕФИКС ЗАКАЗА (12345) с этим НОМЕРОМ СТИКЕРА (XXXXX),
                        # то вам нужно либо:
                        #   а) Изменить extract_sticker_data_from_pdf, чтобы она извлекала то, что нужно сопоставлять (например, если стикер в PDF тоже состоит из цифр и как-то связано с номером заказа).
                        #   б) Или, что более вероятно, вам нужно, чтобы данные в PDF тоже содержали номер заказа, а не стикер, и вы извлекали его.
                        #
                        # Предполагая, что "стикер" из PDF (XXXXX) как-то связан с "префиксом заказа" (12345)
                        # и ВЫ ХОТИТЕ СОПОСТАВЛЯТЬ ИХ НАПРЯМУЮ:

                        found_page = None
                        # Ищем ИДЕНТИФИКАТОР (префикс заказа) в ИЗВЛЕЧЕННЫХ СТИКЕРАХ PDF
                        for page_num, pdf_sticker_value in pdf_sticker_data.items():
                            # !!!!! ПРОВЕРЬТЕ ЭТО СОПОСТАВЛЕНИЕ !!!!!
                            # Сейчас сравнивается ПРЕФИКС ЗАКАЗА (csv_identifier) со СТИКЕРОМ ИЗ PDF (pdf_sticker_value).
                            # Это сработает ТОЛЬКО ЕСЛИ эти значения ДОЛЖНЫ быть РАВНЫ.
                            if pdf_sticker_value == csv_identifier:  # <-- Вот тут происходит сопоставление!
                                found_page = (page_num, pdf_sticker_value)
                                # Удаляем найденный стикер из словаря, чтобы не использовать его повторно
                                del pdf_sticker_data[page_num]
                                break  # Нашли соответствие, переходим к следующей строке CSV

                        if found_page:
                            pdf_pages_in_csv_order.append(found_page)
                        else:
                            # Если не нашли соответствия, добавляем в список отсутствующих
                            missing_pdf_pages.append(
                                csv_identifier)  # Теперь добавляем сам идентификатор (префикс заказа)

                    if missing_pdf_pages:
                        st.warning(
                            f"Следующие идентификаторы (префиксы заказов) из отсортированного CSV не были найдены в PDF: {', '.join(missing_pdf_pages)}. Страницы с соответствующими стикерами не будут включены в новый PDF.")

                    # Проверяем, остались ли стикеры в PDF, которые не были найдены в CSV
                    if pdf_sticker_data:
                        st.info(
                            f"В PDF файле остались стикеры, которые не были найдены в CSV: {', '.join(pdf_sticker_data.values())}. Эти страницы не будут использованы.")

                    if not pdf_pages_in_csv_order:
                        st.error(
                            "Не удалось найти соответствие между идентификаторами из CSV и стикерами из PDF. Переупорядочивание PDF невозможно.")
                    else:
                        st.write("Порядок страниц PDF для нового файла (исходная_страница_PDF, стикер_из_PDF):")
                        st.write(pdf_pages_in_csv_order)

                        # Переупорядочиваем страницы PDF
                        reordered_pdf_writer = reorder_pdf_pages(uploaded_pdf_file, pdf_pages_in_csv_order)

                        if reordered_pdf_writer:
                            st.success("Страницы PDF успешно переупорядочены!")

                            # --- НОВЫЙ БЛОК: Подготовка и скачивание отсортированного CSV ---
                            st.header("4. Результат")

                            # 1. Выбираем нужные колонки из отсортированного DataFrame
                            columns_to_display_base = ['Номер отправления', 'Наименование товара', 'Артикул',
                                                       'Стикер']

                            # Создаем словарь с данными для вывода, проверяя наличие колонок
                            display_data_values = {}
                            for col in columns_to_display_base:
                                if col in df_sorted.columns:
                                    display_data_values[col] = df_sorted[col]
                                elif col == 'Номер отправления':  # Если 'Номер отправления' отсутствует
                                    st.warning(
                                        "Колонка 'Номер отправления' не найдена. Для отображения будет использоваться 'Артикул'.")
                                    display_data_values['Номер отправления'] = df_sorted.get('Артикул',
                                                                                             pd.Series(dtype='str'))
                                else:
                                    st.warning(f"Колонка '{col}' не найдена в данных CSV.")
                                    display_data_values[col] = pd.Series(dtype='str')

                            # Добавляем колонку 'Код' с порядковыми номерами
                            num_rows = len(df_sorted)
                            display_data_values['Код'] = pd.Series(range(1, num_rows + 1), index=df_sorted.index)

                            # Создаем DataFrame для отображения и скачивания
                            df_display = pd.DataFrame(display_data_values)

                            # !!! НОВОЕ: Переупорядочиваем колонки, чтобы 'Код' была первой !!!
                            # Составляем новый список колонок в желаемом порядке
                            # Начинаем с 'Код', затем добавляем остальные колонки, которые есть в df_display
                            desired_column_order = ['Код']
                            for col in df_display.columns:
                                if col != 'Код':  # Добавляем остальные колонки, кроме 'Код', чтобы избежать дублирования
                                    desired_column_order.append(col)

                            # Применяем новый порядок к DataFrame
                            df_display = df_display[desired_column_order]

                            st.write("Отсортированные данные (выбранные колонки):")
                            st.dataframe(df_display)

                            # --- БЛОК для скачивания Excel ---
                            st.header("5. Excel экспорт")

                            # Функция для извлечения последних 4 цифр
                            def get_last_4_digits(value):
                                if pd.isna(value):
                                    return ""

                                # Преобразуем в строку, если это число
                                value_str = str(value)

                                # Ищем последовательность из 4 цифр в конце строки
                                match = re.search(r'(\d{4})$', value_str)

                                if match:
                                    return match.group(0)  # Возвращаем найденные 4 цифры
                                else:
                                    # Если не нашли 4 цифры в конце, попробуем найти любые 4 цифры
                                    # или просто вернуть пустую строку, если ничего не подходит
                                    # Можно выбрать:
                                    # 1. Вернуть пустую строку, если нет 4 цифр в конце
                                    # 2. Вернуть последние 4 символа, если они цифры (более гибко)

                                    # Вариант 2: Найти любые 4 последние цифры, если они есть
                                    digits_only = "".join(filter(str.isdigit, value_str))
                                    if len(digits_only) >= 4:
                                        return digits_only[-4:]
                                    else:
                                        return ""  # Или вернуть value_str, если хотим видеть полный текст без 4 последних цифр
                                        # Но по вашему запросу "удалим остальное", поэтому пустая строка лучше

                            # Применяем функцию к колонке 'Стикер' в df_display
                            # Создаем копию, чтобы не изменять исходный DataFrame, если он еще нужен
                            df_for_excel = df_display.copy()
                            df_for_excel['Стикер'] = df_for_excel['Стикер'].apply(get_last_4_digits)

                            # Создаем буфер для сохранения Excel файла в памяти
                            excel_output_buffer = io.BytesIO()

                            # Сохраняем DataFrame в Excel
                            df_for_excel.to_excel(excel_output_buffer, index=False,
                                                  sheet_name='Последние 4 цифры стикера')

                            excel_output_buffer.seek(0)  # Перемещаем указатель в начало буфера

                            # Добавляем кнопку для скачивания Excel файла
                            st.download_button(
                                label="Скачать Excel (только последние 4 цифры стикера)",
                                data=excel_output_buffer,
                                file_name = f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            # --- КОНЕЦ БЛОКА Excel ---

                            st.header("- CSV файл -")
                            # 2. Создаем CSV для скачивания
                            csv_output_buffer = io.StringIO()
                            df_display.to_csv(csv_output_buffer, index=False, sep=';', encoding='utf-8-sig')
                            csv_output_buffer.seek(0)

                            # 3. Добавляем кнопку для скачивания CSV
                            st.download_button(
                                label="Скачать отсортированный CSV",
                                data=csv_output_buffer.getvalue(),
                                file_name = f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.csv",
                                mime="text/csv"
                            )
                            # --- КОНЕЦ НОВОГО БЛОКА ---

                            # ... (код для скачивания PDF остался здесь) ...
                            pdf_output_buffer = io.BytesIO()
                            reordered_pdf_writer.write(pdf_output_buffer)
                            pdf_output_buffer.seek(0)
                            st.header("- PDF файл -")
                            st.write("Ваш новый PDF файл с переупорядоченными страницами:")
                            st.download_button(
                                label="Скачать переупорядоченный PDF",
                                data=pdf_output_buffer,
                                file_name = f"Repeats_Ozon-{datetime.now().strftime('%H-%M-%S')}.pdf",
                                mime="application/pdf"
                            )
        except Exception as e:
                    st.error(f"Произошла ошибка при обработке файлов: {e}")
                    st.exception(e)




if __name__ == "__main__":
 main()

