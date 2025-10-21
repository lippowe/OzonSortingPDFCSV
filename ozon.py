import streamlit as st
import pandas as pd
import re
import csv

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
    # Сначала сортируем по убыванию количества
    df = df.sort_values(by='Количество', ascending=False)

    def get_priority(row):
        article = str(row['Артикул']).lower()
        match = re.search(r'k(\d+)', article)  # Ищем "k" с одной или более цифрами после него и запоминаем число
        if match:
            return int(match.group(1))  # Возвращаем число после "k" как приоритет
        else:
            return 0  # Наименьший приоритет для тех, у кого нет "k"

    df['Приоритет'] = df.apply(get_priority, axis=1)

    # Сортируем сначала по приоритету (убыванию), затем по наименованию товара (возрастанию), и в конце по артикулу (возрастанию)
    df = df.sort_values(by=['Приоритет', 'Наименование товара', 'Артикул'], ascending=[False, True, True])

    df = df.drop('Приоритет', axis=1)

    return df

def main():
    st.title("Обработка и отображение данных из CSV")

    uploaded_file = st.file_uploader("Загрузите CSV файл", type="csv")

    if uploaded_file is not None:
        # Попробуем прочитать файл с разными параметрами и обработкой ошибок
        df = None  # Инициализируем df как None
        try:
            # 1. Чтение с указанием разделителя (определите правильный разделитель!)
            df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')  # По умолчанию - запятая

        except pd.errors.ParserError as e:
            st.error(f"Ошибка разбора CSV файла (попытка 1): {e}. Попробуйте указать правильный разделитель.")
            try:
                # 2. Чтение с другим разделителем (например, точка с запятой)
                df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
                st.info("Файл успешно прочитан с разделителем ';'")  # Сообщение об успехе
            except pd.errors.ParserError as e2:
                st.error(f"Ошибка разбора CSV файла (попытка 2): {e2}.  Попробуйте указать другой разделитель или исправить файл.")
                try:
                    # 3. Чтение с помощью csv.reader и ручной обработкой
                    data = []
                    with open(uploaded_file, 'r', encoding='utf-8') as csvfile:
                        reader = csv.reader(csvfile)  # Автоматическое определение разделителя
                        for row in reader:
                            data.append(row)
                    df = pd.DataFrame(data[1:], columns=data[0])  # Создаем DataFrame
                    st.info("Файл успешно прочитан с помощью csv.reader")
                except Exception as e3:
                    st.error(f"Ошибка при чтении файла с помощью csv.reader: {e3}. Проверьте структуру файла.")
                    return  # Выходим, если ничего не получилось


        except UnicodeDecodeError:
            try:
                df = pd.read_csv(uploaded_file, encoding='latin1')
            except UnicodeDecodeError:
                st.error("Не удалось декодировать CSV файл. Попробуйте другую кодировку (utf-8, latin1, cp1251).")
                return
        except Exception as e:
            st.error(f"Произошла общая ошибка при чтении файла: {e}")
            return

        # Дальше - ваш код для обработки DataFrame (если он был успешно прочитан)
        if df is not None:  # Проверяем, что DataFrame был успешно прочитан
            try:
                selected_columns = ['Номер заказа', 'Наименование товара', 'Артикул', 'Количество']
                df = df[selected_columns]
            except KeyError as e:
                st.error(f"Одна или несколько колонок отсутствуют в файле: {e}. Убедитесь, что названия колонок в CSV файле совпадают с ожидаемыми.")
                return


            try:
                df['Стикер'] = df['Номер заказа'].apply(extract_sticker)
            except Exception as e:
                st.error(f"Произошла ошибка при создании столбца 'Стикер': {e}. Убедитесь, что столбец 'Номер заказа' существует и имеет строковый тип.")
                return

            df['Артикул'] = df['Артикул'].fillna('')
            sorted_df = sort_dataframe(df.copy())

            st.dataframe(sorted_df)

            csv = sorted_df.to_csv(index=False)
            st.download_button(
                label="Скачать отсортированный CSV",
                data=csv,
                file_name='sorted_data.csv',
                mime='text/csv',
            )


if __name__ == "__main__":
    main()
