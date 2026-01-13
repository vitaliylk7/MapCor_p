# data.py
import pandas as pd
import numpy as np
import os
import logging  # Для лога ошибок

# Настройка логирования
logging.basicConfig(filename='data_load.log', level=logging.WARNING, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class TStat:
    def __init__(self, min_val=0, max_val=0, mean=0, mean_l=0, min_bigger_zero=0):
        self.min = min_val
        self.max = max_val
        self.mean = mean
        self.mean_l = mean_l
        self.min_bigger_zero = min_bigger_zero

class TData:
    def __init__(self):
        self.filename = ""
        self.df = None  # Pandas DataFrame для данных
        self.stats = []  # Список TStat для каждого столбца
        self.invalid_columns = []  # Список индексов столбцов с невалидными данными
        self.is_loaded = False

    def load_file(self, fname):
        """
        Чтение файла в формате:
        Опционально: 1-я строка — количество столбцов (целое число)
        Следующая строка — названия столбцов (разделены пробелами/табами)
        Далее — строки с числами (float, разделитель пробел/таб)
        Улучшения:
        - Автоопределение, если нет строки с кол-вом столбцов
        - Замена ',' на '.' для десятичных разделителей
        - Пропуск комментариев (#, //, ;)
        - Поддержка Excel (.xlsx)
        - Обработка невалидных значений: лог + пометка столбца
        - Проверка на большие данные (столбцы <=200, строки <=65000)
        """
        try:
            ext = os.path.splitext(fname)[1].lower()

            if ext == '.xlsx':
                # Excel-поддержка
                self.df = pd.read_excel(fname, engine='openpyxl', dtype='float32', header=0)
                logging.info(f"Загружен Excel файл: {fname}, строк: {len(self.df)}, столбцов: {len(self.df.columns)}")
            else:
                # Текстовый файл (CSV/TXT)
                with open(fname, 'r', encoding='utf-8', errors='replace') as f:
                    lines = []
                    for line in f:
                        stripped = line.strip()
                        if stripped and not stripped.startswith(('#', '//', ';')):
                            lines.append(stripped)  # Пропуск комментариев

                if len(lines) < 1:
                    raise ValueError("Файл пустой или содержит только комментарии")

                # Автоопределение формата
                try:
                    expected_cols = int(lines[0])
                    header_line_idx = 1
                    data_start_idx = 2
                except ValueError:
                    expected_cols = None
                    header_line_idx = 0
                    data_start_idx = 1

                # Заголовок
                header_parts = lines[header_line_idx].split()
                actual_header_cols = len(header_parts)

                if expected_cols is None:
                    expected_cols = actual_header_cols
                elif actual_header_cols != expected_cols:
                    logging.warning(f"Предупреждение: указано {expected_cols} столбцов, но в заголовке {actual_header_cols}")

                column_names = header_parts[:expected_cols]

                # Данные
                data_lines = lines[data_start_idx:]
                if not data_lines:
                    raise ValueError("Нет строк с данными")

                data = []
                bad_rows = []
                for i, line in enumerate(data_lines, start=data_start_idx + 1):
                    line = line.replace(',', '.')  # Десятичный разделитель
                    try:
                        row = []
                        for val_str in line.split():
                            try:
                                row.append(float(val_str))
                            except ValueError:
                                row.append(np.nan)  # Невалидное значение → NaN
                                logging.warning(f"Невалидное значение '{val_str}' в строке {i}, заменено на NaN")
                        actual_cols = len(row)
                        if actual_cols != expected_cols:
                            bad_rows.append((i, f"Ожидалось {expected_cols} столбцов, найдено {actual_cols}"))
                        data.append(row)
                    except Exception as e:
                        bad_rows.append((i, str(e)))
                        continue

                if bad_rows:
                    logging.warning("Обнаружены проблемные строки:")
                    for row_info in bad_rows[:5]:
                        logging.warning(row_info)
                    if len(bad_rows) > 5:
                        logging.warning(f"... и ещё {len(bad_rows) - 5} строк с ошибками")

                if not data:
                    raise ValueError("Нет корректных строк данных")

                self.df = pd.DataFrame(data, columns=column_names[:len(data[0])], dtype='float32')

            # Проверка на большие данные
            num_rows, num_cols = self.df.shape
            if num_rows > 65000 or num_cols > 200:
                logging.warning(f"Данные превышают лимит: строк {num_rows} (>65000), столбцов {num_cols} (>200)")
                raise ValueError("Данные слишком большие для обработки")

            # Помечаем invalid столбцы (где >10% NaN или все NaN)
            for col_idx, col in enumerate(self.df.columns):
                nan_percent = self.df[col].isna().mean()
                if nan_percent > 0.1:  # Или все NaN
                    self.invalid_columns.append(col_idx)
                    logging.warning(f"Столбец '{col}' помечен invalid: {nan_percent*100:.1f}% NaN")

            self.filename = fname
            self.df = self.df[sorted(self.df.columns)]  # Сортировка по алфавиту
            self.calc_stat()
            self.is_loaded = True
            return True

        except Exception as e:
            logging.error(f"Ошибка загрузки {fname}: {e}")
            import traceback
            traceback.print_exc()
            return False

    def calc_stat(self):
        self.stats = []
        for col in self.df.columns:
            values = self.df[col].dropna()  # Игнорируем NaN
            if len(values) == 0:
                continue

            min_val = values.min()
            max_val = values.max()
            mean = values.mean()

            # Min bigger zero
            positive = values[values > 0]
            min_bigger_zero = positive.min() if not positive.empty else 0.1

            # Mean log (mean_l)
            log_values = np.log10(values.clip(lower=min_bigger_zero / 2))  # Избежать log(0)
            mean_l = log_values.mean()

            self.stats.append(TStat(min_val, max_val, mean, mean_l, min_bigger_zero))

    def get_data(self, col, rec):
        return self.df.iloc[rec, col]

    def get_data_l(self, col, rec):
        value = self.get_data(col, rec)
        min_bz = self.get_min_bigger_zero(col)
        return np.log10(value) if value > 0 else np.log10(min_bz / 2)

    def get_count_column(self):
        return len(self.df.columns) if self.df is not None else 0

    def get_count_record(self):
        return len(self.df) if self.df is not None else 0

    def get_column_name(self, col):
        return self.df.columns[col]

    def get_number_for_column_name(self, col_name):
        try:
            return self.df.columns.get_loc(col_name)
        except KeyError:
            return -1

    def get_min(self, col):
        return self.stats[col].min

    def get_max(self, col):
        return self.stats[col].max

    def get_mean(self, col):
        return self.stats[col].mean

    def get_mean_l(self, col):
        return self.stats[col].mean_l

    def get_min_bigger_zero(self, col):
        return self.stats[col].min_bigger_zero

    def get_file_name(self):
        return self.filename

    def get_column_names(self):
        return list(self.df.columns)

    def round_b(self, value, m):
        e = 10 ** m
        if value * e > 2147483647:
            return value
        return np.ceil(value * e) / e if value > 0 else np.floor(value * e) / e