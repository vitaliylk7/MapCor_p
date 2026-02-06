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
                self.df = pd.read_excel(fname, engine='openpyxl', dtype='float64', header=0)
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

                self.df = pd.DataFrame(data, columns=column_names[:len(data[0])], dtype='float64')

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


    def get_full_statistics(self):
        """
        Возвращает pandas DataFrame со статистикой по всем столбцам.
        Используется для отображения в GUI и отчётах.
        """
        if self.df is None or self.df.empty:
            return None

        import pandas as pd
        import numpy as np
        from scipy.stats import skew, kurtosis

        desc = self.df.describe(percentiles=[0.05, 0.25, 0.5, 0.75, 0.95]).T

        # Дополнительные метрики
        desc['skew']              = self.df.skew(numeric_only=True).round(3)
        desc['kurtosis']          = self.df.kurtosis(numeric_only=True).round(3)
        desc['nan_percent']       = (self.df.isna().mean() * 100).round(2)

        desc['below_lod_percent'] = ((self.df <= 0.03) | self.df.isna()).mean() * 100
        desc['below_lod_percent'] = desc['below_lod_percent'].round(1)

        # Процент повторяющихся минимальных значений
        def repeating_min_ratio(col):
            if col.isna().all():
                return 0.0
            min_val = col.min()
            if pd.isna(min_val):
                return 0.0
            count_min = (col == min_val).sum()
            total = len(col)
            return (count_min / total * 100) if total > 0 else 0.0

        desc['repeating_min_percent'] = self.df.apply(repeating_min_ratio).round(1)

        desc['zero_percent']      = ((self.df == 0) | (self.df <= 0.03)).mean() * 100
        desc['zero_percent']      = desc['zero_percent'].round(1)

        desc['unique_count']      = self.df.nunique()
        desc['variance']          = self.df.var(numeric_only=True).round(6)

        # ─── Новая метрика: CV, % ───────────────────────────────────────────────
        # CV = std / mean * 100, с защитой от деления на 0
        desc['CV_percent'] = np.where(
            desc['mean'] != 0,
            (desc['std'] / desc['mean'].abs() * 100).round(1),
            np.nan
        )


        # ── J (информативность по Шеннону, нормированная на 6 интервалов) ────────────────
        def compute_entropy_norm(series, n_bins=6):
            """
            Информативность J — мера однородности геологического признака.
            Диапазон: [0, 1]
            J → 0: монолитный пласт (все значения в одном интервале)
            J → 1: равномерное распределение по всем 6 интервалам (макс. гетерогенность)
            """
            series = series.dropna()
            if len(series) < 2:
                return np.nan
            
            try:
                # Гистограмма по фиксированным 6 интервалам
                hist, _ = np.histogram(series, bins=n_bins)
                total = hist.sum()
                if total == 0:
                    return np.nan
                
                # Вероятности только для непустых интервалов (защита от log2(0))
                p = hist[hist > 0] / total
                
                # Энтропия Шеннона
                H = -np.sum(p * np.log2(p))
                
                # Нормировка НА ФИКСИРОВАННОЕ число интервалов (6), а не на количество непустых!
                H_max = np.log2(n_bins)
                J = H / H_max
                return J
            
            except Exception:
                return np.nan

        desc['J'] = self.df.apply(lambda col: compute_entropy_norm(col, n_bins=6)).round(3)

        # Округление
        desc = desc.round({
            'mean': 3, 'std': 3, 'min': 3, 'max': 3, '50%': 3,
            'below_lod_percent': 1, 'repeating_min_percent': 1,
            'zero_percent': 1, 'CV_percent': 1,
            'variance': 6, 'skew': 3, 'kurtosis': 3
        })

        # Переименование
        desc = desc.rename(columns={
            '50%': 'median',
            '25%': 'Q1',
            '75%': 'Q3'
        })

        # ─── Порядок столбцов (CV ставим после std и mean) ────────────────────────
        columns_order = [
            'count', 'nan_percent',
            'min', 'repeating_min_percent',
            'below_lod_percent',
            'zero_percent',
            '5%', 'Q1', 'median', 'Q3', '95%', 'max',
            'mean', 'std', 'CV_percent',          # ← новая позиция
            'variance',
            'skew', 'kurtosis',
            'unique_count', 'J'
        ]

        existing_cols = [c for c in columns_order if c in desc.columns]
        desc = desc[existing_cols]

        return desc


    def get_geo_recommendations(self):
        """
        Вычисляет минимальный набор параметров и генерирует текстовые рекомендации
        для каждой характеристики с точки зрения геолога.
        
        Возвращает: dict {имя_столбца: {'params': dict_параметров, 'recommendation': str}}
        """
        if self.df is None or self.df.empty:
            return {}

        import pandas as pd
        import numpy as np

        df = self.df.select_dtypes(include=[np.number])  # только числовые столбцы

        recommendations = {}

        for col in df.columns:
            s = df[col]
            n = len(s.dropna())

            if n < 10:
                recommendations[col] = {
                    'params': {},
                    'recommendation': "Мало данных (<10 значений) — характеристика неинформативна."
                }
                continue

            # Основные параметры
            params = {
                'count': int(n),
                'min': float(s.min()),
                'p5': float(s.quantile(0.05)),
                'median': float(s.median()),
                'p95': float(s.quantile(0.95)),
                'max': float(s.max()),
                'geometric_mean': float(np.exp(np.log(s[s > 0]).mean())) if (s > 0).any() else np.nan,
                'cv_percent': float((s.std() / s.mean() * 100)) if s.mean() != 0 else np.nan,
                'anomaly_ratio': float(s.max() / s.median()) if s.median() != 0 else np.nan,
                'below_lod_percent': float((s <= 0.03).mean() * 100),
                'above_lod_count': int((s > 0.03).sum()),
                'skewness': float(s.skew()),
                'kurtosis': float(s.kurtosis()),
                'unique_percent': float(s.nunique() / n * 100) if n > 0 else 0
            }

            # Формирование рекомендации
            rec_parts = []

            # 1. Неинформативность
            if params['unique_percent'] < 5 or params['below_lod_percent'] > 70 or params['above_lod_count'] < max(20, n * 0.2):
                rec_parts.append("Неинформативна: очень мало уникальных значений / редко детектируется выше LOD / почти константа.")

            # 2. Стабильность / фон
            elif params['cv_percent'] < 30 and params['anomaly_ratio'] < 5:
                rec_parts.append("Стабильная фоновая характеристика, низкая изменчивость.")

            # 3. Высокая изменчивость / аномалии
            elif params['cv_percent'] > 80 or params['anomaly_ratio'] > 10:
                rec_parts.append("Высокая изменчивость и/или сильные аномалии — потенциально интересна для поиска рудных зон.")

            # 4. Трансформация
            if params['skewness'] > 1.5 or params['kurtosis'] > 6:
                rec_parts.append("Рекомендуется лог-трансформация (log10) перед расчётом корреляций из-за сильной асимметрии и тяжёлых хвостов.")

            # 5. Редкие детекции
            if params['below_lod_percent'] > 50:
                rec_parts.append(f"Часто ниже LOD ({params['below_lod_percent']:.1f}%) — корреляции могут быть шумными.")

            # Итоговая рекомендация
            if not rec_parts:
                rec_parts.append("Информативная характеристика, умеренная изменчивость, подходит для корреляционного анализа без специальной предобработки.")

            recommendations[col] = {
                'params': params,
                'recommendation': " ".join(rec_parts)
            }

        return recommendations

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
        
    def save_statistics_to_csv(self, filename=None):
        """
        Сохраняет полную статистику всех характеристик в TXT-файл с табуляцией в качестве разделителя,
        включая столбец с рекомендациями.
        
        Параметры:
            filename — путь к файлу (если None — открывается диалог сохранения)
        
        Возвращает:
            True — если сохранено успешно, False — при ошибке или отмене
        """
        import pandas as pd
        from pathlib import Path
        from PySide6.QtWidgets import QFileDialog

        # Получаем основную статистику
        stats_df = self.get_full_statistics()
        if stats_df is None or stats_df.empty:
            return False

        # Добавляем рекомендации
        recs = self.get_geo_recommendations()
        if recs:
            rec_series = pd.Series(
                {col: info['recommendation'] for col, info in recs.items()}
            )
            stats_df['Рекомендация'] = rec_series.reindex(stats_df.index).fillna("—")

        # Если имя файла не указано — открываем диалог
        if filename is None:
            default_name = str(Path(self.filename).with_suffix('.statistics.txt')) if self.filename else "statistics.txt"
            fname, _ = QFileDialog.getSaveFileName(
                None,
                "Сохранить статистику характеристик",
                default_name,
                "Текстовые файлы (*.txt);;Все файлы (*.*)"
            )
            if not fname:
                return False
            filename = fname

        # Убеждаемся, что расширение .txt
        if not str(filename).lower().endswith('.txt'):
            filename = str(filename) + '.txt'

        try:
            stats_df.to_csv(
                filename,
                sep='\t',                  # табуляция как разделитель
                encoding='utf-8-sig',
                index_label="Признак",
                float_format='%.6f',
                na_rep='—'
            )
            return True
        except Exception as e:
            print(f"Ошибка сохранения статистики: {e}")
            return False