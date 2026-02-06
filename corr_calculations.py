# corr_calculations.py
"""
Модуль расчёта корреляций Спирмена, DIST10 и RR (мета-корреляции).
Все функции адаптированы под отсутствие log_scale — параметр удалён.
"""

import numpy as np
from scipy.stats import spearmanr
from stat_corr_types import TStatCorr, TColumnPair

def rank_array(values):
    """
    Ранжирование массива с обработкой связок (средний ранг).
    Порог сравнения 1e-9 для учёта погрешности float (как в оригинале Delphi).
    """
    values = np.asarray(values)
    n = len(values)
    if n == 0:
        return np.array([])

    indices = np.argsort(values)
    ranks = np.empty(n, dtype=float)

    i = 0
    while i < n:
        j = i
        while j < n - 1 and abs(values[indices[j]] - values[indices[j + 1]]) < 1e-9:
            j += 1
        temp_rank = (i + j + 2) / 2.0  # ранги начинаются с 1
        for k in range(i, j + 1):
            ranks[indices[k]] = temp_rank
        i = j + 1

    return ranks


def join_percent(stat_corr, col_a, col_b, get_data, num_records, percent):
    """
    Коэффициент пересечения топ-% значений (DIST10).
    """
    #return 0.0 #заглушка

    if num_records < 2 or percent <= 0 or percent > 100:
        return 0.0

    cnt_sel = int(num_records * percent / 100.0)
    if cnt_sel == 0:
        return 0.0

    indices_a = np.argsort([get_data(col_a, i) for i in range(num_records)])
    indices_b = np.argsort([get_data(col_b, i) for i in range(num_records)])

    top_a = set(indices_a[-cnt_sel:])
    top_b = set(indices_b[-cnt_sel:])

    cnt_11 = len(top_a.intersection(top_b))
    denominator = cnt_sel * 2 - cnt_11

    return (cnt_11 * 100.0 / denominator) if denominator != 0 else 0.0


def calculate_rr_for_pair(stat_corr, pair_idx, get_data, num_records):
    """
    Расчёт мета-корреляции RR для одной пары (Spearman между векторами корреляций).
    log_scale удалён — всегда без логарифмирования.
    """
    pair = stat_corr.get_pair(pair_idx)
    col_a, col_b = pair.col1, pair.col2

    num_features = len(stat_corr.column_names)
    if num_features < 3:
        stat_corr.set_rr(pair_idx, 0.0)
        return

    corr_vec_a = []
    corr_vec_b = []
    for i in range(num_features):
        if i == col_a or i == col_b:
            continue
        idx_ac = stat_corr.get_pair_index(min(col_a, i), max(col_a, i))
        idx_bc = stat_corr.get_pair_index(min(col_b, i), max(col_b, i))
        if idx_ac == -1 or idx_bc == -1:
            continue
        corr_vec_a.append(stat_corr.get_corr(idx_ac))
        corr_vec_b.append(stat_corr.get_corr(idx_bc))

    common_count = len(corr_vec_a)
    if common_count < 2:
        stat_corr.set_rr(pair_idx, 0.0)
        return

    # Всегда без log_scale → используем scipy (или manual, если нужно)
    value = spearmanr(corr_vec_a, corr_vec_b)[0]
    stat_corr.set_rr(pair_idx, value)


def calculate_all_correlations(
    stat_corr: TStatCorr,
    get_data,
    num_records: int,
    percent10: int = 10,
):
    """
    Основная функция расчёта всех корреляций.

    """
    # 1. Расчёт DIST10 (не зависит от режима)
    for i in range(stat_corr.count()):
        pair = stat_corr.get_pair(i)
        stat_corr.set_dist10(i, join_percent(stat_corr, pair.col1, pair.col2, get_data, num_records, percent10))

    # 2. Расчёт Spearman R
    for i in range(stat_corr.count()):
        pair = stat_corr.get_pair(i)
        col_a, col_b = pair.col1, pair.col2

        vals_a = np.array([get_data(col_a, rec) for rec in range(num_records)])
        vals_b = np.array([get_data(col_b, rec) for rec in range(num_records)])

        # Расчёт корреляции
        corr, _ = spearmanr(vals_a, vals_b, nan_policy='omit')

        stat_corr.set_corr(i, corr)

    # 3. Расчёт RR (мета-корреляция)
    for i in range(stat_corr.count()):
        calculate_rr_for_pair(stat_corr, i, get_data, num_records)

    # 4. Обновление всех статистик
    stat_corr.update_all_statistics()
    stat_corr.update_feature_statistics()