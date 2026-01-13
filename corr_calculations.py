# corr_calculations.py
import numpy as np
from scipy.stats import spearmanr
from stat_corr_types import TStatCorr, TColumnPair

class TSpearmanDataSource:
    def __init__(self, vec_a, vec_b):
        self.vec_a = np.array(vec_a)
        self.vec_b = np.array(vec_b)

    def get_data(self, col, rec):
        if rec < 0 or rec >= len(self.vec_a):
            return 0.0
        return self.vec_a[rec] if col == 0 else self.vec_b[rec]

def rank_array(values):
    # Ранжирование с обработкой ties (как в Delphi)
    n = len(values)
    if n == 0:
        return np.array([])

    indices = np.argsort(values)
    ranks = np.empty(n)

    i = 0
    while i < n:
        j = i
        while j < n - 1 and np.abs(values[indices[j]] - values[indices[j + 1]]) < 1e-9:
            j += 1
        temp_rank = (i + j + 2) / 2.0  # Ранги с 1
        for k in range(i, j + 1):
            ranks[indices[k]] = temp_rank
        i = j + 1

    return ranks

def spearman_corr(stat_corr: TStatCorr, col_a, col_b, get_data, num_records, log_scale):
    if num_records < 2:
        return np.nan

    vals_a = np.array([get_data(col_a, i) for i in range(num_records)])
    vals_b = np.array([get_data(col_b, i) for i in range(num_records)])

    if log_scale:
        vals_a = np.log10(np.clip(vals_a, 1e-10, None))
        vals_b = np.log10(np.clip(vals_b, 1e-10, None))

    # Используем scipy для Spearman (он обрабатывает ties автоматически)
    corr, _ = spearmanr(vals_a, vals_b)
    return corr

def join_percent(stat_corr: TStatCorr, col_a, col_b, get_data, num_records, percent):
    if num_records < 2 or percent <= 0 or percent > 100:
        return 0.0

    cnt_sel = int(num_records * percent / 100.0)
    if cnt_sel == 0:
        return 0.0

    # Сортировка индексов по значениям (ascending)
    indices_a = np.argsort([get_data(col_a, i) for i in range(num_records)])
    indices_b = np.argsort([get_data(col_b, i) for i in range(num_records)])

    # Топ cnt_sel (последние, т.к. ascending)
    top_a = set(indices_a[-cnt_sel:])
    top_b = set(indices_b[-cnt_sel:])

    cnt_11 = len(top_a.intersection(top_b))

    denominator = cnt_sel * 2 - cnt_11
    return (cnt_11 * 100.0 / denominator) if denominator != 0 else 0.0

def calculate_rr_for_pair(stat_corr: TStatCorr, pair_idx, get_data, num_records, log_scale):
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
        idx_ac = stat_corr.get_pair_index(col_a, i)
        idx_bc = stat_corr.get_pair_index(col_b, i)
        if idx_ac == -1 or idx_bc == -1:
            continue
        corr_vec_a.append(stat_corr.get_corr(idx_ac))
        corr_vec_b.append(stat_corr.get_corr(idx_bc))

    common_count = len(corr_vec_a)
    if common_count < 2:
        stat_corr.set_rr(pair_idx, 0.0)
        return

    data_source = TSpearmanDataSource(corr_vec_a, corr_vec_b)
    value = spearman_corr(stat_corr, 0, 1, data_source.get_data, common_count, False)
    stat_corr.set_rr(pair_idx, value)

def calculate_all_correlations(stat_corr: TStatCorr, get_data, num_records, percent50=50, percent10=10, log_scale=False):
    # Пропускаем пары с invalid столбцами (проверяем перед расчётом)
    for i in range(stat_corr.count()):
        pair = stat_corr.get_pair(i)
        if pair.col1 in stat_corr.invalid_columns or pair.col2 in stat_corr.invalid_columns:  # Добавьте invalid из TData
            stat_corr.set_corr(i, np.nan)
            stat_corr.set_dist50(i, np.nan)
            stat_corr.set_dist10(i, np.nan)
            continue
    # 1. Расчёт Corr, Dist50, Dist10
    for i in range(stat_corr.count()):
        pair = stat_corr.get_pair(i)
        stat_corr.set_corr(i, spearman_corr(stat_corr, pair.col1, pair.col2, get_data, num_records, log_scale))
        stat_corr.set_dist50(i, join_percent(stat_corr, pair.col1, pair.col2, get_data, num_records, percent50))
        stat_corr.set_dist10(i, join_percent(stat_corr, pair.col1, pair.col2, get_data, num_records, percent10))

    # 2. RR
    for i in range(stat_corr.count()):
        calculate_rr_for_pair(stat_corr, i, get_data, num_records, log_scale)

    # 3. Статистики
    stat_corr.update_all_statistics()
    stat_corr.update_feature_statistics()