# associations.py
import pandas as pd
import numpy as np
from stat_corr_types import TStatCorr

def build_associations(stat_corr: TStatCorr, params: dict):
    """
    Формирует ассоциации по алгоритму.
    params: {'threshold_root': float, 'threshold_avg': float, 'max_iters': int, 'convergence_epsilon': float}
    Возвращает: list[dict] кластеров {'features': list[int], 'root': int, 'internal_avg_r': float, 'external_avg_r': float}
    """
    num_features = len(stat_corr.column_names)
    if num_features < 2 or stat_corr.count() == 0:
        return []

    # Матрица R (симметричная, положительные только)
    corr_matrix = pd.DataFrame(np.zeros((num_features, num_features)), dtype=float)
    for i in range(stat_corr.count()):
        pair = stat_corr.get_pair(i)
        r = stat_corr.get_corr(i)
        if np.isnan(r):
            r = 0.0
        r = max(r, 0.0)  # Только положительные
        corr_matrix.at[pair.col1, pair.col2] = r
        corr_matrix.at[pair.col2, pair.col1] = r
    np.fill_diagonal(corr_matrix.values, 1.0)

    # Доступные фичи
    available = set(range(num_features))
    clusters = []
    iter_count = 0
    while available and iter_count < params['max_iters']:
        iter_count += 1
        moved = 0

        # Фаза 1: Новые кластеры из available
        while available:
            # Найти max пару в available
            max_r = -1
            best_pair = None
            for a in list(available):
                for b in list(available):
                    if a < b and corr_matrix.at[a, b] > max_r and corr_matrix.at[a, b] >= params['threshold_root']:
                        max_r = corr_matrix.at[a, b]
                        best_pair = (a, b)
            if best_pair is None:
                break  # Нет сильных пар

            a, b = best_pair
            # Выбрать root: тот с выше средней R ко всем
            mean_a = corr_matrix.iloc[a].mean()
            mean_b = corr_matrix.iloc[b].mean()
            root = a if mean_a > mean_b else b
            current_cluster = [a, b] if root == a else [b, a]

            # Добавление кандидатов
            candidates = sorted(available - set(current_cluster),
                                key=lambda c: corr_matrix.at[c, root], reverse=True)
            for cand in candidates:
                r_root = corr_matrix.at[cand, root]
                avg_cluster = corr_matrix.loc[cand, current_cluster].mean()
                if r_root >= params['threshold_root'] and avg_cluster >= params['threshold_avg']:
                    current_cluster.append(cand)

            # Добавить кластер
            clusters.append({'features': sorted(current_cluster), 'root': root})
            available -= set(current_cluster)

        # Фаза 2: Перераспределение weak
        weak_features = []
        for cl in clusters:
            root = cl['root']
            weak = [f for f in cl['features'] if f != root and corr_matrix.at[f, root] < params['threshold_root']]
            weak_features.extend(weak)
            cl['features'] = [f for f in cl['features'] if f not in weak]  # Удалить weak

        # Добавить weak в available
        available.update(weak_features)

        # Переместить weak/free в лучшие кластеры
        for f in list(available):
            current_r = 0.0  # Для free
            best_cl = None
            best_potential_r = -1
            for cl_idx, cl in enumerate(clusters):
                if len(cl['features']) == 0:
                    continue
                potential_r = corr_matrix.at[f, cl['root']]
                avg_cl = corr_matrix.loc[f, cl['features']].mean()
                if (potential_r > best_potential_r and potential_r > current_r and
                    potential_r >= params['threshold_root'] and avg_cl >= params['threshold_avg']):
                    best_potential_r = potential_r
                    best_cl = cl_idx
            if best_cl is not None:
                clusters[best_cl]['features'].append(f)
                available.remove(f)
                moved += 1

        # Проверка сходимости
        if moved < params['convergence_epsilon'] * num_features:
            break

    # Постобработка: одиночные + averages
    for f in available:
        clusters.append({'features': [f], 'root': f})

    for cl in clusters:
        feats = cl['features']
        if len(feats) == 1:
            cl['internal_avg_r'] = 1.0
            external_mask = ~corr_matrix.index.isin(feats)
            cl['external_avg_r'] = corr_matrix.loc[feats[0], external_mask].mean()
        else:
            # Internal: mean верхнего треугольника
            internal_rs = [corr_matrix.at[i, j] for i in feats for j in feats if i < j]
            cl['internal_avg_r'] = np.mean(internal_rs) if internal_rs else 0.0
            # External: mean с остальными
            external_mask = ~corr_matrix.index.isin(feats)
            external_avg = corr_matrix.loc[feats, external_mask].values.flatten().mean()
            cl['external_avg_r'] = external_avg if not np.isnan(external_avg) else 0.0

    # Сортировка: по размеру desc, затем internal_avg_r desc
    clusters.sort(key=lambda c: (-len(c['features']), -c['internal_avg_r']))

    # Удалить пустые
    clusters = [c for c in clusters if c['features']]

    return clusters