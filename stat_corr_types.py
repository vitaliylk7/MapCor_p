# stat_corr_types.py
import numpy as np
from dataclasses import dataclass

@dataclass
class TColumnPair:
    col1: int  # меньший индекс
    col2: int  # больший индекс

    @classmethod
    def create(cls, a, b):
        return cls(min(a, b), max(a, b))

    def contains(self, index):
        return index == self.col1 or index == self.col2

    def __str__(self):
        return f"{self.col1}-{self.col2}"

@dataclass
class TExtendedStat:
    min: float = 0.0
    max: float = 0.0
    mean: float = 0.0
    std_dev: float = 0.0

@dataclass
class TFeatureStat:
    feature_idx: int = 0
    pair_count: int = 0
    avg_corr: float = 0.0
    avg_dist10: float = 0.0
    avg_rr: float = 0.0

class TStatCorr:
    def __init__(self):
        self.column_names = []
        self.pairs = []
        self.pair_names = []
        self.corr = []
        self.dist10 = []
        self.rr = []
        self.reserve1 = []  # Зарезервировано
        self.reserve2 = []  # Зарезервировано
        self.all_pairs_stat = {
            'corr': TExtendedStat(),
            'dist10': TExtendedStat(),
            'rr': TExtendedStat()
        }
        self.feature_stats = []

    def initialize(self, column_names):
        self.clear()
        self.column_names = column_names.copy()

    def clear(self):
        self.column_names = []
        self.pairs = []
        self.pair_names = []
        self.corr = []
        self.dist10 = []
        self.rr = []
        self.reserve1 = []
        self.reserve2 = []
        self.feature_stats = []
        self.all_pairs_stat = {
            'corr': TExtendedStat(),
            'dist10': TExtendedStat(),
            'rr': TExtendedStat()
        }

    def count(self):
        return len(self.pairs)

    def add_or_get_pair(self, col1, col2):
        if col1 == col2:
            return -1
        p = TColumnPair.create(col1, col2)
        idx = self.find_pair_index(p.col1, p.col2)
        if idx >= 0:
            return idx
        # Добавляем
        idx = len(self.pairs)
        self.pairs.append(p)
        self.pair_names.append(self.generate_pair_name(p.col1, p.col2))
        self.corr.append(0.0)
        self.dist10.append(0.0)
        self.rr.append(0.0)
        self.reserve1.append(0.0)
        self.reserve2.append(0.0)
        return idx

    def find_pair_index(self, col1, col2):
        for i, p in enumerate(self.pairs):
            if p.col1 == col1 and p.col2 == col2:
                return i
        return -1

    def generate_pair_name(self, col1, col2):
        n1 = self.column_names[col1]
        n2 = self.column_names[col2]
        return f"{n1} _ {n2}"

    def set_corr(self, index, value):
        if 0 <= index < len(self.corr):
            self.corr[index] = value

    def set_dist10(self, index, value):
        if 0 <= index < len(self.dist10):
            self.dist10[index] = value

    def set_rr(self, index, value):
        if 0 <= index < len(self.rr):
            self.rr[index] = value

    def get_column_name(self, idx):
        return self.column_names[idx] if 0 <= idx < len(self.column_names) else ""

    def get_pair_name(self, index):
        return self.pair_names[index] if 0 <= index < len(self.pair_names) else ""

    def get_pair(self, index):
        return self.pairs[index] if 0 <= index < len(self.pairs) else TColumnPair(-1, -1)

    def get_corr(self, index):
        return self.corr[index] if 0 <= index < len(self.corr) else 0.0

    def get_dist10(self, index):
        return self.dist10[index] if 0 <= index < len(self.dist10) else 0.0

    def get_rr(self, index):
        return self.rr[index] if 0 <= index < len(self.rr) else 0.0

    def get_pair_index(self, col1, col2):
        return self.find_pair_index(min(col1, col2), max(col1, col2))

    def update_all_statistics(self):
        n = self.count()
        if n == 0:
            return

        # Векторизация с Numpy
        corr_arr = np.array(self.corr)
        dist10_arr = np.array(self.dist10)
        rr_arr = np.array(self.rr)

        self.all_pairs_stat['corr'] = TExtendedStat(corr_arr.min(), corr_arr.max(), corr_arr.mean(), corr_arr.std())
        self.all_pairs_stat['dist10'] = TExtendedStat(dist10_arr.min(), dist10_arr.max(), dist10_arr.mean(), dist10_arr.std())
        self.all_pairs_stat['rr'] = TExtendedStat(rr_arr.min(), rr_arr.max(), rr_arr.mean(), rr_arr.std())

    def update_feature_statistics(self):
        n_features = len(self.column_names)
        if n_features == 0:
            return

        sums = [{'sum_corr': 0.0, 'sum_d10': 0.0, 'sum_rr': 0.0, 'count': 0} for _ in range(n_features)]

        for i in range(self.count()):
            col1 = self.pairs[i].col1
            col2 = self.pairs[i].col2

            # Для col1
            sums[col1]['sum_corr'] += self.corr[i]
            sums[col1]['sum_d10'] += self.dist10[i]
            sums[col1]['sum_rr'] += self.rr[i]
            sums[col1]['count'] += 1

            # Для col2
            sums[col2]['sum_corr'] += self.corr[i]
            sums[col2]['sum_d10'] += self.dist10[i]
            sums[col2]['sum_rr'] += self.rr[i]
            sums[col2]['count'] += 1

        self.feature_stats = []
        for i in range(n_features):
            s = sums[i]
            count = s['count']
            fs = TFeatureStat(i, count)
            if count > 0:
                fs.avg_corr = s['sum_corr'] / count
                fs.avg_dist10 = s['sum_d10'] / count
                fs.avg_rr = s['sum_rr'] / count
            self.feature_stats.append(fs)