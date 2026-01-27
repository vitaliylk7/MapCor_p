# main.py
"""
MAPCOR_M — Python-версия программы для расчёта ранговых корреляций
(перенос с Delphi → PySide6 + pandas + numpy + scipy)

Основные функции:
- Загрузка данных (txt/csv/xlsx)
- Выбор признаков через чек-лист (QListWidget с галочками)
- Расчёт Spearman, DIST10, RR (мета-корреляция)
- Отображение исходных данных в QTableView
- Таблица результатов в отдельном окне
- Генерация и просмотр HTML-отчёта
"""

import sys
import os
from pathlib import Path
import datetime
import pandas as pd
import numpy as np
import tkinter as tk
import json
import subprocess

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QTableView,
    QStatusBar,  QFileDialog, QMessageBox,
    QHeaderView, QDialog, QLabel, QCheckBox, QPushButton, QScrollArea, QGridLayout
)
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QStandardItemModel, QStandardItem, QDesktopServices, QFont, QColor

from data import TData
from stat_corr_types import TStatCorr, TExtendedStat
from corr_calculations import calculate_all_correlations


# ────────────────────────────────────────────────────────────────
# Вспомогательные функции для цветовой кодировки (должны быть ДО классов!)
# ────────────────────────────────────────────────────────────────
COLOR_SCALE = [
    '#7bb8c2',   #  0 — мягкий бирюзово-голубой (самый лучший)
    '#8cc9d3',   #  1
    '#a0e0e0',   #  2
    '#c2f0eb',   #  3 — очень светлый
    '#e0f7f4',   #  4 — почти белый с оттенком
    '#80cca8',   #  5 — мятный зелёный
    '#a8d97f',   #  6 — светло-салатовый
    '#d9ec5f',   #  7 — лимонно-зелёный

    '#f2e96b',   #  8 — мягкий жёлтый (не кричащий)
    '#f5cf63',   #  9 — жёлто-оранжевый
    '#f7b05b',   # 10 — светлый оранжевый
    '#f28b55',   # 11 — оранжево-красный
    '#e36a52',   # 12 — приглушённый красно-оранжевый
    '#d65a54',   # 13 — бледно-красный (финал, без агрессии)
]

def get_color_index(value, min_val, max_val, median=None):
    """
    Основная функция вычисления индекса цвета (0..13)
    Логика идентична getColorOfMedian в Delphi
    
    Параметры:
        value   — текущее значение
        min_val — минимум диапазона
        max_val — максимум диапазона
        median  — медиана (если None — линейное деление)
    
    Возвращает: индекс 0..13
    """
    if np.isnan(value):
        return 7  # середина, серый
    
    # Защита от выхода за границы
    value = max(min_val, min(max_val, value))
    
    if median is not None:
        # Деление относительно медианы (как в Delphi)
        if value >= median:
            # Выше медианы → индексы 7..13 (7 интервалов)
            portion = (value - median) / (max_val - median) if max_val > median else 0
            ind = round(portion * 7) + 6
            ind = min(13, max(7, ind))
        else:
            # Ниже медианы → индексы 0..6
            portion = (value - min_val) / (median - min_val) if median > min_val else 0
            ind = round(portion * 7)
            ind = min(6, max(0, ind))
    else:
        # Линейное деление всего диапазона на 14 частей
        portion = (value - min_val) / (max_val - min_val) if max_val > min_val else 0
        ind = round(portion * 13)
        ind = min(13, max(0, ind))
    
    return ind


def get_color_for_r(value, median=None):
    """Цвет для Spearman R (диапазон -1..1)"""
    return COLOR_SCALE[get_color_index(value, -1.0, 1.0, median)]


def get_color_for_rr(value, median=None):
    """Цвет для RR (мета-корреляция, тоже -1..1)"""
    return COLOR_SCALE[get_color_index(value, -1.0, 1.0, median)]


def get_color_for_dist10(value, median=None):
    """Цвет для DIST10 (0..100)"""
    return COLOR_SCALE[get_color_index(value, 0.0, 100.0, median)]


# ────────────────────────────────────────────────────────────────
# Дальше идут классы и остальной код
# ────────────────────────────────────────────────────────────────

class StatisticsDialog(QDialog):
    def __init__(self, data: TData, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Статистика характеристик")
        self.resize(1400, 800)
        self.data = data

        layout = QVBoxLayout(self)

        # Таблица
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        # Кнопки внизу
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("Сохранить в CSV...")
        btn_save.clicked.connect(self.save_to_csv)
        btn_layout.addWidget(btn_save)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        self._fill_table()

    def _fill_table(self):
        stats_df = self.data.get_full_statistics()
        if stats_df is None or stats_df.empty:
            QMessageBox.information(self, "Нет данных", "Данные не загружены или пустые.")
            return

        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(stats_df.columns.tolist())

        for row_idx, (feature, row) in enumerate(stats_df.iterrows()):
            items = [QStandardItem(feature)]  # первый столбец — имя признака

            for val in row:
                item = QStandardItem(str(val) if not pd.isna(val) else "—")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

                # Выделяем потенциально неинформативные красным
                if feature == 'Z' or (row['variance'] < 1e-5) or (row['unique_count'] < len(self.data.df) * 0.05):
                    item.setBackground(QColor(255, 220, 220))  # светло-красный

                items.append(item)

            model.appendRow(items)

        # Добавляем заголовок "Признак" в начало
        full_headers = ["Признак"] + stats_df.columns.tolist()
        model.setHorizontalHeaderLabels(full_headers)

        self.table.setModel(model)

        # Авто-размер первых столбцов
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)

    def save_to_csv(self):
        fname, _ = QFileDialog.getSaveFileName(
            self, "Сохранить статистику", "statistics.csv", "CSV (*.csv);;Все файлы (*.*)"
        )
        if not fname:
            return
        if not fname.lower().endswith('.csv'):
            fname += '.csv'

        stats_df = self.data.get_full_statistics()
        stats_df.to_csv(fname, encoding='utf-8-sig')
        QMessageBox.information(self, "Сохранено", f"Статистика сохранена в {fname}")


class ResultsDialog(QDialog):
    """Диалоговое окно с таблицей результатов корреляций"""
    def __init__(self, stat_corr, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Таблица результатов корреляций")
        self.resize(1000, 700)
        self.stat_corr = stat_corr

        layout = QVBoxLayout(self)

        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        layout.addWidget(self.table)

        self._fill_table()

    def _fill_table(self):
        model = QStandardItemModel()
        headers = ["Пара", "Spearman R", "DIST_10", "RR (мета-корр)"]
        model.setHorizontalHeaderLabels(headers)

        for i in range(self.stat_corr.count()):
            pair_name = self.stat_corr.get_pair_name(i)
            r = self.stat_corr.get_corr(i)
            d10 = self.stat_corr.get_dist10(i)
            rr = self.stat_corr.get_rr(i)

            row = [
                QStandardItem(pair_name),
                QStandardItem(f"{r:.3f}" if not np.isnan(r) else "—"),
                QStandardItem(f"{d10:.1f}" if not np.isnan(d10) else "—"),
                QStandardItem(f"{rr:.3f}" if not np.isnan(rr) else "—")
            ]

            model.appendRow(row)

        self.table.setModel(model)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MAPCOR — P версия")
        self.resize(1280, 800)
        self.closeEvent = self._on_close

        self.data = TData()
        self.stat_corr = TStatCorr()

        # Главный горизонтальный layout (левая + правая часть)
        main_layout = QHBoxLayout()
        central = QWidget()
        central.setLayout(main_layout)
        self.setCentralWidget(central)

        # Левая панель — выбор характеристик
        left_group = QGroupBox("Выбор характеристик")
        left_layout = QVBoxLayout(left_group)

        # ─── Кнопки управления выбором ──────────────────────────────────────
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        btn_all = QPushButton("Выбрать все")
        btn_none = QPushButton("Снять все")
        btn_inv = QPushButton("Инвертировать")

        btn_all.clicked.connect(lambda: self._set_all_checked(True))
        btn_none.clicked.connect(lambda: self._set_all_checked(False))
        btn_inv.clicked.connect(self._invert_checks)

        btn_layout.addWidget(btn_all)
        btn_layout.addWidget(btn_none)
        btn_layout.addWidget(btn_inv)
        btn_layout.addStretch()

        left_layout.addLayout(btn_layout)

        # ─── Прокручиваемая область с сеткой чекбоксов ──────────────────────
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.check_container = QWidget()
        self.grid_layout = QGridLayout(self.check_container)
        self.grid_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.grid_layout.setContentsMargins(12, 8, 12, 12)
        self.grid_layout.setSpacing(8)

        self.scroll_area.setWidget(self.check_container)
        left_layout.addWidget(self.scroll_area)

        # Добавляем в основной layout с желаемой шириной
        main_layout.addWidget(left_group, 5)   # ← 5 частей — левая панель шире

        # ─── Правая часть (таблица + настройки под ней) ──────────────────
        right_column = QVBoxLayout()  # вертикальный контейнер для правой стороны

        # Правая панель — исходные данные
        right_group = QGroupBox("Исходные данные")
        right_layout = QVBoxLayout(right_group)
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.verticalHeader().setVisible(False)
        right_layout.addWidget(self.table_view)

        right_column.addWidget(right_group, stretch=1)  # таблица растягивается

        # ─── Секция настроек — под таблицей ──────────────────────────────
        settings_group = QGroupBox("Настройки расчёта")
        settings_layout = QVBoxLayout(settings_group)

        self.delphi_compat_checkbox = QCheckBox(
            "Совместимость для старых отчётов"
        )
        # Сразу после создания чекбокса
        self.delphi_compat_checkbox.stateChanged.connect(self._save_delphi_setting) 

        # Загружаем сохранённое состояние чекбокса (если файл существует)
        try:
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                saved_value = settings.get("delphi_compat", False)
                self.delphi_compat_checkbox.setChecked(saved_value)
        except (FileNotFoundError, json.JSONDecodeError, KeyError):
            # Если файла нет или он повреждён — оставляем по умолчанию (False)
            self.delphi_compat_checkbox.setChecked(False)  # по умолчанию — scipy


        settings_layout.addWidget(self.delphi_compat_checkbox)

        # устанавливаем tooltip
        self.delphi_compat_checkbox.setToolTip(
            "Включено — точная копия старой версии\n"
            "Выключено — модернизированный метод (scipy) — рекомендуется"
        )
        settings_layout.addWidget(self.delphi_compat_checkbox)

        btn_stats = QPushButton("Статистика характеристик")
        btn_stats.setStyleSheet("font-weight: bold; padding: 8px; background-color: #6c757d; color: white;")
        btn_stats.clicked.connect(self.show_statistics)
        settings_layout.addWidget(btn_stats)

        btn_save_stats = QPushButton("Сохранить статистику в CSV")
        btn_save_stats.setStyleSheet("font-weight: bold; padding: 8px; background-color: #28a745; color: white;")
        btn_save_stats.clicked.connect(self.act_save_statistics)

        settings_layout.addWidget(btn_save_stats)

        # После btn_stats или после чекбокса в settings_layout

        btn_geo_rec = QPushButton("Геологические рекомендации")
        btn_geo_rec.setStyleSheet("font-weight: bold; padding: 8px; background-color: #17a2b8; color: white;")
        btn_geo_rec.clicked.connect(self.show_geo_recommendations)
        settings_layout.addWidget(btn_geo_rec)
        
        right_column.addWidget(settings_group, stretch=0)  # настройки не растягиваются


        # Добавляем правую колонку в главный layout
        main_layout.addLayout(right_column, stretch=4)  # правая часть уже левой


        # Статусбар
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Готов к открытию файла...")

        # Создаём меню
        self._create_menu()

    def _save_delphi_setting(self):
        """Сохраняет состояние чекбокса в файл settings.json"""
        settings = {
            "delphi_compat": self.delphi_compat_checkbox.isChecked()
        }
        try:
            with open("settings.json", "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")
            # Можно показать QMessageBox, но обычно достаточно лога в консоль

    def _on_close(self, event):
        # Сохраняем перед закрытием
        self._save_delphi_setting()
        event.accept()

    def get_selected_columns(self) -> list[int]:
        """Возвращает индексы выбранных (и активных) признаков"""
        selected = []
        for i, cb in enumerate(self.check_boxes):
            if cb.isChecked() and cb.isEnabled():
                selected.append(i)
        return selected

    def _set_all_checked(self, checked: bool):
        """Устанавливает состояние всем чекбоксам"""
        for cb in self.check_boxes:
            if cb.isEnabled():  # не трогаем отключённые
                cb.setChecked(checked)


    def _invert_checks(self):
        """Инвертирует состояние всех доступных чекбоксов"""
        for cb in self.check_boxes:
            if cb.isEnabled():
                cb.setChecked(not cb.isChecked())

    def act_save_statistics(self):
        """Сохраняет статистику характеристик в CSV"""
        if not hasattr(self, 'data') or self.data.df is None or self.data.df.empty:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите файл данных.")
            return
        
        success = self.data.save_statistics_to_csv()
        if success:
            QMessageBox.information(self, "Успех", "Статистика характеристик сохранена в CSV-файл.")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить файл статистики.")

    def show_geo_recommendations(self):
        if not hasattr(self, 'data') or self.data.df is None or self.data.df.empty:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите файл данных.")
            return

        recs = self.data.get_geo_recommendations()
        if not recs:
            QMessageBox.information(self, "Нет данных", "Нет числовых характеристик для анализа.")
            return

        # Формируем красивый текст для QMessageBox (или можно в отдельное окно)
        text = "Рекомендации по характеристикам (геологическая интерпретация):\n\n"
        for col, info in sorted(recs.items()):
            text += f"• {col}:\n"
            text += f"  {info['recommendation']}\n\n"

        # Показываем в большом окне с прокруткой
        msg = QMessageBox(self)
        msg.setWindowTitle("Геологические рекомендации по характеристикам")
        msg.setText(text)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.setDefaultButton(QMessageBox.Ok)

        # Делаем окно большим и с прокруткой
        msg.setSizeGripEnabled(True)
        msg.setMinimumSize(800, 600)
        msg.exec()

    def show_statistics(self):
        if not self.data.is_loaded:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите файл данных.")
            return

        dialog = StatisticsDialog(self.data, self)
        dialog.exec()

    def fill_features_list(self):
        """Заполняет сетку чекбоксов в 3 столбца + подсветка отключённых"""
        
        # 1. Очищаем старую сетку
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.check_boxes = []  # список всех QCheckBox

        features = self.data.get_column_names()
        if not features:
            lbl = QLabel("Нет загруженных признаков")
            self.grid_layout.addWidget(lbl, 0, 0, 1, 3, Qt.AlignCenter)
            return

        COLUMNS = 3  # фиксированное количество столбцов

        for i, col_name in enumerate(features):
            cb = QCheckBox(col_name)
            cb.setChecked(True)

            # ─── Подсветка и отключение проблемных признаков ───────────────
            is_invalid = False
            if hasattr(self.data, 'invalid_columns') and i in self.data.invalid_columns:
                is_invalid = True

            if is_invalid:
                cb.setChecked(False)
                cb.setEnabled(False)
                cb.setStyleSheet("""
                    color: #888888;
                    font-style: italic;
                """)

            row = i // COLUMNS
            col = i % COLUMNS

            self.grid_layout.addWidget(cb, row, col, Qt.AlignLeft | Qt.AlignVCenter)
            self.check_boxes.append(cb)

        # Растягиваем столбцы равномерно
        for c in range(COLUMNS):
            self.grid_layout.setColumnStretch(c, 1)

        # Добавляем пространство внизу
        self.grid_layout.setRowStretch(self.grid_layout.rowCount(), 1)

        # Сбрасываем скроллбар вверх
        self.scroll_area.verticalScrollBar().setValue(0)

    def get_selected_columns(self) -> list[int]:
        """Возвращает индексы выбранных и активных признаков"""
        return [i for i, cb in enumerate(self.check_boxes) if cb.isChecked() and cb.isEnabled()]


    def get_all_columns_count(self) -> int:
        """Количество всех признаков (включая отключённые)"""
        return len(self.check_boxes)

    def _create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")
        file_menu.addAction("Открыть данные...", self.act_open, shortcut="Ctrl+O")
        file_menu.addSeparator()
        file_menu.addAction("Выход", self.close, shortcut="Alt+F4")

        operation_menu = menubar.addMenu("Операции")
        operation_menu.addAction("Вычислить корреляции", self.act_run, shortcut="F9")

        view_menu = menubar.addMenu("Отчеты")
        
        view_menu.addAction("Отчёт корреляций расширенный", self.act_view_report_ext)
        view_menu.addAction("Отчёт статистики", self.act_view_stat_report)
        view_menu.addAction("Отчёт корреляций матричный", self.act_view_report_old)
        view_menu.addAction("Таблица результатов", self.act_view_result)

        # Новый раздел "Экспорт"
        export_menu = menubar.addMenu("Экспорт")
        export_menu.addAction("Экспорт расширенного отчёта в WORD", self.act_export_extended_report_to_word)
        #export_menu.addAction("Экспорт отчёта в WORD", self.act_export_report_to_word)

        save_menu = menubar.addMenu("Сохранить")
        save_menu.addAction("Сохранить таблицу результатов...", self.act_save_result)
        save_menu.addAction("Сохранить HTML-отчёт...", self.act_save_report)

    def act_open(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Открыть файл данных",
            "",
            "Данные (*.elnm *.txt *.dat);;Все файлы (*.*)"
        )
        if not fname:
            return

        self.stat_corr.clear()
        self.table_view.setModel(None)

        self.statusBar.showMessage("Загрузка файла... Подождите...")

        success = self.data.load_file(fname)

        if success:
            # Заполняем таблицу данных
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(self.data.get_column_names())

            for row_idx in range(self.data.get_count_record()):
                for col_idx, col_name in enumerate(self.data.get_column_names()):
                    val = self.data.df.iloc[row_idx, col_idx]
                    item = QStandardItem(f"{val:.4g}" if not pd.isna(val) else "—")
                    item.setEditable(False)
                    model.setItem(row_idx, col_idx, item)

            self.table_view.setModel(model)

            # Заполняем чекбоксы в левой панели
            self.fill_features_list()

            self.statusBar.showMessage(
                f"Файл загружен: {Path(fname).name}   •   строк: {self.data.get_count_record()}   •   признаков: {self.data.get_count_column()}"
            )
        else:
            QMessageBox.warning(self, "Ошибка загрузки",
                                "Не удалось загрузить файл.\nПодробности в data_load.log")
            self.statusBar.showMessage("Ошибка загрузки файла")

    def act_run(self):
        selected_cols = self.get_selected_columns()

        if len(selected_cols) < 2:
            QMessageBox.warning(self, "Предупреждение", "Выберите хотя бы два признака")
            return

        self.stat_corr.clear()
        self.stat_corr.initialize(self.data.get_column_names())
        self.stat_corr.invalid_columns = self.data.invalid_columns.copy()

        # Создаём все пары из выбранных столбцов
        for i in range(len(selected_cols) - 1):
            for j in range(i + 1, len(selected_cols)):
                self.stat_corr.add_or_get_pair(selected_cols[i], selected_cols[j])

        self.statusBar.showMessage("Расчёт корреляций...")

        # ← Вот здесь получаем текущее состояние чекбокса
        delphi_mode = self.delphi_compat_checkbox.isChecked()
        #Показываем пользователю, что используется
        mode_text = "Delphi-совместимый" if delphi_mode else "Правильный (scipy)"
        self.statusBar.showMessage(f"Запуск расчёта в режиме: {mode_text}")
        
        calculate_all_correlations(
            self.stat_corr,
            lambda col, rec: self.data.get_data(col, rec),
            self.data.get_count_record(),
            percent10=10,
            delphi_compatible=delphi_mode
        )

        QMessageBox.information(self, "Расчёт завершён",
                                f"Рассчитано {self.stat_corr.count()} пар корреляций.")
        self.statusBar.showMessage(f"Готово. Рассчитано {self.stat_corr.count()} пар.")

    def act_view_result(self):
        if self.stat_corr.count() == 0:
            QMessageBox.information(self, "Нет результатов",
                                    "Сначала выполните расчёт (F9).")
            return

        dialog = ResultsDialog(self.stat_corr, self)
        dialog.exec()

    def act_view_report_old(self):
        """Вариант отчета из старой версии"""
        if self.stat_corr.count() == 0:
            QMessageBox.information(self, "Нет результатов", "Сначала выполните расчёт.")
            return

        html_content = self._generate_old_report()

        report_path = Path("mapcor_report.html")
        report_path.write_text(html_content, encoding="utf-8")

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path.absolute())))

    def act_view_report_ext(self):
        if self.stat_corr.count() == 0:
            QMessageBox.information(self, "Нет результатов", "Сначала выполните расчёт (F9).")
            return

        html_content = self._generate_extended_report()

        report_path = Path("mapcor_extended_report.html")
        report_path.write_text(html_content, encoding="utf-8")

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path.absolute())))

        self.statusBar.showMessage("Расширенный отчёт открыт в браузере")

    def act_view_stat_report(self):
        """Генерирует и открывает геостатистический отчёт"""
        if self.data.df is None or self.data.df.empty:
            QMessageBox.warning(self, "Нет данных", "Сначала загрузите файл данных.")
            return


        # Получаем имена выбранных столбцов
        selected_cols = self.get_selected_columns()
        
        html = self._generate_stats_report(self.data, self.get_selected_columns())  #
        
        report_path = Path("geo_stats_report.html")
        report_path.write_text(html, encoding="utf-8")
        
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path.absolute())))
        
        self.statusBar.showMessage("Геостатистический отчёт открыт в браузере", 5000)


    def act_export_extended_report_to_word(self):
        html_content = self._generate_extended_report()
        
        # Модифицируем HTML для лучшей совместимости с Pandoc
        html_content = self._prepare_html_for_pandoc(html_content)
        
        fname, _ = QFileDialog.getSaveFileName(self, "Сохранить", "report.docx", "Word (*.docx)")
        if not fname: return
        if not fname.endswith('.docx'): fname += '.docx'
        
        temp_html = "temp_report.html"
        with open(temp_html, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        try:
            subprocess.run([
                "pandoc",
                temp_html,
                "-o", fname,
                "--from=html+raw_html",
                "--to=docx",
                "--standalone"
            ], check=True)
            
            os.remove(temp_html)
            QMessageBox.information(self, "Успех", f"Сохранено:\n{fname}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Pandoc не запустился:\n{e}\nУбедитесь, что Pandoc установлен.")

    def _prepare_html_for_pandoc(self, html_content):
        """
        Более агрессивный подход: заменяем стили на старые HTML-атрибуты
        """
        import re
        from bs4 import BeautifulSoup
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Обрабатываем все ячейки с цветным фоном
        for cell in soup.find_all(['td', 'th']):
            # Проверяем inline style
            if cell.has_attr('style'):
                style = cell['style']
                
                # Извлекаем hex-цвет
                match = re.search(r'background(?:-color)?:\s*#([0-9a-fA-F]{6})', style)
                if match:
                    hex_color = match.group(1)
                    
                    # Конвертируем в RGB для Word
                    r = int(hex_color[0:2], 16)
                    g = int(hex_color[2:4], 16)
                    b = int(hex_color[4:6], 16)
                    
                    # Используем bgcolor (старый HTML атрибут, который Pandoc понимает лучше)
                    cell['bgcolor'] = f"#{hex_color}"
                    
                    # Также оборачиваем содержимое в span с цветом
                    content = cell.decode_contents()
                    cell.clear()
                    new_span = soup.new_tag('span', style=f'background-color:#{hex_color};display:block;padding:2px;')
                    new_span.append(BeautifulSoup(content, 'html.parser'))
                    cell.append(new_span)
            
            # Обрабатываем классы
            if cell.has_attr('class'):
                classes = cell['class']
                
                if 'diag' in classes:
                    cell['bgcolor'] = '#e8e8e8'
                elif 'diag-header' in classes:
                    cell['bgcolor'] = '#b3e0ff'
                elif 'row-header' in classes:
                    cell['bgcolor'] = '#f0f4ff'
        
        return str(soup)

    def _generate_extended_report(self):
        """Генерация расширенного HTML-отчёта — версия с крупным названием и полусерым диагональным текстом"""
        
        # Шаг 1: Собираем выбранные валидные признаки
        selected_indices = []
        for i in range(len(self.check_boxes)):
            item = self.check_boxes[i]
            if item.checkState() == Qt.Checked:
                col_name = item.text().replace(" [invalid]", "")
                col_idx = self.data.get_number_for_column_name(col_name)
                if col_idx >= 0 and col_idx not in self.data.invalid_columns:
                    selected_indices.append(col_idx)

        if len(selected_indices) < 2:
            return "<html><body><h2>Ошибка: выберите хотя бы 2 валидных признака</h2></body></html>"

        BLOCK_SIZE = 8

        # Шаг 2: HTML-заголовок и стили (обновлены размеры и цвета для ch10-стиля)
        lines = [
            "<!DOCTYPE html>",
            "<html lang='ru'>",
            "<head>",
            "  <meta charset='UTF-8'>",
            "  <title>Отчет программы MapCor</title>",
            "  <style>",
            "    body {font-family: 'Segoe UI', Arial, sans-serif; margin: 32px; background: #f9f9f9; color: #222; font-size: 1.05em;}",
            "    h1 {color: #1a3c5e; text-align: center; font-size: 2.3em; margin-bottom: 0.5em;}",
            "    h2 {color: #2c5282; text-align: center; font-size: 1.7em; margin: 2em 0 0.8em;}",
            "    .feature-caption {text-align: center; margin: 1.8em 0 0.9em;}",
            "    .feature-caption .name {font-size: 2.4em; font-weight: bold; color: #0f2a6e; letter-spacing: -0.5px;}",
            "    .feature-caption .stats {font-size: 1.05em; font-weight: bold; color: #555; margin-left: 28px;}",
            "    table {border-collapse: collapse; margin: 0 auto 2.4em auto; width: auto; box-shadow: 0 2px 8px rgba(0,0,0,0.07);}",
            "    th, td {border: 1px solid #d0d0d0; padding: 11px 15px; text-align: center; font-size: 1.1em;;-webkit-print-color-adjust: exact; color-adjust: exact;}",
            "    th {background: #e8f0ff; color: #1e3a8a; font-weight: 600;}",
            "    td {font-weight: bold;}",
            "    .diag-header {",
            "      font-size: 1.25em !important;",
            "      font-weight: bold !important;",
            "      background: #b3e0ff !important;",
            "      border: 2px solid #80c0ff !important;",
            "      padding: 12px 16px !important;",
            "    }",
            "    .diag {",
            "      background: #e8e8e8 !important;",
            "      color: #B3B3B3 !important;",           # ← полусерый текст на диагонали
            "      font-style: italic;",
            "      font-weight: bold;",
            "      font-size: 1.0em !important;",        # ← тот же размер, что и остальные значения
            "    }",
            "    .na {color: #888; font-style: italic; font-weight: normal;}",
            "    .row-header {background: #f0f4ff; font-weight: bold; text-align: left; min-width: 60px;}",
            "    td.num {font-family: Consolas, 'Courier New', monospace;-webkit-print-color-adjust: exact; color-adjust: exact;}",
            "    hr {border: 0; height: 1px; background: #ddd; margin: 2.4em 0;}",
            "  </style>",
            "</head>",
            "<body>",
            "<h1>Корреляционный анализ</h1>",
            f"<p align='center' style='font-size:1.25em; margin-bottom:2.2em;'>",
            f"<b>Программа:</b> MapCor ;  ",
            f"<b>Файл:</b> {Path(self.data.filename).name} ;  ",
            f"<b>Число объектов:</b> {self.data.get_count_record()} ;  ",
            f"<b>Число характеристик:</b> {len(selected_indices)} ;  ",


            
        ]

        # Шаг 3: Общая статистика — первая
        lines.append('    <h2>Общая статистика по всем выбранным парам</h2>')
        lines.append('    <table style="width:72%; max-width:950px;">')
        lines.append('      <tr><th>Показатель</th><th>Минимум</th><th>Максимум</th><th>Среднее</th></tr>')
        corr_stat = self.stat_corr.all_pairs_stat['corr']
        lines.append(f'      <tr><td><b>R</b></td><td>{corr_stat.min:.3f}</td><td>{corr_stat.max:.3f}</td><td>{corr_stat.mean:.3f}</td></tr>')
        #dist10_stat = self.stat_corr.all_pairs_stat['dist10']
        #lines.append(f'      <tr><td><b>DIST_10</b></td><td>{dist10_stat.min:.1f}</td><td>{dist10_stat.max:.1f}</td><td>{dist10_stat.mean:.1f}</td></tr>')
        rr_stat = self.stat_corr.all_pairs_stat['rr']
        lines.append(f'      <tr><td><b>RR</b></td><td>{rr_stat.min:.3f}</td><td>{rr_stat.max:.3f}</td><td>{rr_stat.mean:.3f}</td></tr>')
        lines.append('    </table>')
        lines.append('<p style="text-align:center; color:#555; font-size:1.05em; margin: 0.8em 0 2em 0;">')
        lines.append('<b>R</b> — коэффициент ранговой корреляции<br> <b>RR</b> — корреляция корреляций')
        lines.append('</p>')

        #lines.append('    <hr>')

        # Шаг 4: Таблицы по каждой характеристике
        for feature_idx in selected_indices:
            feature_name = self.data.get_column_name(feature_idx)
            fs = self.stat_corr.feature_stats[feature_idx]

            # Одна строка над таблицей: крупное имя + маленькая статистика
            lines.append(
                '    <div class="feature-caption">'
                f'<span class="name">{feature_name}</span>  '
                f'<span class="stats">'
                f'M(R) = {fs.avg_corr:.3f} ;  '
                #f'M(DIST10) = {fs.avg_dist10:.1f} ;  '
                f'M(RR) = {fs.avg_rr:.3f}'
                f'</span>'
                '</div>'
            )

            # Блочные таблицы
            block_start = 0
            while block_start < len(selected_indices):
                block_end = min(block_start + BLOCK_SIZE - 1, len(selected_indices) - 1)

                lines.append('    <table>')
                lines.append('      <tr><th class="row-header"></th>')

                # Заголовки столбцов — диагональ выделена сильнее
                for j in range(block_start, block_end + 1):
                    col_name = self.data.get_column_name(selected_indices[j])
                    if selected_indices[j] == feature_idx:
                        lines.append(f'        <th class="diag-header">{col_name}</th>')
                    else:
                        lines.append(f'        <th>{col_name}</th>')
                lines.append('      </tr>')

                # R
                lines.append('      <tr><td class="row-header"><b>R</b></td>')
                for j in range(block_start, block_end + 1):
                    other_idx = selected_indices[j]
                    if feature_idx == other_idx:
                        lines.append('        <td class="diag">1.000</td>')
                    else:
                        pair_idx = self.stat_corr.get_pair_index(feature_idx, other_idx)
                        if pair_idx >= 0:
                            val = self.stat_corr.get_corr(pair_idx)
                            color = get_color_for_r(val)
                            lines.append(f'        <td style="background:{color};" class="num">{val:.3f}</td>')
                        else:
                            lines.append('        <td class="na">—</td>')
                lines.append('      </tr>')

                # DIST_10
               # lines.append('      <tr><td class="row-header"><b>DIST_10</b></td>')
               # for j in range(block_start, block_end + 1):
               #     other_idx = selected_indices[j]
               #     if feature_idx == other_idx:
               #         lines.append('        <td class="diag">100</td>')
               #     else:
               #         pair_idx = self.stat_corr.get_pair_index(feature_idx, other_idx)
               #         if pair_idx >= 0:
               #             val = self.stat_corr.get_dist10(pair_idx)
               #             color = get_color_for_dist10(val)
               #             lines.append(f'        <td style="background:{color};" class="num">{val:.1f}</td>')
               #         else:
               #             lines.append('        <td class="na">—</td>')
               # lines.append('      </tr>')

                # RR
                lines.append('      <tr><td class="row-header"><b>RR</b></td>')
                for j in range(block_start, block_end + 1):
                    other_idx = selected_indices[j]
                    if feature_idx == other_idx:
                        lines.append('        <td class="diag">—</td>')
                    else:
                        pair_idx = self.stat_corr.get_pair_index(feature_idx, other_idx)
                        if pair_idx >= 0:
                            val = self.stat_corr.get_rr(pair_idx)
                            color = get_color_for_rr(val)
                            lines.append(f'        <td style="background:{color};" class="num">{val:.3f}</td>')
                        else:
                            lines.append('        <td class="na">—</td>')
                lines.append('      </tr>')

                lines.append('    </table>')
                block_start = block_end + 1

            lines.append('    <hr>')

        lines.append('  </body></html>')
        return '\n'.join(lines)

    def _generate_old_report(self):
        """Генерация исходной версии HTML-отчёта — версия с крупным названием и полусерым диагональным текстом"""
        
        # Шаг 1: Собираем выбранные валидные признаки
        selected_indices = []
        for i in range(len(self.check_boxes)):
            item = self.check_boxes[i]
            if item.checkState() == Qt.Checked:
                col_name = item.text().replace(" [invalid]", "")
                col_idx = self.data.get_number_for_column_name(col_name)
                if col_idx >= 0 and col_idx not in self.data.invalid_columns:
                    selected_indices.append(col_idx)

        if len(selected_indices) < 2:
            return "<html><body><h2>Ошибка: выберите хотя бы 2 валидных признака</h2></body></html>"

        BLOCK_SIZE = 8

        # Шаг 2: HTML-заголовок и стили (обновлены размеры и цвета для ch10-стиля)
        lines = [
            "<!DOCTYPE html>",
            "<html lang='ru'>",
            "<head>",
            "  <meta charset='UTF-8'>",
            "  <title>Отчет программы MapCor</title>",
            "  <style>",
            "    body {font-family: 'Segoe UI', Arial, sans-serif; margin: 32px; background: #f9f9f9; color: #222; font-size: 1.05em;}",
            "    h1 {color: #1a3c5e; text-align: center; font-size: 2.3em; margin-bottom: 0.5em;}",
            "    h2 {color: #2c5282; text-align: center; font-size: 1.7em; margin: 2em 0 0.8em;}",
            "    .feature-caption {text-align: center; margin: 1.8em 0 0.9em;}",
            "    .feature-caption .name {font-size: 2.4em; font-weight: bold; color: #0f2a6e; letter-spacing: -0.5px;}",
            "    .feature-caption .stats {font-size: 1.05em; font-weight: bold; color: #555; margin-left: 28px;}",
            "    table {border-collapse: collapse; margin: 0 auto 2.4em auto; width: auto; box-shadow: 0 2px 8px rgba(0,0,0,0.07);}",
            "    th, td {border: 1px solid #d0d0d0; padding: 11px 15px; text-align: center; font-size: 1.1em;-webkit-print-color-adjust: exact; color-adjust: exact;}",
            "    th {background: #e8f0ff; color: #1e3a8a; font-weight: 600;}",
            "    td {font-weight: bold;}",
            "    .diag-header {",
            "      font-size: 1.25em !important;",
            "      font-weight: bold !important;",
            "      background: #b3e0ff !important;",
            "      border: 2px solid #80c0ff !important;",
            "      padding: 12px 16px !important;",
            "    }",
            "    .diag {",
            "      background: #e8e8e8 !important;",
            "      color: #B3B3B3 !important;",           # ← полусерый текст на диагонали
            "      font-style: italic;",
            "      font-weight: bold;",
            "      font-size: 1.0em !important;",        # ← тот же размер, что и остальные значения
            "    }",
            "    .na {color: #888; font-style: italic; font-weight: normal;}",
            "    .row-header {background: #f0f4ff; font-weight: bold; text-align: left; min-width: 150px;}",
            "    td.num {font-family: Consolas, 'Courier New', monospace;}",
            "    hr {border: 0; height: 1px; background: #ddd; margin: 2.4em 0;}",
            "  </style>",
            "</head>",
            "<body>",
            "<h1>Отчет программы MapCor</h1>",
            f"<p align='center' style='font-size:1.25em; margin-bottom:2.2em;'>",
            f"<b>Файл:</b> {Path(self.data.filename).name} ;  ",
            f"<b>Выбрано признаков:</b> {len(selected_indices)} ;  ",
            f"<b>Записей:</b> {self.data.get_count_record()} ;  ",
            f"<b>Количество пар:</b> {self.stat_corr.count()}",
            "<hr>"
        ]

        # Шаг 3: Общая статистика — первая
        lines.append('    <h2>Общая статистика по всем выбранным парам</h2>')
        lines.append('    <table style="width:72%; max-width:950px;">')
        lines.append('      <tr><th>Показатель</th><th>Минимум</th><th>Максимум</th><th>Среднее</th></tr>')
        corr_stat = self.stat_corr.all_pairs_stat['corr']
        lines.append(f'      <tr><td><b>Corr (R)</b></td><td>{corr_stat.min:.3f}</td><td>{corr_stat.max:.3f}</td><td>{corr_stat.mean:.3f}</td></tr>')
        dist10_stat = self.stat_corr.all_pairs_stat['dist10']
        lines.append(f'      <tr><td><b>DIST_10</b></td><td>{dist10_stat.min:.1f}</td><td>{dist10_stat.max:.1f}</td><td>{dist10_stat.mean:.1f}</td></tr>')
        rr_stat = self.stat_corr.all_pairs_stat['rr']
        lines.append(f'      <tr><td><b>RR</b></td><td>{rr_stat.min:.3f}</td><td>{rr_stat.max:.3f}</td><td>{rr_stat.mean:.3f}</td></tr>')
        lines.append('    </table>')
        lines.append('    <hr>')

        # Шаг 4: Таблицы по каждой характеристике
        # =============================================================================
        # Матрица корреляций + DIST10 (верхний треугольник = R, нижний = DIST10)
        # =============================================================================
        lines.append('<h2 style="margin-top: 3.5em; text-align: center;">Матрица корреляций Спирмена и DIST₁₀</h2>')
        lines.append('<p style="text-align:center; color:#555; font-size:1.05em; margin: 0.8em 0 2em 0;">')
        lines.append('Выше диагонали — <b>R (Spearman)</b><br>Ниже диагонали — <b>DIST₁₀</b>')
        lines.append('</p>')

        lines.append('<table style="margin: 0 auto 3em auto; border-collapse: collapse; box-shadow: 0 3px 12px rgba(0,0,0,0.08);">')

        # ── Заголовочная строка (имена признаков) ────────────────────────────────────
        lines.append('  <tr>')
        lines.append('    <th style="background:#e8f0ff; min-width:180px; font-weight:bold; padding:12px;"></th>')
        for col_idx in selected_indices:
            col_name = self.data.get_column_name(col_idx)
            lines.append(f'    <th title="{col_name}" style="background:#e8f0ff; padding:10px 14px;">{col_name}</th>')
        lines.append('  </tr>')

        # ── Строки матрицы ───────────────────────────────────────────────────────────
        for row_i, row_idx in enumerate(selected_indices):
            row_name = self.data.get_column_name(row_idx)
            lines.append('  <tr>')
            # Левый столбец — имена строк
            lines.append(f'    <td class="row-header" style="min-width:180px;">{row_name}</td>')

            for col_j, col_idx in enumerate(selected_indices):
                if row_i == col_j:
                    # Главная диагональ — всегда 1.000 (для корреляции)
                    lines.append('    <td class="diag" style="background:#d0e0ff; font-weight:bold; color:#1a3c5e;">1.000</td>')
                elif row_i < col_j:
                    # Выше диагонали → Spearman R
                    pair_idx = self.stat_corr.get_pair_index(row_idx, col_idx)
                    if pair_idx >= 0:
                        val = self.stat_corr.get_corr(pair_idx)
                        color = get_color_for_r(val)   # ← твоя функция цвета для R
                        lines.append(f'    <td style="background:{color};" class="num">{val:.3f}</td>')
                    else:
                        lines.append('    <td class="na">—</td>')
                else:
                    # Ниже диагонали → DIST10
                    pair_idx = self.stat_corr.get_pair_index(row_idx, col_idx)
                    if pair_idx >= 0:
                        val = self.stat_corr.get_dist10(pair_idx)
                        color = get_color_for_dist10(val)   # ← твоя функция цвета для DIST10
                        lines.append(f'    <td style="background:{color};" class="num">{val:.1f}</td>')
                    else:
                        lines.append('    <td class="na">—</td>')

            lines.append('  </tr>')

        lines.append('</table>')
        lines.append('<hr style="margin: 3em 0;">')

        lines.append('  </body></html>')
        return '\n'.join(lines)


    def _generate_stats_report(self, data, selected_columns=None):
        """
        Генерирует HTML-отчёт со статистикой выбранных характеристик.
        Добавлены: 5%, Q1, Q3, 95%, J (информативность по Шеннону, 6 интервалов)
        """
        import pandas as pd
        import numpy as np
        from pathlib import Path

        if data.df is None or data.df.empty:
            return "<h2>Нет данных для отчёта</h2>"

        # Фильтруем по выбранным индексам
        if selected_columns:
            valid_indices = [idx for idx in selected_columns if 0 <= idx < len(data.df.columns)]
            if not valid_indices:
                return "<h2>Нет выбранных числовых характеристик</h2>"
            df = data.df.iloc[:, valid_indices].select_dtypes(include=[np.number])
        else:
            df = data.df.select_dtypes(include=[np.number])

        if df.empty:
            return "<h2>Нет выбранных числовых характеристик</h2>"

        # Вычисляем статистику
        stats = pd.DataFrame(index=df.columns)

        stats['Min']    = df.min()
        stats['5%']     = df.quantile(0.05)
        stats['Q1']     = df.quantile(0.25)
        stats['Median'] = df.median()
        stats['Q3']     = df.quantile(0.75)
        stats['95%']    = df.quantile(0.95)
        stats['Max']    = df.max()
        stats['Range']  = stats['Max'] - stats['Min']
        stats['Mean']   = df.mean()
        stats['SD']     = df.std()
        stats['CV, %']  = (stats['SD'] / stats['Mean'] * 100).fillna(0)

        # ── J (Шеннон, нормированный, 6 интервалов) ────────────────────────────────
        def compute_J(series, n_bins=6):
            if len(series.dropna()) < 2:
                return np.nan
            hist, _ = np.histogram(series.dropna(), bins=n_bins)
            total = hist.sum()
            if total == 0:
                return np.nan
            pi = hist / total
            pi = pi[pi > 0]  # исключаем пустые интервалы
            if len(pi) == 0:
                return np.nan
            H = -np.sum(pi * np.log2(pi))
            J = 1 - H / np.log2(n_bins)
            return J

        stats['J'] = [compute_J(df[col]) for col in df.columns]

        # Порядок столбцов
        stat_order = [
            'Min', '5%', 'Q1', 'Median', 'Q3', '95%', 'Max', 'Range',
            'Mean', 'SD', 'CV, %', 'J'
        ]
        stats = stats[stat_order]

        stats.index.name = 'Признак'

        # ── HTML ────────────────────────────────────────────────────────────────────
        ROWS_PER_TABLE = 300
        num_rows = len(stats)  # количество признаков

        lines = [
            "<!DOCTYPE html>",
            "<html lang='ru'>",
            "<head>",
            "<meta charset='UTF-8'>",
            "<title>Статистический отчёт MapCor</title>",
            "<style>",
            "  body {font-family: Arial, sans-serif; line-height: 1.5; color: #333; max-width: 1400px; margin: 0 auto; padding: 20px;}",
            "  h1 {text-align: center; color: #1a3c5e; margin-bottom: 0.6em;}",
            "  h2 {color: #2c5282; border-bottom: 2px solid #e2e8f0; padding-bottom: 0.3em; margin: 1.8em 0 0.8em;}",
            "  table {border-collapse: collapse; margin: 1.2em auto 2.8em; width: 100%; max-width: 100%; box-shadow: 0 2px 10px rgba(0,0,0,0.08);}",
            "  th, td {border: 1px solid #d0d0d0; padding: 9px 12px; text-align: right; font-size: 0.98em; min-width: 90px;}",
            "  th {background: #e8f0ff; color: #1e3a8a; font-weight: 600; white-space: nowrap; vertical-align: bottom;}",
            "  .row-header {text-align: left; font-weight: bold; background: #f0f4ff; min-width: 160px; padding-left: 14px;}",
            "  .na {color: #777; font-style: italic;}",
            "  .small {font-size: 0.92em; color: #555;}",
            "  .percentile {background: #f9f9ff; font-weight: 500;}",
            "  .j-col {font-weight: bold; background: #fff7e6;}",
            "</style>",
            "</head>",
            "<body>",
            "<h1>Статистический отчёт по выбранным признакам</h1>",
            f"<p style='text-align: center; color: #555; font-size: 1.1em; margin-bottom: 2.2em;'>",
            f"Файл: <b>{Path(data.filename).name}</b> ;  ",
            f"Записей: <b>{data.get_count_record()}</b> ;  ",
            f"Признаков: <b>{num_rows}</b>",
            "</p>",
            "<hr style='border: 1px solid #aaa; margin: 2em 0;'>"
        ]

        # Разбиение на таблицы (если очень много признаков)
        row_groups = []
        start = 0
        while start < num_rows:
            end = min(start + ROWS_PER_TABLE, num_rows)
            row_groups.append(stats.iloc[start:end])
            start = end

        for tbl_idx, group_df in enumerate(row_groups, 1):
            lines.append("<table>")

            lines.append("  <tr>")
            lines.append("    <th class='row-header'>Признак</th>")
            for stat_name in group_df.columns:
                cls = ""
                if stat_name in ['5%', 'Q1', 'Q3', '95%']:
                    cls = " class='percentile'"
                if stat_name == 'J':
                    cls = " class='j-col'"
                display_name = stat_name
                lines.append(f"    <th{cls} title='{stat_name}'>{display_name}</th>")
            lines.append("  </tr>")

            for feature_name, row in group_df.iterrows():
                lines.append("  <tr>")
                lines.append(f"    <td class='row-header'>{feature_name}</td>")
                for stat_name, val in row.items():
                    if pd.isna(val):
                        val_str = "<span class='na'>—</span>"
                    elif stat_name in ['Min', '5%', 'Q1', 'Median', 'Q3', '95%', 'Max', 'Range', 'Mean']:
                        val_str = f"{val:.3f}"
                    elif stat_name == 'SD':
                        val_str = f"{val:.4f}"
                    elif stat_name == 'CV, %':
                        val_str = f"{val:.1f}"
                    elif stat_name == 'J':
                        val_str = f"{val:.3f}"
                        if val >= 0.65:
                            val_str = f"<b>{val_str}</b>"
                    else:
                        val_str = f"{val:.3f}"
                    lines.append(f"    <td>{val_str}</td>")
                lines.append("  </tr>")

            lines.append("</table>")

        lines.append("<hr style='margin: 3em 0 1.5em;'>")

        lines.append("<h2 style='text-align:center; color:#2c5282; margin-bottom:1em;'>Расшифровка показателей</h2>")
        lines.append("<ul style='max-width:960px; margin:0 auto 2em; font-size:0.98em; line-height:1.7; list-style:none; padding-left:0;'>")
        lines.append("  <li><strong>Min / Max</strong> — минимальное и максимальное значение</li>")
        lines.append("  <li><strong>5% / 95%</strong> — 5-й и 95-й перцентили</li>")
        lines.append("  <li><strong>Q1 / Q3</strong> — первый и третий квартили (25% и 75%)</li>")
        lines.append("  <li><strong>Median</strong> — медиана (50%)</li>")
        lines.append("  <li><strong>Range</strong> — размах (Max − Min)</li>")
        lines.append("  <li><strong>Mean</strong> — арифметическое среднее</li>")
        lines.append("  <li><strong>SD</strong> — стандартное отклонение</li>")
        lines.append("  <li><strong>CV, %</strong> — коэффициент вариации (SD / Mean × 100)</li>")
        lines.append("  <li><strong>J</strong> — нормированная информативность по Шеннону (6 интервалов). J ∈ [0;1]</li>")
        lines.append("  <li style='margin-left:2em;'>J → 0 — полная гетерогенность (равномерное распределение)</li>")
        lines.append("  <li style='margin-left:2em;'>J → 1 — монолитный пласт (все значения в одном интервале)</li>")
        lines.append("  <li style='margin-left:2em;'>Рекомендуемый порог однородности: <b>J ≥ 0.65</b></li>")
        lines.append("</ul>")

        lines.append("</body>")
        lines.append("</html>")

        return '\n'.join(lines)

    def act_save_result(self):
        if self.stat_corr.count() == 0:
            QMessageBox.information(self, "Нет данных", "Нет рассчитанных результатов для сохранения.")
            return

        # Диалог сохранения — по умолчанию .txt
        fname, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить таблицу результатов",
            "results.txt",
            "Текстовые файлы (*.txt);;Все файлы (*.*)"
        )
        if not fname:
            return

        # Добавляем расширение .txt, если пользователь его не указал
        if not fname.lower().endswith('.txt'):
            fname += '.txt'

        try:
            with open(fname, "w", encoding="utf-8") as f:
                        # 1. Первая строка — количество пар (а не количество признаков!)
                        f.write(f"4\n")

                        # 2. Заголовок таблицы
                        f.write("n\tname\tR\tRR\n")

                        # 3. Данные по всем парам
                        for i in range(self.stat_corr.count()):
                            pair_name = self.stat_corr.get_pair_name(i)
                            r_value   = self.stat_corr.get_corr(i)
                            rr_value  = self.stat_corr.get_rr(i)

                            # Порядковый номер начиная с 1
                            num = i + 1

                            # Форматирование значений с учётом возможных NaN
                            r_str  = f"{r_value:.3f}"  if not np.isnan(r_value)  else "—"
                            rr_str = f"{rr_value:.3f}" if not np.isnan(rr_value) else "—"

                            line = f"{num}\t{pair_name}\t{r_str}\t{rr_str}"
                            f.write(line + "\n")

            QMessageBox.information(self, "Сохранено", f"Результаты сохранены в файл:\n{fname}")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка сохранения",
                                f"Не удалось сохранить файл:\n{str(e)}")
            
            
    def act_save_report(self):
        fname, _ = QFileDialog.getSaveFileName(
            self, "Сохранить HTML-отчёт", "mapcor_report.html", "HTML (*.html);;Все файлы (*.*)"
        )
        if not fname:
            return

        html = self._generate_simple_html_report()
        Path(fname).write_text(html, encoding="utf-8")

        QMessageBox.information(self, "Сохранено", f"Отчёт сохранён в {fname}")


if __name__ == "__main__":
    app = QApplication(sys.argv)

# Самый современный и чистый вид на Windows 10/11
    app.setStyle("Fusion")

    # Лёгкая тёмная/светлая тема с хорошей читаемостью
    app.setStyleSheet("""
        QMainWindow {
            background-color: #f8f9fa;
        }
        QGroupBox {
            font-weight: bold;
            border: 1px solid #ced4da;
            border-radius: 4px;
            margin-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px;
        }
        QTableView {
            gridline-color: #dee2e6;
            alternate-background-color: #f1f3f5;
        }
        QHeaderView::section {
            background-color: #e9ecef;
            padding: 6px;
            border: 1px solid #ced4da;
        }
        QListWidget {
            border: 1px solid #ced4da;
            border-radius: 4px;
            background-color: white;
        }
        QStatusBar {
            background-color: #e9ecef;
            color: #495057;
        }
    """)

    # Шрифт (очень важно для профессионального вида)
    font = QFont("Segoe UI", 12)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())