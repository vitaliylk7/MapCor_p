# main.py
"""
MAPCOR_M — Python-версия программы для расчёта ранговых корреляций
(перенос с Delphi → PySide6 + pandas + numpy + scipy)

Основные функции:
- Загрузка данных (txt/csv/xlsx)
- Выбор признаков через чек-лист (QListWidget с галочками)
- Расчёт Spearman, DIST50, DIST10, RR (мета-корреляция)
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

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QListWidget, QListWidgetItem, QTableView,
    QStatusBar, QMenuBar, QFileDialog, QMessageBox,
    QHeaderView, QDialog, QLabel, QPushButton, QTextEdit
)
from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QStandardItemModel, QStandardItem, QDesktopServices, QFont

from data import TData
from stat_corr_types import TStatCorr
from corr_calculations import calculate_all_correlations


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
        headers = ["Пара", "Spearman R", "DIST_50", "DIST_10", "RR (мета-корр)"]
        model.setHorizontalHeaderLabels(headers)

        for i in range(self.stat_corr.count()):
            pair_name = self.stat_corr.get_pair_name(i)
            r = self.stat_corr.get_corr(i)
            d50 = self.stat_corr.get_dist50(i)
            d10 = self.stat_corr.get_dist10(i)
            rr = self.stat_corr.get_rr(i)

            row = [
                QStandardItem(pair_name),
                QStandardItem(f"{r:.3f}" if not np.isnan(r) else "—"),
                QStandardItem(f"{d50:.1f}" if not np.isnan(d50) else "—"),
                QStandardItem(f"{d10:.1f}" if not np.isnan(d10) else "—"),
                QStandardItem(f"{rr:.3f}" if not np.isnan(rr) else "—")
            ]

            model.appendRow(row)

        self.table.setModel(model)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MAPCOR_M — PySide6 версия")
        self.resize(1280, 800)

        self.data = TData()
        self.stat_corr = TStatCorr()

        # Центральный виджет
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)

        # Левая панель — выбор признаков
        left_group = QGroupBox("Выбор характеристик")
        left_layout = QVBoxLayout(left_group)
        self.check_list = QListWidget()
        self.check_list.setAlternatingRowColors(True)
        self.check_list.setSelectionMode(QListWidget.NoSelection)
        left_layout.addWidget(self.check_list)

        main_layout.addWidget(left_group, 1)

        # Правая панель — исходные данные
        right_group = QGroupBox("Исходные данные")
        right_layout = QVBoxLayout(right_group)
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.verticalHeader().setVisible(False)
        right_layout.addWidget(self.table_view)

        main_layout.addWidget(right_group, 3)

        # Статусбар
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Готов к открытию файла...")

        # Создаём меню
        self._create_menu()

    def _create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")
        file_menu.addAction("Открыть данные...", self.act_open, shortcut="Ctrl+O")
        file_menu.addSeparator()
        file_menu.addAction("Выход", self.close, shortcut="Alt+F4")

        operation_menu = menubar.addMenu("Операции")
        operation_menu.addAction("Вычислить корреляции", self.act_run, shortcut="F9")

        view_menu = menubar.addMenu("Результаты")
        view_menu.addAction("Таблица результатов", self.act_view_result)
        view_menu.addAction("HTML-отчёт", self.act_view_report)

        save_menu = menubar.addMenu("Сохранить")
        save_menu.addAction("Сохранить таблицу результатов...", self.act_save_result)
        save_menu.addAction("Сохранить HTML-отчёт...", self.act_save_report)

    def act_open(self):
        fname, _ = QFileDialog.getOpenFileName(
            self, "Открыть файл данных",
            "",
            "Данные (*.csv *.txt *.xlsx);;Все файлы (*.*)"
        )
        if not fname:
            return

        self.stat_corr.clear()
        self.check_list.clear()
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

            # Заполняем чек-лист
            self.check_list.clear()
            for col_name in self.data.get_column_names():
                item = QListWidgetItem(col_name)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)

                col_idx = self.data.get_number_for_column_name(col_name)
                if col_idx in self.data.invalid_columns:
                    item.setCheckState(Qt.Unchecked)
                    item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                    item.setText(f"{col_name} [invalid]")

                self.check_list.addItem(item)

            self.statusBar.showMessage(
                f"Файл загружен: {Path(fname).name}   •   строк: {self.data.get_count_record()}   •   признаков: {self.data.get_count_column()}"
            )
        else:
            QMessageBox.warning(self, "Ошибка загрузки",
                                "Не удалось загрузить файл.\nПодробности в data_load.log")
            self.statusBar.showMessage("Ошибка загрузки файла")

    def act_run(self):
        selected_cols = []
        for i in range(self.check_list.count()):
            item = self.check_list.item(i)
            if item.checkState() == Qt.Checked:
                col_name = item.text().replace(" [invalid]", "")
                col_idx = self.data.get_number_for_column_name(col_name)
                if col_idx >= 0 and col_idx not in self.data.invalid_columns:
                    selected_cols.append(col_idx)

        if len(selected_cols) < 2:
            QMessageBox.warning(self, "Недостаточно признаков",
                                "Выберите хотя бы два валидных признака.")
            return

        self.stat_corr.clear()
        self.stat_corr.initialize(self.data.get_column_names())
        self.stat_corr.invalid_columns = self.data.invalid_columns.copy()

        # Создаём все пары из выбранных столбцов
        for i in range(len(selected_cols) - 1):
            for j in range(i + 1, len(selected_cols)):
                self.stat_corr.add_or_get_pair(selected_cols[i], selected_cols[j])

        self.statusBar.showMessage("Расчёт корреляций...")

        calculate_all_correlations(
            self.stat_corr,
            lambda col, rec: self.data.get_data(col, rec),
            self.data.get_count_record(),
            percent50=50,
            percent10=10,
            log_scale=False  # можно сделать параметр из GUI
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

    def act_view_report(self):
        """Простая заглушка — можно расширить до полноценного HTML-отчёта"""
        if self.stat_corr.count() == 0:
            QMessageBox.information(self, "Нет результатов", "Сначала выполните расчёт.")
            return

        html_content = self._generate_simple_html_report()

        report_path = Path("mapcor_report.html")
        report_path.write_text(html_content, encoding="utf-8")

        QDesktopServices.openUrl(QUrl.fromLocalFile(str(report_path.absolute())))

    def _generate_simple_html_report(self):
        """Пример простого HTML-отчёта — можно сильно улучшить"""
        lines = [
            "<!DOCTYPE html>",
            "<html lang='ru'>",
            "<head><meta charset='UTF-8'>",
            "<title>MAPCOR — Результаты</title>",
            "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:30px;}",
            "table{border-collapse:collapse;width:100%;margin:20px 0;}",
            "th,td{border:1px solid #ccc;padding:8px;text-align:center;}",
            "th{background:#e6f2ff;}</style></head>",
            "<body>",
            f"<h2>Результаты MAPCOR — {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}</h2>",
            f"<p>Файл: {Path(self.data.filename).name}<br>",
            f"Строк: {self.data.get_count_record()} • Признаков: {self.data.get_count_column()}</p>",
            "<h3>Таблица корреляций</h3>",
            "<table><tr><th>Пара</th><th>Spearman R</th><th>DIST50</th><th>DIST10</th><th>RR</th></tr>"
        ]

        for i in range(self.stat_corr.count()):
            pair = self.stat_corr.get_pair_name(i)
            r = self.stat_corr.get_corr(i)
            d50 = self.stat_corr.get_dist50(i)
            d10 = self.stat_corr.get_dist10(i)
            rr = self.stat_corr.get_rr(i)
            lines.append(f"<tr><td>{pair}</td>"
                         f"<td>{r:.3f}</td><td>{d50:.1f}</td><td>{d10:.1f}</td><td>{rr:.3f}</td></tr>")

        lines.extend(["</table>", "</body></html>"])
        return "\n".join(lines)

    def act_save_result(self):
        if self.stat_corr.count() == 0:
            QMessageBox.information(self, "Нет данных", "Нет результатов для сохранения.")
            return

        fname, _ = QFileDialog.getSaveFileName(
            self, "Сохранить таблицу результатов", "results.csv", "CSV (*.csv);;Все файлы (*.*)"
        )
        if not fname:
            return

        with open(fname, "w", encoding="utf-8") as f:
            f.write("Pair;SpearmanR;DIST50;DIST10;RR\n")
            for i in range(self.stat_corr.count()):
                f.write(f"{self.stat_corr.get_pair_name(i)};"
                        f"{self.stat_corr.get_corr(i):.3f};"
                        f"{self.stat_corr.get_dist50(i):.1f};"
                        f"{self.stat_corr.get_dist10(i):.1f};"
                        f"{self.stat_corr.get_rr(i):.3f}\n")

        QMessageBox.information(self, "Сохранено", f"Таблица сохранена в {fname}")

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

    # Можно задать стиль оформления (опционально)
    # app.setStyle("Fusion")           # или "Windows"
    # app.setStyleSheet("...")         # тёмная тема и т.д.

    window = MainWindow()
    window.show()
    sys.exit(app.exec())