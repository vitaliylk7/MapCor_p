# main.py
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser
import os
import datetime
from data import TData
from stat_corr_types import TStatCorr
from corr_calculations import calculate_all_correlations

class FrmMain(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MAPCOR_M")
        self.geometry("1200x600")
        self.data = TData()
        self.stat_corr = TStatCorr()

        # Группа для параметров (выбор признаков)
        group1 = tk.LabelFrame(self, text="Параметры")
        group1.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.ch_criterion = tk.Listbox(group1, selectmode=tk.MULTIPLE, height=20)
        self.ch_criterion.pack(fill=tk.BOTH, expand=True)

        # Контекстное меню для чеклиста
        self.pm_criterions = tk.Menu(self, tearoff=0)
        self.pm_criterions.add_command(label="Отметить все", command=self.select_all)
        self.pm_criterions.add_command(label="Снять все отметки", command=self.deselect_all)
        self.pm_criterions.add_command(label="Инвертировать отметки", command=self.invert_selection)
        self.ch_criterion.bind("<Button-3>", self.popup_menu)

        # Группа для данных (таблица)
        group2 = tk.LabelFrame(self, text="Исходные данные")
        group2.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.sg = ttk.Treeview(group2, show="headings")
        self.sg.pack(fill=tk.BOTH, expand=True)

        # Статусбар
        self.sb_main = tk.Label(self, text="", relief=tk.SUNKEN, anchor=tk.W)
        self.sb_main.pack(side=tk.BOTTOM, fill=tk.X)

        # Меню
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть", command=self.act_open)
        file_menu.add_command(label="Выход", command=self.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)

        operation_menu = tk.Menu(menubar, tearoff=0)
        operation_menu.add_command(label="Вычислить", command=self.act_run)
        menubar.add_cascade(label="Операции", menu=operation_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Таблица", command=self.act_view_result)
        view_menu.add_command(label="Отчет", command=self.act_view_report)
        menubar.add_cascade(label="Результаты", menu=view_menu)

        save_menu = tk.Menu(menubar, tearoff=0)
        save_menu.add_command(label="Сохранить таблицу", command=self.act_save_result)
        save_menu.add_command(label="Сохранить отчет", command=self.act_save_report)
        menubar.add_cascade(label="Сохранить", menu=save_menu)

        self.config(menu=menubar)

    def popup_menu(self, event):
        self.pm_criterions.post(event.x_root, event.y_root)

    def select_all(self):
        self.ch_criterion.selection_set(0, tk.END)

    def deselect_all(self):
        self.ch_criterion.selection_clear(0, tk.END)

    def invert_selection(self):
        selected = self.ch_criterion.curselection()
        all_items = range(self.ch_criterion.size())
        for i in all_items:
            if i in selected:
                self.ch_criterion.selection_clear(i)
            else:
                self.ch_criterion.selection_set(i)

    def act_open(self):
        fname = filedialog.askopenfilename(filetypes=[("CSV/TXT files", "*.csv *.txt")])
        if not fname:
            return

        self.stat_corr.clear()
        # Очистка GUI
        for item in self.sg.get_children():
            self.sg.delete(item)
        self.ch_criterion.delete(0, tk.END)
        self.sb_main['text'] = "Загрузка файла..."

        if self.data.load_file(fname):
            # Заполнение таблицы
            columns = self.data.get_column_names()
            self.sg["columns"] = columns
            for col in columns:
                self.sg.heading(col, text=col)
            for i in range(self.data.get_count_record()):
                row = [self.data.df.iloc[i][col] for col in columns]
                self.sg.insert("", tk.END, values=row)

            # Заполнение чеклиста
            for name in columns:
                self.ch_criterion.insert(tk.END, name)
            self.select_all()  # Все отмечены по умолчанию

            self.sb_main['text'] = f"Файл: {os.path.basename(fname)}   n = {self.data.get_count_record()}"
            # после загрузки, для invalid столбцов disable чекбоксы
            for col_name in columns:
                var = tk.IntVar(value=1)
                self.check_vars[col_name] = var
                chk = tk.Checkbutton(self.check_frame, text=col_name, variable=var, anchor="w", justify="left")
                col_idx = self.data.get_number_for_column_name(col_name)
                if col_idx in self.data.invalid_columns:
                    chk.config(state="disabled", text=f"{col_name} [invalid]")
                    var.set(0)
                chk.pack(fill=tk.X, padx=5, pady=1)
                self.check_buttons[col_name] = chk

    def act_run(self):
        self.stat_corr.invalid_columns = self.data.invalid_columns.copy()
        selected = self.ch_criterion.curselection()
        if len(selected) < 2:
            messagebox.showwarning("Ошибка", "Выберите хотя бы два признака")
            return

        selected_cols = list(selected)

        self.stat_corr.clear()
        self.stat_corr.initialize(self.data.get_column_names())

        for i in range(len(selected_cols) - 1):
            for j in range(i + 1, len(selected_cols)):
                self.stat_corr.add_or_get_pair(selected_cols[i], selected_cols[j])

        calculate_all_correlations(self.stat_corr, self.data.get_data, self.data.get_count_record(), 50, 10, False)

        messagebox.showinfo("Готово", f"Рассчитано {self.stat_corr.count()} пар")
        self.sb_main['text'] = f"Рассчитано {self.stat_corr.count()} пар признаков"

    def act_view_result(self):
        # Показ таблицы результатов (аналог ViewResult)
        result_win = tk.Toplevel(self)
        result_win.title("Таблица результатов")
        tree = ttk.Treeview(result_win, columns=("Pair", "Corr", "DIST50", "DIST10", "RR"))
        tree.pack(fill=tk.BOTH, expand=True)
        tree.heading("Pair", text="Pair")
        tree.heading("Corr", text="Corr")
        tree.heading("DIST50", text="DIST50")
        tree.heading("DIST10", text="DIST10")
        tree.heading("RR", text="RR")

        for i in range(self.stat_corr.count()):
            pair_name = self.stat_corr.get_pair_name(i)
            corr = self.stat_corr.get_corr(i)
            d50 = self.stat_corr.get_dist50(i)
            d10 = self.stat_corr.get_dist10(i)
            rr = self.stat_corr.get_rr(i)
            tree.insert("", tk.END, values=(pair_name, f"{corr:.3f}", f"{d50:.1f}", f"{d10:.1f}", f"{rr:.3f}"))

    def act_view_report(self):
        self.generate_extended_report()

    def act_save_result(self):
        fname = filedialog.asksaveasfilename(defaultextension=".csv")
        if not fname:
            return
        with open(fname, 'w') as f:
            f.write("Pair,Corr,DIST50,DIST10,RR\n")
            for i in range(self.stat_corr.count()):
                pair_name = self.stat_corr.get_pair_name(i)
                corr = self.stat_corr.get_corr(i)
                d50 = self.stat_corr.get_dist50(i)
                d10 = self.stat_corr.get_dist10(i)
                rr = self.stat_corr.get_rr(i)
                f.write(f"{pair_name},{corr:.3f},{d50:.1f},{d10:.1f},{rr:.3f}\n")

    def act_save_report(self):
        fname = filedialog.asksaveasfilename(defaultextension=".html")
        if not fname:
            return
        html = self.generate_extended_report(save=True)
        with open(fname, 'w', encoding='utf-8') as f:
            f.write(html)

    def generate_extended_report(self, save=False):
        # Генерация HTML (аналог GenerateExtendedReport)
        selected = self.ch_criterion.curselection()
        if len(selected) < 2:
            messagebox.showwarning("Ошибка", "Выберите хотя бы два признака")
            return

        selected_indices = list(selected)

        html_lines = [
            '<!DOCTYPE html>',
            '<html lang="ru">',
            '<head><meta charset="UTF-8"><title>Отчёт MAPCOR</title>',
            # Стили (скопировать из Delphi, упрощённо)
            '<style>body {font-family: Arial;} table {border-collapse: collapse;} th, td {border: 1px solid black; padding: 5px;}</style>',
            '</head><body>'
        ]

        html_lines.append(f"<h2>Расширенный отчёт</h2>")
        html_lines.append(f"<p>Файл: {os.path.basename(self.data.get_file_name())}<br>Признаков: {len(selected_indices)}<br>Записей: {self.data.get_count_record()}<br>Дата: {datetime.datetime.now()}</p>")

        # Для каждого признака
        BLOCK_SIZE = 10  # Пример
        for i, feature_idx in enumerate(selected_indices):
            html_lines.append(f"<h3>{self.data.get_column_name(feature_idx)}</h3>")
            fs = self.stat_corr.feature_stats[feature_idx]
            html_lines.append(f"M(R) = {fs.avg_corr:.3f} • M(DIST10) = {fs.avg_dist10:.1f} • M(RR) = {fs.avg_rr:.3f}")

            # Блоки таблиц
            block_start = 0
            while block_start < len(selected_indices):
                block_end = min(block_start + BLOCK_SIZE, len(selected_indices) - 1)
                html_lines.append(f"<table><tr><th></th>")
                for j in range(block_start, block_end + 1):
                    html_lines.append(f"<th>{self.data.get_column_name(selected_indices[j])}</th>")
                html_lines.append("</tr>")

                # R, DIST10, RR строки (аналогично Delphi, с цветами)
                # ... (добавьте логику GetColorForCorr и т.д., упрощённо)
                # Пример для R:
                html_lines.append("<tr><td>R</td>")
                for j in range(block_start, block_end + 1):
                    other_idx = selected_indices[j]
                    if feature_idx == other_idx:
                        html_lines.append("<td>1.000</td>")
                    else:
                        pair_idx = self.stat_corr.get_pair_index(feature_idx, other_idx)
                        if pair_idx >= 0:
                            val = self.stat_corr.get_corr(pair_idx)
                            color = "ffffff"  # Замените на GetColorForCorr
                            html_lines.append(f'<td style="background:#{color}">{val:.3f}</td>')
                        else:
                            html_lines.append("<td>—</td>")
                html_lines.append("</tr>")

                # Аналогично для DIST10 и RR

                html_lines.append("</table>")
                block_start = block_end + 1

        # Общая статистика
        html_lines.append("<h2>Общая статистика</h2><table>")
        # Добавьте строки для Corr, DIST10, RR из all_pairs_stat
        html_lines.append("</table></body></html>")

        html = "\n".join(html_lines)

        if save:
            return html

        html_file = "report.html"
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html)
        webbrowser.open(html_file)

if __name__ == "__main__":
    app = FrmMain()
    app.mainloop()