import openpyxl
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import warnings

warnings.filterwarnings('ignore', message='Data Validation extension is not supported')
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


class PayrollCalculator:
    """Калькулятор расчета фонда оплаты труда (ФОТ)"""

    def __init__(self):
        self.employees_data = []
        self.original_employees_data = []
        self.modified_employees = {}
        self.reference_data = {}
        self.motivation_map = {}
        self.changes_log = []

        self.hypotheses_data = {
            'working_hours': {}, 'staffing': {}, 'tariff_coefficients': {},
            'field_coefficients': {}, 'additional_coefficients': {}, 'bonus_payment': {},
            'night_share': 0.267
        }

        self.months = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
                       'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
        self.months_short = ['янв', 'фев', 'мар', 'апр', 'май', 'июн',
                             'июл', 'авг', 'сен', 'окт', 'ноя', 'дек']
        self.year = 2026

        self.total_res_a_excel = 0.0

    def _parse_number(self, value):
        """Точный парсинг чисел из Excel"""
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return 0.0 if np.isnan(value) or np.isinf(value) else float(value)

        s = str(value).strip()
        if not s:
            return 0.0

        is_percent = s.endswith('%')
        if is_percent:
            s = s[:-1].strip()

        for ch in ('\u00a0', '\u202f', '\u2009', '\u2007', ' '):
            s = s.replace(ch, '')

        s = s.replace('#REF!', '0').replace('ERROR', '0')
        if not s or s in ('-', '—'):
            return 0.0

        try:
            if ',' in s and '.' not in s:
                left, _, right = s.rpartition(',')
                if right.isdigit() and len(right) <= 2 and left.replace('.', '').replace('-', '').isdigit():
                    s = left + '.' + right
                else:
                    s = s.replace(',', '')
            elif ',' in s and '.' in s:
                s = s.replace('.', '').replace(',', '.') if s.rfind(',') > s.rfind('.') else s.replace(',', '')
            return float(s) / 100.0 if is_percent else float(s)
        except (TypeError, ValueError):
            return 0.0

    def load_employee_data(self, filepath):
        """Загрузка данных из Excel"""
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            if 'Справочники' in wb.sheetnames:
                self.reference_data = self._parse_directory_sheet(wb['Справочники'])

            employee_sheet = next((s for s in wb.sheetnames if s.upper() == 'ФОРМАТ'), wb.sheetnames[0])
            ws = wb[employee_sheet]
            self._parse_employee_sheet(ws)
            self.original_employees_data = self.employees_data.copy()
            self.modified_employees = {}
            self.changes_log = []

            ws_hyp = next((s for s in wb.sheetnames if 'Гипотез' in s), None)
            if ws_hyp:
                self.hypotheses_data.update(self._parse_hypotheses_sheet(wb[ws_hyp]))

            return self
        except Exception as e:
            raise Exception(f"Ошибка загрузки файла: {str(e)}")

    def _parse_directory_sheet(self, ws):
        """Парсинг справочников"""
        directories = {'legal_entities': {}, 'cities': {}}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            d = dict(zip([str(c) if c else '' for c in ws[1]], row))
            motiv_key = str(d.get('Мес. Мотивация', '')).strip()
            motiv_val = self._parse_number(d.get('Премия vs Оклад', 0.0))
            if motiv_key:
                self.motivation_map[motiv_key] = motiv_val
            if d.get('Юрлицо'):
                directories['legal_entities'][d['Юрлицо']] = {'sv_group': d.get('Группа для СВ', 'А')}
            if d.get('Город базирования'):
                directories['cities'][d['Город базирования']] = {'schedule': d.get('График работы', '')}
        return directories

    def _parse_employee_sheet(self, ws):
        """Парсинг сотрудников с учетом прямых значений СВ"""
        employees = []
        current_dir = ''
        self.total_res_a_excel = 0.0

        for row in ws.iter_rows(min_row=6, values_only=True):
            if not row or not any(row):
                continue

            col_a = row[0] if len(row) > 0 else None
            if col_a and 'ДИРЕКЦИЯ' in str(col_a).upper():
                current_dir = str(col_a).strip()
                if not (len(row) > 8 and row[8] and str(row[8]).strip()):
                    continue

            fio = row[8] if len(row) > 8 else None
            position = row[7] if len(row) > 7 else (row[8] if len(row) > 8 else None)
            if not fio or str(fio).strip() == '':
                continue
            fio_str = str(fio).strip().upper()
            if any(x in fio_str for x in ['ФИО', 'ДОЛЖНОСТЬ', 'ДИРЕКЦИЯ', 'СТОЛБЕЦ', 'ИТОГО']):
                continue

            # ✅ ЧТЕНИЕ ОКЛАДОВ ИЗ BJ-BU (ИНДЕКСЫ 61-72)
            base_salary_monthly = [self._parse_number(row[i]) if len(row) > i else 0.0 for i in range(61, 73)]
            if all(v == 0.0 for v in base_salary_monthly):
                base_sal = self._parse_number(row[14]) if len(row) > 14 else 0.0
                if base_sal == 0.0 and len(row) > 15:
                    base_sal = self._parse_number(row[15])
                base_salary_monthly = [base_sal] * 12

            motiv_raw = str(row[39]).strip() if len(row) > 39 else ''
            monthly_bonus_pct = self.motivation_map.get(motiv_raw,
                                                        self._parse_number(row[39]) if len(row) > 39 else 0.0)

            # ✅ Численность из AV-BG (индексы 47-58)
            headcounts = [self._parse_number(row[i]) if len(row) > i else 1.0 for i in range(47, 59)]
            headcounts = [v if v > 0 else 1.0 for v in headcounts]

            compensation_summer = self._parse_number(row[37]) if len(row) > 37 else 0.0
            compensation_winter = self._parse_number(row[38]) if len(row) > 38 else 0.0

            # ✅ ЧТЕНИЕ СВ НАПРЯМУЮ ИЗ ЯЧЕЕК (CZ-DK, индексы 103-114)
            sv_excel = [round(self._parse_number(row[i]), 2) for i in range(103, 115)]
            sv_total_excel = self._parse_number(row[115]) if len(row) > 115 else 0.0

            comp_excel = [self._parse_number(row[i]) if len(row) > i else 0.0 for i in range(116, 129)]
            comp_total_excel = self._parse_number(row[129]) if len(row) > 129 else 0.0

            # ✅ Резервы квартальных из EA-EL (индексы 130-141)
            res_q_excel = [self._parse_number(row[i]) if len(row) > i else 0.0 for i in range(130, 142)]
            res_q_total_excel = self._parse_number(row[142]) if len(row) > 142 else 0.0

            # ✅ Резервы годовых из EN-EY (индексы 143-154)
            res_a_excel = [self._parse_number(row[i]) if len(row) > i else 0.0 for i in range(143, 155)]
            res_a_total_excel = self._parse_number(row[155]) if len(row) > 155 else 0.0

            self.total_res_a_excel += res_a_total_excel

            emp = {
                'directorate': current_dir, 'fio': str(fio).strip(),
                'position': str(position).strip() if position else '',
                'base_salary': base_salary_monthly[0] if base_salary_monthly else 0.0,
                'base_salary_monthly': base_salary_monthly,
                'allowance_cat1': self._parse_number(row[16]) if len(row) > 16 else 0.0,
                'allowance_cat2': self._parse_number(row[17]) if len(row) > 17 else 0.0,
                'travel_allowance': self._parse_number(row[18]) if len(row) > 18 else 0.0,
                'rk_percent': self._parse_number(row[19]) if len(row) > 19 else 0.0,
                'sn_percent': self._parse_number(row[20]) if len(row) > 20 else 0.0,
                'harmfulness_percent': self._parse_number(row[21]) if len(row) > 21 else 0.0,
                'night_percent': self._parse_number(row[22]) if len(row) > 22 else 0.0,
                'equipment_compensation': self._parse_number(row[25]) if len(row) > 25 else 0.0,
                'car_rent': self._parse_number(row[26]) if len(row) > 26 else 0.0,
                'crew_rent': self._parse_number(row[27]) if len(row) > 27 else 0.0,
                'fuel_limit_summer': self._parse_number(row[28]) if len(row) > 28 else 0.0,
                'fuel_limit_winter': self._parse_number(row[29]) if len(row) > 29 else 0.0,
                'transport_expenses': self._parse_number(row[30]) if len(row) > 30 else 0.0,
                'car_compensation': self._parse_number(row[31]) if len(row) > 31 else 0.0,
                'phone_limit': self._parse_number(row[32]) if len(row) > 32 else 0.0,
                'internet_limit': self._parse_number(row[33]) if len(row) > 33 else 0.0,
                'dms': self._parse_number(row[34]) if len(row) > 34 else 0.0,
                'housing_compensation': self._parse_number(row[35]) if len(row) > 35 else 0.0,
                'travel_compensation': self._parse_number(row[36]) if len(row) > 36 else 0.0,
                'compensation_summer': compensation_summer,
                'compensation_winter': compensation_winter,
                'monthly_bonus_percent': monthly_bonus_pct,
                'quarterly_bonus_percent': self._parse_number(row[40]) if len(row) > 40 else 0.0,
                'annual_bonus_percent': self._parse_number(row[41]) if len(row) > 41 else 0.0,
                'salary_type': str(row[13]).lower().strip() if len(row) > 13 else '',
                'last_employee': str(row[10]).strip().lower() if len(row) > 10 else '',
                'headcounts': headcounts,
                'excel_sv_monthly': sv_excel, 'excel_sv_total': sv_total_excel,
                'excel_comp_monthly': comp_excel, 'excel_comp_total': comp_total_excel,
                'excel_res_q_monthly': res_q_excel, 'excel_res_q_total': res_q_total_excel,
                'excel_res_a_monthly': res_a_excel, 'excel_res_a_total': res_a_total_excel,
                'gross_additions': 0.0,
                'comment': ''
            }
            employees.append(emp)
        self.employees_data = employees

    def _parse_hypotheses_sheet(self, ws):
        """Парсинг гипотез (оставлен для совместимости, но не используется для СВ)"""
        hyp = {'tariff_coefficients': {}, 'field_coefficients': {}, 'additional_coefficients': {}, 'bonus_payment': {},
               'staffing': {}}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            d = dict(zip([str(c) if c else '' for c in ws[1]], row))
            if d.get('Год') == '2026' and d.get('Месяц'):
                m = int(self._parse_number(d.get('Номер месяца', 0)))
                if 1 <= m <= 12:
                    hyp['staffing'][m] = self._parse_number(d.get('Укомплект-ть ЭТК', 0.9))
                    hyp['tariff_coefficients'][m] = self._parse_number(d.get('Коэффициент тарифа', 1.0))
                    hyp['field_coefficients'][m] = self._parse_number(d.get('Полевой коэффициент', 1.0))
                    hyp['additional_coefficients'][m] = self._parse_number(d.get('Дополнительный коэффициент', 1.0))
            if 'ночных' in str(d.get('Месяц2', '')).lower():
                hyp['night_share'] = self._parse_number(d.get('Доля ночных', 0.267))
            q = str(d.get('Месяц', ''))
            rate = self._parse_number(d.get('Процент выплаты кв. премии')) if d.get(
                'Процент выплаты кв. премии') else None
            if rate:
                if '1 кв' in q:
                    hyp['bonus_payment']['q1'] = rate
                elif '2 кв' in q:
                    hyp['bonus_payment']['q2'] = rate
                elif '3 кв' in q:
                    hyp['bonus_payment']['q3'] = rate
                elif '4 кв' in q:
                    hyp['bonus_payment']['q4'] = rate
        return hyp

    def get_employee_for_calculation(self, index):
        if index in self.modified_employees:
            entry = self.modified_employees[index]
            return entry['employee'] if isinstance(entry, dict) and 'employee' in entry else entry
        if 0 <= index < len(self.employees_data):
            return self.employees_data[index]
        return None

    def get_employee_for_month(self, index, month_index):
        if not (0 <= index < len(self.employees_data)):
            return None
        base = self.employees_data[index]
        if index not in self.modified_employees:
            return base
        entry = self.modified_employees[index]
        if isinstance(entry, dict) and 'employee' in entry:
            eff = max(1, min(12, int(entry.get('effective_from_month', 1))))
            emp = entry['employee']
        else:
            eff, emp = 1, entry
        return base if month_index + 1 < eff else emp

    def get_modification_effective_month(self, index):
        entry = self.modified_employees.get(index)
        if not entry:
            return 1
        if isinstance(entry, dict) and 'employee' in entry:
            return max(1, min(12, int(entry.get('effective_from_month', 1))))
        return 1

    def update_employee(self, index, updated_data, effective_from_month=1):
        eff = max(1, min(12, int(effective_from_month)))
        clean = dict(updated_data)
        clean.pop('effective_from_month', None)
        self.modified_employees[index] = {'employee': clean, 'effective_from_month': eff}
        self.changes_log.append(f"Изменен: {clean.get('fio', 'Сотрудник')} (с {self.months[eff - 1]})")
        return True

    def set_zero_headcount(self, index, month_index):
        emp = self.get_employee_for_calculation(index)
        if not emp:
            return False
        headcounts = list(emp.get('headcounts', [1.0] * 12))
        if 0 <= month_index < len(headcounts):
            headcounts[month_index] = 0.0
        clean = dict(emp)
        clean['headcounts'] = headcounts
        clean.pop('original_index', None)
        self.update_employee(index, clean, month_index + 1)
        self.changes_log[-1] = f"Выведена позиция: {clean.get('fio', '')} ({self.months[month_index]} -> 0 чел)"
        return True

    def reset_employee_changes(self, index):
        if index in self.modified_employees:
            del self.modified_employees[index]
            return True
        return False

    def search_employees(self, fio_query='', position_query=''):
        results = []
        if not self.employees_data:
            return results
        fio_q = fio_query.lower().strip() if fio_query else ''
        pos_q = position_query.lower().strip() if position_query else ''
        for idx, emp in enumerate(self.employees_data):
            if (fio_q and fio_q in emp['fio'].lower()) or (pos_q and pos_q in emp['position'].lower()):
                res = emp.copy()
                res['original_index'] = idx
                results.append(res)
        return results

    def calculate_monthly_fot_detailed(self, month_index, employee_index=None):
        """Расчёт ФОТ за конкретный месяц с учетом считывания СВ из ячеек"""
        if employee_index is None:
            return None
        emp = self.get_employee_for_month(employee_index, month_index)
        if emp is None:
            return None

        comp_s = emp.get('compensation_summer', 0.0)
        comp_w = emp.get('compensation_winter', 0.0)

        # Получаем массив СВ из Excel (CZ-DK)
        sv_monthly_data = emp.get('excel_sv_monthly', [0.0] * 12)

        income_ytd, sv_ytd = 0.0, 0.0
        target = None
        base_salary_arr = emp.get('base_salary_monthly', [emp.get('base_salary', 0.0)] * 12)

        for m in range(1, 13):
            base_m = base_salary_arr[m - 1] if m - 1 < len(base_salary_arr) else 0.0
            hc = emp['headcounts'][m - 1] if m - 1 < len(emp['headcounts']) else 1.0

            if hc > 0:
                ms = base_m
                mb_pct = emp.get('monthly_bonus_percent', 0.0) / 100 if emp.get('monthly_bonus_percent',
                                                                                0.0) > 1 else emp.get(
                    'monthly_bonus_percent', 0.0)
                mb = ms * mb_pct

                income_ytd += ms + mb

                # ✅ СЧИТЫВАНИЕ СВ НАПРЯМУЮ ИЗ EXCEL (CZ-DK)
                sv_m = sv_monthly_data[m - 1]
                sv_ytd += sv_m
            else:
                ms = 0.0
                mb = 0.0
                sv_m = 0.0

            comp_m = comp_w if m >= 10 or m <= 3 else comp_s
            res_q = emp['excel_res_q_monthly'][m - 1] if m - 1 < len(emp['excel_res_q_monthly']) else 0.0
            res_a = emp['excel_res_a_monthly'][m - 1] if m - 1 < len(emp['excel_res_a_monthly']) else 0.0

            fot_m = ms + mb + sv_m + res_q + res_a + comp_m

            if m == month_index + 1:
                target = {
                    'month_index': month_index, 'month_name': self.months[month_index],
                    'monthly_salary': ms, 'monthly_bonus': mb,
                    'cumulative_income': income_ytd, 'social_contributions': sv_m,
                    'compensations': comp_m, 'quarterly_reserve': res_q,
                    'annual_reserve': res_a,
                    'fot_total': fot_m
                }
        return target

    def create_empty_employee(self):
        return {
            'directorate': '', 'fio': '', 'position': '', 'base_salary': 0.0,
            'base_salary_monthly': [0.0] * 12, 'allowance_cat1': 0.0, 'allowance_cat2': 0.0,
            'travel_allowance': 0.0, 'rk_percent': 0.0, 'sn_percent': 0.0,
            'harmfulness_percent': 0.0, 'night_percent': 0.0, 'equipment_compensation': 0.0,
            'car_rent': 0.0, 'crew_rent': 0.0, 'fuel_limit_summer': 0.0, 'fuel_limit_winter': 0.0,
            'transport_expenses': 0.0, 'car_compensation': 0.0, 'phone_limit': 0.0, 'internet_limit': 0.0,
            'dms': 0.0, 'housing_compensation': 0.0, 'travel_compensation': 0.0,
            'compensation_summer': 0.0, 'compensation_winter': 0.0,
            'monthly_bonus_percent': 0.0, 'quarterly_bonus_percent': 0.0, 'annual_bonus_percent': 0.0,
            'salary_type': '', 'last_employee': '', 'headcounts': [1.0] * 12,
            'excel_sv_monthly': [0.0] * 12, 'excel_sv_total': 0.0,
            'excel_comp_monthly': [0.0] * 12, 'excel_comp_total': 0.0, 'excel_res_q_monthly': [0.0] * 12,
            'excel_res_q_total': 0.0, 'excel_res_a_monthly': [0.0] * 12, 'excel_res_a_total': 0.0,
            'gross_additions': 0.0, 'comment': ''
        }

    def add_employee(self, emp_data):
        self.employees_data.append(emp_data)
        self.changes_log.append(f"Добавлен сотрудник: {emp_data.get('fio', 'Без ФИО')}")


class EmployeeEditDialog:
    """Диалог редактирования"""
    _PERCENT_FIELD_LABELS = {
        'rk_percent': 'РК (%)', 'sn_percent': 'СН (%)',
        'monthly_bonus_percent': 'Месячная премия (%)',
        'quarterly_bonus_percent': 'Квартальная премия (%)',
        'annual_bonus_percent': 'Годовая премия (%)',
    }

    def __init__(self, parent, calculator):
        self.calculator = calculator
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Редактирование данных")
        self.dialog.geometry("950x850")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.selected_employee = None
        self.original_index = None
        self.entries = {}
        self.edit_widgets = []
        self.effective_month_combo = None
        self.is_adding_new = False
        self._create_widgets()
        self.dialog.wait_window()

    @staticmethod
    def _format_salary_display(value):
        try:
            return f"{float(value):.0f}"
        except:
            return "0"

    @staticmethod
    def _format_percent_display(value):
        try:
            v = float(value)
        except:
            return "0"
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        s = f"{v:.4f}".rstrip('0').rstrip('.')
        return s if s else "0"

    @staticmethod
    def _format_text_display(value):
        return " " if value is None else str(value).strip()

    def _create_widgets(self):
        mf = ttk.Frame(self.dialog, padding="10")
        mf.pack(fill=tk.BOTH, expand=True)
        sf = ttk.LabelFrame(mf, text="Поиск сотрудника", padding="10")
        sf.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(sf, text="ФИО:   ").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.search_fio_entry = ttk.Entry(sf, width=40)
        self.search_fio_entry.grid(row=0, column=1, sticky="ew", padx=(0, 20))
        ttk.Label(sf, text="Должность:   ").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=(10, 0))
        self.search_position_entry = ttk.Entry(sf, width=40)
        self.search_position_entry.grid(row=1, column=1, sticky="ew", padx=(0, 20), pady=(10, 0))
        btn_f = ttk.Frame(sf)
        btn_f.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(btn_f, text="Найти", command=self._search_employees).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_f, text="Очистить", command=self._clear_search).pack(side=tk.LEFT)

        rf = ttk.LabelFrame(mf, text="Результаты поиска", padding="10")
        rf.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        cols = ('fio', 'position', 'directorate', 'salary', 'status')
        self.tree = ttk.Treeview(rf, columns=cols, show='headings', height=8)
        for c, w, t in zip(cols, [250, 200, 150, 120, 100], ('ФИО', 'Должность', 'Дирекция', 'Оклад', 'Статус')):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor='e' if c == 'salary' else 'center' if c == 'status' else 'w')
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.item_to_employee = {}
        self.tree.bind('<Double-1>', self._select_employee)

        ef = ttk.LabelFrame(mf, text="Редактирование данных", padding="10")
        ef.pack(fill=tk.X, pady=(0, 10))
        self.edit_fields_frame = ttk.Frame(ef)
        self.edit_fields_frame.pack(fill=tk.X)
        self.no_sel_label = ttk.Label(self.edit_fields_frame, text="Выберите сотрудника из таблицы")
        self.no_sel_label.pack(pady=20)

        af = ttk.Frame(mf)
        af.pack(fill=tk.X)
        self.btn_save = ttk.Button(af, text="Сохранить изменения", command=self._save_changes, state=tk.DISABLED)
        self.btn_save.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_reset = ttk.Button(af, text="Сбросить изменения", command=self._reset_changes, state=tk.DISABLED)
        self.btn_reset.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_zero = ttk.Button(af, text="Вывести позицию (0 в выбранном мес.)",
                                   command=self._set_zero_headcount, state=tk.DISABLED)
        self.btn_zero.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_add = ttk.Button(af, text="Добавить позицию", command=self._start_add_mode)
        self.btn_add.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(af, text="Закрыть", command=self.dialog.destroy).pack(side=tk.RIGHT)

    def _modification_status_label(self, idx):
        if idx < 0 or idx not in self.calculator.modified_employees:
            return "Оригинал"
        m = self.calculator.get_modification_effective_month(idx)
        return f"Изменен (с {self.calculator.months[m - 1]})"

    def _current_employee_row_for_edit(self, emp):
        idx = emp.get('original_index', -1)
        if idx < 0: return emp
        row = self.calculator.get_employee_for_calculation(idx)
        return emp if row is None else {**row, 'original_index': idx}

    def _search_employees(self):
        fio = self.search_fio_entry.get().strip()
        pos = self.search_position_entry.get().strip()
        if not fio and not pos:
            return messagebox.showwarning("Предупреждение", "Введите хотя бы один параметр")
        if not self.calculator.employees_data:
            return messagebox.showwarning("Предупреждение", "Загрузите файл перед поиском")
        res = self.calculator.search_employees(fio, pos)
        for i in self.tree.get_children(): self.tree.delete(i)
        self.item_to_employee.clear()
        if not res: return messagebox.showinfo("Результат", "Сотрудники не найдены")
        for emp in res:
            idx = emp.get('original_index', -1)
            self.tree.insert('', tk.END, values=(emp['fio'], emp['position'], emp['directorate'],
                                                 self._format_salary_display(emp['base_salary']),
                                                 self._modification_status_label(idx)))
            self.item_to_employee[self.tree.get_children()[-1]] = emp

    def _clear_search(self):
        self.search_fio_entry.delete(0, tk.END)
        self.search_position_entry.delete(0, tk.END)
        for i in self.tree.get_children(): self.tree.delete(i)
        self.item_to_employee.clear()
        self._clear_edit_form()

    def _select_employee(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        emp = self.item_to_employee.get(sel[0])
        if not emp: return
        self.selected_employee = self._current_employee_row_for_edit(emp)
        self.original_index = self.selected_employee.get('original_index', -1)
        self._show_edit_form(self.selected_employee)
        self.btn_save.config(state=tk.NORMAL)
        self.btn_reset.config(state=tk.NORMAL)
        self.btn_zero.config(state=tk.NORMAL)
        self.btn_add.config(state=tk.DISABLED)

    def _show_edit_form(self, emp):
        self._clear_edit_form()
        self.no_sel_label.pack_forget()
        eff = self.calculator.get_modification_effective_month(
            self.original_index) if self.original_index is not None and self.original_index >= 0 else 1
        mr = ttk.Frame(self.edit_fields_frame)
        mr.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(mr, text="Изменения с месяца:   ", width=25, anchor="w").pack(side=tk.LEFT, padx=(0, 10))
        self.effective_month_combo = ttk.Combobox(mr, width=22, state="readonly", values=self.calculator.months)
        self.effective_month_combo.current(eff - 1)
        self.effective_month_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.edit_widgets.append(mr)

        for lbl, key in [("ФИО:   ", 'fio'), ("Должность:   ", 'position'), ("Оклад (руб.):   ", 'base_salary'),
                         ("РК (%):   ", 'rk_percent'), ("СН (%):   ", 'sn_percent'),
                         ("Месячная премия (%):   ", 'monthly_bonus_percent'),
                         ("Квартальная премия (%):   ", 'quarterly_bonus_percent'),
                         ("Годовая премия (%):   ", 'annual_bonus_percent')]:
            f = ttk.Frame(self.edit_fields_frame)
            f.pack(fill=tk.X, pady=5)
            ttk.Label(f, text=lbl, width=25, anchor="w").pack(side=tk.LEFT, padx=(0, 10))
            e = ttk.Entry(f, width=35)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True)
            val = emp.get(key, '' if key in ('fio', 'position') else 0)
            if key == 'base_salary':
                e.insert(0, self._format_salary_display(val))
            elif key in self._PERCENT_FIELD_LABELS:
                e.insert(0, self._format_percent_display(val * 100 if val < 1 else val))
            else:
                e.insert(0, self._format_text_display(val))
            self.entries[key] = e
            self.edit_widgets.append(f)

        f_gross = ttk.Frame(self.edit_fields_frame)
        f_gross.pack(fill=tk.X, pady=5)
        ttk.Label(f_gross, text="Доплаты ГРОСС (руб.):   ", width=25, anchor="w").pack(side=tk.LEFT, padx=(0, 10))
        e_gross = ttk.Entry(f_gross, width=35)
        e_gross.pack(side=tk.LEFT, fill=tk.X, expand=True)
        e_gross.insert(0, self._format_salary_display(emp.get('gross_additions', 0)))
        self.entries['gross_additions'] = e_gross
        self.edit_widgets.append(f_gross)

        f_comm = ttk.Frame(self.edit_fields_frame)
        f_comm.pack(fill=tk.X, pady=5)
        ttk.Label(f_comm, text="Комментарий:   ", width=25, anchor="w").pack(side=tk.LEFT, padx=(0, 10))
        e_comm = ttk.Entry(f_comm, width=35)
        e_comm.pack(side=tk.LEFT, fill=tk.X, expand=True)
        e_comm.insert(0, emp.get('comment', ''))
        self.entries['comment'] = e_comm
        self.edit_widgets.append(f_comm)

    def _clear_edit_form(self):
        for w in self.edit_widgets: w.destroy()
        self.edit_widgets.clear()
        self.entries.clear()
        self.effective_month_combo = None
        self.no_sel_label.pack(pady=20)
        self.btn_save.config(state=tk.DISABLED)
        self.btn_reset.config(state=tk.DISABLED)
        if hasattr(self, 'btn_zero'): self.btn_zero.config(state=tk.DISABLED)

    def _save_changes(self):
        if self.original_index is None and not self.is_adding_new: return
        upd = self.calculator.create_empty_employee() if self.is_adding_new else self.selected_employee.copy()
        for k, e in self.entries.items():
            v = e.get().strip()
            if k in ['base_salary', 'rk_percent', 'sn_percent', 'monthly_bonus_percent',
                     'quarterly_bonus_percent', 'annual_bonus_percent', 'gross_additions']:
                num = self.calculator._parse_number(v) if v else 0
                if k not in ('base_salary', 'gross_additions') and num > 100: num /= 100
                upd[k] = num
            else:
                upd[k] = v
        if self.is_adding_new:
            if not upd.get('fio', '').strip():
                return messagebox.showwarning("Внимание", "Введите ФИО сотрудника")
            self.calculator.add_employee(upd)
            messagebox.showinfo("Успех", f"Сотрудник {upd.get('fio', '')} добавлен.")
        else:
            try:
                eff = self.calculator.months.index(self.effective_month_combo.get()) + 1
            except:
                eff = 1
            self.calculator.update_employee(self.original_index, upd, eff)
            messagebox.showinfo("Успех", f"Изменения сохранены: {upd.get('fio', '')}")
        self.dialog.destroy()

    def _set_zero_headcount(self):
        if self.original_index is None: return messagebox.showwarning("Внимание", "Сначала выберите сотрудника")
        try:
            eff_month = self.calculator.months.index(self.effective_month_combo.get()) + 1
        except:
            eff_month = 1
        if self.calculator.set_zero_headcount(self.original_index, eff_month - 1):
            messagebox.showinfo("Успех", f"Позиция выведена для '{self.calculator.months[eff_month - 1]}'")
            self.dialog.destroy()

    def _reset_changes(self):
        if self.is_adding_new:
            self.is_adding_new = False
            self.original_index = None
            self._clear_edit_form()
            self.btn_add.config(state=tk.NORMAL)
            return
        if self.original_index is None: return
        if messagebox.askyesno("Подтверждение", "Сбросить изменения?"):
            self.calculator.reset_employee_changes(self.original_index)
            self.dialog.destroy()

    def _start_add_mode(self):
        self.is_adding_new = True
        self.original_index = None
        self.search_fio_entry.delete(0, tk.END)
        self.search_position_entry.delete(0, tk.END)
        for i in self.tree.get_children(): self.tree.delete(i)
        self.item_to_employee.clear()
        self._show_edit_form(self.calculator.create_empty_employee())
        self.btn_save.config(state=tk.NORMAL)
        self.btn_reset.config(state=tk.NORMAL)
        self.btn_zero.config(state=tk.DISABLED)
        self.btn_add.config(state=tk.DISABLED)


class PayrollGUI:
    """Графический интерфейс"""

    def __init__(self, calculator):
        self.calculator = calculator
        self.root = tk.Tk()
        self.root.title("Расчет ФОТ 2026")
        self.root.geometry("1200x700")
        self.results_text = " "
        self._create_widgets()

    def _create_widgets(self):
        mf = ttk.Frame(self.root, padding="10")
        mf.pack(fill=tk.BOTH, expand=True)

        lp = ttk.Frame(mf, width=350)
        lp.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        lp.pack_propagate(False)

        ttk.Button(lp, text="Загрузить Файл", command=self._load_file).pack(fill=tk.X, pady=(0, 5))
        ttk.Button(lp, text="Добавить изменение", command=self._add_change).pack(fill=tk.X, pady=(0, 10))

        hist_frame = ttk.LabelFrame(lp, text="История изменений", padding=5)
        hist_frame.pack(fill=tk.BOTH, expand=True)
        self.history_list = tk.Listbox(hist_frame, height=25, font=("Arial", 9), width=40)
        self.history_list.pack(fill=tk.BOTH, expand=True)

        rp = ttk.Frame(mf)
        rp.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rf = ttk.LabelFrame(rp, text="Результаты:   ", padding="10")
        rf.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.txt = tk.Text(rf, wrap=tk.WORD, font=("Courier New", 9))
        self.txt.pack(fill=tk.BOTH, expand=True)

        bp = ttk.Frame(rp)
        bp.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(bp, text="Сохранить", command=self._save_results).pack(side=tk.RIGHT)
        ttk.Button(bp, text="Рассчитать", command=self._calculate).pack(side=tk.RIGHT, padx=(0, 10))
        self._update_results()

    def _load_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if fp:
            try:
                self.calculator.load_employee_data(fp)
                messagebox.showinfo("Успех", f"Загружено сотрудников: {len(self.calculator.employees_data)}")
                self._refresh_history()
                self._update_results()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def _add_change(self):
        if not self.calculator.employees_data:
            return messagebox.showwarning("Внимание", "Сначала загрузите файл")
        EmployeeEditDialog(self.root, self.calculator)
        self._refresh_history()

    def _refresh_history(self):
        self.history_list.delete(0, tk.END)
        for idx, change in enumerate(self.calculator.changes_log, 1):
            self.history_list.insert(tk.END, f"{idx}. {change}")

    def _calculate(self):
        if not self.calculator.employees_data:
            return messagebox.showwarning("Внимание", "Загрузите файл перед расчетом")
        try:
            text = "=" * 80 + "\nОБЩИЙ РАСЧЕТ ФОТ\n" + "=" * 80 + "\n\n"
            total = 0.0
            for idx in range(len(self.calculator.employees_data)):
                cur = self.calculator.get_employee_for_calculation(idx)
                # Суммируем месячный ФОТ за год (СВ берется из ячеек)
                base_emp_fot = sum(
                    self.calculator.calculate_monthly_fot_detailed(m, employee_index=idx)['fot_total'] for m in
                    range(12))
                gross_add = cur.get('gross_additions', 0.0)
                emp_fot = base_emp_fot + gross_add

                total += emp_fot
                marker = "*   " if idx in self.calculator.modified_employees else "    "

                comment = cur.get('comment', '')
                comment_str = f"    ({comment})" if comment else ""

                text += f"{marker}{cur['fio']:<40} {emp_fot:>15,.2f}{comment_str}\n"
            text += "=" * 80 + f"\nИТОГО: {total:>15,.2f}\n"
            self.txt.delete(1.0, tk.END)
            self.txt.insert(tk.END, text)
            self.results_text = text
            messagebox.showinfo("Готово", f"Расчет завершен. ФОТ: {total:,.0f}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _update_results(self):
        self.txt.delete(1.0, tk.END)
        self.txt.insert(tk.END, "Загрузите файл и нажмите 'Рассчитать'.")

    def _save_results(self):
        if not self.results_text:
            return messagebox.showwarning("Внимание", "Нет результатов")
        d = filedialog.askdirectory()
        if not d:
            return
        try:
            rows = []
            for idx in range(len(self.calculator.employees_data)):
                emp = self.calculator.get_employee_for_calculation(idx)
                gross = emp.get('gross_additions', 0.0)
                for m in range(12):
                    r = self.calculator.calculate_monthly_fot_detailed(m, employee_index=idx)
                    rows.append({'ФИО': emp['fio'], 'Месяц': r['month_name'], 'Оклад': r['monthly_salary'],
                                 'Премия': r['monthly_bonus'], 'СВ': r['social_contributions'],
                                 'Доплаты ГРОСС (разово)': gross, 'ФОТ': r['fot_total']})
            import pandas as pd
            pd.DataFrame(rows).to_excel(os.path.join(d, "ФОТ_Расчет.xlsx"), index=False)
            messagebox.showinfo("Успех", "Экспорт завершен")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def run(self):
        self.root.mainloop()


def run_gui():
    calculator = PayrollCalculator()
    app = PayrollGUI(calculator)
    app.run()


if __name__ == "__main__":
    run_gui()