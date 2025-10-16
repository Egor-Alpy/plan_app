import sys
from datetime import datetime, timedelta
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QComboBox,
                             QTableWidget, QTableWidgetItem, QMessageBox,
                             QFileDialog, QHeaderView, QFrame, QScrollArea)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont


class EducationalScheduleApp:
    def __init__(self):
        # Русские названия месяцев
        self.month_names_ru = {
            1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель',
            5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август',
            9: 'Сентябрь', 10: 'Октябрь', 11: 'Ноябрь', 12: 'Декабрь'
        }

        # Праздничные дни России по годам
        self.holidays = {
            2025: ['2025-01-01', '2025-01-02', '2025-01-03', '2025-01-04', '2025-01-06',
                   '2025-01-07', '2025-01-08', '2025-02-23', '2025-03-08', '2025-05-01',
                   '2025-05-02', '2025-05-08', '2025-05-09', '2025-06-12', '2025-06-13',
                   '2025-11-03', '2025-11-04'],
            2026: ['2026-01-01', '2026-01-02', '2026-01-05', '2026-01-06', '2026-01-07',
                   '2026-01-08', '2026-01-09', '2026-02-23', '2026-03-09', '2026-05-01',
                   '2026-05-09', '2026-05-11', '2026-06-12', '2026-11-04'],
            2027: ['2027-01-01', '2027-01-04', '2027-01-05', '2027-01-06', '2027-01-07',
                   '2027-01-08', '2027-02-22', '2027-02-23', '2027-03-08', '2027-05-03',
                   '2027-05-10', '2027-06-14', '2027-11-04'],
            2028: ['2028-01-03', '2028-01-04', '2028-01-05', '2028-01-06', '2028-01-07',
                   '2028-02-23', '2028-03-08', '2028-05-01', '2028-05-09', '2028-06-12',
                   '2028-11-04']
        }

    def get_monday_of_week(self, date):
        days_since_monday = date.weekday()
        return date - timedelta(days=days_since_monday)

    def is_holiday(self, date):
        year = date.year
        date_str = date.strftime('%Y-%m-%d')
        return date_str in self.holidays.get(year, [])

    def is_working_day(self, date):
        return date.weekday() < 5 and not self.is_holiday(date)

    def calculate_academic_weeks(self, start_date, weeks_float):
        current_date = start_date
        working_days_needed = int(weeks_float * 5)
        working_days_count = 0
        schedule_days = []

        while working_days_count < working_days_needed:
            if self.is_working_day(current_date):
                schedule_days.append(current_date)
                working_days_count += 1
            current_date += timedelta(days=1)

        while not self.is_working_day(current_date):
            current_date += timedelta(days=1)

        return schedule_days, current_date

    def generate_schedule(self, periods_data, start_year):
        start_date = datetime(start_year, 9, 1)
        current_date = self.get_monday_of_week(start_date)

        generated_schedule = []

        for row in periods_data:
            year = int(row['Год'])
            semester = int(row['Семестр'])
            activity_type = row['Тип']
            weeks = float(row['Недели'])

            period_days, next_date = self.calculate_academic_weeks(current_date, weeks)

            period_info = {
                'year': year,
                'semester': semester,
                'type': activity_type,
                'weeks': weeks,
                'start_date': current_date,
                'end_date': period_days[-1] if period_days else current_date,
                'days': period_days
            }

            generated_schedule.append(period_info)
            current_date = next_date

        return generated_schedule

    def create_excel_file(self, generated_schedule, start_year, program_type):
        wb = Workbook()
        program_years = 2 if "Ординатура" in program_type else 3

        # Стили
        header_font = Font(bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

        activity_fills = {
            'Т': PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid"),
            'П': PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid"),
            'ПА': PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid"),
            'ГИА': PatternFill(start_color="DDA0DD", end_color="DDA0DD", fill_type="solid"),
            'К': PatternFill(start_color="F0E68C", end_color="F0E68C", fill_type="solid")
        }

        weekend_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
        holiday_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")

        # Создать листы для каждого года
        for academic_year in range(program_years):
            actual_year = start_year + academic_year

            if academic_year == 0:
                ws = wb.active
                ws.title = f"{actual_year}-{actual_year + 1}"
            else:
                ws = wb.create_sheet(f"{actual_year}-{actual_year + 1}")

            self.create_academic_year_calendar(ws, actual_year, header_font,
                                               weekend_fill, holiday_fill, activity_fills,
                                               thin_border, generated_schedule)

        # Лист с обозначениями
        legend_ws = wb.create_sheet("Обозначения")
        self.create_legend_sheet(legend_ws, header_font, activity_fills,
                                 weekend_fill, holiday_fill, thin_border)

        return wb

    def create_academic_year_calendar(self, ws, start_year, header_font,
                                      weekend_fill, holiday_fill, activity_fills,
                                      thin_border, generated_schedule):

        # Заголовок
        ws.merge_cells('A1:AH1')
        ws['A1'] = f"Календарный учебный график {start_year}-{start_year + 1} г."
        ws['A1'].font = Font(size=16, bold=True)
        ws['A1'].alignment = Alignment(horizontal='center')

        # Месяцы учебного года (сентябрь-август)
        academic_months = [(start_year, m) for m in range(9, 13)] + [(start_year + 1, m) for m in range(1, 9)]

        # Строка 2 - названия месяцев
        current_col = 2

        for year, month in academic_months:
            month_name = self.month_names_ru[month]
            cal = calendar.monthcalendar(year, month)
            month_weeks = len(cal)

            # Заголовок месяца
            if month_weeks > 1:
                ws.merge_cells(f'{get_column_letter(current_col)}2:{get_column_letter(current_col + month_weeks - 1)}2')

            cell = ws.cell(row=2, column=current_col)
            cell.value = month_name
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

            current_col += month_weeks

        # Дни недели в первой колонке
        days_of_week = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс']

        ws['A2'] = 'Мес'
        ws['A2'].font = header_font
        ws['A2'].alignment = Alignment(horizontal='center')
        ws['A2'].border = thin_border

        for row_idx, day_name in enumerate(days_of_week, 3):
            cell = ws['A{}'.format(row_idx)]
            cell.value = day_name
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border

        ws['A10'] = 'Нед'
        ws['A10'].font = header_font
        ws['A10'].alignment = Alignment(horizontal='center')
        ws['A10'].border = thin_border

        # Заполняем календарную сетку
        current_col = 2
        week_number = 1

        for year, month in academic_months:
            cal = calendar.monthcalendar(year, month)

            for week_idx, week in enumerate(cal):
                col = current_col + week_idx

                # Номер недели
                ws.cell(row=10, column=col).value = week_number
                ws.cell(row=10, column=col).alignment = Alignment(horizontal='center')
                ws.cell(row=10, column=col).border = thin_border
                week_number += 1

                # Дни недели
                for day_idx, day in enumerate(week):
                    row = 3 + day_idx
                    cell = ws.cell(row=row, column=col)
                    cell.border = thin_border

                    if day == 0:
                        cell.value = ""
                        cell.fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
                    else:
                        date = datetime(year, month, day)
                        cell.value = day
                        cell.alignment = Alignment(horizontal='center')

                        # Применяем цвета и обозначения
                        if self.is_holiday(date):
                            cell.fill = holiday_fill
                        elif date.weekday() >= 5:
                            cell.fill = weekend_fill
                        else:
                            activity_type = self.get_activity_for_date(date, generated_schedule)
                            if activity_type and activity_type in activity_fills:
                                cell.fill = activity_fills[activity_type]
                                cell.value = f"{day}\n{activity_type}"
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                cell.font = Font(size=9)

            current_col += len(cal)

        # Настройка размеров
        ws.column_dimensions['A'].width = 6
        for col in range(2, current_col):
            ws.column_dimensions[get_column_letter(col)].width = 5

        for row in range(3, 10):
            ws.row_dimensions[row].height = 25

    def get_activity_for_date(self, date, generated_schedule):
        for period in generated_schedule:
            if date in period['days']:
                return period['type']
        return None

    def create_legend_sheet(self, ws, header_font, activity_fills, weekend_fill, holiday_fill, thin_border):
        ws['A1'] = "Условные обозначения"
        ws['A1'].font = Font(size=16, bold=True)

        ws['A3'] = "Типы занятий:"
        ws['A3'].font = header_font

        activity_names = ['Т', 'П', 'ПА', 'ГИА', 'К']
        activity_descriptions = ['Теоретическая подготовка', 'Практика', 'Промежуточная аттестация',
                                 'Государственная итоговая аттестация', 'Каникулы']

        for i, name in enumerate(activity_names):
            col = chr(66 + i)
            ws[f'{col}4'] = name
            ws[f'{col}4'].font = header_font
            ws[f'{col}4'].fill = activity_fills[name]
            ws[f'{col}4'].border = thin_border
            ws[f'{col}4'].alignment = Alignment(horizontal='center')

            ws[f'{col}5'] = activity_descriptions[i]
            ws[f'{col}5'].border = thin_border
            ws[f'{col}5'].alignment = Alignment(horizontal='center')

        ws['A7'] = "Прочие обозначения:"
        ws['A7'].font = header_font

        ws['B8'] = "Выходные"
        ws['B8'].fill = weekend_fill
        ws['B8'].border = thin_border
        ws['B8'].alignment = Alignment(horizontal='center')

        ws['C8'] = "Праздники"
        ws['C8'].fill = holiday_fill
        ws['C8'].border = thin_border
        ws['C8'].alignment = Alignment(horizontal='center')

        for col_letter in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col_letter].width = 20


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.app = EducationalScheduleApp()
        self.periods_data = []
        self.generated_schedule = None
        self.start_year = 2025
        self.program_type = "Ординатура (2 года)"

        self.init_ui()
        self.apply_styles()

    def init_ui(self):
        self.setWindowTitle('Учебный график')
        self.setGeometry(100, 100, 1500, 900)

        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Скролл область
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        scroll_content = QWidget()
        scroll_content.setStyleSheet("background-color: #0e1117;")
        content_layout = QVBoxLayout(scroll_content)
        content_layout.setSpacing(0)
        content_layout.setContentsMargins(0, 0, 0, 0)

        # Основной контент с padding
        content_container = QWidget()
        content_container.setStyleSheet("background-color: #0e1117;")
        container_layout = QVBoxLayout(content_container)
        container_layout.setContentsMargins(50, 40, 50, 50)
        container_layout.setSpacing(32)

        # Логотип и заголовок
        header_row = QHBoxLayout()
        header_row.setSpacing(24)

        # Логотип (если есть файл logo.png)
        logo_label = QLabel()
        logo_label.setFixedSize(100, 100)
        logo_label.setScaledContents(True)
        logo_label.setStyleSheet("""
            background-color: transparent;
            border-radius: 12px;
        """)
        # Попытка загрузить логотип (если файл существует)
        try:
            from PyQt6.QtGui import QPixmap
            pixmap = QPixmap('logo.png')
            if not pixmap.isNull():
                logo_label.setPixmap(pixmap)
                header_row.addWidget(logo_label)
        except:
            pass

        # Текстовый блок заголовка
        title_layout = QVBoxLayout()
        title_layout.setSpacing(8)

        # Заголовок
        title = QLabel('Учебный график')
        title.setObjectName("mainTitle")

        subtitle = QLabel('Создание календарного учебного графика')
        subtitle.setObjectName("subtitle")

        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)

        header_row.addLayout(title_layout)
        header_row.addStretch()

        container_layout.addLayout(header_row)

        # Декоративная линия под хедером
        header_line = QFrame()
        header_line.setFrameShape(QFrame.Shape.HLine)
        header_line.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 transparent, stop:0.2 #34d399, stop:0.8 #34d399, stop:1 transparent);
                border: none;
                height: 3px;
                margin: 16px 0px;
            }
        """)
        container_layout.addWidget(header_line)
        container_layout.addSpacing(8)

        # Настройки (без заголовка)
        settings_row = QHBoxLayout()
        settings_row.setSpacing(20)

        # Тип программы
        program_layout = QVBoxLayout()
        program_layout.setSpacing(8)
        program_label = QLabel('Тип программы')
        program_label.setObjectName("inputLabel")
        self.program_combo = QComboBox()
        self.program_combo.addItems(['Ординатура (2 года)', 'Аспирантура (3 года)'])
        self.program_combo.currentTextChanged.connect(self.on_program_changed)
        program_layout.addWidget(program_label)
        program_layout.addWidget(self.program_combo)

        # Начальный год
        year_layout = QVBoxLayout()
        year_layout.setSpacing(8)
        year_label = QLabel('Начальный год')
        year_label.setObjectName("inputLabel")
        self.year_combo = QComboBox()
        self.year_combo.addItems(['2025', '2026', '2027'])
        self.year_combo.currentTextChanged.connect(self.on_year_changed)
        year_layout.addWidget(year_label)
        year_layout.addWidget(self.year_combo)

        settings_row.addLayout(program_layout)
        settings_row.addLayout(year_layout)
        settings_row.addStretch()

        container_layout.addLayout(settings_row)

        # Кнопки примера
        button_row = QHBoxLayout()
        button_row.setSpacing(12)

        example_btn = QPushButton('📋 Пример ординатуры')
        example_btn.setObjectName("secondaryButton")
        example_btn.clicked.connect(self.load_example)

        clear_btn = QPushButton('🗑️ Очистить')
        clear_btn.setObjectName("secondaryButton")
        clear_btn.clicked.connect(self.clear_data)

        button_row.addWidget(example_btn)
        button_row.addWidget(clear_btn)
        button_row.addStretch()

        container_layout.addLayout(button_row)
        container_layout.addSpacing(16)

        # Периоды обучения
        periods_label = QLabel('Периоды обучения')
        periods_label.setObjectName("sectionTitle")
        container_layout.addWidget(periods_label)

        # Таблица
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['Год', 'Семестр', 'Тип', 'Недели'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(82)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setMinimumHeight(400)
        container_layout.addWidget(self.table)

        # Кнопки таблицы
        table_btn_row = QHBoxLayout()
        table_btn_row.setSpacing(12)

        add_row_btn = QPushButton('➕ Добавить строку')
        add_row_btn.setObjectName("secondaryButton")
        add_row_btn.clicked.connect(self.add_row)

        remove_row_btn = QPushButton('➖ Удалить строку')
        remove_row_btn.setObjectName("secondaryButton")
        remove_row_btn.clicked.connect(self.remove_row)

        table_btn_row.addWidget(add_row_btn)
        table_btn_row.addWidget(remove_row_btn)
        table_btn_row.addStretch()

        container_layout.addLayout(table_btn_row)
        container_layout.addSpacing(16)

        # Кнопки действий
        action_row = QHBoxLayout()
        action_row.setSpacing(16)

        generate_btn = QPushButton('🎓 Сгенерировать график')
        generate_btn.setObjectName("primaryButton")
        generate_btn.clicked.connect(self.generate_schedule)

        self.download_btn = QPushButton('📥 Скачать Excel')
        self.download_btn.setObjectName("downloadButton")
        self.download_btn.clicked.connect(self.download_excel)
        self.download_btn.setEnabled(False)

        action_row.addWidget(generate_btn, 1)
        action_row.addWidget(self.download_btn, 1)

        container_layout.addLayout(action_row)

        # Предварительный просмотр
        self.preview_section = QWidget()
        self.preview_section.setStyleSheet("background-color: #0e1117;")
        preview_layout = QVBoxLayout(self.preview_section)
        preview_layout.setContentsMargins(0, 32, 0, 0)
        preview_layout.setSpacing(16)

        preview_label = QLabel('Предварительный просмотр')
        preview_label.setObjectName("sectionTitle")
        preview_layout.addWidget(preview_label)

        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(7)
        self.preview_table.setHorizontalHeaderLabels(['Год', 'Семестр', 'Тип', 'Недели', 'Начало', 'Конец', 'Дней'])
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.verticalHeader().setDefaultSectionSize(60)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setShowGrid(False)
        self.preview_table.setMinimumHeight(350)
        preview_layout.addWidget(self.preview_table)

        self.preview_section.setVisible(False)
        container_layout.addWidget(self.preview_section)

        # Footer с авторами
        footer_layout = QVBoxLayout()
        footer_layout.setContentsMargins(0, 32, 0, 0)
        footer_layout.setSpacing(0)

        authors_label = QLabel('Разработчики: Бахмутов Е., Клюев П.')
        authors_label.setObjectName("authorsLabel")
        authors_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer_layout.addWidget(authors_label)

        container_layout.addLayout(footer_layout)

        content_layout.addWidget(content_container)

        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0e1117;
            }

            QLabel#mainTitle {
                font-size: 42px;
                font-weight: 800;
                color: #fafafa;
                margin: 0;
                padding: 0;
                letter-spacing: -0.5px;
            }

            QLabel#subtitle {
                font-size: 18px;
                color: #34d399;
                margin-top: 4px;
                font-weight: 500;
                letter-spacing: 0.3px;
            }

            QLabel#sectionTitle {
                font-size: 20px;
                font-weight: 600;
                color: #fafafa;
                margin-bottom: 8px;
            }

            QLabel#inputLabel {
                font-size: 14px;
                font-weight: 600;
                color: #fafafa;
            }

            QLabel#authorsLabel {
                font-size: 13px;
                color: #6b7280;
                font-weight: 400;
                opacity: 0.7;
            }

            QComboBox {
                padding: 12px 16px;
                border: 1px solid #464a5e;
                border-radius: 8px;
                background-color: #262730;
                font-size: 15px;
                min-width: 240px;
                color: #fafafa;
                min-height: 44px;
                font-weight: 400;
            }

            QComboBox:hover {
                border-color: #2d8659;
            }

            QComboBox:focus {
                border-color: #2d8659;
                outline: none;
            }

            QComboBox::drop-down {
                border: none;
                width: 32px;
            }

            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #a3a8b4;
                margin-right: 10px;
            }

            QComboBox QAbstractItemView {
                background-color: #262730;
                border: 1px solid #464a5e;
                selection-background-color: #2d8659;
                selection-color: #ffffff;
                outline: none;
                padding: 6px;
                font-size: 15px;
                color: #fafafa;
            }

            QComboBox QAbstractItemView::item {
                padding: 10px 12px;
                min-height: 36px;
                color: #fafafa;
            }

            QComboBox QAbstractItemView::item:hover {
                background-color: #2d8659;
                color: #ffffff;
            }

            QPushButton#primaryButton {
                padding: 14px 32px;
                border: 2px solid #34d399;
                border-radius: 10px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #34d399, stop:0.5 #2d8659, stop:1 #10b981);
                color: white;
                font-size: 17px;
                font-weight: 700;
                min-height: 56px;
                letter-spacing: 1px;
            }

            QPushButton#primaryButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #6ee7b7, stop:0.5 #34d399, stop:1 #10b981);
                border-color: #6ee7b7;
                border-width: 3px;
                padding: 13px 31px;
            }

            QPushButton#primaryButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #22c55e, stop:0.5 #16a34a, stop:1 #15803d);
                border-color: #16a34a;
                border-width: 2px;
                padding: 14px 32px;
            }

            QPushButton#downloadButton {
                padding: 14px 32px;
                border: 2px solid #3b82f6;
                border-radius: 10px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3b82f6, stop:0.5 #2563eb, stop:1 #1d4ed8);
                color: white;
                font-size: 17px;
                font-weight: 700;
                min-height: 56px;
                letter-spacing: 0.5px;
            }

            QPushButton#downloadButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #93c5fd, stop:0.5 #60a5fa, stop:1 #3b82f6);
                border-color: #93c5fd;
                border-width: 3px;
                padding: 13px 31px;
            }

            QPushButton#downloadButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2563eb, stop:0.5 #1d4ed8, stop:1 #1e40af);
                border-color: #1d4ed8;
                border-width: 2px;
                padding: 14px 32px;
            }

            QPushButton#downloadButton:disabled {
                background: #1a1c24;
                color: #464a5e;
                border-color: #31343f;
                border-width: 1px;
            }

            QPushButton#secondaryButton {
                padding: 10px 20px;
                border: 2px solid #464a5e;
                border-radius: 8px;
                background-color: #262730;
                color: #fafafa;
                font-size: 15px;
                font-weight: 500;
                min-height: 40px;
            }

            QPushButton#secondaryButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3d4150, stop:1 #2d3038);
                border-color: #34d399;
                border-width: 2px;
                color: #34d399;
            }

            QPushButton#secondaryButton:pressed {
                background-color: #1c1f26;
                border-color: #2d8659;
                color: #2d8659;
            }

            QTableWidget {
                border: 1px solid #31343f;
                border-radius: 8px;
                background-color: #1a1c24;
                font-size: 16px;
                color: #fafafa;
            }

            QTableWidget::item {
                padding: 16px;
                color: #fafafa;
                background-color: #1a1c24;
                font-size: 16px;
                border: none;
            }

            QTableWidget::item:selected {
                background-color: #262730;
                color: #fafafa;
            }

            QHeaderView::section {
                background-color: #262730;
                padding: 16px;
                border: none;
                border-bottom: 2px solid #31343f;
                font-weight: 600;
                font-size: 15px;
                color: #fafafa;
            }

            QTableWidget::item:alternate {
                background-color: #14161d;
            }

            QScrollArea {
                border: none;
                background-color: #0e1117;
            }

            QScrollBar:vertical {
                border: none;
                background: #1a1c24;
                width: 12px;
                margin: 0px;
            }

            QScrollBar::handle:vertical {
                background: #464a5e;
                border-radius: 6px;
                min-height: 30px;
            }

            QScrollBar::handle:vertical:hover {
                background: #5a5f75;
            }

            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)

    def on_program_changed(self, text):
        self.program_type = text

    def on_year_changed(self, text):
        self.start_year = int(text)

    def load_example(self):
        self.periods_data = [
            {"Год": 1, "Семестр": 1, "Тип": "Т", "Недели": 10},
            {"Год": 1, "Семестр": 1, "Тип": "П", "Недели": 12},
            {"Год": 1, "Семестр": 1, "Тип": "ПА", "Недели": 1},
            {"Год": 1, "Семестр": 2, "Тип": "Т", "Недели": 4},
            {"Год": 1, "Семестр": 2, "Тип": "П", "Недели": 16},
            {"Год": 1, "Семестр": 2, "Тип": "ПА", "Недели": 1},
            {"Год": 1, "Семестр": 2, "Тип": "К", "Недели": 6},
            {"Год": 2, "Семестр": 1, "Тип": "Т", "Недели": 10},
            {"Год": 2, "Семестр": 1, "Тип": "П", "Недели": 12},
            {"Год": 2, "Семестр": 1, "Тип": "ПА", "Недели": 1},
            {"Год": 2, "Семестр": 2, "Тип": "Т", "Недели": 9},
            {"Год": 2, "Семестр": 2, "Тип": "П", "Недели": 8},
            {"Год": 2, "Семестр": 2, "Тип": "ПА", "Недели": 1},
            {"Год": 2, "Семестр": 2, "Тип": "ГИА", "Недели": 2},
            {"Год": 2, "Семестр": 2, "Тип": "К", "Недели": 6}
        ]
        self.update_table()

    def clear_data(self):
        self.periods_data = []
        self.update_table()
        self.preview_section.setVisible(False)
        self.download_btn.setEnabled(False)

    def add_row(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)

        # Стиль для комбобоксов в таблице
        table_combo_style = """
            QComboBox {
                padding: 11px 16px;
                border: 1px solid #464a5e;
                border-radius: 6px;
                background-color: #262730;
                font-size: 16px;
                color: #fafafa;
            }
            QComboBox:hover {
                border-color: #34d399;
                border-width: 2px;
                background-color: #2d3038;
            }
            QComboBox::drop-down {
                border: none;
                width: 28px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #a3a8b4;
                margin-right: 10px;
            }
            QComboBox:hover::down-arrow {
                border-top-color: #34d399;
            }
            QComboBox QAbstractItemView {
                background-color: #262730;
                border: 1px solid #464a5e;
                selection-background-color: #2d8659;
                selection-color: #ffffff;
                font-size: 16px;
            }
            QComboBox QAbstractItemView::item {
                padding: 12px 14px;
                min-height: 38px;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: #34d399;
            }
        """

        # Год
        year_combo = QComboBox()
        year_combo.addItems(['1', '2', '3'])
        year_combo.setStyleSheet(table_combo_style)

        year_container = QWidget()
        year_container.setStyleSheet("background-color: transparent;")
        year_layout = QHBoxLayout(year_container)
        year_layout.addWidget(year_combo)
        year_layout.setContentsMargins(6, 0, 6, 0)
        year_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setCellWidget(row_position, 0, year_container)

        # Семестр
        semester_combo = QComboBox()
        semester_combo.addItems(['1', '2'])
        semester_combo.setStyleSheet(table_combo_style)

        semester_container = QWidget()
        semester_container.setStyleSheet("background-color: transparent;")
        semester_layout = QHBoxLayout(semester_container)
        semester_layout.addWidget(semester_combo)
        semester_layout.setContentsMargins(6, 0, 6, 0)
        semester_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setCellWidget(row_position, 1, semester_container)

        # Тип
        type_combo = QComboBox()
        type_combo.addItems(['Т', 'П', 'ПА', 'ГИА', 'К'])
        type_combo.setStyleSheet(table_combo_style)

        type_container = QWidget()
        type_container.setStyleSheet("background-color: transparent;")
        type_layout = QHBoxLayout(type_container)
        type_layout.addWidget(type_combo)
        type_layout.setContentsMargins(6, 0, 6, 0)
        type_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.table.setCellWidget(row_position, 2, type_container)

        # Недели
        weeks_item = QTableWidgetItem('1.0')
        weeks_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
        font = QFont()
        font.setPointSize(16)
        weeks_item.setFont(font)
        self.table.setItem(row_position, 3, weeks_item)

        self.table.setRowHeight(row_position, 82)

    def remove_row(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)

    def update_table(self):
        self.table.setRowCount(0)

        # Стиль для комбобоксов в таблице
        table_combo_style = """
            QComboBox {
                padding: 11px 16px;
                border: 1px solid #464a5e;
                border-radius: 6px;
                background-color: #262730;
                font-size: 16px;
                color: #fafafa;
            }
            QComboBox:hover {
                border-color: #34d399;
                border-width: 2px;
                background-color: #2d3038;
            }
            QComboBox::drop-down {
                border: none;
                width: 28px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 7px solid #a3a8b4;
                margin-right: 10px;
            }
            QComboBox:hover::down-arrow {
                border-top-color: #34d399;
            }
            QComboBox QAbstractItemView {
                background-color: #262730;
                border: 1px solid #464a5e;
                selection-background-color: #2d8659;
                selection-color: #ffffff;
                font-size: 16px;
            }
            QComboBox QAbstractItemView::item {
                padding: 12px 14px;
                min-height: 38px;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: #34d399;
            }
        """

        for data in self.periods_data:
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)

            # Год
            year_combo = QComboBox()
            year_combo.addItems(['1', '2', '3'])
            year_combo.setCurrentText(str(data['Год']))
            year_combo.setStyleSheet(table_combo_style)

            year_container = QWidget()
            year_container.setStyleSheet("background-color: transparent;")
            year_layout = QHBoxLayout(year_container)
            year_layout.addWidget(year_combo)
            year_layout.setContentsMargins(6, 0, 6, 0)
            year_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setCellWidget(row_position, 0, year_container)

            # Семестр
            semester_combo = QComboBox()
            semester_combo.addItems(['1', '2'])
            semester_combo.setCurrentText(str(data['Семестр']))
            semester_combo.setStyleSheet(table_combo_style)

            semester_container = QWidget()
            semester_container.setStyleSheet("background-color: transparent;")
            semester_layout = QHBoxLayout(semester_container)
            semester_layout.addWidget(semester_combo)
            semester_layout.setContentsMargins(6, 0, 6, 0)
            semester_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setCellWidget(row_position, 1, semester_container)

            # Тип
            type_combo = QComboBox()
            type_combo.addItems(['Т', 'П', 'ПА', 'ГИА', 'К'])
            type_combo.setCurrentText(data['Тип'])
            type_combo.setStyleSheet(table_combo_style)

            type_container = QWidget()
            type_container.setStyleSheet("background-color: transparent;")
            type_layout = QHBoxLayout(type_container)
            type_layout.addWidget(type_combo)
            type_layout.setContentsMargins(6, 0, 6, 0)
            type_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setCellWidget(row_position, 2, type_container)

            # Недели
            weeks_item = QTableWidgetItem(str(data['Недели']))
            weeks_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
            font = QFont()
            font.setPointSize(16)
            weeks_item.setFont(font)
            self.table.setItem(row_position, 3, weeks_item)

            self.table.setRowHeight(row_position, 82)

    def get_table_data(self):
        data = []
        for row in range(self.table.rowCount()):
            year_container = self.table.cellWidget(row, 0)
            semester_container = self.table.cellWidget(row, 1)
            type_container = self.table.cellWidget(row, 2)
            weeks_item = self.table.item(row, 3)

            if year_container and semester_container and type_container and weeks_item:
                try:
                    # Извлекаем комбобоксы из контейнеров
                    year_combo = year_container.findChild(QComboBox)
                    semester_combo = semester_container.findChild(QComboBox)
                    type_combo = type_container.findChild(QComboBox)

                    if year_combo and semester_combo and type_combo:
                        weeks = float(weeks_item.text())
                        data.append({
                            'Год': int(year_combo.currentText()),
                            'Семестр': int(semester_combo.currentText()),
                            'Тип': type_combo.currentText(),
                            'Недели': weeks
                        })
                except ValueError:
                    pass
        return data

    def generate_schedule(self):
        periods_data = self.get_table_data()

        if not periods_data:
            QMessageBox.warning(self, 'Внимание', 'Добавьте периоды обучения')
            return

        try:
            self.generated_schedule = self.app.generate_schedule(periods_data, self.start_year)

            # Обновление таблицы предварительного просмотра
            self.preview_table.setRowCount(0)
            for period in self.generated_schedule:
                row_position = self.preview_table.rowCount()
                self.preview_table.insertRow(row_position)

                items = [
                    str(period['year']),
                    str(period['semester']),
                    period['type'],
                    f"{period['weeks']:.1f}",
                    period['start_date'].strftime('%d.%m.%Y'),
                    period['end_date'].strftime('%d.%m.%Y'),
                    str(len(period['days']))
                ]

                for col, text in enumerate(items):
                    item = QTableWidgetItem(text)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter)
                    font = QFont()
                    font.setPointSize(16)
                    item.setFont(font)
                    self.preview_table.setItem(row_position, col, item)

                self.preview_table.setRowHeight(row_position, 60)

            self.preview_section.setVisible(True)
            self.download_btn.setEnabled(True)

            QMessageBox.information(self, 'Успех', f'График готов!\nСоздано периодов: {len(self.generated_schedule)}')

        except Exception as e:
            QMessageBox.critical(self, 'Ошибка', f'Ошибка при генерации графика:\n{str(e)}')

    def download_excel(self):
        if not self.generated_schedule:
            QMessageBox.warning(self, 'Внимание', 'Сначала сгенерируйте график')
            return

        filename, _ = QFileDialog.getSaveFileName(
            self,
            'Сохранить Excel файл',
            f'график_{self.start_year}-{self.start_year + 1}.xlsx',
            'Excel Files (*.xlsx)'
        )

        if filename:
            try:
                wb = self.app.create_excel_file(self.generated_schedule, self.start_year, self.program_type)
                wb.save(filename)
                QMessageBox.information(self, 'Успех', f'Файл успешно сохранен:\n{filename}')
            except Exception as e:
                QMessageBox.critical(self, 'Ошибка', f'Ошибка при сохранении файла:\n{str(e)}')


def main():
    app = QApplication(sys.argv)

    # Установка шрифта
    font = QFont()
    font.setPointSize(10)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()