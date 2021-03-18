import xlrd
import xlwt
import pylightxl as xl
import os
import sys
import re
import PySimpleGUI as sg
from collections import OrderedDict


class SalaryCounter:

    def __init__(self, values):
        self.values = values
        self.payrolls = {}
        self.headers = []
        self.counter = False
        self.main_col_num = None
        self.sum = None

    @staticmethod
    def find_path():
        if getattr(sys, 'frozen', False):
            # one-file
            application_path = os.path.dirname(sys.executable)
            # one-folder
            # application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        return application_path

    def parse_directory(self):
        application_path = self.find_path()
        for dir_path, _, file_names in os.walk(application_path):
            if dir_path == application_path:
                for file_name in file_names:
                    self.check_format(file_name=file_name)

    def check_format(self, file_name: str):
        pattern = re.compile(r"^([^\\]{1,50}).(xls|xlsx)$")
        result = re.search(pattern=pattern, string=file_name)
        if result:
            if result[2] == 'xls' or result[2] == 'xlsx':
                self.payrolls[file_name] = result[2]

    def read(self, payroll: str, format_: str):
        to_count = []
        if format_ == 'xls':
            book = xlrd.open_workbook(payroll)
            sheet = book.sheet_by_index(0)
            for row in range(sheet.nrows):
                new_row = []
                for element in sheet.row(row):
                    value = self._read_helper(value=element.value, row=sheet.row(row), format_=format_)
                    if value:
                        new_row.append(value)
                if new_row:
                    to_count.append(new_row)
        elif format_ == 'xlsx':
            db = xl.readxl(fn=payroll)
            ws_name = db.ws_names[0]
            for row in db.ws(ws=ws_name).rows:
                new_row = []
                for element in row:
                    value = self._read_helper(value=element, row=row, format_=format_)
                    if value:
                        new_row.append(value)
                if new_row:
                    to_count.append(new_row)
        to_count.pop(0)
        to_count.pop(0)
        return to_count

    def _read_helper(self, value, row, format_: str):
        if value == "Фамилия, имя, отчество":
            if format_ == 'xls':
                self.headers = [el.value for el in row if
                                el.value and el.value != "Расписка в получении"]
            elif format_ == 'xlsx':
                self.headers = [el for el in row if el and el != "Расписка в получении"]
            self.main_col_num = ([self.headers.index(i) for i in self.headers if i == "Сумма"])[0]
            self.headers.extend([str(i) + "р" for i in self.values])
            self.counter = True
        elif value == "Итого":
            self.counter = False
        if self.counter and value:
            return value

    def count(self, data: list):
        result = []
        for_person = None
        for row in data:
            for num, element in enumerate(row):
                if num == self.main_col_num:
                    for_person = self._count_one_instance(salary=element)
            if for_person:
                row.extend(for_person)
            result.append(row)
        return result

    def _count_one_instance(self, salary: float):
        result = []
        for value in self.values:
            amount = 0
            if value not in self.sum:
                self.sum.setdefault(value, 0)
            while salary >= value:
                salary -= value
                amount += 1
                self.sum[value] += 1
            result.append(amount)
        return result

    def save(self, data: list, format_: str, name: str):
        data.append(self.sum.keys())
        data.append(self.sum.values())
        name = name[:len(name) - (len(format_) + 1)]
        file_name = name + "_расчёт" + "." + format_
        sheetname = "Зарплаты"
        if format_ == 'xls':
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet(sheetname=sheetname)
            for col_num, column in enumerate(self.headers):
                sheet.write(r=0, c=col_num, label=column)
            for num, row in enumerate(data, start=1):
                for col_num, column in enumerate(row):
                    sheet.write(r=num, c=col_num, label=column)
            workbook.save(filename_or_stream=file_name)
        elif format_ == 'xlsx':
            new_db = xl.Database()
            new_db.add_ws(ws=sheetname)
            for col_num, column in enumerate(self.headers, start=1):
                new_db.ws(ws=sheetname).update_index(row=1, col=col_num, val=column)
            for num, row in enumerate(data, start=2):
                for col_num, column in enumerate(row, start=1):
                    new_db.ws(ws=sheetname).update_index(row=num, col=col_num, val=column)
            xl.writexl(db=new_db, fn=file_name)

    def run(self):
        self.parse_directory()
        if not self.payrolls:
            raise ImportError("Файлов с расширением .xls/.xlsx в папке нет")
        for payroll, format_ in self.payrolls.items():
            self.sum = OrderedDict()
            to_count = self.read(payroll=payroll, format_=format_)
            result = self.count(data=to_count)
            self.save(data=result, name=payroll, format_=format_)


class Interface:

    def __init__(self):
        self.title = 'Счётчик купюр/монет'
        self.theme = 'DarkAmber'
        self.layout = list()
        self.main_window = None
        self.values = [5000, 2000, 1000, 500, 200, 100, 50, 10, 5, 2, 1, 0.50, 0.10]

    def run(self):
        self.start_window()
        while True:
            event, values = self.main_window.read()
            if event == sg.WIN_CLOSED or event == "Отмена":
                break
            elif event == "Посчитать":
                values_to_use = []
                for value in self.values:
                    if values[f'check{value}']:
                        values_to_use.append(value)
                try:
                    counter = SalaryCounter(values=values_to_use)
                    counter.run()
                except Exception as exc:
                    exc_popup = self.popup_window(title="Ошибка", text=exc)
                    pop_event, pop_value = exc_popup.read()
                    if pop_event == sg.WIN_CLOSED:
                        exc_popup.close()
            self.main_window.close()

    def popup_window(self, text: Exception, title: str):
        sg.theme(self.theme)
        layout = [[sg.Text(text=text)]]

        size = 100
        popup = sg.Window(
            title=title,
            layout=layout,
            default_button_element_size=(10, 2),
            size=(size * 5, size),
            element_padding=(10, 10),
            auto_size_buttons=False,
        )
        return popup

    def start_window(self):
        sg.theme(self.theme)

        self.layout.append([sg.Text(text="Пересчитать платёжные ведомости формата .xls/.xlsx, "
                                    "находящиеся в папке?")])
        self.layout.append([sg.Text(text="Используемые купюры:")])
        for value in self.values:
            spaces = (4 - len(str(value))) * " "
            self.layout.append([sg.Text(f"{value}{spaces}"), sg.Checkbox(text="", default=True, key=f'check{value}')])
        self.layout.append([sg.Button("Посчитать"), sg.Button("Отмена")])

        window = sg.Window(
            title=self.title,
            layout=self.layout,
            default_button_element_size=(10, 2),
            size=(550, 500),
            element_padding=(2, 2),
            auto_size_buttons=False,
        )
        self.main_window = window


if __name__ == '__main__':
    ui = Interface()
    ui.run()
