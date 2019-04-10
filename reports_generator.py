# -*- coding: utf-8 -*-
import MySQLdb
import pandas as pd
import numpy as np
import configparser
import xlsxwriter
import datetime
import os


config = configparser.ConfigParser()
config.read('settings.ini')
BASE_DIR = os.path.dirname(os.path.realpath(__file__))

db = MySQLdb.connect(config['CONNECTION']['HOST'],
                     config['CONNECTION']['USER'],
                     config['CONNECTION']['PASSWORD'],
                     config['CONNECTION']['DATABASE'],
                     charset='utf8',
                     init_command='SET NAMES UTF8')

class Report:
    def __init__(self):
        self.cursor = db.cursor(MySQLdb.cursors.DictCursor)
        self.data = None
        self.form = None
        self.themes = [
            'Выбор и покупка приемного оборудования (телевизор, приставка, антенна)',
            'Социальная поддержка льготных категорий граждан',
            'Вещание на территориях вне зоны цифрового сигнала',
            'Вызов волонтеров на подключение оборудования',
            'Подключение к системе коллективного приема телевидения (СКПТ)',
            'Вещание региональных каналов',
            'Иное'
        ]

    def get_data_from_db(self, filename):
        sql_form = open(filename).read()
        self.cursor.execute(sql_form)
        self.data = self.cursor.fetchall()

    def data_to_form_template(self):
        pass

    def form_to_excel(self):
        pass
    

class ReportForm01(Report):
    def __init__(self):
        super().__init__()
        self.folder_name = 'Форма_1'

    def get_data_from_db(self):
        super(ReportForm01, self).get_data_from_db('form_01.sql')

    def data_to_form_template(self):
        _values = list(set(map(lambda x: x['name'], self.data)))
        _index = list(set(map(lambda x: x['value_text'], self.data)))
        df = pd.DataFrame(0, index=_index, columns=_values)
        for row in self.data:
            df.at[row['value_text'], row['name']] = row['frequency']
        for key in df.keys():
            df.at['Итого', key] = sum(df[key][:-1])
        self.form = df

    def form_to_excel(self):
        file_name = datetime.date.today().strftime("%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(BASE_DIR, self.folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:H', 20)
        worksheet.set_row(0, 80)
        header_format = workbook.add_format()
        header_format.set_bold()
        header_format.set_text_wrap()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        worksheet.write(0, 0, 'Наименование ОМСУ', header_format)
        for idx, theme in enumerate(self.themes, start=1):
            worksheet.write(0, idx, theme, header_format)
        for row_idx, form_row in enumerate(self.form.iterrows(), start=1):
            worksheet.write(row_idx, 0, form_row[0])
            for col_idx, row in enumerate(form_row[1], start=1):
                worksheet.write(row_idx, col_idx, row)
        workbook.close()


class ReportFacade:
    reports = None

    @classmethod
    def create_reports(cls):
        cls.reports = [
            ReportForm01(),
        ]

    @classmethod
    def data_to_excel(cls):
        for report in cls.reports:
            report.get_data_from_db()
            report.data_to_form_template()
            report.form_to_excel()


if __name__ == '__main__':
    ReportFacade.create_reports()
    ReportFacade.data_to_excel()
