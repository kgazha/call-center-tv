# -*- coding: utf-8 -*-
import MySQLdb
import pandas as pd
import numpy as np
import configparser
import xlsxwriter
import datetime
import os
from collections import defaultdict


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

    def get_data_from_db(self, filename, *args):
        sql_form = open(filename).read()
        if args:
            self.cursor.execute(sql_form.format(*args))
        else:
            self.cursor.execute(sql_form)
        self.data = self.cursor.fetchall()

    def data_to_form_template(self):
        pass

    def form_to_excel(self):
        pass

    @staticmethod
    def get_header_format(workbook):
        header_format = workbook.add_format()
        header_format.set_bold()
        header_format.set_text_wrap()
        header_format.set_align('center')
        header_format.set_align('vcenter')
        return header_format

    @staticmethod
    def get_row_format(workbook):
        row_format = workbook.add_format()
        row_format.set_text_wrap()
        row_format.set_align('left')
        row_format.set_align('top')
        return row_format


class RecordForm02:
    def __init__(self):
        self.name = ''
        self.locality = ''
        self.street = ''
        self.product_type = ''
        self.store = ''
        self.post_office = ''


class RecordForm03:
    def __init__(self):
        self.name = ''
        self.locality = ''
        self.product_type = ''
        self.social_category = ''
        self.address = ''


class RecordForm04:
    def __init__(self):
        self.name = ''
        self.locality = ''
        self.address = ''
        self.phone_number = ''
        self.create_time = ''
        self.empty_field = ''
        self.operator = ''


class ReportForm01(Report):
    def __init__(self):
        super().__init__()
        self.form_name = 'Форма_1'
        self.themes = {
            9: 'Выбор и покупка приемного оборудования (телевизор, приставка, антенна)',
            8: 'Социальная поддержка льготных категорий граждан',
            10: 'Вещание на территориях вне зоны цифрового сигнала',
            11: 'Вызов волонтеров на подключение оборудования',
            12: 'Подключение к системе коллективного приема телевидения (СКПТ)',
            13: 'Вещание региональных каналов',
            14: 'Иное',
        }

    def get_data_from_db(self):
        start_date = config['REPORT_FORM_01']['START_DATE']
        super(ReportForm01, self).get_data_from_db('form_01.sql', start_date)

    def data_to_form_template(self):
        _values = list(set(map(lambda x: self.themes[x['ticket_type_id']], self.data)))
        _index = list(set(map(lambda x: x['value_text'], self.data)))
        df = pd.DataFrame(0, index=_index, columns=_values)
        for row in self.data:
            df.at[row['value_text'], self.themes[row['ticket_type_id']]] = row['frequency']
        df.at['Итого'] = 0
        for key in df.keys():
            df.at['Итого', key] = sum(df[key])
        self.form = df

    def form_to_excel(self):
        file_name = self.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(BASE_DIR, self.form_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:H', 20)
        worksheet.set_row(0, 80)
        header_format = self.get_header_format(workbook)
        row_format = self.get_row_format(workbook)
        worksheet.write(0, 0, 'Наименование ОМСУ', header_format)
        for idx, theme_key in enumerate(self.themes, start=1):
            worksheet.write(0, idx, self.themes[theme_key], header_format)
        for row in range(1, len(self.form) + 1):
            for col in range(1, len(self.themes) + 1):
                worksheet.write(row, col, 0, row_format)
        for idx_row, (row_name, row_data) in enumerate(self.form.iterrows(), start=1):
            worksheet.write(idx_row, 0, row_name)
            for col_name, col_value in row_data.items():
                for idx_theme, (theme_key, theme_value) in enumerate(self.themes.items(), start=1):
                    if theme_value == col_name:
                        worksheet.write(idx_row, idx_theme, col_value, row_format)
                        break
        workbook.close()


class ReportForm02(Report):
    def __init__(self):
        super().__init__()
        self.form_name = 'Форма_2'
        self.header = [
            'ФИО',
            'Наименование населенного пункта',
            'Наименование улицы (номер дома)',
            'Наименование товара, который хотел приобрести Абонент',
            'Рекомендованные торговые сети',
            'Рекомендованные почтовые отделения',
        ]

    def get_data_from_db(self):
        start_date = config['REPORT_FORM_02']['START_DATE']
        super(ReportForm02, self).get_data_from_db('form_02.sql', start_date)

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        ticket_ids = list(set(df['ticket_id']))
        data = defaultdict(list)
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            record = RecordForm02()
            if not ticket_df[ticket_df['field_id'] == 12]['value_text'].empty:
                record.name = ticket_df[ticket_df['field_id'] == 12]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 15]['value_text'].empty:
                record.locality = ticket_df[ticket_df['field_id'] == 15]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 23]['value_text'].empty:
                record.street = ticket_df[ticket_df['field_id'] == 23]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 13]['value_text'].empty:
                record.product_type = ticket_df[ticket_df['field_id'] == 13]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 19]['value_text'].empty:
                record.store += 'Магазин 1: ' + ticket_df[ticket_df['field_id'] == 19]['value_text'].item() + '\n'
            if not ticket_df[ticket_df['field_id'] == 20]['value_text'].empty:
                record.store += 'Магазин 2: ' + ticket_df[ticket_df['field_id'] == 20]['value_text'].item() + '\n'
            if not ticket_df[ticket_df['field_id'] == 21]['value_text'].empty:
                record.store += 'Магазин 3: ' + ticket_df[ticket_df['field_id'] == 21]['value_text'].item() + '\n'
            if not ticket_df[ticket_df['field_id'] == 22]['value_text'].empty:
                record.post_office = ticket_df[ticket_df['field_id'] == 22]['value_text'].item()
            record.store = record.store.strip()
            data[ticket_df[ticket_df['field_id'] == 14]['value_text'].item()].append(record.__dict__)
        self.form = dict(data)

    def form_to_excel(self):
        file_name = self.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(BASE_DIR, self.form_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        for record in self.form.items():
            worksheet = workbook.add_worksheet(name=record[0])
            worksheet.set_column('A:F', 20)
            worksheet.set_row(0, 30)
            worksheet.set_row(1, 80)
            header_format = self.get_header_format(workbook)
            row_format = self.get_row_format(workbook)
            worksheet.merge_range('A1:F1', record[0], header_format)
            for idx, key in enumerate(self.header):
                worksheet.write(1, idx, key, header_format)
            for row_idx, data in enumerate(record[1], start=2):
                for col_idx, (key, value) in enumerate(data.items()):
                    worksheet.write(row_idx, col_idx, value, row_format)
        workbook.close()


class ReportForm03(Report):
    def __init__(self):
        super().__init__()
        self.form_name = 'Форма_3'
        self.header = [
            'ФИО',
            'Наименование населенного пункта',
            'Наименование товара, (ТВ-приставка, спутниковое обрудование)',
            'Социальная категория от Абонента',
            'Наименование УСЗН, в которое направили Абонента',
        ]

    def get_data_from_db(self):
        start_date = config['REPORT_FORM_03']['START_DATE']
        super(ReportForm03, self).get_data_from_db('form_03.sql', start_date)

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        ticket_ids = list(set(df['ticket_id']))
        data = defaultdict(list)
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            record = RecordForm03()
            if not ticket_df[ticket_df['field_id'] == 12]['value_text'].empty:
                record.name = ticket_df[ticket_df['field_id'] == 12]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 15]['value_text'].empty:
                record.locality = ticket_df[ticket_df['field_id'] == 15]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 13]['value_text'].empty:
                record.product_type = ticket_df[ticket_df['field_id'] == 13]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 24]['value_text'].empty:
                record.social_category = ticket_df[ticket_df['field_id'] == 24]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 36]['value_text'].empty:
                record.address = ticket_df[ticket_df['field_id'] == 36]['value_text'].item()
            data[ticket_df[ticket_df['field_id'] == 14]['value_text'].item()].append(record.__dict__)
        self.form = dict(data)

    def form_to_excel(self):
        file_name = self.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(BASE_DIR, self.form_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        for record in self.form.items():
            worksheet = workbook.add_worksheet(name=record[0])
            worksheet.set_column('A:E', 20)
            worksheet.set_row(0, 30)
            worksheet.set_row(1, 80)
            header_format = self.get_header_format(workbook)
            row_format = self.get_row_format(workbook)
            worksheet.merge_range('A1:E1', record[0], header_format)
            for idx, key in enumerate(self.header):
                worksheet.write(1, idx, key, header_format)
            for row_idx, data in enumerate(record[1], start=2):
                for col_idx, (key, value) in enumerate(data.items()):
                    worksheet.write(row_idx, col_idx, value, row_format)
        workbook.close()


class ReportForm04(Report):
    def __init__(self):
        super().__init__()
        self.form_name = 'Форма_4'
        self.header = [
            'ФИО',
            'Наименование населенного пункта',
            'Точный адрес',
            'Телефонный номер',
            'Дата формирования заявки',
            'Отработанные заявки',
            'Оператор',
        ]

    def get_data_from_db(self):
        start_date = config['REPORT_FORM_04']['START_DATE']
        super(ReportForm04, self).get_data_from_db('form_04.sql', start_date)

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        if df.empty:
            return
        ticket_ids = list(set(df['ticket_id']))
        data = defaultdict(list)
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            record = RecordForm04()
            if not ticket_df[ticket_df['field_id'] == 12]['value_text'].empty:
                record.name = ticket_df[ticket_df['field_id'] == 12]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 15]['value_text'].empty:
                record.locality = ticket_df[ticket_df['field_id'] == 15]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 17]['value_text'].empty:
                record.address = ticket_df[ticket_df['field_id'] == 17]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 16]['value_text'].empty:
                record.phone_number = ticket_df[ticket_df['field_id'] == 16]['value_text'].item()
            if not ticket_df[ticket_df['field_id'] == 18]['value_text'].empty:
                record.operator = ticket_df[ticket_df['field_id'] == 18]['value_text'].item()
            record.create_time = str(ticket_df['create_time'].iloc[0])
            data[ticket_df[ticket_df['field_id'] == 14]['value_text'].item()].append(record.__dict__)
        self.form = dict(data)

    def form_to_excel(self):
        if not self.form:
            return
        file_name = self.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(BASE_DIR, self.form_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        for record in self.form.items():
            worksheet = workbook.add_worksheet(name=record[0])
            worksheet.set_column('A:G', 20)
            worksheet.set_row(0, 30)
            worksheet.set_row(1, 80)
            header_format = self.get_header_format(workbook)
            row_format = self.get_row_format(workbook)
            worksheet.merge_range('A1:G1', record[0], header_format)
            for idx, key in enumerate(self.header):
                worksheet.write(1, idx, key, header_format)
            for row_idx, data in enumerate(record[1], start=2):
                for col_idx, (key, value) in enumerate(data.items()):
                    worksheet.write(row_idx, col_idx, value, row_format)
        workbook.close()


class ReportForm51(Report):
    def __init__(self):
        super().__init__()
        self.form_name = 'Форма_5_1'
        self.header = [
            'Наименование ОМСУ',
            'Всего обращений',
            'Закрытые заявки',
            'В работе',
            'Не взяты в работу в течение 3 дней',
            'Взяты в работу, но не закрыты в течение 10 дней',
            'Повторные обращения',
        ]

    def get_data_from_db(self):
        start_date = config['REPORT_FORM_51']['START_DATE']
        super(ReportForm51, self).get_data_from_db('form_51.sql', start_date)

    def data_to_form_template(self):
        pass

    def form_to_excel(self):
        pass


class ReportFacade:
    reports = None

    @classmethod
    def create_reports(cls):
        cls.reports = [
            ReportForm01(),
            ReportForm02(),
            ReportForm03(),
            ReportForm04(),
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
