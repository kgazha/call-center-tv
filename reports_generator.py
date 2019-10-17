# -*- coding: utf-8 -*-
import MySQLdb
import pandas as pd
import configparser
import xlsxwriter
import datetime
import os
import argparse
from collections import defaultdict
from working_time import compute_working_time


config = configparser.ConfigParser()
config.read('settings.ini')
BASE_DIR = 'Z:\Отчеты OTRS\CallCenter'
report_dates = configparser.ConfigParser()
report_dates.read(os.path.join(BASE_DIR, 'report_dates.ini'))
CURRENT_DATE = datetime.datetime.now().strftime('%Y-%m-%d')
CURRENT_DATE_TIME = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
TOMORROW = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
YESTERDAY = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')

MAX_WORKING_DAYS_DICT = {
    ('2019-04-11', '2019-07-29'): 10,
    ('2019-07-29', '2019-07-16'): 5,
    ('2019-07-29', TOMORROW): 3,
}

db = MySQLdb.connect(config['CONNECTION']['HOST'],
                     config['CONNECTION']['USER'],
                     config['CONNECTION']['PASSWORD'],
                     config['CONNECTION']['DATABASE'],
                     charset='utf8',
                     init_command='SET NAMES UTF8')


class RecordTypes:
    VOLUNTEERS = {'code': 11, 'name': 'Вызов волонтеров на подключение оборудования'}
    OTHER = {'code': 14, 'name': 'Иное'}
    PURCHASE = {'code': 9, 'name': 'Выбор и покупка приемного оборудования (телевизор, приставка, антенна)'}
    SOCIAL = {'code': 8, 'name': 'Социальная поддержка льготных категорий граждан'}
    COMPLAINTS = {'code': 15, 'name': 'Жалобы'}
    BROADCASTING_OUTSIDE = {'code': 10, 'name': 'Вещание на территориях вне зоны цифрового сигнала'}
    CONNECTION = {'code': 12, 'name': 'Подключение к системе коллективного приема телевидения (СКПТ)'}
    BROADCASTING_REGIONAL = {'code': 13, 'name': 'Вещание региональных каналов'}
    TOTAL = {'code': None, 'name': 'Итого'}
    DATA = {'code': None, 'name': 'Дата'}

    @staticmethod
    def get_record_types():
        return [k for k in RecordTypes.__dict__.keys() if not k.startswith('_') and k.isupper()]

    @staticmethod
    def get_queues():
        record_types = RecordTypes.get_record_types()
        return [k for k in record_types if getattr(RecordTypes, k)['code']]

    @staticmethod
    def get_record_queue_by_code(code):
        queues = RecordTypes.get_queues()
        for k in queues:
            if getattr(RecordTypes, k)['code'] == code:
                return getattr(RecordTypes, k)
        return None


def get_max_working_days(date):
    for key in MAX_WORKING_DAYS_DICT.keys():
        if datetime.datetime.strptime(key[0], '%Y-%m-%d') < datetime.datetime.strptime(date, '%Y-%m-%d %H:%M:%S') < \
                datetime.datetime.strptime(key[1], '%Y-%m-%d'):
            return MAX_WORKING_DAYS_DICT[key]
    raise ValueError("Incorrect date")


class Report:
    def __init__(self, **kwargs):
        self.cursor = db.cursor(MySQLdb.cursors.DictCursor)
        self.data = None
        self.form = None
        self.form_name = None
        self.form_verbose_name = None
        self.header = None
        self.start_date = None
        self.end_date = None
        self.daily = False
        self.result_folder_path = BASE_DIR
        if 'daily' in kwargs:
            if kwargs['daily']:
                self.daily = kwargs['daily']
        if 'path' in kwargs:
            if kwargs['path']:
                self.result_folder_path = kwargs['path']

    def get_data_from_db(self, filename, *args):
        sql_form = open(filename).read()
        if args:
            self.cursor.execute(sql_form.format(*args))
        else:
            self.cursor.execute(sql_form)
        self.data = self.cursor.fetchall()

    def init_dates(self):
        if self.daily:
            self.start_date = CURRENT_DATE
        elif 'START_DATE' in report_dates[self.form_name]:
            self.start_date = report_dates[self.form_name]['START_DATE']
        if 'END_DATE' not in report_dates[self.form_name]:
            self.end_date = TOMORROW
        else:
            self.end_date = report_dates[self.form_name]['END_DATE']

    def data_to_form_template(self):
        pass

    def form_to_file(self):
        pass

    def form_to_excel_by_territory(self, column_range, header_merge_range):
        if not self.form:
            return
        file_name = self.form_verbose_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        for record in self.form.items():
            worksheet = workbook.add_worksheet(name=record[0])
            worksheet.set_column(column_range, 20)
            worksheet.set_row(0, 30)
            worksheet.set_row(1, 80)
            header_format = self.get_header_format(workbook)
            row_format = self.get_row_format(workbook)
            worksheet.merge_range(header_merge_range, record[0], header_format)
            for idx, key in enumerate(self.header):
                worksheet.write(1, idx, key, header_format)
            for row_idx, data in enumerate(record[1], start=2):
                for col_idx, (key, value) in enumerate(data.items()):
                    worksheet.write(row_idx, col_idx, value, row_format)
        workbook.close()

    def form_to_excel_aggregated(self, column_range):
        if not self.form:
            return
        file_name = self.form_verbose_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:B', 40)
        worksheet.set_column(column_range, 20)
        worksheet.set_row(0, 80)
        header_format = self.get_header_format(workbook)
        row_format = self.get_row_format(workbook)
        for idx, key in enumerate(self.header):
            worksheet.write(0, idx, key, header_format)
        row_idx = 1
        for record in self.form.items():
            for data in record[1]:
                worksheet.write(row_idx, 0, record[0], row_format)
                for col_idx, (key, value) in enumerate(data.items(), start=1):
                    worksheet.write(row_idx, col_idx, value, row_format)
                row_idx += 1
        workbook.close()

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


class RecordForm542:
    def __init__(self):
        self.name = ''
        self.locality = ''
        self.address = ''
        self.phone_number = ''
        self.create_time = ''
        self.closed = ''
        self.complaint = 0
        self.ticket_number = ''
        self.volunteers = ''


class RecordBadGuysForm:
    def __init__(self):
        self.ticket_number = ''
        self.theme = ''
        self.create_time = ''
        self.reopened_dates = ''


class ReportForm01(Report):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.form_name = 'REPORT_FORM_01'
        self.init_dates()
        self.form_verbose_name = 'Форма_1'
        self.records = [RecordTypes.PURCHASE, RecordTypes.SOCIAL, RecordTypes.BROADCASTING_OUTSIDE,
                        RecordTypes.VOLUNTEERS, RecordTypes.CONNECTION, RecordTypes.BROADCASTING_REGIONAL,
                        RecordTypes.OTHER, RecordTypes.TOTAL, RecordTypes.COMPLAINTS]

    def get_data_from_db(self, filename='form_01.sql', *args):
        super(ReportForm01, self).get_data_from_db(filename, self.start_date, self.end_date)

    def data_to_form_template(self):
        _values = list(set(map(lambda x: x['name'], self.records)))
        _index = list(set(map(lambda x: x['value_text'], self.data)))
        df = pd.DataFrame(0, index=_index, columns=_values)
        for row in self.data:
            if row['complaints'] is not None:
                df.at[row['value_text'], RecordTypes.COMPLAINTS['name']] = row['complaints']
            record_type = RecordTypes.get_record_queue_by_code(row['ticket_type_id'])
            df.at[row['value_text'], record_type['name']] = row['frequency']
        df.at[RecordTypes.TOTAL['name']] = 0
        df.T.at[RecordTypes.TOTAL['name']] = 0
        for key in df.keys():
            df.at[RecordTypes.TOTAL['name'], key] = sum(df[key])
        total_df = df.drop(RecordTypes.COMPLAINTS['name'], axis=1)
        for row in total_df.T.keys():
            df.at[row, RecordTypes.TOTAL['name']] = sum(total_df.T[row])
        self.form = df

    def form_to_file(self):
        self.form_to_excel()

    def form_to_csv(self):
        pass

    def form_to_excel(self):
        if self.form is None:
            self.data_to_form_template()
        file_name = self.form_verbose_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:H', 20)
        worksheet.set_row(0, 80)
        header_format = self.get_header_format(workbook)
        row_format = self.get_row_format(workbook)
        worksheet.write(0, 0, 'Наименование ОМСУ', header_format)
        for idx, record in enumerate(self.records, start=1):
            worksheet.write(0, idx, record['name'], header_format)
        for row in range(1, len(self.form) + 1):
            for col in range(1, len(self.records) + 1):
                worksheet.write(row, col, 0, row_format)
        for idx_row, (row_name, row_data) in enumerate(self.form.iterrows(), start=1):
            worksheet.write(idx_row, 0, row_name)
            for col_name, col_value in row_data.items():
                for idx_theme, record in enumerate(self.records, start=1):
                    if record['name'] == col_name:
                        worksheet.write(idx_row, idx_theme, col_value, row_format)
                        break
        workbook.close()


class ReportForm542(Report):
    """
    Tickets that was closed for ten days
    """
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.form_name = 'REPORT_FORM_54_2'
        self.init_dates()
        self.form_verbose_name = 'Форма_5_4_2'
        self.header = [
            'ФИО',
            'Населенный пункт',
            'Точный адрес',
            'Телефонный номер',
            'Дата открытия',
            'Дата закрытия',
            'Жалоба',
            'Номер заявки',
            'Волонтёры',
        ]
        self.municipalities = 'Муниципальные образования'

    def get_data_from_db(self, filename='form_54_2.sql', *args):
        super(ReportForm542, self).get_data_from_db(filename, self.start_date, self.end_date)

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        if df.empty:
            return
        ticket_ids = list(set(df[df['closed'].notnull()]['ticket_id']))
        data = defaultdict(list)
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            create_time = ticket_df[ticket_df['field_id'] == 14]['create_time'].astype(str).iloc[0]
            max_working_days = get_max_working_days(create_time)
            closed = ticket_df[ticket_df['field_id'] == 14]['closed'].astype(str).iloc[0]
            if compute_working_time(create_time, closed) > int(max_working_days) * 24:
                continue
            record = RecordForm542()
            if not ticket_df[ticket_df['field_id'] == 12]['value_text'].empty:
                record.name = ticket_df[ticket_df['field_id'] == 12]['value_text'].iloc[0]
            if not ticket_df[ticket_df['field_id'] == 15]['value_text'].empty:
                record.locality = ticket_df[ticket_df['field_id'] == 15]['value_text'].iloc[0]
            if not ticket_df[ticket_df['field_id'] == 17]['value_text'].empty:
                record.address = ticket_df[ticket_df['field_id'] == 17]['value_text'].iloc[0]
            if not ticket_df[ticket_df['field_id'] == 16]['value_text'].empty:
                record.phone_number = ticket_df[ticket_df['field_id'] == 16]['value_text'].iloc[0]
            if not ticket_df[ticket_df['field_id'] == 37]['value_int'].empty:
                record.complaint = ticket_df[ticket_df['field_id'] == 37]['value_int'].iloc[0]
            if not ticket_df[ticket_df['field_id'] == 39]['value_text'].empty:
                record.volunteers = ticket_df[ticket_df['field_id'] == 39]['value_text'].iloc[0]
            if not ticket_df[ticket_df['field_id'] == 40]['value_text'].empty:
                record.volunteers += '; ' + ticket_df[ticket_df['field_id'] == 40]['value_text'].iloc[0]
            record.ticket_number = str(ticket_df['tn'].iloc[0])
            record.create_time = str(ticket_df['create_time'].iloc[0])
            record.closed = str(ticket_df['closed'].iloc[0])
            data[ticket_df[ticket_df['field_id'] == 14]['value_text'].iloc[0]].append(record.__dict__)
        self.form = dict(data)

    def form_to_file(self):
        column_range = 'A:I'
        header_merge_range = 'A1:I1'
        self.form_to_excel_by_territory(column_range, header_merge_range)
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name, self.municipalities)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        for record in self.form.items():
            file_name = record[0] + '.xlsx'
            workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
            worksheet = workbook.add_worksheet()
            worksheet.set_column(column_range, 20)
            worksheet.set_row(0, 30)
            worksheet.set_row(1, 80)
            header_format = self.get_header_format(workbook)
            row_format = self.get_row_format(workbook)
            worksheet.merge_range(header_merge_range, record[0], header_format)
            for idx, key in enumerate(self.header):
                worksheet.write(1, idx, key, header_format)
            for row_idx, data in enumerate(record[1], start=2):
                for col_idx, (key, value) in enumerate(data.items()):
                    worksheet.write(row_idx, col_idx, value, row_format)
            workbook.close()


class ReportForm543(Report):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.form_name = 'REPORT_FORM_54_3'
        self.init_dates()
        self.form_verbose_name = 'Форма_5_4_3'
        self.header = [
            'Наименование ОМСУ',
            'Количество заявок',
            'Количество закрытых заявок',
            'Количество вовремя закрытых заявок',
            'Количество просроченных закрытых заявок',
            'Количество просроченных открытых заявок',
            'Количество открытых непросроченных заявок',
            'Процент закрытых заявок в течение 5 дней от количества закрытых заявок',
        ]

    def get_data_from_db(self, filename='form_54_2.sql', *args):
        super(ReportForm543, self).get_data_from_db(filename, self.start_date, self.end_date)

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        if df.empty:
            return
        ticket_ids = list(set(df['ticket_id']))
        data = dict()
        data['Итого'] = {'total_tickets': 0, 'closed': 0, 'closed_on_time': 0,
                         'expired_closed': 0, 'expired_open': 0, 'opened_not_expired': 0,
                         'percent': 0}
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            if not ticket_df[ticket_df['field_id'] == 14]['value_text'].empty:
                name = ticket_df[ticket_df['field_id'] == 14]['value_text'].item()
                if name not in data.keys():
                    data[name] = {'total_tickets': 0, 'closed': 0, 'closed_on_time': 0,
                                  'expired_closed': 0, 'expired_open': 0, 'opened_not_expired': 0,
                                  'percent': 0}
            else:
                continue
            create_time = ticket_df[ticket_df['field_id'] == 14]['create_time'].astype(str).item()
            is_new_ticket = False
            if self.start_date.split().__len__() == 2:
                if datetime.datetime.strptime(create_time, '%Y-%m-%d %H:%M:%S') > datetime.datetime.strptime(self.start_date, '%Y-%m-%d %H-%M'):
                    data[name]['total_tickets'] += 1
                    data['Итого']['total_tickets'] += 1
                    is_new_ticket = True
            else:
                if datetime.datetime.strptime(create_time, '%Y-%m-%d %H:%M:%S') > datetime.datetime.strptime(self.start_date, '%Y-%m-%d'):
                    data[name]['total_tickets'] += 1
                    data['Итого']['total_tickets'] += 1
                    is_new_ticket = True
            max_working_days = get_max_working_days(create_time)
            last_action_time = ticket_df[ticket_df['field_id'] == 14]['last_action_time'].astype(str).item()
            ticket_state_id = ticket_df[ticket_df['field_id'] == 14]['ticket_state_id'].item()
            if ticket_state_id in (2, 3, 10) and last_action_time:
                data[name]['closed'] += 1
                data['Итого']['closed'] += 1
                if compute_working_time(create_time, last_action_time) <= int(max_working_days) * 24:
                    data[name]['closed_on_time'] += 1
                    data['Итого']['closed_on_time'] += 1
                else:
                    data[name]['expired_closed'] += 1
                    data['Итого']['expired_closed'] += 1
            elif compute_working_time(create_time, CURRENT_DATE_TIME) > int(max_working_days) * 24 and is_new_ticket:
                data[name]['expired_open'] += 1
                data['Итого']['expired_open'] += 1
            elif is_new_ticket:
                data[name]['opened_not_expired'] += 1
                data['Итого']['opened_not_expired'] += 1
            if data[name]['closed'] == 0:
                data[name]['percent'] = 0
            else:
                data[name]['percent'] = data[name]['closed_on_time'] / data[name]['closed']
        if data['Итого']['closed'] != 0:
            data['Итого']['percent'] = data['Итого']['closed_on_time'] / data['Итого']['closed']
        else:
            data['Итого']['percent'] = 0
        self.form = data

    def form_to_file(self):
        if not self.form:
            return
        file_name = self.form_verbose_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:B', 40)
        worksheet.set_column('B:H', 20)
        worksheet.set_row(0, 80)
        header_format = self.get_header_format(workbook)
        row_format = self.get_row_format(workbook)
        for idx, key in enumerate(self.header):
            worksheet.write(0, idx, key, header_format)
        for row_idx, (name, data) in enumerate(self.form.items()):
            if name != 'Итого':
                worksheet.write(row_idx, 0, name, row_format)
                for col_idx, (key, value) in enumerate(data.items(), start=1):
                    worksheet.write(row_idx, col_idx, value, row_format)
        worksheet.write(len(self.form), 0, 'Итого', row_format)
        for col_idx, (key, value) in enumerate(self.form['Итого'].items(), start=1):
            worksheet.write(len(self.form), col_idx, value, row_format)
        workbook.close()


class HourlyTotals(Report):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.form_name = 'HOURLY_TOTALS'
        self.init_dates()
        self.form_verbose_name = 'HourlyTotals'
        self.header = [
            'Время',
            'Количество открытых заявок',
            'Количество закрытых заявок',
            'Заявок в работе',
        ]

    def get_data_from_db(self, *args):
        start_date = datetime.datetime.strptime(self.start_date, '%Y-%m-%d %H-%M')
        end_date = datetime.datetime.strptime(self.end_date, '%Y-%m-%d %H-%M')
        sql_form = open('form_54_2.sql').read()
        sql_opened_tickets = open('opened_tickets.sql').read()
        next_hour = start_date
        self.data = []
        while next_hour < end_date:
            self.cursor.execute(sql_form.format(next_hour, next_hour + datetime.timedelta(hours=1)))
            data = self.cursor.fetchall()
            self.cursor.execute(sql_opened_tickets.format(next_hour + datetime.timedelta(hours=1)))
            total = self.cursor.fetchall()[0]['_count']
            df = pd.DataFrame.from_records(data)
            closed = 0
            opened = 0
            if not df.empty:
                closed = df[df['ticket_state_id'].isin([2, 3, 10])].groupby(['ticket_id'])['ticket_id'].count().shape[0]
                opened = df[(df['ticket_state_id'] == 4)
                            & (df['create_time'] >= next_hour)].groupby(['ticket_id'])['ticket_id'].count().shape[0]
            next_hour += datetime.timedelta(hours=1)
            self.data.append((next_hour, opened, closed, total))

    def data_to_form_template(self):
        self.form = pd.DataFrame.from_records(self.data, columns=['Время', 'Открытых заявок',
                                                                  'Закрытых заявок', 'В работе'])

    def form_to_file(self):
        file_name = self.form_verbose_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        self.form.to_excel(os.path.join(folder_path, file_name), index=0)


class BadGuysReportForm(Report):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.form_name = 'BAD_GUYS'
        self.init_dates()
        self.form_verbose_name = 'Bad_Guys'
        self.header = [
            'Населенный пункт',
            'Номер заявки',
            'Тема',
            'Дата создания',
            'Даты переоткрытия',
        ]
        self.ticket_history = 'select * from ticket_history where ticket_id = {0};'

    def get_data_from_db(self, filename='tickets_by_location.sql', *args):
        super(BadGuysReportForm, self).get_data_from_db(filename, self.start_date, self.end_date)

    def ticket_is_reopened(self, ticket_id):
        bad_guys = False
        self.cursor.execute(self.ticket_history.format(ticket_id))
        data = list(map(lambda x: x['state_id'], self.cursor.fetchall()))
        first_closed_index = min(list(map(lambda x: data.index(x) if x in data else len(data), [2, 3, 10])))
        for idx in range(first_closed_index, len(data)):
            if data[idx] == 4:
                bad_guys = True
        return bad_guys

    def get_reopened_dates(self, ticket_id):
        dates = []
        self.cursor.execute(self.ticket_history.format(ticket_id))
        data = self.cursor.fetchall()
        ids = list(map(lambda x: x['state_id'], data))
        first_closed_index = min(list(map(lambda x: ids.index(x) if x in ids else len(ids), [2, 3, 10])))
        for idx in range(first_closed_index, len(data)):
            if ids[idx] == 4 and ids[idx-1] != 4:
                dates.append(data[idx]['create_time'].strftime('%Y-%m-%d'))
        return dates

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        if df.empty:
            return
        ticket_ids = list(set(df['ticket_id']))
        data = defaultdict(list)
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            bad_guys = self.ticket_is_reopened(_id)
            if bad_guys:
                record = RecordBadGuysForm()
                record.theme = str(ticket_df['title'].iloc[0])
                record.create_time = str(ticket_df['create_time'].iloc[0])
                record.ticket_number = str(ticket_df['tn'].iloc[0])
                record.reopened_dates = '; '.join(self.get_reopened_dates(_id))
                data[ticket_df[ticket_df['field_id'] == 14]['value_text'].item()].append(record.__dict__)
        self.form = dict(data)

    def form_to_file(self):
        self.form_to_excel_aggregated('A:D')


class VolunteerRatingForm(Report):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.form_name = 'VOLUNTEER_RATING_FORM'
        self.init_dates()
        self.form_verbose_name = 'Рейтинг_волонтёров'
        self.header = [
            'ФИО',
            'Муниципальный район',
            'Количество баллов',
            'Количество заявок',
        ]
        self.volunteer_ids = [39, 40]
        self.volunteer_ticket_rate = {3: 1, 5: 2}

    def get_data_from_db(self, filename='volunteer_rating.sql', *args):
        super(VolunteerRatingForm, self).get_data_from_db(filename, self.start_date, self.end_date)

    def get_volunteers_rating(self, df):
        volunteers_rating = defaultdict(dict)
        region = df[df['field_id'] == 14]['value_text'].item()
        volunteers = {}
        for _id in self.volunteer_ids:
            if not df[df['field_id'] == _id]['value_text'].empty:
                ticket_priority_id = df[df['field_id'] == _id]['ticket_priority_id'].iloc[0]
                volunteer_score = self.volunteer_ticket_rate[ticket_priority_id]
                volunteers.update({df[df['field_id'] == _id]['value_text'].iloc[0]: volunteer_score})
        for name, score in volunteers.items():
            volunteers_rating.update({'name': name, 'region': region, 'score': score, 'tickets': 1})
        return volunteers_rating

    def data_to_form_template(self):
        df = pd.DataFrame.from_records(self.data)
        if df.empty:
            return
        ticket_ids = list(set(df['ticket_id']))
        volunteers_rating = []
        for _id in ticket_ids:
            ticket_df = df[df['ticket_id'] == _id]
            if ticket_df[ticket_df['field_id'] == 14]['value_text'].empty:
                continue
            create_time = ticket_df[ticket_df['field_id'] == 14]['create_time'].astype(str).item()
            max_working_days = get_max_working_days(create_time)
            closed = ticket_df[ticket_df['field_id'] == 14]['closed'].astype(str).item()
            ticket_state_id = ticket_df[ticket_df['field_id'] == 14]['ticket_state_id'].item()
            if ticket_state_id in (2, 3, 10):
                working_time = compute_working_time(create_time, closed)
                if working_time <= int(max_working_days) * 24:
                    volunteers_rating.append(self.get_volunteers_rating(ticket_df))
        vdf = pd.DataFrame(volunteers_rating)
        self.form = vdf.groupby(['name', 'region']).sum().reset_index().sort_values(by=['score'], ascending=False)

    def form_to_file(self):
        file_name = self.form_verbose_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(self.result_folder_path, self.form_verbose_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        workbook = xlsxwriter.Workbook(os.path.join(folder_path, file_name))
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:D', 20)
        worksheet.set_row(0, 80)
        header_format = self.get_header_format(workbook)
        row_format = self.get_row_format(workbook)
        for idx, key in enumerate(self.header):
            worksheet.write(0, idx, key, header_format)

        for row_idx, (idx, row) in enumerate(self.form.iterrows(), start=1):
            for col_idx, value in enumerate(row.values):
                worksheet.write(row_idx, col_idx, value, row_format)
        workbook.close()


class ReportFacade:
    reports = None

    @classmethod
    def create_reports(cls, daily, path):
        cls.reports = [
            # ReportForm01(daily=daily, path=path),
            # ReportForm543(daily=daily, path=path),
            VolunteerRatingForm(daily=daily, path=path),
            # BadGuysReportForm(daily=daily, path=path),
            # HourlyTotals(daily=daily, path=path),
        ]

    @classmethod
    def data_to_excel(cls):
        for report in cls.reports:
            report.get_data_from_db()
            report.data_to_form_template()
            report.form_to_file()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--type')
    parser.add_argument('-p', '--path')
    args = parser.parse_args()
    _daily = False
    _path = None
    if args.type == 'daily':
        _daily = True
    if args.path:
        _path = args.path
    ReportFacade.create_reports(daily=_daily, path=_path)
    ReportFacade.data_to_excel()
