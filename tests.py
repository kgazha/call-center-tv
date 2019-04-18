import unittest
import reports_generator
import os.path
import datetime


class TestReportForm01(unittest.TestCase):
    report = reports_generator.ReportForm01()
    report.data = (
        {'value_text': 'Агаповский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 2, 'value_int': None, 'complaint_field_id': None},
        {'value_text': 'Ашинский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 3, 'value_int': 0, 'complaint_field_id': 37},
        {'value_text': 'Златоустовский ГО', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 1, 'value_int': 1, 'complaint_field_id': 37},
        {'value_text': 'Златоустовский ГО', 'name': 'МинСоц', 'ticket_type_id': 8,
         'frequency': 2, 'value_int': None, 'complaint_field_id': None},
        {'value_text': 'Златоустовский ГО', 'name': 'Покупка', 'ticket_type_id': 9,
         'frequency': 1, 'value_int': None, 'complaint_field_id': None},
        {'value_text': 'Карабашский ГО', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 1, 'value_int': None, 'complaint_field_id': None},
        {'value_text': 'Карталинский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 1, 'value_int': 1, 'complaint_field_id': 37},
        {'value_text': 'Каслинский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 2, 'value_int': 0, 'complaint_field_id': 37},
        {'value_text': 'Копейский ГО', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 2, 'value_int': 0, 'complaint_field_id': 37},
        {'value_text': 'Копейский ГО', 'name': 'Иное', 'ticket_type_id': 14,
         'frequency': 1, 'value_int': None, 'complaint_field_id': None},
        {'value_text': 'Коркинский район', 'name': 'Иное', 'ticket_type_id': 14,
         'frequency': 1, 'value_int': 1, 'complaint_field_id': 37},
        {'value_text': 'Красноармейский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 3, 'value_int': None, 'complaint_field_id': None},
        {'value_text': 'Магнитогорский ГО', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 1, 'value_int': 0, 'complaint_field_id': 37},
        {'value_text': 'Миасский ГО', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 1, 'value_int': 0, 'complaint_field_id': 37},
        {'value_text': 'Саткинский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 2, 'value_int': 0, 'complaint_field_id': 37},
        {'value_text': 'Сосновский район', 'name': 'Волонтеры', 'ticket_type_id': 11,
         'frequency': 3, 'value_int': 0, 'complaint_field_id': 37},
    )

    def test_data_to_template(self):
        TestReportForm01.report.data_to_form_template()
        self.assertIsNotNone(TestReportForm01.report.form)

    def test_form_to_excel(self):
        TestReportForm01.report.form_to_excel()
        file_name = TestReportForm01.report.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(reports_generator.BASE_DIR, TestReportForm01.report.form_name)
        self.assertTrue(os.path.isfile(os.path.join(folder_path, file_name)))


class TestReportForm53(unittest.TestCase):
    report = reports_generator.ReportForm53()
    report.data = (
        {'field_id': 12, 'value_text': 'иван', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 9, 10, 27, 5), 'ticket_state_id': 1},
        {'field_id': 14, 'value_text': 'Аgapovski', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 9, 10, 27, 5), 'ticket_state_id': 1},
        {'field_id': 15, 'value_text': 'агаповка', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 9, 10, 27, 5), 'ticket_state_id': 1},
        {'field_id': 16, 'value_text': '88005553533', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 9, 10, 27, 5), 'ticket_state_id': 1},
        {'field_id': 17, 'value_text': 'улица: пушкина, дом: колотушкина, кв.: /3', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 9, 10, 27, 5), 'ticket_state_id': 1},
        {'field_id': 12, 'value_text': 'иванов', 'ticket_id': 20,
         'create_time': datetime.datetime(2019, 4, 9, 10, 30, 1), 'ticket_state_id': 1},
        {'field_id': 14, 'value_text': 'Агаповский', 'ticket_id': 20,
         'create_time': datetime.datetime(2019, 4, 9, 10, 30, 1), 'ticket_state_id': 1},
        {'field_id': 17, 'value_text': 'улица: , дом: , кв.:', 'ticket_id': 20,
         'create_time': datetime.datetime(2019, 4, 9, 10, 30, 1), 'ticket_state_id': 1},
        {'field_id': 12, 'value_text': 'абонент', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 1},
        {'field_id': 14, 'value_text': 'Еткульский район', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 1},
        {'field_id': 15, 'value_text': 'еткуль', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 1},
        {'field_id': 16, 'value_text': '+7123456789', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 1},
    )

    def test_data_to_template(self):
        TestReportForm53.report.data_to_form_template()
        self.assertIsNotNone(TestReportForm53.report.form)

    def test_date_filtering(self):
        TestReportForm53.report.data_to_form_template()
        self.assertEqual(len(TestReportForm53.report.form), 2)

    def test_form_to_excel(self):
        TestReportForm53.report.form_to_excel()
        file_name = TestReportForm53.report.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(reports_generator.BASE_DIR, TestReportForm53.report.form_name)
        self.assertTrue(os.path.isfile(os.path.join(folder_path, file_name)))


class TestReportForm54(unittest.TestCase):
    report = reports_generator.ReportForm54()
    report.data = (
        {'field_id': 12, 'value_text': 'иван', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 2, 10, 27, 5), 'ticket_state_id': 4},
        {'field_id': 14, 'value_text': 'Аgapovski', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 2, 10, 27, 5), 'ticket_state_id': 4},
        {'field_id': 15, 'value_text': 'агаповка', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 2, 10, 27, 5), 'ticket_state_id': 4},
        {'field_id': 16, 'value_text': '88005553533', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 2, 10, 27, 5), 'ticket_state_id': 4},
        {'field_id': 17, 'value_text': 'улица: пушкина, дом: колотушкина, кв.: /3', 'ticket_id': 19,
         'create_time': datetime.datetime(2019, 4, 3, 10, 27, 5), 'ticket_state_id': 4},
        {'field_id': 12, 'value_text': 'иванов', 'ticket_id': 20,
         'create_time': datetime.datetime(2019, 4, 3, 10, 30, 1), 'ticket_state_id': 4},
        {'field_id': 14, 'value_text': 'Агаповский', 'ticket_id': 20,
         'create_time': datetime.datetime(2019, 4, 3, 10, 30, 1), 'ticket_state_id': 4},
        {'field_id': 17, 'value_text': 'улица: , дом: , кв.:', 'ticket_id': 20,
         'create_time': datetime.datetime(2019, 4, 3, 10, 30, 1), 'ticket_state_id': 4},
        {'field_id': 12, 'value_text': 'абонент', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 4},
        {'field_id': 14, 'value_text': 'Еткульский район', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 4},
        {'field_id': 15, 'value_text': 'еткуль', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 4},
        {'field_id': 16, 'value_text': '+7123456789', 'ticket_id': 24,
         'create_time': datetime.datetime.now(), 'ticket_state_id': 4},
    )

    def test_data_to_template(self):
        TestReportForm54.report.data_to_form_template()
        self.assertIsNotNone(TestReportForm54.report.form)

    def test_date_filtering(self):
        TestReportForm54.report.data_to_form_template()
        self.assertEqual(len(TestReportForm54.report.form), 2)

    def test_form_to_excel(self):
        TestReportForm54.report.form_to_excel()
        file_name = TestReportForm54.report.form_name + datetime.date.today().strftime("_%d_%m_%Y") + '.xlsx'
        folder_path = os.path.join(reports_generator.BASE_DIR, TestReportForm54.report.form_name)
        self.assertTrue(os.path.isfile(os.path.join(folder_path, file_name)))


if __name__ == '__main__':
    unittest.main()
