# -*- coding: utf-8 -*-
"""
Created on Wed Apr 10 10:24:54 2019

@author: gazhakv
"""
import reports_generator


if __name__ == '__main__':
    rf_01 = ReportForm01()
    data = rf_01.get_data_from_db()
    print(data)