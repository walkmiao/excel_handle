#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2018/11/5 20:00
# @Author  : lch
# @File    : demo1.py
import xlrd
from xlutils.copy import copy
# from config import input_excel_path, output_excel_path, input_path_list
from threading import Thread
from logger import logger
import random
import os
import re
import configparser

def get_conf(filepath='./config.ini'):
    conf = configparser.ConfigParser()
    conf.read(filepath, encoding='utf-8')
    path_list = conf.get('PATH', 'path_list')
    path_list = path_list.split(',')
    output_excel_path = conf.get('PATH', 'output_path')
    return path_list, output_excel_path


def get_table(path_list):
    table_dict = {}
    for path in path_list:
        path_name = os.path.basename(path)
        sheet_dict = get_sheet(path)
        table_dict[path_name] = sheet_dict
    return table_dict


def get_sheet(path):
    wb = xlrd.open_workbook(path)
    sheets = wb.sheets()
    sheets_name = wb.sheet_names()
    if len(sheets)>3:
        sheet_dict={}
        sheets = sheets[1:]
        sheets_name = sheets_name[1:]
        for index, sheet in enumerate(sheets):
            sheet_dict[sheet] =sheets_name[index]
        return sheet_dict
    else:
        return sheets[0]


def get_wb_copy(path):
    wb = xlrd.open_workbook(path)
    wb_copy = copy(wb)
    return wb_copy


class ExcelHandle(Thread):
    def __init__(self,input_sheet,output_sheet, wb_copy, sheet_name):
        super(ExcelHandle, self).__init__(name=sheet_name)
        self.input_sheet = input_sheet
        self.output_sheet = output_sheet
        self.wb_copy = wb_copy
        self.sheet_name = sheet_name
        logger.info('THREAD[{}] START!'.format(self.getName()))

    def run(self):
        self.insert_to_excle()

    def get_out_put_excel_info(self):
        mrid_col_value = self.output_sheet.col_values(1)
        return mrid_col_value

    def what_type(self, value_str):
        if '.' in value_str:
            value_str = float(value_str)
        else:
            value_str = int(value_str)
        return value_str

    def get_write_sheet(self):
        sheet = self.wb_copy.get_sheet(0)
        return sheet

    def complex_re(self, warn_value):
        high1='high1'
        low1='low1'
        high2='high2'
        low2='low2'
        result={}
        if '/' in warn_value:
            r = warn_value.split('/')
            low_value, high_value = r
            if self.what_type(low_value) >self.what_type(high_value):
                low_value,high_value=high_value,low_value
            result[high1]=high_value
            result[low1]=low_value
            return result

        if '>' in warn_value and '<' in warn_value:
            if warn_value.count('>')==1:
                if warn_value.count('<')==1:
                    high_value = re.findall(r'.*?>(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    low_value = re.findall(r'.*?<(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    result[high1] = high_value
                    result[low1] = low_value
                else:
                    high_value = re.findall(r'.*?>(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    low_value1,low_value2 = re.findall(r'.*?<(-?\d+\.?\d{0,4}).*<(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    result[high1] = high_value
                    result[low2] = low_value2
                    result[low1] = low_value1
            else:
                if warn_value.count('<') == 1:
                    r= re.findall(r'.*?>(-?\d+\.?\d{0,4}).*>(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    high_value1, high_value2 = r
                    low_value = re.findall(r'.*?<(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    result[high1] = high_value1
                    result[high2] = high_value2
                    result[low1] = low_value
                else:
                    high_value1, high_value2 = re.findall(r'.*?>(-?\d+\.?\d{0,4}).*>(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    low_value1, low_value2 = re.findall(r'.*?<(-?\d+\.?\d{0,4}).*<(-?\d+\.?\d{0,4}).*', warn_value)[0]
                    result[high1] = high_value1
                    result[high2] = high_value2
                    result[low1] = low_value1
                    result[low2] = low_value2
            return result

        if '>' in warn_value and '<' not in warn_value:
            if warn_value.count('>')==1:
                high_value = re.findall(r'.*?>(-?\d+\.?\d{0,4}).*', warn_value)[0]
                result[high1] = high_value
            else:
                high_value1, high_value2 = re.findall(r'.*?>(-?\d+\.?\d{0,4}).*>(-?\d+\.?\d{0,4}).*', warn_value)[0]
                result[high1] = high_value1
                result[high2] = high_value2
            return result
        if '<' in warn_value and '>' not in warn_value:
            if warn_value.count('<')==1:
                low_value = re.findall(r'.*?<(-?\d+\.?\d{0,4}).*', warn_value)[0]
                result[low1] = low_value
            else:
                low_value1, low_value2 = re.findall(r'.*?<(-?\d+\.?\d{0,4}).*<(-?\d+\.?\d{0,4}).*', warn_value)[0]
                result[low1] = low_value1
                result[low2] = low_value2
            return result

        if '>=' in warn_value:
            if '<=' in warn_value:
                high_value1, low_value1 = re.findall(r'.*?>=(-?\d+\.?\d{0,4}).*<=(-?\d+\.?\d{0,4}).*', warn_value)[0] \
                    if re.findall(r'.*?>=(-?\d+\.?\d{0,4}).*<=(-?\d+\.?\d{0,4}).*', warn_value) else None
                result[high1] = high_value1
                result[low1] = low_value1
            else:
                high_value1 = re.findall(r'.*?>=(-?\d+\.?\d{0,4}).*', warn_value)[0] \
                    if re.findall(r'.*?>=(-?\d+\.?\d{0,4}).*', warn_value) else None
                result[high1] = high_value1
            return result

        if '≥' in warn_value:
            if '≤' in warn_value:
                high_value1, low_value1 = re.findall(r'.*?≥(-?\d+\.?\d{0,4}).*≤(-?\d+\.?\d{0,4}).*', warn_value)[0]
                result[high1] = high_value1
                result[low1] = low_value1
            else:
                high_value1 = re.findall(r'.*?≥(-?\d+\.?\d{0,4}).*', warn_value)[0] \
                    if re.findall(r'.*?≥(-?\d+\.?\d{0,4}).*', warn_value) else None
                result[high1] = high_value1
            return result

        if ',' in warn_value and '>':
            low_value1, high_value1 = warn_value.split(',')
            if self.what_type(low_value1) >self.what_type(high_value1):
                low_value1,high_value1=high_value1,low_value1
            result[high1] = high_value1
            result[low1] = low_value1
            return result
        if '，' in warn_value:
            low_value1, high_value1 = warn_value.split('，')
            if self.what_type(low_value1) >self.what_type(high_value1):
                low_value1,high_value1=high_value1,low_value1
            result[high1] = high_value1
            result[low1] = low_value1
            return result
        return warn_value

    def insert_to_excle(self):
        mrid_col_value = self.get_out_put_excel_info()[1:]
        w_sheet = self.get_write_sheet()
        for index_input, mrid in enumerate(self.input_sheet.col_values(3)[1:], 1): #循环每个列mrid的值是否在指定表中匹配
            is_find = False
            for index_output, out_mrid in enumerate(mrid_col_value,1):
                if mrid in out_mrid:
                    is_find = True
                    warn_value = str(self.input_sheet.cell_value(index_input, 6))
                    result = self.complex_re(warn_value)
                    try:
                        if result:
                            if isinstance(result, dict):
                                w_sheet.write(index_output, 8, 1)
                                if 'high1'in result:
                                    w_sheet.write(index_output, 9, self.what_type(result['high1']))
                                if 'high2' in result:
                                    w_sheet.write(index_output, 12, self.what_type(result['high2']))
                                if 'low1' in result:
                                    w_sheet.write(index_output, 10, self.what_type(result['low1']))
                                if 'low2' in result:
                                    w_sheet.write(index_output, 13, self.what_type(result['low1']))
                            else:
                                try:
                                    w_sheet.write(index_output, 9, self.what_type(result))
                                    w_sheet.write(index_output, 8, 1)
                                except Exception:
                                    w_sheet.write(index_output, 9, result)
                                    w_sheet.write(index_output, 8, 1)
                        else:
                            pass

                    except Exception as e:
                        logger.info('[ERROR] CAUSE BY{},result={}'.format(e, result))


                    break
            if not is_find:
                logger.info('输入excel sheet:[{}] 未找到序号{} MRID={} 描述:{}'
                      .format(self.sheet_name, index_input+1,mrid,str(self.input_sheet.cell_value(index_input, 1))))


def main():
    path_list, output_excel_path = get_conf()
    tab_dict = get_table(path_list)

    out_sheet = get_sheet(output_excel_path)
    wb_copy = get_wb_copy(output_excel_path)
    thread_list = []
    for tab_name, sheet_dict in tab_dict.items():
        for input_sheet, sheet_name in sheet_dict.items():
            excel_ins = ExcelHandle(input_sheet, out_sheet, wb_copy, tab_name+'::'+sheet_name)
            thread_list.append(excel_ins)
    for t in thread_list:
        t.start()
    for t in thread_list:
        t.join()
    out_base_name = os.path.basename(output_excel_path) +'输出结果'
    wb_copy.save('./result.xlsx')
    try:
        os.rename('./result.xlsx', './{}.xls'.format(out_base_name))
    except Exception:
        os.rename('./result.xlsx','./result'+ str(random.randint(1,10))+'.xls')
    logger.info('执行完毕')


if __name__ == '__main__':
    main()
    # print(get_conf())