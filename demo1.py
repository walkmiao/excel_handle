#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2018/11/5 20:00
# @Author  : lch
# @File    : demo1.py
import xlrd
from xlutils.copy import copy
from config import input_excel_path, output_excel_path, input_path_list
from threading import Thread
from logger import logger
import random
import os
import re


def get_table(path_list):
    table_dict = {}
    for path in path_list:
        path_name = os.path.basename(path)
        sheet_dict = get_sheet(path)
        table_dict[path_name] = sheet_dict
    return table_dict

def get_sheet(path=input_excel_path):
    wb = xlrd.open_workbook(path)
    sheets = wb.sheets()
    sheets_name = wb.sheet_names()
    if path!=output_excel_path:
        sheet_dict={}
        sheets = sheets[1:]
        sheets_name = sheets_name[1:]
        for index, sheet in enumerate(sheets):
            sheet_dict[sheet] =sheets_name[index]
        return sheet_dict
    else:
        return sheets[0]


def get_wb_copy(path=output_excel_path):
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
        if '/' in warn_value:
            result = warn_value.split('/')
            return result

        if '>' in warn_value:
            if '<' in warn_value:
                result = re.findall(r'.*<(-?\d+\.?\d?).*>(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*<(-?\d+\.?\d?).*>(-?\d+\.?\d?).*', warn_value) else None
            else:
                result = re.findall(r'.*>(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*>(-?\d+\.?\d?).*', warn_value) else None
            return result

        if'<' in warn_value:
            if '>' in warn_value:
                result = re.findall(r'.*<(-?\d+\.?\d?).*>(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*<(-?\d+\.?\d?).*>(-?\d+\.?\d?).*', warn_value) else None
            else:
                result = re.findall(r'.*<(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*<(-?\d+\.?\d?).*', warn_value) else None
            return result

        if '>=' in warn_value:
            if '<=' in warn_value:
                result = re.findall(r'.*>=(-?\d+\.?\d?).*<=(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*>=(-?\d+\.?\d?).*<=(-?\d+\.?\d?).*', warn_value) else None
            else:
                result = re.findall(r'.*>=(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*>=(-?\d+\.?\d?).*', warn_value) else None
            return result

        if '≥' in warn_value:
            if '≤' in warn_value:
                result = re.findall(r'.*≥(-?\d+\.?\d?).*≤(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*≥(-?\d+\.?\d?).*≤(-?\d+\.?\d?).*', warn_value) else None
            else:
                result = re.findall(r'.*≥(-?\d+\.?\d?).*', warn_value)[0] \
                    if re.findall(r'.*≥(-?\d+\.?\d?).*', warn_value) else None
            return result

        if ',' in warn_value and '>':
            result = warn_value.split(',')
            return result
        if '，' in warn_value:
            result = warn_value.split('，')
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
                        if not isinstance(result, str) and result:
                            if len(result) > 1:
                                value_low, value_high = result
                                if self.what_type(value_low) > self.what_type(value_high):
                                    w_sheet.write(index_output, 10, self.what_type(value_high))
                                    w_sheet.write(index_output, 9, self.what_type(value_low))
                                else:
                                    w_sheet.write(index_output, 10, self.what_type(value_low))
                                    w_sheet.write(index_output, 9, self.what_type(value_high))
                            else:
                                w_sheet.write(index_output, 10, self.what_type(result))
                        else:
                            try:
                                w_sheet.write(index_output, 10, self.what_type(result))
                            except Exception:
                                w_sheet.write(index_output, 10, result)

                    except Exception as e:
                        logger.info('[ERROR] CAUSE BY{},result={}'.format(e, result))


                    break
            if not is_find:
                logger.info('输入excel sheet:[{}] 未找到序号{} MRID={} 描述:{}'
                      .format(self.sheet_name, index_input+1,mrid,str(self.input_sheet.cell_value(index_input, 1))))


def main():
    tab_dict = get_table(input_path_list)
    # sheet_dict = get_sheet()
    out_sheet = get_sheet(output_excel_path)
    wb_copy = get_wb_copy()
    thread_list = []
    for tab_name, sheet_dict in tab_dict.items():
        for input_sheet, sheet_name in sheet_dict.items():
            excel_ins = ExcelHandle(input_sheet, out_sheet, wb_copy, tab_name+'::'+sheet_name)
            thread_list.append(excel_ins)
    for t in thread_list:
        t.start()
    for t in thread_list:
        t.join()
    wb_copy.save('./result.xlsx')
    try:
        os.rename('./result.xlsx', './result.xls')
    except Exception:
        os.rename('./result.xlsx','./result'+ str(random.randint(1,10))+'.xls')
    logger.info('执行完毕')

if __name__ == '__main__':
        main()