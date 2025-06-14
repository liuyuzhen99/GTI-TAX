import pandas as pd
from datetime import datetime
import os
from exception_handler import exception_signal_instance
from MatchException import MatchError
import re
from readExcelClass import ReadExcelClass


class BindClass:
    def __init__(self, wb_reader=ReadExcelClass(), df_gti_reader=ReadExcelClass()):
        self.gti_excel = df_gti_reader
        self.wb_excel = wb_reader

    def format_number(self, number, point_num, length_num):
        formatted_number = str(round(number, point_num))
        if len(formatted_number) > length_num:
            formatted_number = formatted_number[:length_num]
        return formatted_number

    def restrain_len(self, string, length):
        if len(string) > length:
            return string[:length]
        else:
            return string

    def restrain_tax_rate(self, number, length):
        if len(str(number)) > length:
            return float(str(number)[:length])
        else:
            return number

    def convert_excel(self):
        start_row_sheet2 = self.wb_excel.wb['2-发票明细信息'].max_row + 1
        start_row_sheet1 = self.wb_excel.wb['1-发票基本信息'].max_row + 1
        current_doc_num = 0
        for i in range(len(self.gti_excel.dataframe)):
            self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=1,
                                              value=self.restrain_len(str(self.gti_excel.dataframe.iloc[i, 1]),
                                                                      20))  # 2-发票明细信息  发票流水号
            self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=2,
                                              value=self.restrain_len(self.gti_excel.dataframe.iloc[i, 18],
                                                                      100))  # 2-发票明细信息 项目名称
            self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=3,
                                              value=self.restrain_len(str(self.gti_excel.dataframe.iloc[i, 28]),
                                                                      20))  # 2-发票明细信息 商品和服务税收编码
            if self.gti_excel.dataframe.iloc[i, 20] == "千克":
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=5,
                                                  value="吨")  # 2-发票明细信息 单位(千克->吨)
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=6,
                                                  value=self.format_number(self.gti_excel.dataframe.iloc[i, 21] / 1000,
                                                                           13,
                                                                           16))  # 2-发票明细信息 数量
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=7,
                                                  value=self.restrain_len(
                                                      format((self.gti_excel.dataframe.iloc[i, 22] -
                                                              self.gti_excel.dataframe.iloc[i, 24]) /
                                                             (self.gti_excel.dataframe.iloc[i, 21]), '.2f'),
                                                      16))  # 2-发票明细信息 单价 (self.gti_excel.dataframe.iloc[i, 21] / 1000), '.2f')
            else:
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=5,
                                                  value=self.restrain_len(self.gti_excel.dataframe.iloc[i, 20],
                                                                          22))  # 2-发票明细信息 单位
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=6,
                                                  value=self.format_number(self.gti_excel.dataframe.iloc[i, 21], 13,
                                                                           16))  # 2-发票明细信息 数量
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=7,
                                                  value=self.restrain_len(
                                                      format((self.gti_excel.dataframe.iloc[i, 22] -
                                                              self.gti_excel.dataframe.iloc[i, 24]) /
                                                             (self.gti_excel.dataframe.iloc[i, 21]), '.2f'),
                                                      16))  # 2-发票明细信息 单价 (self.gti_excel.dataframe.iloc[i, 21] / 1000), '.2f'),
            self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=8,
                                              value=self.restrain_len(format(self.gti_excel.dataframe.iloc[i, 22] -
                                                                             self.gti_excel.dataframe.iloc[i, 24],
                                                                             '.2f'),
                                                                      16))  # 2-发票明细信息 金额
            self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=9,
                                              value=self.restrain_tax_rate(self.gti_excel.dataframe.iloc[i, 23],
                                                                           8))  # 2-发票明细信息 税率
            if self.gti_excel.dataframe.iloc[i, 25] - self.gti_excel.dataframe.iloc[i, 26] != 0:
                self.wb_excel.wb['2-发票明细信息'].cell(row=start_row_sheet2, column=10,
                                                  value=self.format_number((self.gti_excel.dataframe.iloc[i, 25] -
                                                                            self.gti_excel.dataframe.iloc[i, 26]), 2,
                                                                           16))  # 2-发票明细信息 折扣金额
            # wb['2-发票明细信息'].cell(row=start_row_sheet2, column=11, value=format_number(df.iloc[i, 25], 2, 16)) # 2-发票明细信息 是否使用优惠政策
            start_row_sheet2 += 1
            if self.gti_excel.dataframe.iloc[i, 1] != current_doc_num:
                self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=1,
                                                  value=self.restrain_len(str(self.gti_excel.dataframe.iloc[i, 1]),
                                                                          20))  # 1-发票基本信息 发票流水号
                if self.gti_excel.dataframe.iloc[i, 0] == "增票":
                    self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=2, value='增值税专用发票')  # 1-发票基本信息 发票类型
                elif self.gti_excel.dataframe.iloc[i, 0] == "普票":
                    self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=2, value='普通发票')  # 1-发票基本信息 发票类型
                self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=4, value='否')  # 1-发票基本信息 是否含税
                # (re.sub("[A-Za-z0-9\,\。]", "", str) 提取中文字符
                self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=6,
                                                  value=self.restrain_len(re.sub("[A-Za-z0-9\,\。]", "",
                                                                                 self.gti_excel.dataframe.iloc[i, 4]),
                                                                          100))  # 1-发票基本信息 购买方名称
                self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=8,
                                                  value=self.restrain_len(str(self.gti_excel.dataframe.iloc[i, 5]),
                                                                          20))  # 1-发票基本信息 购买人纳税人识别号
                self.wb_excel.wb['1-发票基本信息'].cell(row=start_row_sheet1, column=9,
                                                  value=self.restrain_len(self.gti_excel.dataframe.iloc[i, 6],
                                                                          100))  # 1-发票基本信息 购买方地址
                start_row_sheet1 += 1
                current_doc_num = self.gti_excel.dataframe.iloc[i, 1]
        self.wb_excel.wb.save(os.path.dirname(self.wb_excel.path) + '\\file_for_tax_website.xlsx')
