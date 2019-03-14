#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : InfoPro.py
# Author: Wangyuan
# Date  : 2019-1-4
from re import match,sub
from pandas import read_excel,DataFrame,ExcelWriter
class InfoPro:

    def __init__(self,addr):
        self.form_out = []
        self.omit_value = []
        self.addr = addr
        df = read_excel(addr)
        for line in df.values:
            if line[1] == line[1] and line[1]!='通径' and line[3] == line[3]:
                form_out = {}
                # 对应excel中的序号
                form_out[0] = line[0] # list_num
                # 对应excel中的产品代号
                form_out[1] = line[1] # pro_name
                # 对应excel中的零件名称
                form_out[2] = line[3] # part_name
                # 对应excel中的材料
                form_out[3] = line[4] # material
                # 对应excel中的材料厚度
                form_out[4] = line[5]  # thickness
                # 长度、外径、环板外直径、以号代图
                form_out[5]= line[6]  # p1
                form_out[6]=line[7]  # p2
                # 对应excel中的总数
                if line [9] == line[9]:
                    form_out[7]=line[9]  # mark_sum
                else:
                    form_out[7] = line[2] * line[8]
                form_out[8] = 'Whatever'  # remark
                form_out[9]=line[10] # remark

                self.form_out.append(form_out)

    def to_excel(self):
        df = DataFrame(self.form_out)
        df.to_excel('重整化'+self.addr)

if __name__ == '__main__':
    i = InfoPro('3032017XXXX_排料清单.xlsx')
    print(i.form_out)

