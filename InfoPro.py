#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : InfoPro.py
# Author: Wangyuan
# Date  : 2019-1-4
from re import split,match,sub
from pandas import read_excel,DataFrame
# addr = 'MM-2108  0-0.xlsx'
class InfoPro:

    def __init__(self,addr):
        self.form_out = []
        self.addr = addr
        df = read_excel(addr, skipfooter= 1)
        cal = 0
        for line in reversed(df.values):
            form_out = {}
            # 对应excel中的通径
            form_out[1] = addr.split(' ')[0]
            # 对应excel中的件数
            form_out[2] = 1
            # 对应excel中的零件名称
            form_out[3] = line[2]
            # 对应excel中的数量
            form_out[9]= line[7]
            # 对应excel中的材料
            form_out[4] = line[3]
            # 对应excel中的零件数
            form_out[8] = 1
            if line[1] == line[1]:
                if line[5] == line[5] and line[6] == line[6]:
                    cal += 1;
                    form_out[0] = cal
                    form_out[6] = line[5]
                    form_out[7]= line[6]
                    form_out[5]= line[4]
                    self.form_out.append(form_out)
                elif  str(line[1]).find('*')!= -1 and match('^[HWNZJ]',line[1]) and not match('^[J|j][w|W]',line[1]) :
                    cal += 1
                    form_out[0] = cal
                    form_out[6] = line[1]
                    form_out[7] = 0
                    form_out[5] = 0
                    self.form_out.append(form_out)
    def to_excel(self):
        df = DataFrame(self.form_out)
        df.to_excel('重整化'+self.addr)
# i = InfoPro(addr)
# i.to_excel()
# # # # for line in i.form_out:
# # # #     print(line)