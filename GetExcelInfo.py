#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : Creator0.5.py
# Author: WangYuan
# Date  : 2018-12-20

from pandas import read_excel

from os import getcwd
from InfoPro import *
from re import match
from str2map import *
dirc =getcwd()
global table_finish,table_result



class GetExcelInfo:
    def __init__(self,address):
        # 设定输出到下料清单的数据
        self.docx_list = []
        self.output = []
        # bellows_weight_1219 = 0
        # bellows_weight_1000 = 0

        #判断模板Excel的类型
        if address.find('排料清单')!=-1:
            inform = read_excel(address,skiprows=1).values
        else:
            inform  = InfoPro(address).form_out
        for line in inform:
            docx_input = {}
            format_input = {}
            format_input['parameter'] = []
            if match('^[a-zA-Z]',str(line[6])):  #如果列6是以号代图,则识别以号代图
                b = Str2map(line[6])
                format_input['parameter'] =b.parameter
                format_input['type'] =b.type
                format_input['thickness'] = b.thickness
                format_input['material'] = str(line[4])
                #如果是波纹管就只修改数量
                if format_input['type'] == '波纹管':
                    format_input['sum'] = int(b.num * line[2] * line[8])
                #如果是其他类型的零件就将文件输出
                else:

                    docx_input['product_code'] = line[1]
                    docx_input['part_name'] = line[3]
                    docx_input['thickness'] = format_input['thickness']
                    docx_input['material'] = format_input['material']
                    if line[9] != line[2] * line[8]:
                        if line[9] != line[9]:
                            docx_input['sum'] = format_input['sum'] = int(line[2] * line[8])
                        else:
                            docx_input['sum'] = format_input['sum'] = int(line[9])
                            print('请检查%s行数据是否异常' % line[0])
                    else:
                        docx_input['sum'] = format_input['sum'] = int(line[9])
                    #单纯的为了之后代码简单
                    p1 = format_input['parameter'][0]
                    p2 = format_input['parameter'][1]
                    if format_input['type'] == '整圆':
                        docx_input['size'] = '直径: ' + str(p1)
                    elif format_input['type'] == '接管':
                        docx_input['size'] = '展开长: ' + str(int(p1) + 1) + ' 宽度: ' + str(p2)
                    elif format_input['type'] == '圆环':
                        docx_input['size'] = '外直径: ' + str(p1) + ' 内直径: ' + str(p2)
                    else:
                        docx_input['size'] = '长度: ' + str(int(p1) + 1) + ' 宽度: ' + str(p2)
                    self.docx_list.append(docx_input)
                format_input['name'] = str(line[0]) + '_' + line[1] + '_' + line[3] + '_' + str(format_input['sum'])

            else:
                # 判断除以号代图外的其他输入类型
                # 初始化下清单的参数
                docx_input['product_code'] = line[1]
                docx_input['part_name'] = line[3]
                docx_input['material'] = str(line[4])
                docx_input['thickness'] = line[5]
                if line[7] != line[7] or line[7] ==0 or type(line[7]) == str :
                    format_input['type']='整圆'
                    format_input['parameter'].append(line[6])
                    docx_input['size'] = '直径: ' + str(line[6])
                elif line[3].find("接管") != -1 or (str(line[6]).find("径")!= -1 and str(line[7]).find("径")== -1):
                    format_input['type']='接管'
                    line[6] = (line[6]- line[5])*3.1415926
                    format_input['parameter'].append(line[6])
                    docx_input['size'] = '展开长: ' + str(int(line[6])+1) + ' 宽度: ' +str(line[7])
                elif line[3].find("环")!= -1 or line[3].find("B板")!= -1 or (str(line[6]).find("径")!= -1 and str(line[7]).find("径")!= -1):
                    format_input['type'] = '圆环'
                    docx_input['size'] = '外直径: ' + str(line[6]) + ' 内直径: ' + str(line[7])
                    format_input['parameter'].append(line[6])
                else:
                    format_input['type'] ='搭板'
                    format_input['parameter'].append(line[6])
                    docx_input['size'] = '长度: ' + str(int(line[6]) + 1) + ' 宽度: ' + str(line[7])

                #除波纹管外的其他的类型提供parameter参数和name参数
                if type(line[7]) == int or type(line[6]) == float:
                    if line[7] != 0:
                        format_input['parameter'].append(line[7])

                docx_input['thickness'] = format_input['thickness'] = line[5]
                format_input['material'] = str(line[4])
                format_input['sum'] = line[9]
                if line[9] != line[2]*line[8]:
                    if line[9] != line[9]:
                        docx_input['sum'] = format_input['sum'] =int(line[2] * line[8])
                    else:
                        docx_input['sum'] = format_input['sum'] = int(line[9])
                        print('请检查%s行数据是否异常'%line[0])
                else:
                    docx_input['sum'] =format_input['sum'] =int(line[9])
                # print(line[0])
                format_input['name']=str(line[0]) + '_' + line[1] + '_' + line[3]+ '_' +str(format_input['sum'])
                self.docx_list.append(docx_input)
            self.output.append(format_input)
    # # 波纹管板材简算
    # global result_success,result_1219,result_1000,estimate_1219
    # result_success = '所有图形均已绘制完毕~'
    # print(result_success)
    # if bellows_weight_1219 != 0:
    #     result_1219 = '未绘制的波纹管所需板幅1219总重量为: %d kg' %bellows_weight_1219
    #     print(result_1219)
    # if bellows_weight_1000 != 0:
    #     result_1000 = '未绘制的波纹管所需板幅1000 总重量为: %d kg'%bellows_weight_1000
    #     print(result_1000)
    #     estimate_1219 = '均使用板幅1219所需重量:%d kg' %(bellows_weight_1000*1.22+bellows_weight_1219)
    #     print(estimate_1219)
#