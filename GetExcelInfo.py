#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : Creator0.5.py
# Author: WangYuan
# Date  : 2018-12-20

from pandas import read_excel

from os import getcwd
from InfoPro import *
from re import match,findall
from str2map import *
import math
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
            format_input['parameter'] = {}

            list_num = line[0]
            pro_name = line[1]
            pro_sum  = line[2]
            part_name = line[3]
            material = line[4]
            thickness = line[5]
            p1 = line[6]
            # 长度、外径、环板外直径、以号代图
            p2 = line[7]
            # 宽度、接管长、环板内直径
            sum_num = int(line[8])
            mark_sum = line[9]
            remark = line[10]

            if findall('\*', str(p1)):  #如果列6是以号代图,则识别以号代图
                b = Str2map(p1)
                format_input['parameter'] =b.parameter
                format_input['type'] =b.type
                format_input['thickness'] = b.thickness
                format_input['material'] = str(material)
                #如果是波纹管就只修改数量
                if format_input['type'] == '波纹管':
                    format_input['sum'] = int(b.num * pro_sum * sum_num)
                #如果是其他类型的零件就将文件输出
                else:

                    docx_input['product_code'] = pro_name
                    docx_input['part_name'] = part_name
                    docx_input['thickness'] = format_input['thickness']
                    docx_input['material'] = format_input['material']
                    if mark_sum != pro_sum * sum_num:
                        if mark_sum != mark_sum:
                            docx_input['sum'] = format_input['sum'] = int(pro_sum * sum_num)
                        else:
                            docx_input['sum'] = format_input['sum'] = int(mark_sum)
                            # print('请检查%s行数据是否异常' % list_num)
                    else:
                        docx_input['sum'] = format_input['sum'] = int(mark_sum)


                #     if format_input['type'] == '整圆':
                #         docx_input['size'] = '直径: ' + str(format_input['parameter']['od'])
                #     elif format_input['type'] == '接管':
                #         docx_input['size'] = '展开长: ' + str(int(format_input['parameter']['w']) + 1) + ' 宽度: ' + str(format_input['parameter']['l'])
                #     elif format_input['type'] == '圆环':
                #         docx_input['size'] = '外直径: ' + str(format_input['parameter']['od']) + ' 内直径: ' + str(format_input['parameter']['id'])
                #     else:
                #         docx_input['size'] = '长度: ' + str(int(format_input['parameter']['l']) + 1) + ' 宽度: ' + str(format_input['parameter']['w'])
                    self.docx_list.append(docx_input)
                format_input['name'] = str(list_num) + '_' + pro_name + '_' + part_name + '_' + str(format_input['sum'])

            else:
                # 判断除以号代图外的其他输入类型
                # 初始化下清单的参数
                docx_input['product_code'] = pro_name
                docx_input['part_name'] = part_name
                docx_input['material'] = str(material)
                docx_input['thickness'] = thickness
                format_input['sum'] = int(sum_num)
                if p2 != p2 or p2 ==0 or type(p2) == str :
                    # 内直径为空
                    format_input['type']='整圆'
                    format_input['parameter'].update({'od': p1})
                    docx_input['size'] = '直径: ' + str(p1)

                elif part_name.find("接管") != -1 or (str(p1).find("径")!= -1 and str(p2).find("径")== -1):
                    format_input['type']='接管'
                    format_input['parameter'].update({'l': (p1 - thickness) * 3.1415926})
                    docx_input['size'] = '展开长: ' + str(int(p1)+1) + ' 宽度: ' +str(p2)
                    #判断是否拼接
                    if str(remark).find('拼') == -1:
                        format_input['parameter'].update({'w': p2})
                        docx_input['size'] = '展开长: ' + str(int(p1) + 1) + ' 宽度: ' + str(p2)
                    else:
                        weld_part=[]
                        for i in findall('[^\d]*(\d+)\.?[^\d]*',remark):
                            weld_part.append(i)
                        format_input['ping']=[]
                        for i in range(len(weld_part)):
                            p={}
                            p['parameter']={}
                            p['parameter'].update({'l': (p1-thickness)*3.1415926})
                            p['parameter'].update({'w': weld_part[i]})
                            p['type'] = '接管'
                            p['thickness'] = thickness
                            p['material'] = str(material)
                            p['sum'] = sum_num
                            p['name'] = str(list_num) + '_' + pro_name + '_' + part_name + '_' + str(sum_num)+'第'+str(i+1)+'拼，共'+str(len(weld_part))+'拼'
                            format_input['ping'].append(p)

                elif part_name.find("环")!= -1 or part_name.find("B板")!= -1 or (str(p1).find("径")!= -1 and str(p2).find("径")!= -1) or (str(p1).find("φ")!= -1 and str(p2).find("φ")!= -1):

                    format_input['type'] = '圆环'
                    docx_input['size'] = '外直径: ' + str(p1) + ' 内直径: ' + str(p2)
                    format_input['parameter'].update({'od': p1})
                    format_input['parameter'].update({'id': p2})
                    if str(remark).find('拼') != -1:
                        for i in findall('[^\d]*(\d+)[^\d]*', remark):
                            degree = 360/int(i[0])
                            format_input['sum'] = format_input['sum'] * int(i[0])
                        format_input['parameter'].update({'angle': degree})
                #判断是否为弧板
                elif part_name.find('HB')!=-1:
                    format_input['type'] = '弧板'
                    if  thickness<=80:
                        format_input['parameter'].update({'od': (p1 + p2)*2})
                        format_input['parameter'].update({'id': p1*2})
                        format_input['parameter'].update({'angle': float(re.findall('[^\d]*(\d+)\.?[^\d]*',remark)[0])})
                    else:
                        format_input['parameter'].update({'w': (p1*2 + p2)*3.1415926})
                        n=sum_num/math.floor(360/float(re.findall('[^\d]*(\d+)\.?[^\d]*',remark)[0]))
                        format_input['parameter'].update({'l': thickness*n})
                        format_input['thickness']=float(p2)
                        format_input['sum']=1

                else:
                    format_input['type'] = '搭板'
                    format_input['parameter'].update({'w': p1})
                    format_input['parameter'].update({'l': p2})
                    docx_input['size'] = '长度: ' + str(int(p1) + 1) + ' 宽度: ' + str(p2)

                #除波纹管外的其他的类型提供parameter参数和name参数
                # if type(p2) == int or type(p1) == float:
                #     if p2 != 0:
                #         format_input['parameter'].append(p2)

                docx_input['thickness'] = format_input['thickness'] = thickness
                format_input['material'] = str(material)
                format_input['sum'] = mark_sum
                if mark_sum != pro_sum*sum_num:
                    if mark_sum != mark_sum:
                        docx_input['sum'] = format_input['sum'] =int(pro_sum * sum_num)
                    else:
                        docx_input['sum'] = format_input['sum'] = int(mark_sum)
                        # print('请检查%s行数据是否异常'%list_num)
                else:
                    docx_input['sum'] =format_input['sum'] =int(mark_sum)
                # print(list_num)
                format_input['name']=str(list_num) + '_' + pro_name + '_' + part_name+ '_' +str(format_input['sum'])
                self.docx_list.append(docx_input)
            if 'ping' in format_input:
                for i in format_input['ping']:
                    self.output.append(i)
                    print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'], i['sum'])
            else:
                self.output.append(format_input)
                print(format_input['name'], format_input['type'], format_input['parameter'],
                          format_input['material'], format_input['thickness'], format_input['sum'])
        for i in self.output:
            print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'], i['sum'])
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
# g = GetExcelInfo('E:/Program Files/feiq/Recv Files/V2/MM-2108  0-0.xlsx')
