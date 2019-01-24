#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2019/1/3 20:44
# @Author  : WangYuan
# @Site    : 
# @File    : str2map.py
# @Software: PyCharm
from re import split,match,sub
import re

class Str2map:
    # 判断输入类型是否为波纹管
    def __init__(self,str):
        self.num = 1
        self.error = {}
        if match('^b|B[0-9]', str):
            self.type ='波纹管'
            bellows = []
            for i in split('[*-]',str.strip('B').strip('-')):#用re.split可以分割多个分割符
                bellows.append(float(i))  #波纹管参数汇总
            self.parameter = {}
            self.parameter.update({'w': (bellows[0] + bellows[-2]) * 3.14})
            self.parameter.update({"l":2*bellows[1]*bellows[3]+2*90+(bellows[3]*3.14159-2*bellows[3]-1)*bellows[2]/2})
            self.thickness = bellows[-2]
            self.num =int( bellows[-3])
            return
            # #波纹管重量计算 如果波纹管材料_纵向展开过短将产生dxf文件
            # if  format_input['parameter'][1] > 1219:
            #     bellows_weight_1219 = format_input['parameter'][0] * format_input['parameter'][1] * bellows[-2] * 0.00000785 * format_input['sum'] + bellows_weight_1219
            # elif 950 < format_input['parameter'][1] < 1219 :
            #     bellows_weight_1219 = format_input['parameter'][0] * 1219 * bellows[-2] * 0.00000785*format_input['sum'] + bellows_weight_1219
            # elif 700 < format_input['parameter'][1] < 950:
            #     bellows_weight_1000 = format_input['parameter'][0] * 1000 * bellows[-2] * 0.00000785 * format_input['sum'] + bellows_weight_1000
            # else:
            #     DrawCreator(format_input['name'], format_input['type'], format_input['parameter'],format_input['material'],format_input['thickness'],address)

        elif match('^[J|j][G|g]',str) or match('^[Z|z][J|j][G|g]',str) or match('^[J|j][L|l][H|h]',str) or match('^[D|d][H|h]',str) :   #接管、中间接管、剪力环以号代图
            self.type='接管'
            self.parameter={}
            temp=[]
            for i in re.findall('[^\d]*(\d+)\.?[^\d]*',str):
                temp.append(float(i))
            if(re.findall('[N|n]',str)):
                self.parameter.update({'w': (temp[0] + temp[1]) * 3.1415926})
            else:
                self.parameter.update({'w': (temp[0] - temp[1]) * 3.1415926})
            self.parameter.update({'l': temp[2]})
            self.thickness=float(temp[1])
            return

        elif match('[E|e][B|b]',str) :   #耳板以号代图
            self.type = '耳板'
            self.parameter = {}
            temp = []
            for i in re.findall('[^\d]*(\d+)\.?[^\d]*',str):
                temp.append(float(i))  # 接管参数汇总
            self.parameter.update({'w': temp[1]})
            self.parameter.update({'l': temp[2]})
            self.thickness = float(temp[4])
            return

        elif match('[D|d][E|e]',str) :          #吊耳以号代图
            self.type = '吊耳'
            self.parameter = {}
            temp = []
            for i in re.findall('[^\d]*(\d+)\.?[^\d]*', str):
                temp.append(float(i))  # 接管参数汇总
            self.parameter.update({'w':temp[0]})
            self.parameter.update({'l': temp[1]})
            self.thickness = float(temp[3])
            return

        elif match('^[H|h][B|b]',str):          #环板以号代图
            self.type = '环板'
            self.parameter = {}
            temp = []
            for i in re.findall('[^\d]*(\d+)\.?[^\d]*', str):
                temp.append(float(i))
            self.parameter.update({'od':temp[1]})
            self.parameter.update({'id':temp[2]})
            self.thickness = float(temp[3])
            return

        elif match('^[N|n][C|c]', str) :   #内衬筒类以号代图
            self.type = '内衬筒'
            self.parameter = {}
            temp = []
            for i in  re.findall('[^\d]*(\d+)\.?[^\d]*', str):
                temp.append(float(i))
            if(len(re.findall('\*',str))==1):
                self.parameter.update({'w':(temp[0]-temp[1])*3.1415926})
                self.parameter.update({'l':temp[2]})
                self.thickness=float(temp[1])
            else:
                if((temp[0]-temp[1])/2>30):
                    self.parameter.update({'l': temp[3] + 10})
                else:
                    self.parameter.update({'l': temp[3] + 5})
                self.parameter.update({'w': (temp[1]-temp[2]) * 3.1415926})
                self.thickness=float(temp[2])
            return
        elif match('^[W|w][T|t]',str):  #外护套类以号代图
            self.type = '外护套'
            self.parameter = {}
            temp = []
            for i in re.findall('[^\d]*(\d+)\.?[^\d]*', str):
                temp.append(float(i))
            self.parameter.update({'w':(temp[0]-temp[1])*3.1415926})
            self.parameter.update({'l':temp[2]})
            self.thickness=float(temp[1])
            return
        elif match('^[K|k][W|w][T|t]',str):  #可拆卸外护套类以号代图
            self.type = '可拆外护套'
            self.parameter = {}
            temp = []
            for i in re.findall('[^\d]*(\d+)\.?[^\d]*', str):
                temp.append(float(i))
            if temp[0]==1:
                self.parameter.update({'w':(temp[1]+temp[2])*3.14154926/2+60})
            else:
                self.parameter.update({'w':(temp[1] + temp[2]) * 3.14154926 / 2})
            self.parameter.update({'l':temp[3]})
            self.thickness=float(temp[2])
            self.num=2
            return
        else:
            self.error['code'] = "003"
            self.error['content'] = "该以号代图未收录:"+str



