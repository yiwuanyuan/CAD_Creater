#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : DrawCreator.py
# Author: Wangyuan
# Date  : 2019-1-3

from dxfwrite import DXFEngine as dxf #导入模块
from os import getcwd,mkdir,path,chdir
dirc =getcwd()
def DrawCreator(name,type,parameter,material,thickness=0,address =dirc):

    file_name = name + '.dxf' #加一步判据

    drawing = dxf.drawing(file_name)
    if type == '整圆':
        drawing.add(dxf.circle(radius=(parameter[0])/2,center=(0,0)))
        # drawing.add(dxf.text(file_name,height=parameter[0]/4,alignpoint =(0,0)))
    elif type =='环板':
        drawing.add(dxf.circle(radius=(parameter[0])/2, center=(0, 0)))
        drawing.add(dxf.circle(radius=(parameter[1])/2, center=(0, 0)))
        # drawing.add(dxf.text(file_name,height=parameter[1]/4, alignpoint =(0,0)))
    elif type =='接管':
        length = parameter[0]
        width = parameter[1]
        drawing.add(
            dxf.polyline(
                points=[(0,0),(0,width+3),(length,width+3),(length,0),(0,0)]
            )
        )
    else:
        drawing.add(
            dxf.polyline(
                points=[(0, 0), (0, parameter[0]), (parameter[1], parameter[0]), (parameter[1], 0), (0, 0)]
            )
        )
        # 创建对应_厚度_的文件夹
    if not path.exists(address + '//' + '材料_' + material + '_厚度_' + str(thickness)):
            #如果文件夹不存在创建
            mkdir(address + '//' + '材料_' + material + '_厚度_' + str(thickness))
    chdir(address + '\\材料_' + material + '_厚度_'+str(thickness))

    drawing.save()
