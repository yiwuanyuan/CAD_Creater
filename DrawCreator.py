
from dxfwrite import DXFEngine as dxf #导入模块
from os import getcwd,mkdir,path,chdir,close
import math
import re
from ProcessError import *
dirc = getcwd()
def DrawCreator(name ,parameter , material , thickness , addr , address=dirc):

    file_name = name + '.dxf' #加一步判据
    file_name = re.sub('[\\\]|[\/]', '_', file_name)
    file_name = re.sub('[\*]', 'X', file_name)
    drawing = dxf.drawing(file_name)
    if 'od'in parameter and not('id'in parameter):

        drawing.add(dxf.circle(radius=(parameter['od'])/2,center=(0,0)))
    elif 'od'in parameter and 'id'in parameter and not('angle'in parameter):

        drawing.add(dxf.circle(radius=(parameter['od'])/2, center=(0, 0)))
        drawing.add(dxf.circle(radius=(parameter['id'])/2, center=(0, 0)))
    elif 'od' in parameter and 'id' in parameter and 'angle' in parameter:

        drawing.add(dxf.arc(radius=parameter['od']/2,center=(0, 0),startangle=0,endangle=parameter['angle']))
        drawing.add(dxf.arc(radius=(parameter['id'])/2,center=(0, 0),startangle=0,endangle=parameter['angle']))
        drawing.add(dxf.polyline(points=[(parameter['od']/2, 0), ((parameter['id'])/2,0)]))
        drawing.add(dxf.polyline(points=[((parameter['od'])/2*math.cos(parameter['angle']*math.pi/180), (parameter['od'])/2*math.sin(parameter['angle']*math.pi/180)),
                                         ((parameter['id']) / 2 * math.cos(parameter['angle'] * math.pi / 180),(parameter['id']) / 2 * math.sin(parameter['angle'] * math.pi / 180))]))
    elif 'w' in parameter and 'l' in parameter:

        drawing.add(
            dxf.polyline(
                points=[(0,0),(0,parameter['w']),(parameter['l'],parameter['w']),(parameter['l'],0),(0,0)]
            )
        )
    elif 'array' in parameter:
        for i in range(360):
            drawing.add(dxf.polyline(points=[(parameter['array'][i]['x'], parameter['array'][i]['y']),(parameter['array'][i+1]['x'], parameter['array'][i+1]['y'])]))
            if not parameter['H&T']:
                drawing.add(dxf.polyline(points=[(parameter['array'][i]['x'], -(parameter['array'][i]['y'])),(parameter['array'][i + 1]['x'], -(parameter['array'][i + 1]['y']))]))
        if not parameter['H&T']:
            drawing.add(dxf.polyline(points=[(parameter['array'][0]['x'], parameter['array'][0]['y']),
                                             (parameter['array'][0]['x'], -(parameter['array'][0]['y']))]))
            drawing.add(dxf.polyline(points=[(parameter['array'][360]['x'], parameter['array'][360]['y']),
                                             (parameter['array'][360]['x'], -(parameter['array'][360]['y']))]))
        else:
            drawing.add(dxf.polyline(points=[(parameter['array'][0]['x'], parameter['array'][0]['y']),
                                             (parameter['array'][0]['x'], 0)]))
            drawing.add(dxf.polyline(points=[(parameter['array'][360]['x'], parameter['array'][360]['y']),
                                             (parameter['array'][360]['x'], 0)]))
            drawing.add(dxf.polyline(points=[(parameter['array'][0]['x'], 0),
                                             (parameter['array'][360]['x'], 0)]))
    else:
        raise ProcessError(file_name+'- - -无对应图形', False)

    # 创建对应dxf文件夹
    if not path.exists(address + '//' + 'dxf'):
        mkdir(address + '//' + 'dxf')

        # 创建对应Excel名字的文件夹
    if not path.exists(address + '//' + 'dxf' + '//' + addr):
        mkdir(address + '//' + 'dxf' + '//' + addr)

    # 创建对应_厚度_的文件夹
    if not path.exists(address + '//' + 'dxf' + '//' + addr + '//' + '材料_' + material + '_厚度_' + str(thickness)):
        mkdir(address + '//' + 'dxf' + '//' + addr + '//' + '材料_' + material + '_厚度_' + str(thickness))

    # 切换当前工作地址到相应厚度材质的文件夹下
    chdir(address + '\\dxf' + '\\' + addr + '\\材料_' + material + '_厚度_'+str(thickness))
    drawing.save()
    # 切换当前工作地址到该项目的根目录，解除对dxf及其子文件夹的占用
    chdir(address)
