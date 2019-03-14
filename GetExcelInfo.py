from pandas import read_excel
from ProcessError import *
from os import getcwd
from re import match, findall
from StrToGraphy import *
import math
from SolvePing import *
from InfoPro import *

dirc = getcwd()
global table_finish, table_result


def a2r(angle):
    # 角度换弧度公式
    return angle / 180 * math.pi



class GetExcelInfo:

    def __init__(self, address):
        # 设定输出到下料清单的数据

        self.output = []
        self.error = []
        self.filter = []
        print(address)
        # 判断模板Excel的类型

        #新增加判断模板类型
        try:
            if address.find('排料清单') != -1:
                inform = InfoPro(address).form_out
            else:
                inform = read_excel(address).values
        #
        except Exception:
            temp = 'excel表转换异常！！！'
            self.error.push(temp)
            print(temp)

        address = re.findall('(.*).xlsx',address.split('/')[-1])[0]
        # 地址信息简化（只保留文件名）

        for line in inform:
            try:

                format_input = {}
                format_input['parameter'] = {}

                list_num = line[0]
                pro_name = line[1]
                part_name = line[2]
                material = line[3]
                thickness = line[4]
                p1 = line[5]
                # 长度、外径、环板外直径、以号代图
                p2 = line[6]
                # 宽度、接管长、环板内直径
                mark_sum = line[7]
                remark = line[9]

                if str(pro_name).find('*') == -1 and p1 != p1 and p2 != p2:
                    raise ProcessError(pro_name + part_name + '- - -该零件不会绘制图形', True)

                if material != material:
                    raise ProcessError('材料未定义！', False)
                else:
                    material = str(material)

                if math.isnan(mark_sum):
                    raise ProcessError('数量未定义！', False)
                else:
                    mark_sum = int(mark_sum)

                if findall('\*', str(pro_name)):  # 如果列6是以号代图,则识别以号代图
                    strtography = {}
                    strtography.update({'addr': address})
                    strtography.update({'pro_name': pro_name})
                    strtography.update({'list_num': list_num})
                    strtography.update({'part_name': part_name})
                    strtography.update({'material': material})
                    strtography.update({'sum': mark_sum})
                    strtography.update({'remark': remark})
                    format_input = StrToGraphy(strtography).item

                else:
                    if math.isnan(thickness):
                        raise ProcessError('厚度未定义！', False)
                    else:
                        thickness = float(thickness)
                    # 判断除以号代图外的其他输入类型
                    # 初始化下清单的参数
                    format_input['addr'] = address
                    format_input['list_num'] = list_num
                    format_input['pro_name'] = pro_name
                    format_input['part_name'] = part_name
                    format_input['remark'] = remark
                    format_input['sum'] = mark_sum
                    format_input['material'] = str(material)
                    format_input['thickness'] = thickness
                    #王元修改
                    if format_input['addr'].find('排料清单') != -1:
                        format_input['name'] =str(list_num) + '_' + str(pro_name) + '_' + str(part_name) + '(' + str(format_input['sum'])+')'
                    else:
                        format_input['name'] = str(format_input['addr']) + '_' + str(list_num) + '_' + str(
                            pro_name) + '_' + str(part_name) + '(' + str(format_input['sum']) + ')'

                    # 弯头
                    if part_name.find('弯头') != -1:
                        format_input['type'] = '弯头'
                        p1 = float(findall('\d+\.\d+|\d+', str(p1))[0])
                        d = p1 - thickness
                        angle = float(findall('[a|A][n|N][g|G][l|L][e|E]([\d+\.\d+|\d+]+)', str(remark))[0])
                        angle = a2r(angle)
                        start = float(findall('[s|S][t|T][a|A][r|R][t|T]([\d+\.\d+|\d+]+)', str(remark))[0])
                        div = float(findall('[d|D][i|I][v|V]([\d+\.\d+|\d+]+)', str(remark))[0])
                        ping = []
                        for j in range(int(math.pi / 2 / angle + 1)):
                            temp = {}
                            temp['parameter'] = {}
                            array = []
                            for i in range(361):
                                current = a2r(start + i)
                                loc = {}
                                loc.update({'x': a2r(i) * d / 2})
                                loc.update({'y': d / 2 * (2 - math.cos(current)) * math.tan(angle / 2)})
                                array.append(loc)
                            temp['parameter'].update({'array': array})
                            if j == 0 or j == int(math.pi / 2 / angle):
                                temp['parameter'].update({'H&T': True})
                            else:
                                temp['parameter'].update({'H&T': False})
                            # 判断是否为首尾
                            temp['type'] = format_input['type']
                            temp['sum'] = format_input['sum']
                            temp['material'] = format_input['material']
                            temp['thickness'] = format_input['thickness']
                            temp['addr'] = format_input['addr']
                            temp['list_num'] = format_input['list_num']
                            temp['part_name'] = format_input['part_name']
                            temp['pro_name'] = format_input['pro_name']
                            temp['name'] = format_input['name'] + ',第' + str(j + 1) + '拼' + ',共' + str(int(math.pi / 2 / angle + 1)) + '拼'
                            temp['remark'] = temp['name']
                            ping.append(temp)
                            start = start + div
                        format_input.update({'ping': ping})
                    # 整圆
                    elif p2 != p2 or p2 == 0:
                        # 内直径为空
                        p1 = float(findall('\d+\.\d+|\d+', str(p1))[0])
                        format_input['type'] = '整圆'
                        format_input['parameter'].update({'od': p1})

                    # 接管
                    elif part_name.find("接管") != -1 or part_name.find("筒") != -1 or ((str(p1).find("径") != -1 and str(p2).find("径") == -1)):
                        format_input['type'] = '接管'
                        p1 = float(findall('\d+\.\d+|\d+', str(p1))[0])
                        p2 = float(findall('\d+\.\d+|\d+', str(p2))[0])
                        format_input['parameter'].update({'w': (p1 - thickness) * 3.1415926})

                        # 判断是否拼接
                        if str(remark).find('拼') == -1:
                            format_input['parameter'].update({'l': p2})

                        else:
                            format_input['ping'] = SolvePing.pin(p1, p2, format_input)



                    elif part_name.find("环") != -1 or (str(p1).find("径") != -1 and str(p2).find("径") != -1) or (
                            str(p1).find("φ") != -1 and str(p2).find("φ") != -1):
                        p1 = float(findall('\d+\.\d+|\d+', str(p1))[0])
                        p2 = float(findall('\d+\.\d+|\d+', str(p2))[0])
                        format_input['type'] = '圆环'

                        format_input['parameter'].update({'od': p1})
                        format_input['parameter'].update({'id': p2})
                        if str(remark).find('拼') != -1:
                            format_input['ping'] = SolvePing.pin(p1, p2, format_input)



                    # 判断是否为弧板
                    elif str(remark).find('°') != -1:

                        p1 = float(findall('\d+\.\d+|\d+', str(p1))[0])
                        p2 = float(findall('\d+\.\d+|\d+', str(p2))[0])
                        angle = float(findall('(\d+\.\d+|\d+)°', str(remark))[0])
                        format_input['type'] = '弧板'

                        if str(remark).find('卷') == -1:
                            format_input['parameter'].update({'od': (p1 + p2) * 2})
                            format_input['parameter'].update({'id': p1 * 2})
                            format_input['parameter'].update({'angle': angle})
                        else:
                            temp = format_input['sum']
                            new_p1 = (p1 * 2 + p2) * 3.1415926
                            n = math.ceil(format_input['sum'] / math.floor(360 / float(re.findall('\d+\.\d+|\d+', remark)[0])))
                            new_p2 = thickness * n
                            format_input['parameter'].update({'w': new_p1})
                            format_input['parameter'].update({'l': new_p2})
                            format_input['thickness'] = float(p2)
                            format_input['type'] = '接管'
                            format_input['sum'] = 1
                            format_input['name'] =str(format_input['addr'])+'_'+str(list_num) + '_' + str(pro_name) + '_' + str(part_name) + '_' + '(' + str(format_input['sum'])+')' + '一个接管可以做' + str(
                                temp) + '件'
                            if str(remark).find('拼') == -1:
                                format_input['remark'] = format_input['name']
                            else:
                                format_input['ping'] = SolvePing.pin(new_p1, new_p2, format_input)


                    else:  # 其余都作为搭板处理，如还不能处理则报错
                        format_input['type'] = '搭板'
                        p1 = float(findall('\d+\.\d+|\d+', str(p1))[0])
                        p2 = float(findall('\d+\.\d+|\d+', str(p2))[0])
                        format_input['parameter'].update({'w': p1})
                        format_input['parameter'].update({'l': p2})
                        if str(remark).find('拼') != -1:
                            format_input['ping'] = SolvePing.pin(p1, p2, format_input)



            except Exception as e:
                print(address+'序号'+str(list_num)+'行数据获取异常！！！')
                print(e)
                # temp=''
                item = {}
                item.update({'addr': address})
                item.update({'list_num': list_num})
                item.update({'pro_name': pro_name})
                item.update({'part_name': part_name})
                item.update({'sum': mark_sum})
                item.update({'remark': remark})
                item.update({'p1': p1})
                item.update({'p2': p2})
                try:
                    item.update({'error': e.info})
                    # temp=e.info+'\n'
                except Exception:
                    pass
                # temp = temp+address.strip('\'').split('/')[-1]+'序号'+str(list_num)+'行数据获取异常！！！'
                try:
                    if e.filter:
                        # self.filter.append(temp)
                        self.filter.append(item)
                    else:
                        # self.error.append(temp)
                        self.error.append(item)
                except Exception:
                    # self.error.append(temp)
                    self.error.append(item)
                # print(temp)
                continue

            if 'ping' in format_input:
                for i in format_input['ping']:
                    self.output.append(i)
                    # print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'], i['sum'])
            else:
                self.output.append(format_input)
                # print(format_input['name'], format_input['type'], format_input['parameter'],format_input['material'], format_input['thickness'], format_input['sum'])
        # for i in self.output:
        #     if i['type'] == '弯头':
        #         print(i['name'], i['type'], i['material'], i['thickness'], i['sum'])
        #     else:
        #         print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'], i['sum'])

if __name__ == '__main__':
    g = GetExcelInfo('MM-2109C  0-0.xlsx')
    print(g.output[0])