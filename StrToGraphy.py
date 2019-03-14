from re import split, match, sub
from ProcessError import *
from SolvePing import *
import re


class StrToGraphy:
    # 判断输入类型是否为波纹管

    def __init__(self, strToGraphy):
        self.strToGraphy = strToGraphy
        string = strToGraphy['pro_name']
        self.item = {}
        self.item['parameter'] = {}
        self.item['material'] = self.strToGraphy['material']
        self.item['sum'] = self.strToGraphy['sum']
        self.item['addr'] = self.strToGraphy['addr']
        self.item['list_num'] = self.strToGraphy['list_num']
        self.item['pro_name'] = self.strToGraphy['pro_name']
        self.item['part_name'] = self.strToGraphy['part_name']
        self.item['remark'] = self.strToGraphy['remark']
        self.item['name'] = str(self.strToGraphy['addr'])+'_'+str(self.strToGraphy['list_num']) + '_' + str(self.strToGraphy['pro_name']) + '_' + str(self.strToGraphy['part_name']) + '_' + '(' + str(self.strToGraphy['sum'])+')'

        if match('^[J|j][G|g]', string) or match('^[Z|z][J|j][G|g]', string) or match('^[J|j][L|l][H|h]', string) or match(
                '^[D|d][H|h]', string):  # 接管、中间接管、剪力环以号代图
            type = '接管'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            if (re.findall('[N|n]', string)):
                p1 = temp[0] + temp[1] * 2
            else:
                p1 = temp[0]
            p2 = temp[2]
            thickness = float(temp[1])
            self.item['type'] = type
            self.item['thickness'] = thickness

            if str(self.strToGraphy['remark']).find('拼') == -1:
                self.item['parameter'].update({'w': (p1 - thickness) * 3.1415926})
                self.item['parameter'].update({'l': p2})

            else:
                self.item['ping'] = SolvePing.pin(p1, p2, self.item)
                # self.Pin(p1, p2,thickness,type)
            return

        elif match('^[E|e][B|b]', string):  # 耳板以号代图
            self.item['type'] = '耳板'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            p1 = temp[1]
            p2 = temp[2]
            self.item['parameter'].update({'w': p1})
            self.item['parameter'].update({'l': p2})
            self.item['thickness'] = float(temp[4])
            if str(self.strToGraphy['remark']).find('拼') != -1:
                self.item['ping'] = SolvePing.pin(p1, p2, self.item)
            return

        elif match('^[D|d][E|e]', string):  # 吊耳以号代图
            self.item['type'] = '吊耳'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            p1 = temp[0]
            p2 = temp[1]
            self['parameter'].update({'w': p1})
            self['parameter'].update({'l': p2})
            self['thickness'] = float(temp[3])
            if str(self.strToGraphy['remark']).find('拼') != -1:
                self.item['ping'] = SolvePing.pin(p1, p2, self.item)
            return

        elif match('^[H|h][B|b]', string):  # 环板以号代图
            type = '环板 圆环'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            p1=temp[1]
            p2=temp[2]
            thickness=float(temp[3])
            self.item['thickness'] = thickness
            self.item['type'] = type
            if str(self.strToGraphy['remark']).find('拼') == -1:
                self.item['parameter'].update({'od': p1})
                self.item['parameter'].update({'id': p2})
            else:
                # self.Pin(p1,p2,thickness,type)
                self.item['ping'] = SolvePing.pin(p1, p2, self.item)
            return

        elif match('^[N|n][C|c]', string):  # 内衬筒类以号代图
            type = '内衬筒 接管'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            if len(re.findall('\*', string)) == 1:
                p1 = temp[0]-temp[1]
                p2 = temp[2]
                thickness = float(temp[1])
                self.item['thickness'] = thickness
                self.item['type'] = type
                if str(self.strToGraphy['remark']).find('拼') == -1:
                    self.item['parameter'].update({'w': p1 * 3.1415926})
                    self.item['parameter'].update({'l': p2})

                else:
                    # self.Pin(p1,p2,thickness,type)
                    self.item['ping'] = SolvePing.pin(p1, p2, self.item)
            else:
                thickness = float(temp[2])
                if ((temp[0] - temp[1]) / 2 > 30):
                    p2 = temp[3] + 10
                else:
                    p2 = temp[3] + 5
                p1 = temp[1] - temp[2]
                self.item['thickness'] = thickness
                self.item['type'] = type
                if str(self.strToGraphy['remark']).find('拼') == -1:
                    self.item['parameter'].update({'w': p1 * 3.1415926})
                    self.item['parameter'].update({'l': p2})
                else:
                    # self.Pin(p1,p2,thickness,type)
                    self.item['ping'] = SolvePing.pin(p1, p2, self.item)
            return

        elif match('^[W|w][T|t]', string):  # 外护套类以号代图
            type = '外护套 接管'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            p1 = temp[0] - temp[1]
            p2 = temp[2]
            thickness = float(temp[1])
            self.item['type'] = type
            self.item['thickness'] = thickness
            if str(self.strToGraphy['remark']).find('拼') == -1:
                self.item['parameter'].update({'w': p1 * 3.1415926})
                self.item['parameter'].update({'l': p2})
            else:
                # self.Pin(p1, p2, thickness, type)
                self.item['ping'] = SolvePing.pin(p1, p2, self.item)
            return

        elif match('^[K|k][W|w][T|t]', string):  # 可拆卸外护套类以号代图
            self.item['type'] = '可拆外护套'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            if temp[0] == 1:
                self.item['parameter'].update({'w': (temp[1] + temp[2]) * 3.14154926 / 2 + 60})
            else:
                self.item['parameter'].update({'w': (temp[1] + temp[2]) * 3.14154926 / 2})
            self.item['parameter'].update({'l': temp[3]})
            self.item['thickness'] = float(temp[2])
            self.item['sum'] = int(self.item['sum']*2)
            self.item['name'] = str(strToGraphy['addr'])+'_'+str(self.strToGraphy['list_num']) + '_' + self.strToGraphy['pro_name'] + '_' + self.strToGraphy['part_name'] + '_' + '('+str(self.strToGraphy['sum'])+')'
            return
        elif match('^[J|j][B|b]', string):
            self.item['type'] = '铰链板'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            self.item['parameter'].update({'w': temp[0]})
            self.item['parameter'].update({'l': temp[1]})
            self.item['thickness'] = float(temp[2])
            return
        elif match('^[L|l][B|b]', string):
            self.item['type'] = '肋板'
            temp1 = []
            temp2 = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp1.append(float(i))
            for i in re.findall('\*', string):
                temp2.append(i)
            if len(temp2) == 3 or len(temp2) == 6:
                p1 = temp1[0]
                p2 = temp1[2]
                thickness = temp1[3]
            elif len(temp2) == 5:
                p1 = temp1[0]
                p2 = temp1[1]
                thickness = temp1[2]
            self.item['parameter'].update({'w': p1})
            self.item['parameter'].update({'l': p2})
            self.item['thickness'] = float(thickness)
            return
        elif match('^[F|f][B|b]', string):
            self.item['type'] = '封板'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            self.item['parameter'].update({'w': temp[1]})
            self.item['parameter'].update({'l': temp[0]})
            self.item['thickness'] = float(temp[3])
            return
        elif match('^[J|j][L|l][B|b]', string):
            self.item['type'] = '铰链立板'
            temp = []
            for i in re.findall('\d+\.\d+|\d+', string):
                temp.append(float(i))
            self.item['parameter'].update({'w': temp[0]})
            self.item['parameter'].update({'l': temp[1]})
            self.item['thickness'] = float(temp[2])
            return
        else:
            errorInfo = string + "- - -该以号代图未收录"
            # print( errorInfo)
            raise ProcessError(errorInfo, False)

