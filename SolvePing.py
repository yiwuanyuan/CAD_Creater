from re import *
class SolvePing():

    @staticmethod
    def pin(p1, p2, item):
        ping = {}
        ping.update({'p1': p1})
        ping.update({'p2': p2})
        ping.update({'thickness': item['thickness']})
        ping.update({'addr': item['addr']})
        ping.update({'material': item['material']})
        ping.update({'sum': item['sum']})
        ping.update({'list_num': item['list_num']})
        ping.update({'pro_name': item['pro_name']})
        ping.update({'part_name': item['part_name']})
        ping.update({'type': item['type']})
        ping.update({'remark': item['remark']})
        ping.update({'name': item['name']})
        return SolvePing(ping).pingArray

    def __init__(self, ping):
      self.pingArray = []
      if ping['type'].find('接管') != -1:
          self.pipPing(ping)
      elif ping['type'].find('圆环') != -1:
          self.ringPing(ping)
      elif ping['type'].find('搭板') != -1:
          self.platePing(ping)

    def platePing(self,ping):
        weld_partL = []
        weld_partW = []
        for i in findall('[L|l]([\d+\.\d+|\d+]+)', ping['remark']):
            weld_partL.append(float(i))
        for i in findall('[W|w]([\d+\.\d+|\d+]+)', ping['remark']):
            weld_partW.append(float(i))
        if len(weld_partL) != 0:
            for i in range(len(weld_partL)):
                if len(weld_partW) == 0:
                    p = {}
                    p['parameter'] = {}
                    p['parameter'].update({'w': ping['p1']})
                    p['parameter'].update({'l': weld_partL[i]})
                    p['type'] = ping['type']
                    p['thickness'] = ping['thickness']
                    p['material'] = ping['material']
                    p['sum'] = ping['sum']
                    p['addr'] = ping['addr']
                    p['list_num'] = ping['list_num']
                    p['part_name'] = ping['part_name']
                    p['pro_name'] = ping['pro_name']
                    p['name'] = ping['name'] + ' 横缝 第' + str(i + 1) + '拼，共' + str(len(weld_partL)) + '拼'
                    p['remark'] = p['name']
                    self.pingArray.append(p)

                else:
                    for j in range(len(weld_partW)):
                        p = {}
                        p['parameter'] = {}
                        p['parameter'].update({'w': weld_partW[j]})
                        p['parameter'].update({'l': weld_partL[i]})
                        p['type'] = ping['type']
                        p['thickness'] = ping['thickness']
                        p['material'] = ping['material']
                        p['sum'] = ping['sum']
                        p['addr'] = ping['addr']
                        p['list_num'] = ping['list_num']
                        p['part_name'] = ping['part_name']
                        p['pro_name'] = ping['pro_name']
                        p['name'] = ping['name'] + ' 横缝 第' + str(i + 1) + '拼，纵缝 第' + str(j + 1) + '拼，共' + str(
                            len(weld_partL) * len(weld_partW)) + '拼'
                        p['remark'] = p['name']
                        self.pingArray.append(p)
        else:
            for j in range(len(weld_partW)):
                p = {}
                p['parameter'] = {}
                p['parameter'].update({'w': weld_partW[j]})
                p['parameter'].update({'l': ping['p2']})
                p['type'] = ping['type']
                p['thickness'] = ping['thickness']
                p['material'] = ping['material']
                p['sum'] = ping['sum']
                p['addr'] = ping['addr']
                p['list_num'] = ping['list_num']
                p['part_name'] = ping['part_name']
                p['pro_name'] = ping['pro_name']
                p['name'] = ping['name'] + '纵缝 第' + str(j + 1) + '拼，共' + str(len(weld_partW)) + '拼'
                p['remark'] = p['name']
                self.pingArray.append(p)

    def pipPing(self,ping):
        weld_partL = []
        weld_partW = []
        for i in findall('[L|l]([\d+\.\d+|\d+]+)', ping['remark']):
            weld_partL.append(float(i))
        for i in findall('[W|w]([\d+\.\d+|\d+]+)', ping['remark']):
            weld_partW.append(float(i))
        if len(weld_partL) != 0:
            for i in range(len(weld_partL)):
                if len(weld_partW) == 0:
                    p = {}
                    p['parameter'] = {}
                    p['parameter'].update({'w': (ping['p1'] - ping['thickness']) * 3.1415926})
                    p['parameter'].update({'l': weld_partL[i]})
                    p['type'] = ping['type']
                    p['thickness'] = ping['thickness']
                    p['material'] = ping['material']
                    p['sum'] = ping['sum']
                    p['addr'] = ping['addr']
                    p['list_num'] = ping['list_num']
                    p['part_name'] = ping['part_name']
                    p['pro_name'] = ping['pro_name']
                    p['name'] = ping['name'] + ' 环缝 第' + str(i + 1) + '拼，共' + str(len(weld_partL)) + '拼'
                    p['remark'] = p['name']
                    self.pingArray.append(p)

                else:
                    for j in range(len(weld_partW)):
                        p = {}
                        p['parameter'] = {}
                        p['parameter'].update({'w': weld_partW[j]})
                        p['parameter'].update({'l': weld_partL[i]})
                        p['type'] = ping['type']
                        p['thickness'] = ping['thickness']
                        p['material'] = ping['material']
                        p['sum'] = ping['sum']
                        p['addr'] = ping['addr']
                        p['list_num'] = ping['list_num']
                        p['part_name'] = ping['part_name']
                        p['pro_name'] = ping['pro_name']
                        p['name'] = ping['name'] + ' 环缝 第' + str(i + 1) + '拼，纵缝 第' + str(j + 1) + '拼，共' + str(
                            len(weld_partL) * len(weld_partW)) + '拼'
                        p['remark'] = p['name']
                        self.pingArray.append(p)
        else:
            for j in range(len(weld_partW)):
                p = {}
                p['parameter'] = {}
                p['parameter'].update({'w': weld_partW[j]})
                p['parameter'].update({'l': ping['p2']})
                p['type'] = ping['type']
                p['thickness'] = ping['thickness']
                p['material'] = ping['material']
                p['sum'] = ping['sum']
                p['addr'] = ping['addr']
                p['list_num'] = ping['list_num']
                p['part_name'] = ping['part_name']
                p['pro_name'] = ping['pro_name']
                p['name'] = ping['name'] + '纵缝 第' + str(j + 1) + '拼，共' + str(len(weld_partW)) + '拼'
                p['remark'] = p['name']
                self.pingArray.append(p)

    def ringPing(self, ping):
        i = int(findall('\d+\.\d+|\d+', ping['remark'])[0])
        p = {}
        p['parameter'] = {}
        p['parameter'].update({'od': ping['p1']})
        p['parameter'].update({'id': ping['p2']})
        p['parameter'].update({'angle': 360/i})
        p['type'] = ping['type']
        p['thickness'] = ping['thickness']
        p['material'] = ping['material']
        p['sum'] = str(ping['sum'] * i)
        p['addr'] = ping['addr']
        p['list_num'] = ping['list_num']
        p['part_name'] = ping['part_name']
        p['pro_name'] = ping['pro_name']
        p['name'] = ping['addr']+'_'+str(ping['list_num']) + '_' + ping['pro_name'] + '_' + ping['part_name'] +'_'+ str(i) + '拼' + '_' + '(' + p['sum']+')'
        p['remark'] = p['name']
        self.pingArray.append(p)

