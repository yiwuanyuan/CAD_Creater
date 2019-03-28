from PyQt5 import QtWidgets
from PyQt5.Qt import *
from NestUI import Ui_Form
from GetExcelInfo import *
from DrawCreator import *
from FileOperator import *
from OutputDocx import *
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem
from PyQt5.QtCore import QStringListModel
import sys
class MyWindow(QtWidgets.QWidget,Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.infoTable.setColumnWidth(0,100)
        self.infoTable.setColumnWidth(1,40)
        self.infoTable.setColumnWidth(2,100)
        self.infoTable.setColumnWidth(3,100)
        self.infoTable.setColumnWidth(4,100)
        self.infoTable.setColumnWidth(5,80)
        self.infoTable.setColumnWidth(6,80)
        self.infoTable.setColumnWidth(7,50)
        self.infoTable.setColumnWidth(8,50)
        self.infoTable.setColumnWidth(9,100)
        self.infoTable.verticalHeader().setVisible(False)
        self.infoTable.setAlternatingRowColors(True)
        self.filterTable.setColumnWidth(0, 100)
        self.filterTable.setColumnWidth(1, 40)
        self.filterTable.setColumnWidth(2, 100)
        self.filterTable.setColumnWidth(3, 100)
        self.filterTable.setColumnWidth(4, 100)
        self.filterTable.setColumnWidth(5, 100)
        self.filterTable.setColumnWidth(6, 50)
        self.filterTable.setColumnWidth(7, 300)
        self.filterTable.verticalHeader().setVisible(False)
        self.filterTable.setAlternatingRowColors(True)
        self.errorTable.setColumnWidth(0,100)
        self.errorTable.setColumnWidth(1,40)
        self.errorTable.setColumnWidth(2,100)
        self.errorTable.setColumnWidth(3,200)
        self.errorTable.setColumnWidth(4,600)
        self.errorTable.verticalHeader().setVisible(False)
        self.errorTable.setAlternatingRowColors(True)
        # 设置UI界面样式

        self.dxf_addr = ''
        self.addrs=[]

        # 成员变量初始化

    def summarize(self):
        if self.dxf_addr != '':
            flag = True
            path = self.dxf_addr
            target_path = re.sub('dxf', '汇总', path)
            pronest_path = re.sub('汇总', '排料清单', target_path)
            try:
                FileOperator.copy_method(path, target_path)
                if not os.path.exists(pronest_path):
                    os.mkdir(pronest_path)
            except Exception as e:
                print(e)
                flag = False
                QMessageBox.warning(self, "提示", "请关闭ProNest且跳转至排料文件夹根目录后重试")
            if flag:
                QMessageBox.warning(self, "提示", "汇总完毕！")
        else:
            QMessageBox.warning(self, "提示", "请先生成一次dxf")

    def clear(self):
        self.infoTable.clearContents()
        self.infoTable.setRowCount(0)
        self.filterTable.clearContents()
        self.filterTable.setRowCount(0)
        self.errorTable.clearContents()
        self.errorTable.setRowCount(0)
        temp = []
        slm = QStringListModel()
        slm.setStringList(temp)
        self.list.setModel(slm)
        self.addrs = []

    def selectExcel(self):
        self.files, filetype = QtWidgets.QFileDialog.getOpenFileNames(self, "多文件选择", '', "Excel Files (*.xlsx)")
        print(self.files)
        self.addrs = self.files
        temp = []
        slm = QStringListModel()
        for addr in self.addrs:
            temp.append(addr.split('/')[-1])
        print(temp)
        slm.setStringList(temp)
        self.list.setModel(slm)

    def letDxfWork(self):
        if len(self.addrs) != 0:
            flag = True
            rowCountInfo = 0
            rowCountFilter = 0
            rowCountError = 0
            # 所有表格初始化（清空）
            self.infoTable.clearContents()
            self.infoTable.setRowCount(0)
            self.filterTable.clearContents()
            self.filterTable.setRowCount(0)
            self.errorTable.clearContents()
            self.errorTable.setRowCount(0)
            for ad in self.addrs:
                if ad != '':
                    excel_addr = ad.strip().strip('\'')
                    # 获取每张excel表地址

                    temp_array = excel_addr.split('/')
                    temp_path = ''
                    for i in range(len(temp_array)-1):
                        temp_path += temp_array[i]+'\\'
                    self.dxf_addr = temp_path+'dxf'
                    temp_path += 'dxf'+'\\'+re.findall('(.*).xlsx', temp_array[-1])[0]
                    try:
                        FileOperator.del_method(temp_path)
                        # 删除原来已经生成的dxf文件及文件夹
                    except Exception as e:
                        print(e)
                        flag = False
                        QMessageBox.warning(self, "提示", "请跳转至排料文件夹根目录后重试")
                    g = GetExcelInfo(excel_addr)
                    info = g.output
                    error = g.error
                    filter = g.filter
                    doc_addr = '/'.join(ad.strip().strip('\'').split('/')[:-1])

                    for i in g.output:
                        try:
                            DrawCreator(i['name'], i['parameter'], i['material'], i['thickness'], i['addr'], doc_addr)
                        except Exception as e:
                            temp = {}
                            try:
                                temp.update({'error': e.info})
                            except Exception as ex:
                                pass
                            temp.update({'addr': i['addr']})
                            temp.update({'list_num': i['list_num']})
                            temp.update({'part_name': i['part_name']})
                            temp.update({'pro_name': i['pro_name']})
                            error.append(temp)
                            print(e)
                            flag = False
                            QMessageBox.warning(self, "提示", "请关闭CAD后重试")

                    rowCountInfo = self.addinfo(info, rowCountInfo)
                    rowCountFilter = self.addfilter(filter, rowCountFilter)
                    rowCountError = self.adderror(error, rowCountError)
            if flag:
                QMessageBox.warning(self, "提示", "所有excel已解析完毕")
        else:
            QMessageBox.warning(self, "提示", "至少选择一个Excel文件")


    def adderror(self, error , rowCount):
            self.errorTable.setRowCount(len(error)+rowCount)
            for i in range(rowCount, rowCount+len(error)):
                try:
                    self.errorTable.setItem(i, 0, QTableWidgetItem(error[i-rowCount]['addr']))

                    if error[i - rowCount]['list_num'] != error[i - rowCount]['list_num']:
                        self.errorTable.setItem(i, 1, QTableWidgetItem(''))
                    else:
                        self.errorTable.setItem(i, 1, QTableWidgetItem(str(int(error[i - rowCount]['list_num']))))

                    self.errorTable.setItem(i, 2, QTableWidgetItem(str(error[i-rowCount]['part_name'])))

                    if error[i - rowCount]['pro_name'] != error[i - rowCount]['pro_name']:
                        self.errorTable.setItem(i, 3, QTableWidgetItem(''))
                    else:
                        self.errorTable.setItem(i, 3, QTableWidgetItem(str(error[i-rowCount]['pro_name'])))

                    if 'error' in error[i - rowCount]:
                        self.errorTable.setItem(i, 4, QTableWidgetItem((error[i - rowCount]['error'])))
                    else:
                        self.errorTable.setItem(i, 4, QTableWidgetItem('错误未能识别，请联系技术人员！！！'))

                except Exception as e:
                    print(e)
            return rowCount+len(error)

    def addfilter(self, filter , rowCount):
            self.filterTable.setRowCount(len(filter)+rowCount)
            for i in range(rowCount, rowCount+len(filter)):
                try:
                    self.filterTable.setItem(i, 0, QTableWidgetItem(filter[i-rowCount]['addr']))

                    if filter[i-rowCount]['list_num'] != filter[i-rowCount]['list_num']:
                        self.filterTable.setItem(i, 1, QTableWidgetItem(''))
                    else:
                        self.filterTable.setItem(i, 1, QTableWidgetItem(str(int(filter[i-rowCount]['list_num']))))

                    self.filterTable.setItem(i, 2, QTableWidgetItem(str(filter[i-rowCount]['part_name'])))

                    if filter[i - rowCount]['p1'] != filter[i - rowCount]['p1']:
                        self.filterTable.setItem(i, 3, QTableWidgetItem(''))
                    else:
                        self.filterTable.setItem(i, 3, QTableWidgetItem(str(filter[i-rowCount]['p1'])))

                    if filter[i - rowCount]['p2'] != filter[i - rowCount]['p2']:
                        self.filterTable.setItem(i, 4, QTableWidgetItem(''))
                    else:
                        self.filterTable.setItem(i, 4, QTableWidgetItem(str(filter[i-rowCount]['p2'])))

                    if filter[i - rowCount]['sum'] != filter[i - rowCount]['sum']:
                        self.filterTable.setItem(i, 6, QTableWidgetItem(''))
                    else:
                        self.filterTable.setItem(i, 6, QTableWidgetItem(str(filter[i - rowCount]['sum'])))

                    if filter[i-rowCount]['remark'] != filter[i-rowCount]['remark']:
                        self.filterTable.setItem(i, 7, QTableWidgetItem(''))
                    else:
                        self.filterTable.setItem(i, 7, QTableWidgetItem(str(filter[i-rowCount]['remark'])))
                except Exception as e:
                    print(e)
            return rowCount+len(filter)

    def addinfo(self, info, rowCount):
            self.infoTable.setRowCount(len(info)+rowCount)
            for i in range(rowCount, rowCount+len(info)):
                try:
                    self.infoTable.setItem(i, 0, QTableWidgetItem(info[i-rowCount]['addr']))

                    if info[i - rowCount]['list_num'] != info[i - rowCount]['list_num']:
                        self.infoTable.setItem(i, 1, QTableWidgetItem(''))
                    else:
                        self.infoTable.setItem(i, 1, QTableWidgetItem(str(int(info[i - rowCount]['list_num']))))

                    self.infoTable.setItem(i, 2, QTableWidgetItem(str(info[i - rowCount]['part_name'])))

                    if 'od' in info[i-rowCount]['parameter']:
                        self.infoTable.setItem(i, 3, QTableWidgetItem('外径:'+str(info[i-rowCount]['parameter']['od'])))
                    elif 'w' in info[i-rowCount]['parameter']:
                        self.infoTable.setItem(i, 3, QTableWidgetItem('幅宽:'+str(info[i-rowCount]['parameter']['w'])))
                    if 'id' in info[i-rowCount]['parameter']:
                        self.infoTable.setItem(i, 4, QTableWidgetItem('内径:'+str(info[i-rowCount]['parameter']['id'])))
                    elif 'l' in info[i-rowCount]['parameter']:
                        self.infoTable.setItem(i, 4, QTableWidgetItem('幅长:'+str(info[i-rowCount]['parameter']['l'])))
                    if 'angle' in info[i-rowCount]['parameter']:
                        self.infoTable.setItem(i, 5, QTableWidgetItem('角度:'+str(info[i-rowCount]['parameter']['angle'])))

                    self.infoTable.setItem(i, 6, QTableWidgetItem(str(info[i - rowCount]['material'])))

                    self.infoTable.setItem(i, 7, QTableWidgetItem(str(info[i - rowCount]['thickness'])))

                    self.infoTable.setItem(i, 8, QTableWidgetItem(str(info[i-rowCount]['sum'])))

                    if info[i-rowCount]['remark'] != info[i-rowCount]['remark']:
                        self.infoTable.setItem(i, 9, QTableWidgetItem(''))
                    else:
                        self.infoTable.setItem(i, 9, QTableWidgetItem(str(info[i-rowCount]['remark'])))
                except Exception as e:
                    print(e)
            return rowCount+len(info)

    #新增加
    def letDocxWork(self):

        if len(self.addrs) != 0:
            flag = True
            rowCountInfo = 0
            rowCountFilter = 0
            rowCountError = 0
            # 所有表格初始化（清空）
            self.infoTable.clearContents()
            self.infoTable.setRowCount(0)
            self.filterTable.clearContents()
            self.filterTable.setRowCount(0)
            self.errorTable.clearContents()
            self.errorTable.setRowCount(0)
            for ad in self.addrs:
                if ad != '':
                    excel_addr = ad.strip().strip('\'')
                    # 获取每张excel表地址

                    temp_array = excel_addr.split('/')
                    temp_path = ''
                    for i in range(len(temp_array) - 1):
                        temp_path += temp_array[i] + '\\'
                    self.dxf_addr = temp_path + 'dxf'
                    temp_path += 'dxf' + '\\' + re.findall('(.*).xlsx', temp_array[-1])[0]

                    g = GetExcelInfo(excel_addr)
                    info = g.output

                    try:
                        OutputDocx(info, ad,self.isNeedMaterialList.isChecked())
                    except Exception as e:
                        temp = {}
                        try:
                            temp.update({'error': e.info})
                        except Exception as ex:
                            pass
                        temp.update({'addr': i['addr']})
                        temp.update({'list_num': i['list_num']})
                        temp.update({'part_name': i['part_name']})
                        temp.update({'pro_name': i['pro_name']})
                        error.append(temp)
                        print(e)
                        flag = False
                        QMessageBox.warning(self, "提示", "请关闭CAD后重试")

            if flag:
                QMessageBox.warning(self, "提示", "所有excel已解析完毕")
        else:
            QMessageBox.warning(self, "提示", "至少选择一个Excel文件")

    def helpInfo(self):
        pass



if __name__ =='__main__':
    app = QApplication(sys.argv)
    ex = MyWindow()
    ex.show()
    sys.exit(app.exec_())