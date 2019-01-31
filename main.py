#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : main.py
# Author: Wangyuan
# Date  : 2019/1/31
from PyQt5 import QtWidgets
import os
from pyqt5.index import Ui_Form
from OutputDocx import *
from GetExcelInfo import *
from DrawCreator import *
from PyQt5.QtWidgets import QMessageBox,QTableWidgetItem,QHeaderView,QTableView

class MyWindow(QtWidgets.QWidget,Ui_Form):
    def __init__(self,parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.cwd = os.getcwd()
        self.omit_gather = []
        self.ShowInfoTable.setEditTriggers(QTableView.NoEditTriggers)


        # self.ShowInfoTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # self.ShowInfoTable.resizeColumnsToContents()  # 设置列宽高按照内容自适应
        self.ShowInfoTable.resizeRowsToContents()  # 设置行宽和高按照内容自适应

        self.ShowInfoTable.setColumnWidth(0, 40)  # 设置1列的宽度
        self.ShowInfoTable.setColumnWidth(1, 180)  # 设置2列的宽度
        self.ShowInfoTable.setColumnWidth(2, 100)  # 设置3列的宽度
        self.ShowInfoTable.setColumnWidth(3, 100)  # 设置4列的宽度
        self.ShowInfoTable.setColumnWidth(4, 30)  # 设置5列的宽度

        # 设置省略表格的格式
        self.ShowOmitTable.setEditTriggers(QTableView.NoEditTriggers)
        self.ShowOmitTable.setColumnWidth(0, 40)  # 设置1列的宽度
        self.ShowOmitTable.setColumnWidth(1, 100)  # 设置2列的宽度
        self.ShowOmitTable.setColumnWidth(2, 100)  # 设置3列的宽度
        self.ShowOmitTable.setColumnWidth(3, 150)  # 设置4列的宽度
        self.ShowOmitTable.setColumnWidth(4, 40)  # 设置5列的宽度
    def selectExcel(self):
        self.files, filetype = QtWidgets.QFileDialog.getOpenFileNames(self,
                                                       "多文件选择",'',
                                                        # 起始路径
                                                       "Excel Files (*.xlsx)")
        # "All Files (*);;PDF Files (*.pdf);;Text Files (*.txt)"

    def letDxfWork(self):

        # 判断是否获取到路径
        if len(self.files) != 0:

            #获取所有选中表格的总行数
            out_len = 0
            for file in self.files:
                g = GetExcelInfo(file.strip().strip('\''))
                out_len += len(g.output)
                print(out_len)
            out_no = 0 #输出表格从O开始
            # 获取到的地址
            for file in self.files:
                if file != '':
                    g = GetExcelInfo(file.strip().strip('\''))
                    info = g.output
                    error = g.error_gather
                    doc_addr = '/'.join(file.strip().strip('\'').split('/')[:-1])
                    # 表格错误提示
                    if g.stop:
                        QMessageBox.warning(self,
                                            "表格错误",
                                            "该表格式有错，已略过",
                                            QMessageBox.Yes | QMessageBox.No)
                    else:


                        for i in g.output:

                            infotable = self.ShowInfoTable
                            infotable.setRowCount(out_len)
                            infotable.setRowHeight(0,10)
                            infotable.setItem(out_no, 0, QTableWidgetItem(str(i['excel_list'])))
                            infotable.setItem(out_no, 1, QTableWidgetItem(i['name']))
                            parameter = i["parameter"]
                            size1 = ''
                            size2 = ''
                            if 'od' in parameter and not ('id' in parameter):
                                size1 = '外径: ' + str(parameter['od'])

                            elif 'od' in parameter and 'id' in parameter and not ('angle' in parameter):
                                size1 = '外径: ' + str(parameter['od'])
                                size2 =' 内径: ' + str(parameter['id'])

                            elif 'od' in parameter and 'id' in parameter and 'angle' in parameter:
                                size1 = '外径: ' + str(parameter['od'])
                                size2 = ' 内径: ' + str(
                                    parameter['id']) + ' 弧度: ' + str(parameter['angle'])

                            elif 'w' in parameter and 'l' in parameter:
                                size1 = '长度: ' + str(int(parameter['l']))
                                size2 = ' 宽度: ' + str(parameter['w'])

                            infotable.setItem(out_no, 2, QTableWidgetItem(size1))
                            infotable.setItem(out_no, 3, QTableWidgetItem(size2))
                            infotable.setItem(out_no, 4, QTableWidgetItem(str(i['sum'])))
                            out_no +=1
                            try:
                                DrawCreator(i['name'], i['type'], i['parameter'], i['material'], i['thickness'],
                                            address=doc_addr)
                            except BaseException:

                                # error_text.set('绘图错误请检查！！！')
                                print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'])
                                print('绘图错误请检查！！！')
                        # addinfo(info)
                        # adderror(error)
                        if g.omit_value:
                            omit = {}
                            omit['name'] = file.split('/')[-1].strip('\'').strip('.xlsx')
                            omit['value'] = g.omit_value
                            self.omit_gather.append(omit)
            #如果存在被略去数据
            if self.omit_gather != []:
                omit_table = self.ShowOmitTable
                omit_len = 0
                for  index in self.omit_gather:
                    omit_len += len(index['value'])

                omit_table.setRowCount(omit_len)
                omit_table.setRowHeight(0, 10)
                row_num = 0
                for index in self.omit_gather:
                    omit_map = index['name']
                    for i in index['value']:
                        omit_table.setItem(row_num, 0, QTableWidgetItem(str(i[0])))
                        omit_table.setItem(row_num, 1, QTableWidgetItem(omit_map))
                        omit_table.setItem(row_num, 2, QTableWidgetItem(i[1]))
                        omit_table.setItem(row_num, 3, QTableWidgetItem(str(i[2])))
                        omit_table.setItem(row_num, 4, QTableWidgetItem(str(i[3])))
                        row_num +=1

                ans = QMessageBox.question(self,
                                     "提示",
                                     "存在被略去数据，是否导出Excel",
                                     QMessageBox.Yes | QMessageBox.No)

                if ans == 16384:
                    df = []
                    out_addr = '/'.join(self.files[0].strip().strip('\'').split('/')[:-1])
                    i = 0
                    writer = ExcelWriter(out_addr + '/被省略的excel.xlsx')
                    for omit_unit in self.omit_gather:
                        df.append(DataFrame(omit_unit['value']))
                        print(omit_unit['name'])
                        df[i].to_excel(writer, sheet_name=omit_unit['name'])
                        i += 1
                    writer.save()
        else:
            QMessageBox.information(self,
                                    "提示",
                                    "至少选择一个Excel文件",
                                    QMessageBox.Yes | QMessageBox.No)
    def letDocxWork(self):
        pass


# class MyDialog(QtWidgets.QDialog,Ui_Dialog):
#     def __init__(self,last_form):
#         super(MyDialog, self).__init__()
#         self.setupUi(self)
#         self.last_form =last_form
#     def yes(self):
#         self.close()
#         self.last_form.close()
#     def no(self):
#         self.close()

if __name__ =='__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    myshow.show()
    sys.exit(app.exec_())