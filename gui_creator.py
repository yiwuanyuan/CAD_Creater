#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : gui_creator.py
# Author: Wangyuan
# Date  : 2019-1-3
from OutputDocx import *
from GetExcelInfo import *
from DrawCreator import *
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilenames
from re import sub
def gui_creator():
    #图形界面设计
    w = tk.Tk()
    w.title('排料之光')
    w.geometry('300x400')

    #创建显示框
    f1=tk.Frame(w)
    f1.pack()
    f2=tk.Frame(w)
    f2.pack()
    result1 = tk.StringVar()
    result2 = tk.StringVar()
    result3 = tk.StringVar()
    result4 = tk.StringVar()
    global result5,result6
    result5 = tk.StringVar()
    result6 = tk.StringVar()
    # 显示输出结果
    tk.Label(f1, textvariable=result1).pack()
    tk.Label(f1, textvariable=result2).pack()
    tk.Label(f1, textvariable=result3).pack()
    tk.Label(f1, textvariable=result4).pack()
    tk.Label(f1, textvariable=result5).pack()
    tk.Label(f1, textvariable=result6).pack()
    def sel_doc():
        path_ = askopenfilenames(filetypes=[("text file", "*.xlsx"), ("all", "*.*")], )
        path.set(path_)
        path.get()

    def let_work():
        if  path.get():
            addr = path.get().strip('(').strip(')').split(',')
            for ad in addr:
                if ad!= '':
                    g = GetExcelInfo(ad.strip().strip('\''))
                    doc_addr = '/'.join(ad.strip().strip('\'').split('/')[:-1])
                    for i in g.output:
                        print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'],)
                        DrawCreator(i['name'], i['type'], i['parameter'], i['material'], i['thickness'], address=doc_addr)
        else:
            messagebox.showinfo(title='提示', message='至少选择一个Excel文件')

        #初始化显示数据
        # result1.set(result_success)
        # result2.set(result_1219)
        # result3.set(result_1000)
        # result4.set(estimate_1219)

    def let_docx_work():
        if path.get():
            addr = path.get().strip('(').strip(')').split(',')
            for ad in addr:
                if ad != '':
                    xls_addr = ad.strip().strip('\'')
                    g = GetExcelInfo(xls_addr)
                    print(g.docx_list)
                    OutputDocx(g.docx_list,address=xls_addr)

        else:
             messagebox.showinfo(title='提示', message='至少选择一个Excel文件')
        result5.set('文件已生成')

    tk.Label(f2, text="目标路径:").pack()
    path = tk.StringVar()
    tk.Entry(f2, textvariable = path).pack()

    tk.Button(f2,text='选择文件所在位置',command=sel_doc).pack()
    # tk.Button(f2,text='选择文件夹所在位置',command=sel_doc).pack()
    tk.Button(w,text='一键生成dxf数据',command=let_work).pack()
    tk.Button(w,text='一键生成下料清单',command=let_docx_work).pack()

    w.mainloop()

gui_creator()