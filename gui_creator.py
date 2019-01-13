#!/usr/bin/env python
# -*- coding: utf-8 -*-
# File  : gui_creator.py
# Author: Wangyuan
# Date  : 2019-1-3
from OutputDocx import *
from GetExcelInfo import *
from DrawCreator import *
import tkinter as tk
from tkinter import messagebox,N,S,W,E,ttk
# from tkinter import *
from tkinter.filedialog import askopenfilenames
from re import sub
def gui_creator():
    #图形界面设计
    w = tk.Tk()
    w.title('排料之光')
    sw = w.winfo_screenwidth()
    # 得到屏幕宽度
    sh = w.winfo_screenheight()
    ww = 600
    wh = 400
    x = (sw - ww) / 2
    y = (sh - wh) / 2
    # 得到屏幕高度
    w.geometry("%dx%d+%d+%d" %(ww,wh,x,y))

    #创建显示框
    frame_r=ttk.Frame(w,height= 400,width = 300)
    frame_r.grid(row=0 ,column=1,columnspan=1,ipadx=10,padx=10)
    frame_l=ttk.Frame(w,height= 400,width = 300)
    frame_l.grid(row=0, column=0,columnspan=1,ipadx=10,padx=10)
    result1 = tk.StringVar()
    result2 = tk.StringVar()
    result3 = tk.StringVar()
    result4 = tk.StringVar()
    global result5,result6
    result5 = tk.StringVar()
    result6 = tk.StringVar()
    # 显示输出结果
    ttk.Label(frame_r, textvariable=result1).grid(row =0 )
    ttk.Label(frame_r, textvariable=result2).grid(row =0 )
    ttk.Label(frame_r, textvariable=result3).grid(row =0 )
    ttk.Label(frame_r, textvariable=result4).grid(row =0 )
    ttk.Label(frame_r, textvariable=result5).grid(row =0 )
    ttk.Label(frame_r, textvariable=result6).grid(row =0 )
    def sel_doc():
        path_ = askopenfilenames(filetypes=[("text file", "*.xlsx"), ("all", "*.*")], )
        path.set(path_)

        addr_arr = []
        for i  in path.get().strip('(').strip(')').split(','):
            addr_arr.append(i.split('/')[-1].strip("'"))
        xlsx_addr.set(addr_arr)
        # for i in addr_arr:
        #     excel_list.insert(i)
        # path.get()

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
        messagebox.showinfo(title='提示', message='文件已生成')

    ttk.Label(frame_l, text="目标路径:").grid(row =0,column=0,sticky = N+S )
    path = tk.StringVar()
    xlsx_addr = tk.StringVar()
    excel_list = tk.Listbox(frame_l,listvariable=xlsx_addr,width = 28,height = 4).grid(row =1,column=1,columnspan=2,rowspan =2,sticky=E+W,pady=1)
    ttk.Entry(frame_l, textvariable = path).grid(row =0,column=1,pady=1,padx=2,sticky = N+S)

    ttk.Button(frame_l,text='选择文件所在位置',command=sel_doc).grid(row =0,column=2,sticky = N+S)
    # ttk.Button(frame_l,text='选择文件夹所在位置',command=sel_doc).grid(row=0
    ttk.Button(frame_r,text='一键生成dxf数据',command=let_work).grid(row =0,sticky=E+W,pady=1)
    ttk.Button(frame_r,text='一键生成下料清单',command=let_docx_work).grid(row =1,sticky=E+W,pady=1)

    w.mainloop()

gui_creator()