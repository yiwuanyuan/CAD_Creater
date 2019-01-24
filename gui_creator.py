#!/usr/bin/env python
# coding = gbk
# File  : gui_creator.py
# Author: Wangyuan
# Date  : 2019-1-3
from OutputDocx import *
from GetExcelInfo import *
from DrawCreator import *
from tkinter import *
import tkinter as tk
from tkinter import messagebox, N, S, W, E, ttk
# from tkinter import *
from tkinter.filedialog import askopenfilenames
from re import sub


def gui_creator():
    # 图形界面设计
    w = tk.Tk()
    w.title('排料之光')
    sw = w.winfo_screenwidth()
    # 得到屏幕宽度
    sh = w.winfo_screenheight()
    ww = 800
    wh = 400
    x = (sw - ww) / 2
    y = (sh - wh) / 2
    # 得到屏幕高度
    w.geometry("%dx%d+%d+%d" % (ww, wh, x, y))

    # 创建显示框
    frame_r = tk.Frame(w, width = 300, height = 400)
    frame_r.grid(row=0,rowspan=2, column=1, ipadx=10, padx=5)

    frame_l1 = tk.LabelFrame(w, height=200, width=290,text='文件选择')
    frame_l1.grid(row=0, column=0, ipadx=10,padx=5)
    frame_l2 = tk.LabelFrame(w, height=200, width=290,text='执行')
    frame_l2.grid(row=1, column=0, ipadx=10,padx=5)
    result1 = tk.StringVar()

    document_addr = tk.StringVar()
    result3 = tk.StringVar()
    result4 = tk.StringVar()
    global result5, result6
    result5 = tk.StringVar()
    result6 = tk.StringVar()
    # 显示输出结果
    ttk.Label(frame_r, textvariable=result1).grid(row=0)

    ttk.Label(frame_r, textvariable=result4).grid(row=0)
    ttk.Label(frame_r, textvariable=result5).grid(row=0)
    ttk.Label(frame_r, textvariable=result6).grid(row=0)

    def out_omitexcel(arr, addr):
        pass

    def sel_doc():
        path_ = askopenfilenames(filetypes=[("text file", "*.xlsx"), ("all", "*.*")], )
        path.set(path_)
        document_addr.set('/'.join(path.get().strip('(').strip(')')
                                   .split(',')[0].strip('\'').split('/')[:-1]))

        addr_arr = []
        for i in path.get().strip('(').strip(')').split(','):
            if not i == '':
                addr_arr.append(i.split('/')[-1].strip("'"))
        xlsx_addr.set(addr_arr)
        # for i in addr_arr:
        #     excel_list.insert(i)
        # path.get()

    def let_work():
        # 判断是否获取到路径
        if path.get():
            addr = path.get().strip('(').strip(')').split(',')
            omit_gather = []

            # 获取到的地址
            for ad in addr:
                if ad != '':
                    g = GetExcelInfo(ad.strip().strip('\''))
                    info=g.output
                    error=g.error_gather

                    doc_addr = '/'.join(ad.strip().strip('\'').split('/')[:-1])
                    print(g.output)
                    if g.stop:
                        messagebox.showerror(title='表格错误', message='该表格式有错，已略过')
                    else:
                        # print(g.error_gather)
                        for i in g.output:
                            try:
                                DrawCreator(i['name'], i['type'], i['parameter'], i['material'], i['thickness'],
                                            address=doc_addr)
                            except BaseException:

                                error_text.set('绘图错误请检查！！！')
                                print(i['name'], i['type'], i['parameter'], i['material'], i['thickness'])
                                print('绘图错误请检查！！！')
                        addinfo(info)
                        adderror(error)

                        if g.omit_value:
                            omit = {}
                            omit['name'] = ad.split('/')[-1].strip('\'').strip('.xlsx')
                            omit['value'] = g.omit_value
                            omit_gather.append(omit)

            if omit_gather != []:
                if messagebox.askyesno(title='提示', message='存在被略去数据，是否导出Excel'):
                    df = []
                    out_addr = '/'.join(addr[0].strip().strip('\'').split('/')[:-1])
                    i = 0
                    writer = ExcelWriter(out_addr + '/被省略的excel.xlsx')
                    for omit_unit in omit_gather:
                        df.append(DataFrame(omit_unit['value']))
                        print(omit_unit['name'])
                        df[i].to_excel(writer, sheet_name=omit_unit['name'])
                        i += 1
                    writer.save()
        else:
            messagebox.showinfo(title='提示', message='至少选择一个Excel文件')

    def addinfo(info):
        for i in info:
            infoBox.insert(END,i['name']+str(i['parameter'])+'\n')

    def adderror(error):
        for i in error:
            errorBox.insert(END, i['content']+'\n')
    def clear():
        infoBox.delete('1.0',END)
        errorBox.delete('1.0',END)
        xlsx_addr.set('')
        path.set('')
    def let_docx_work():
        show_info = 0
        if path.get():
            addr = path.get().strip('(').strip(')').split(',')
            for ad in addr:
                if ad != '':
                    xls_addr = ad.strip().strip('\'')
                    g = GetExcelInfo(xls_addr)

                    if g.excel_type == 'HD':
                        OutputDocx(g.output, address=xls_addr)
                        show_info += 1
                    else:
                        messagebox.showinfo(title='提示', message='%s 该文件不需要生成排料清单' % g.addr.split('/')[-1])

        else:
            messagebox.showinfo(title='提示', message='至少选择一个Excel文件')
        if not show_info == 0:
            messagebox.showinfo(title='提示', message='文件已生成')


    ttk.Label(frame_l1, text="目标路径:").grid(row=0, column=0, sticky=N + S)
    ttk.Label(frame_l1, text="输入文件:").grid(row=1, column=0, sticky=N + S)
    path = tk.StringVar()
    xlsx_addr = tk.StringVar()
    excel_list = tk.Listbox(frame_l1, listvariable=xlsx_addr, width=30, height=4).grid(row=1, column=1, columnspan=2,rowspan=2, pady=10)

    # 错误预览框
    tk.Label(frame_r, text="提示信息").grid()
    infoBox = tk.Text(frame_r, width=60, height=10)
    infoBox.grid(padx=10, pady=10)
    tk.Label(frame_r, text="错误报障").grid()
    errorBox = tk.Text(frame_r, width=60, height=10)
    errorBox.grid(padx=10, pady=10)
    tk.Button(frame_r, text='Clear', command=clear).grid()



    ttk.Label(frame_l1, textvariable=document_addr, width=30).grid(row=0, column=1, pady=1, padx=2, sticky=N + S)
    ttk.Button(frame_l2, text='选择文件', command=sel_doc).grid(row=0, column=1, sticky=E + W, padx=5, pady=5)
    # ttk.Button(frame_l,text='选择文件夹所在位置',command=sel_doc).grid(row=0
    ttk.Button(frame_l2, text='生成dxf文件', command=let_work).grid(row=0, column=2, sticky=E + W, pady=10, padx=5)
    ttk.Button(frame_l2, text='生成下料清单', command=let_docx_work).grid(row=0, column=3, sticky=E + W, pady=10, padx=5)
    # ttk.Button(frame_r, text='生成被略去文件清单', command=out_omitexcel).grid(row=2, sticky=E + W, pady=1)

    w.mainloop()


gui_creator()
