import requests
import json
import tkinter as tk
import threading
from tkinter import ttk,Scrollbar, Frame
import tkinter.filedialog
import re
import time
import xlwt
import os
from urllib.parse import urlencode
from tkinter import *

headers = {
    'Host': 'is-hl.snssdk.com',
    'User-Agent': 'Dalvik/2.1.0 (Linux; U; Android 6.0.1; SM-A8000 Build/MMB29M) NewsArticle/7.0.3 cronet/TTNetVersion:a729d5c3',
}
#文章链接汇总
urllist= []
#文章评论ID汇总
makelist = []
#评论ID共享链接汇总
gurllist = list(set())
mindex = 1
gindex = 1


data_dict = {}
key_list = []
value_list = []


class xlsmanager():
    def __init__(self, lst):
        self.outwb = xlwt.Workbook()
        self.outws = self.outwb.add_sheet("sheel")
        for v in range(len(lst)):
            self.outws.write(0, v, lst[v])
        self.index = 1

    def add_data(self, lst, name):
        for v in range(len(lst)):
            self.outws.write(self.index, v, str(lst[v]).replace('\n', '').replace('"', '').replace("'", ""))
        self.outwb.save(name)
        self.index += 1


def get_comment(item_id, offset):
    ts = int(time.time())
    param_data = {
        'offset': offset,
        'group_id': item_id,
        'aggr_type': 1,
        'count': 50,
        'item_id': item_id,
        'ts': ts
    }
    comment_url = 'http://is-hl.snssdk.com/article/v4/tab_comments/?' + urlencode(param_data)
    response = requests.get(comment_url, headers=headers)
    data = response.json()
    if data['data'] == []:
        print("没有评论数据")
        return
    parse_comment(data, item_id, offset)


def parse_comment(data, item_id, offset):
    global mindex
    comments = data['data']
    has_more = data['has_more']
    for comment in comments:
        user_name = comment['comment']['user_name']
        datetime = comment["comment"]["create_time"]
        datetime = time.localtime(int(datetime))
        datetime = time.strftime("%Y-%m-%d %H:%M:%S", datetime)
        make = [mindex, user_name, datetime]
        add_makedata(make)
        mindex += 1
    if has_more:
        offset += 50
        get_comment(item_id, offset)


def check_comment(data, item_id, offset, url, makeid):
    # global gindex
    for key, value in data_dict.items():
        key_list.append(key)
        value_list.append(value)
    if url in value_list:
        get_value_index = value_list.index(url)
        gindex = int(key_list[get_value_index])
    else:
        print("你要查询的值%s不存在" % url)

        # gindex += 1
    comments = data['data']
    has_more = data['has_more']
    for comment in comments:
        user_name = comment['comment']['user_name']
        if (user_name == makeid):
            gurl = [gindex, url]
            gurllist.append(gurl)
            gurl_data.insert("", "end", values=(gurl))
            gindex += 1
        if has_more:
            offset += 50
            get_comment(item_id, offset)


def get_makeID(makeid):
    global gindex
    gindex = 1
    gurllist.clear()
    clear_tree(gurl_data)
    for v in urllist:
        check_url(v, makeid)


def check_url(url, makeid):
    id = get_urlid(url)
    ts = int(time.time())
    param_data = {
        'offset': 0,
        'group_id': id,
        'aggr_type': 1,
        'count': 50,
        'item_id': id,
        'ts': ts
    }
    comment_url = 'http://is-hl.snssdk.com/article/v4/tab_comments/?' + urlencode(param_data)
    response = requests.get(comment_url, headers=headers)
    data = response.json()
    if data['data'] == []:
        print("没有评论数据")
        return
    check_comment(data, id, 0, url, makeid)


def get_re(id, item_id, offset):
    lst = ["ID", "时间"]
    xls = xlsmanager(lst)
    ts = int(time.time())
    param_data = {
        'offset': offset,
        'group_id': item_id,
        'aggr_type': 1,
        'count': 50,
        'item_id': item_id,
        'ts': ts
    }
    comment_url = 'http://is-hl.snssdk.com/article/v4/tab_comments/?' + urlencode(param_data)
    response = requests.get(comment_url, headers=headers)
    data = response.json()
    if data['data'] == []:
        print("没有评论数据")
        return
    parse_re(id, data, item_id,  offset, xls)


def parse_re(id, data, item_id,  offset, xls):
    comments = data['data']
    has_more = data['has_more']
    for comment in comments:
        user_name = comment['comment']['user_name']
        datetime = comment["comment"]["create_time"]
        datetime = time.localtime(int(datetime))
        datetime = time.strftime("%Y-%m-%d %H:%M:%S", datetime)
        make = [user_name, datetime]
        path = id + ".xls"
        xls.add_data(make, path)
        if has_more:
            offset += 50
            get_re(id, item_id, offset)


def start_collection(url):
    global urllist
    # count = len(urllist)+1
    values = url.split(',')
    id = values[0]
    url = values[1]
    urllist.append(url)
    # url_data.insert("", "end", values=(id, url))
    item_id = get_urlid(url)
    get_re(id, item_id, offset=0)
    data_dict[id] = url
    for key, value in data_dict.items():
        url_data.insert('', key, values=(key, value))


def get_urlid(url):
    return re.findall('(\d+)', url, re.I | re.M)[0]


def get_url(url):
    global mindex
    mindex = 1
    id = get_urlid(url)
    clear_tree(make_data)
    makelist.clear()
    get_comment(id, 0)


def datetime_str(timer):
    timeArray = time.localtime(timer)
    otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)
    return otherStyleTime


def clear_tree(tree):
    x = tree.get_children()
    for item in x:
        tree.delete(item)


def add_makedata(lst):
    makelist.append(lst)
    make_data.insert("", "end", values=(lst))


def import_urls():
    global urllist
    selectFileName = tkinter.filedialog.askopenfilename(title="选择文件", filetypes=[('Text file', '*.txt')])
    if(selectFileName != ""):
        with open(selectFileName, "r") as f:
            for li in f.readlines():
                values = li.split(',')
                id = values[0]
                url = values[1]
                urllist.append(url)
                item_id = get_urlid(url)
                get_re(id, item_id, 0)
                data_dict[id] = url
        clear_tree(url_data)
        for key, value in data_dict.items():
            url_data.insert('', key, values=(key, value))


def thread_it(func, *args):
    t = threading.Thread(target=func, args=args)
    t.setDaemon(True)
    t.start()
    # t.join()


def urltreeviewClick(event):
    if len(url_data.selection()) > 0:
        item = url_data.selection()[0]
        url = url_data.item(item, "values")[1]
        get_url(url)


def maketreeviewClick(event):
    if len(make_data.selection()) > 0:
        item = make_data.selection()[0]
        make = make_data.item(item, "values")[1]
        get_makeID(make)


def gurltreeviewClick(event):
    if len(url_data.selection()) > 0:
        item = url_data.selection()[0]
        url = url_data.item(item, "values")[1]
        print(url)


def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    print(tv.get_children(''))
    l.sort(reverse=reverse)  # 排序方式
    # rearrange items in sorted positions
    for index, (val, k) in enumerate(l):  # 根据排序后索引移动
        tv.move(k, '', index)
        print(k)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))


def clear_alldata():
    for url in urllist:
        # urlid = get_urlid(url)
        for key, value in data_dict.items():
            key_list.append(key)
            value_list.append(value)
        if url in value_list:
            get_value_index = value_list.index(url)
            gindex = key_list[get_value_index]
            filename = gindex+'.xls'
            os.remove(filename)
    clear_tree(make_data)
    clear_tree(url_data)
    clear_tree(gurl_data)
    makelist.clear()
    urllist.clear()
    gurllist.clear()


def export_data():
    if len(url_data.selection()) > 0:
        lst = ["编号", "ID", "时间"]
        xlsx = xlsmanager(lst)
        item = url_data.selection()[0]
        # url = url_data.item(item, "values")[1]
        # path = get_urlid(url_data.item(item, "values")[1]) + ".xls"
        # print('123456', path)
        id = get_urlid(url_data.item(item, "values")[1])
        url = url_data.item(item, "values")[1]
        for key, value in data_dict.items():
            key_list.append(key)
            value_list.append(value)
        if url in value_list:
            get_value_index = value_list.index(url)
            gindex = key_list[get_value_index]
            path = gindex + '__' + id + ".xls"
        for v in makelist:
            xlsx.add_data(v, path)


def export_data1():
    if len(make_data.selection()) > 0:
        item = make_data.selection()[0]
        make = make_data.item(item, "values")[1]
        lst = ["编号", "共享链接"]
        xlsx = xlsmanager(lst)
        item = url_data.selection()[0]
        path = make + ".xls"
        for v in gurllist:
            xlsx.add_data(v, path)

def delete_info1():

    item = url_data.selection()[0]
    url = url_data.item(item, "values")[1]
    print('haha', url)
    # url_data.delete(item)
    for key, value in data_dict.items():
        key_list.append(key)
        value_list.append(value)
    if url in value_list:
        get_value_index = value_list.index(url)
        gindex = key_list[get_value_index]
        filename = gindex + '.xls'
        print(filename)
        os.remove(filename)
    Button(window,

               command=url_data.delete(item)
               )


def delete_info2():
    item = make_data.selection()[0]
    # url_data.delete(item)
    Button(window,

           command=make_data.delete(item)
           )
    id = get_urlid(url_data.item(item, "values")[1])
    url = url_data.item(item, "values")[1]
    for key, value in data_dict.items():
        key_list.append(key)
        value_list.append(value)
    if url in value_list:
        get_value_index = value_list.index(url)
        gindex = key_list[get_value_index]
        path = gindex + '__' + id + ".xls"
        os.remove(path)


def delete_info3():
    global gurl_data
    item = gurl_data.selection()[0]
    Button(window,
           command=gurl_data.delete(item)
    )


if __name__ == '__main__':
    window = tk.Tk()
    window.title('今日头条分析工具')
    window.geometry('700x800+700+150')
    window.resizable(False, False)
    tk.Label(window, text="文章链接:").place(x=20, y=20)
    title = tk.StringVar()
    title.set("")
    entry_usr_name = tk.Entry(window, textvariable=title, width=50)
    entry_usr_name.place(x=80, y=20)
    btn_collect = tk.Button(window, text='导入', command=lambda: thread_it(start_collection, title.get()), width=6,
                            height=1)
    btn_collect.place(x=440, y=12)
    btn_import = tk.Button(window, text='批量导入', command=lambda: thread_it(import_urls), width=8, height=1)
    btn_import.place(x=500, y=12)
    btn_alldel = tk.Button(window, text='清空数据', command=clear_alldata, width=8, height=1)
    btn_alldel.place(x=580, y=12)
    btn_export1 = tk.Button(window, text='删除数据', command=delete_info1, width=8, height=1)
    btn_export1.place(x=580, y=120)

    btn_export = tk.Button(window, text='导出数据', command=export_data, width=8, height=1)
    btn_export.place(x=580, y=320)

    btn_export1 = tk.Button(window, text='删除数据', command=delete_info2, width=8, height=1)
    btn_export1.place(x=580, y=400)

    btn_export1 = tk.Button(window, text='导出数据', command=export_data1, width=8, height=1)
    btn_export1.place(x=580, y=520)

    btn_export1 = tk.Button(window, text='删除数据', command=delete_info3, width=8, height=1)
    btn_export1.place(x=580, y=600)

    urlframe = Frame(window)
    urlframe.place(x=70, y=50, width=480, height=200)
    scrollBar = tkinter.Scrollbar(urlframe)
    scrollBar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    url_data = ttk.Treeview(urlframe, show="headings", yscrollcommand=scrollBar.set)
    url_data['columns'] = ['index', 'url']
    url_data.column('index', width=50, anchor='center')
    url_data.column('url', width=400, anchor='center')
    url_data.heading('index', text='编号')
    url_data.heading('url', text='链接')
    url_data.bind("<ButtonRelease-1>", urltreeviewClick)
    url_data.pack(side=tkinter.LEFT, fill=tkinter.Y)
    scrollBar.config(command=url_data.yview)

    tk.Label(window, text="文章评论ID汇总:").place(x=70, y=250)
    dataframe = Frame(window)
    dataframe.place(x=70, y=270, width=480, height=200)
    scrollBar1 = tkinter.Scrollbar(dataframe)
    scrollBar1.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    make_data = ttk.Treeview(dataframe, show="headings", yscrollcommand=scrollBar1.set)
    make_data['columns'] = ['index', 'name', "datetime"]
    make_data.column('index', width=50, anchor='center')
    make_data.column('name', width=200, anchor='center')
    make_data.column('datetime', width=200, anchor='center')
    make_data.heading('name', text='ID')
    make_data.heading('index', text='编号')
    make_data.heading('datetime', text='时间', command=lambda: treeview_sort_column(make_data, 'datetime', False))
    make_data.bind('<ButtonRelease-1>', maketreeviewClick)
    make_data.pack(side=tkinter.LEFT, fill=tkinter.Y)
    scrollBar1.config(command=make_data.yview)

    tk.Label(window, text="评论ID文章链接汇总:").place(x=70, y=470)
    gurlframe = Frame(window)
    gurlframe.place(x=70, y=490, width=480, height=200)
    gscrollBar = tkinter.Scrollbar(gurlframe)
    gscrollBar.pack(side=tkinter.RIGHT, fill=tkinter.Y)
    gurl_data = ttk.Treeview(gurlframe, show="headings", yscrollcommand=gscrollBar.set)
    gurl_data['columns'] = ['index', 'url']
    gurl_data.column('index', width=50, anchor='center')
    gurl_data.column('url', width=400, anchor='center')
    gurl_data.heading('index', text='编号')
    gurl_data.heading('url', text='链接')
    gurl_data.bind("<ButtonRelease-1>", gurltreeviewClick)
    gurl_data.pack(side=tkinter.LEFT, fill=tkinter.Y)
    gscrollBar.config(command=gurl_data.yview)
    window.mainloop()
'return ''.join(re.compile(r"\w://www\.365yg\.com/\w(\d+)", re.S).findall(url))'