# -*- coding: utf-8 -*-
# Author    ： ly
# Datetime  ： 2021/8/12 17:41
import os
import re
import threading
import time
from multiprocessing import Queue, Process

import PySimpleGUI as sg
import win32api
import win32com.client
import win32print


########################################################
#  _____      _       _                             _  #
# |  __ \    (_)     | |     /\                    | | #
# | |__) | __ _ _ __ | |_   /  \   _ __   __ _  ___| | #
# |  ___/ '__| | '_ \| __| / /\ \ | '_ \ / _` |/ _ \ | #
# | |   | |  | | | | | |_ / ____ \| | | | (_| |  __/ | #
# |_|   |_|  |_|_| |_|\__/_/    \_\_| |_|\__, |\___|_| #
#                                         __/ |        #
########################################################


class PrintAngel(object):
    _instance_lock = threading.Lock()

    def __init__(self):
        self.version = '打印精灵v1.0.3'
        self.marked_words = '欢迎使用打印精灵｡◕‿◕｡'
        self.author = 'Copyright @ New year'
        self.contact = '☎ 18571831390'
        self._instance = None
        self.pdf_flag = None
        self.logo = '''
                                                                                  ┏┓      ┏┓
                                                                                ┏┛┻━━━┛┻┓
                                                                                ┃      ☃      ┃
                                                                                ┃  ┳┛  ┗┳  ┃
                                                                                ┃      ┻      ┃
                                                                                ┗━┓      ┏━┛
                                                                                    ┃      ┗━━━┓
                                                                                    ┃  神兽保佑    ┣┓
                                                                                    ┃　永无BUG！   ┏┛
                                                                                    ┗┓┓┏━┳┓┏┛
                                                                                      ┃┫┫  ┃┫┫
                                                                                      ┗┻┛  ┗┻┛
'''

    def sort_by_file_name(self, dirnames):
        """ 文件名排序 已弃用 """
        dirs = dirnames
        try:
            dirs = sorted(dirnames, key=lambda x: (int(re.sub('\D', '', x)), x))
        except:
            self.print_info("sort_by_file_name fail")
        return dirs

    @staticmethod
    def file_name(file_dir):
        """当前路径下所有子目录 已弃用 """
        for root, dirs, files in os.walk(file_dir):
            break
        return files

    @staticmethod
    def select_path(path):
        """查询文件夹下文件数量 已弃用 """
        # read file count
        file_name = os.listdir(path)
        file_dir = [os.path.join(path, x) for x in file_name]
        count = "file count = %d\n" % (len(file_dir))
        return count

    @staticmethod
    def find_all_doc(target_dir, target_suffix="docx"):
        """ 找出所有docx文件 已弃用 """
        find_res = []
        target_suffix_dot = "." + target_suffix
        walk_generator = os.walk(target_dir)
        for root_path, dirs, files in walk_generator:
            if len(files) < 1:
                continue
            for file in files:
                file_name, suffix_name = os.path.splitext(file)
                if suffix_name == target_suffix_dot:
                    find_res.append(os.path.join(root_path, file))
        return find_res

    @staticmethod
    def listdir_nohidden(path):
        """ 忽略隐藏文件并排序 """
        file_lst = []
        for f in os.listdir(path):
            if '~' not in f or '$' not in f or '.' not in f:
                file_lst.append(path + '\\' + f)
        return sorted(file_lst)

    @staticmethod
    def print_info(msg):
        """ Prints """
        f = open("c:\\print_info.log", "a")

        f.write(msg + "\n")

        f.close()

    @staticmethod
    def print_time():
        """ 格式化时间 """
        now = int(time.time())
        time_array = time.localtime(now)
        other_style_time = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
        return other_style_time

    @staticmethod
    def printer_loading(filename):
        """ 调用默认打印接口 """
        current_printer = win32print.GetDefaultPrinter()
        win32api.ShellExecute(0, "print", filename, '/d: "%s"' % current_printer, ".", 0)

    def input_path_widget(self):
        """ 界面组件 """
        sg.theme('BlueMono')  # 设置当前主题
        layout = [
            [sg.Text('', size=(32, 1)),
             sg.Text(self.marked_words, size=(25, 1), text_color='#DA6B38', font=("华文彩云", 16))],
            [sg.Text('您选择打印的文件夹是:', font=("楷体", 14)),
             sg.Text('', key='text1', size=(60, 1), font=("宋体", 12), text_color='#000000'),
             sg.FolderBrowse('打开文件夹', key='folder', target='text1', button_color='#E6A23C')],
            [sg.Submit(tooltip='可点击查看', button_text='查看目录下文件', font=("Helvetica", 10), size=(15, 1),
                       button_color='#409EFF')],
            [sg.Text('')],
            [sg.Multiline(default_text=self.logo, key='files', font=("微软雅黑", 10), text_color='#000000', size=(300, 14),
                          autoscroll=True)],
            [sg.Text('', key='file_count')],
            [sg.Text('', size=(35, 1)),
             sg.Submit(button_text='确认打印', button_color=('white', '#67C23A'), size=(10, 1)),
             sg.Exit(button_text='关闭程序', button_color=('white', '#F56C6C'), key='退出')],
            [sg.Text('', size=(35, 1)), sg.Text(self.author, font=("楷体", 8)),
             sg.Text(self.contact, font=("楷体", 8))],
        ]

        window = sg.Window(
            title=self.version,
            layout=layout,
            ttk_theme='BlueMono',
            default_element_size=[100, ],
            icon=r'D:\projects\个人\HelloWorld\src\打印精灵v1.0\p1.ico',
            size=(850, 500)
        )
        while True:
            event, value = window.Read()
            if event == '查看目录下文件' and value is not None and value != '':
                foldername = value['folder'] or '.'
                if foldername == '.':
                    window['files'].update('当前文件夹下暂无可打印文件')
                else:
                    filenames = self.listdir_nohidden(foldername)
                    # it uses `key='files'` to access `Multiline` widget
                    window['files'].update("\n".join(filenames))
                    window['file_count'].update('共计：%s' % len(filenames))
            else:
                break
        window.Close()
        if value:
            return event, value, value['folder']
        return event, value, ''

    def producer(self, name, q):
        """ 生产者 """
        event, value, folder = self.input_path_widget()

        if folder and folder != '' and event == '确认打印':
            try:
                files = self.listdir_nohidden(folder)
                self.print_info("======================== Start Print ========================")
                time_str = self.print_time()
                self.print_info(time_str)
                for f in files:
                    ext = os.path.splitext(f)[1]
                    if ext.endswith('.x'):
                        # excel
                        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开 EXCEL
                        xlApp.Visible = 0  # 不在后台运行
                        xlApp.EnableEvents = False
                        xlApp.DisplayAlerts = False  # 显示弹窗
                        xlBook = xlApp.Workbooks.Open(f)
                        xlApp.ActiveWorkbook.Sheets(1).PageSetup.Zoom = False
                        xlApp.ActiveWorkbook.Sheets(1).PageSetup.FitToPagesWide = 1
                        xlApp.ActiveWorkbook.Sheets(1).PageSetup.FitToPagesTall = 1
                        xlBook.PrintOut(1, 99, )
                        xlApp.quit()
                    # 忽略pdf 文件
                    elif ext == ('.pdf'):
                        if self.pdf_flag is None:
                            from tkinter import messagebox
                            messagebox.showinfo(title='温馨提示', message="pdf文件已被过滤，如需批量打印PDF文件，请联系作者~\nWeChat：【 刘寅 楚精灵 】")
                            self.pdf_flag = True
                        continue
                    else:
                        # word txt docx
                        time.sleep(2)
                        q.put(f)
                        # print('%s 生产 %s' % (name, f))

            except Exception as e:
                self.print_info(e)

    def consumer(self, name, q):
        """ 消费者 """
        start_time = time.time()
        count = 0
        while True:
            res = q.get()
            if res is None:
                break
            print('%s 打印 %s' % (name, res))
            # 开始打印
            self.printer_loading(res)

            time.sleep(1)
            self.print_info("{} 打印完成".format(res))
            count += 1
        self.print_info("打印完成，共计：{} 份,耗时: {}".format(count, time.time() - start_time))


if __name__ == '__main__':
    # pyinstaller 默认打包不支持多进程程序，需要修改主程序 在main的第一行加入如下代码： multiprocessing.freeze_support()

    multiprocessing.freeze_support()
    angel = PrintAngel()

    q = Queue()
    p1 = Process(target=angel.producer, args=('生产者', q))
    c1 = Process(target=angel.consumer, args=('打印机', q))

    p1.start()
    c1.start()
    p1.join()
    q.put(None)
