# -*- coding: utf-8 -*-
# Author    ： ly
# Datetime  ： 2021/8/12 17:41
import collections
import os
import re
import threading
import time
from multiprocessing import Queue, Process

import PySimpleGUI as sg
import win32api
import win32com.client
import win32print
from PIL import Image, ImageWin
import win32ui


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
        self.version = '打印精灵v2.0.1'
        self.marked_words = '欢迎使用打印精灵｡◕‿◕｡'
        self.author = 'Copyright @ New year'
        self.contact = '☎ 18571831390'
        self.printer_name = win32print.GetDefaultPrinter()
        self.pdf_flag = None
        self.img_list = ['png', 'img', 'jpg', 'jpeg']
        self.file_list = ['txt', 'md', 'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx']
        self.animal = '''
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
        """
        # todo 文件名排序 已弃用
        """
        dirs = dirnames
        try:
            dirs = sorted(dirnames, key=lambda x: (int(re.sub('\D', '', x)), x))
        except:
            self.print_info("sort_by_file_name fail")
        return dirs

    @staticmethod
    def file_name(file_dir):
        """
        # todo 当前路径下所有子目录 已弃用
        """
        for root, dirs, files in os.walk(file_dir):
            break
        return files

    @staticmethod
    def select_path(path):
        """
        # todo 查询文件夹下文件数量 已弃用
        """
        # read file count
        file_name = os.listdir(path)
        file_dir = [os.path.join(path, x) for x in file_name]
        count = "file count = %d\n" % (len(file_dir))
        return count

    @staticmethod
    def find_all_doc(target_dir, target_suffix="docx"):
        """
        # todo 找出所有docx文件 已弃用
        """
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
        """
        # todo 忽略隐藏文件并排序
        """
        file_lst = []
        for f in os.listdir(path):
            if '~' not in f or '$' not in f or '.' not in f:
                file_lst.append(path + '\\' + f)
        return sorted(file_lst)

    def get_all_file(self, path):
        """
        # todo 递归查找文件，并放入到对应的列表排好序 暂时不用
        """
        dl = collections.deque()
        dl.append(path)
        png_path_list = []
        file_path_list = []
        while len(dl) != 0:
            pop = dl.popleft()
            listfile = os.listdir(pop)
            for i in listfile:
                newpath = os.path.join(pop, i)
                if os.path.isdir(newpath):
                    print("目录：", newpath)
                    dl.append(newpath)
                else:
                    if newpath.split(".")[-1] in self.img_list:
                        png_path_list.append(newpath)
                    elif newpath.split(".")[-1] in self.file_list:
                        file_path_list.append(newpath)
        return sorted(png_path_list), sorted(file_path_list)

    @staticmethod
    def print_info(msg):
        """
        # todo Prints
        """
        f = open("c:\\print_info.log", "a")

        f.write(msg + "\n")

        f.close()

    @staticmethod
    def print_time():
        """
        # todo 格式化时间
        """
        now = int(time.time())
        time_array = time.localtime(now)
        other_style_time = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
        return other_style_time

    def printer_default_api(self, filename):
        """
        # todo 调用默认打印接口
        """
        win32api.ShellExecute(0, "print", filename, '/d: "%s"' % self.printer_name, ".", 0)

    def input_path_widget(self):
        """
        # todo 界面组件
        """
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
            [sg.Multiline(default_text=self.animal, key='files', font=("微软雅黑", 10), text_color='#000000',
                          size=(300, 14),
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
            icon=r'D:\projects\my-project\PrintAngel\printAngel_v1.0.3\p1.ico',
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

    def printer_png_loading(self, file_path):
        """
        # todo 打印图片
        """
        try:
            # HORZRES / VERTRES = printable area 可打印的区域
            HORZRES = 8
            VERTRES = 10
            # PHYSICALWIDTH/HEIGHT = total area 总面积
            #
            PHYSICALWIDTH = 110
            PHYSICALHEIGHT = 111
            printer_name = win32print.GetDefaultPrinter()
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)
            printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
            printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
            # 打开图像，如果宽度大于高，计算出要乘的倍数
            # 通过每个像素使它尽可能大
            # 页面不失真。
            bmp = Image.open(file_path)
            if bmp.size[0] > bmp.size[1]:
                bmp = bmp.rotate(90)

            ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
            scale = min(ratios)

            # 开始打印作业，并将位图绘制到
            # 按比例缩放打印机设备。
            hDC.StartDoc(file_path)
            hDC.StartPage()

            dib = ImageWin.Dib(bmp)
            scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
            x1 = int((printer_size[0] - scaled_width) / 2)
            y1 = int((printer_size[1] - scaled_height) / 2)
            x2 = x1 + scaled_width
            y2 = y1 + scaled_height
            dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()

            self.print_info("{} 打印完成".format(file_path))
        except Exception as e:
            self.print_info("{} 打印异常, 异常信息{}".format(file_path, e))

    def producer(self, name, q):
        """ 生产者 """
        event, value, folder = self.input_path_widget()

        if folder and folder != '' and event == '确认打印':
            try:
                # 文件排序
                files = self.listdir_nohidden(folder)
                self.print_info("======================== Start Print ========================")
                time_str = self.print_time()
                self.print_info(time_str)
                for f in files:
                    ext = os.path.splitext(f)[1]
                    # 忽略pdf 文件
                    if ext == '.pdf':
                        if self.pdf_flag is None:
                            from tkinter import messagebox
                            messagebox.showinfo(title='温馨提示', message="pdf文件已被过滤，如需批量打印PDF文件，请联系作者~\nWeChat：【 刘寅 楚精灵 】")
                            self.pdf_flag = True
                        continue
                    time.sleep(1)
                    q.put(f)
                # 生产完成
                q.put(None)
            except Exception as e:
                self.print_info("生产信息有误：{}".format(e))
        else:
            q.put(None)

    def consumer(self, name, q):
        """ 消费者 """
        start_time = time.time()
        count = 0
        while True:
            result = q.get()
            if result is None:
                break
            ext = result.split(".")[-1]
            if ext in self.img_list:
                time.sleep(2)
                self.printer_png_loading(result)
                print('%s 正在打印 %s' % (name, result))
            if ext in self.file_list:
                time.sleep(2)
                self.printer_default_api(result)
                print('%s 正在打印 %s' % (name, result))

            self.print_info("{} 打印完成".format(result))
            count += 1
        self.print_info("打印完成，共计：{} 份,耗时: {}".format(count, time.time() - start_time))


if __name__ == '__main__':
    # pyinstaller 默认打包不支持多进程程序，需要修改主程序 在main的第一行加入如下代码： multiprocessing.freeze_support()

    # multiprocessing.freeze_support()
    angel = PrintAngel()

    q = Queue()
    p1 = threading.Thread(target=angel.producer, args=('生产者', q))
    c1 = threading.Thread(target=angel.consumer, args=(angel.printer_name, q))

    p1.start()
    c1.start()
