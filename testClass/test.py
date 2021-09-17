# -*- coding: utf-8 -*-
# Author    ： ly
# Datetime  ： 2021/9/15 11:44
# 查看当前电脑上安装了哪些打印机：
import os

import win32print
import win32api
import collections

import win32ui
from PIL import Image, ImageWin

printers = win32print.EnumPrinters(3)
# for _ in printers:
# print(_)

# 查看当前电脑上安装了哪些打印机

printer = win32print.GetDefaultPrinter()


# print(printer)


# 递归查找文件，并放入到对应的列表中
def get_file_all(path):
    dl = collections.deque()
    dl.append(path)
    png_path_list = []
    pdf_path_list = []
    while len(dl) != 0:
        pop = dl.popleft()
        listfile = os.listdir(pop)
        for i in listfile:
            newpath = os.path.join(pop, i)
            if os.path.isdir(newpath):
                print("目录：", newpath)
                dl.append(newpath)
            else:
                if newpath.split(".")[-1] in ['png', 'img', 'jpg', 'jpeg']:
                    png_path_list.append(newpath)
                elif newpath.split(".")[-1] in ['pdf', 'txt', 'xlsx', 'md', 'doc', 'docx']:
                    pdf_path_list.append(newpath)
    return png_path_list, pdf_path_list


# 打印图片
def printer_png_loading(file_path):
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
        # #打开图像，如果宽度大于高，计算出要乘的倍数
        # ＃通过每个像素使它尽可能大
        # ＃页面不失真。
        bmp = Image.open(file_path)
        if bmp.size[0] > bmp.size[1]:
            bmp = bmp.rotate(90)

        ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
        scale = min(ratios)

        # ＃开始打印作业，并将位图绘制到
        # ＃按比例缩放打印机设备。
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

        print('-ok!!--The picture prints normally---' + file_path)
    except Exception as ex:
        print('-error!!--Picture prints abnormally---' + file_path)
        print(repr(ex))


# 打印其他文件
def printer_other_loading(self, file_path):
    try:
        open(file_path, "r")
        win32api.ShellExecute(
            0,
            "print",
            file_path,
            '/d:"%s"' % self.printer_name,
            ".",
            0
        )
        print('--ok!!-the pdf prints normally---' + file_path)
    except Exception as ex:
        print(repr(ex))
        print('-error!!--the pdf prints abnormally---', + file_path)


if __name__ == '__main__':
    path = r'D:/打印测试/图片\Big Sur_1.jpg'
    # png_path_list, pdf_path_list = get_file_all(path)
    #
    # print(png_path_list, pdf_path_list)
    # for img in png_path_list:
    #
    printer_png_loading(path)
