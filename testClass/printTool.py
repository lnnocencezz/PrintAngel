# # -*- coding: utf-8 -*-
# # Author    ： ly
# # Datetime  ： 2021/8/25 9:00


from tkinter import *
from tkinter.filedialog import askdirectory
import win32com.client
import win32api
import win32print
import time
import os


def selectPath():
    path_ = askdirectory()
    path.set(path_)

    # read file count
    file_name = os.listdir(path.get())
    file_dir = [os.path.join(path.get(), x) for x in file_name]
    count.set("file count = %d\n" % (len(file_dir)))


def startPritn():
    print("start print %s" % path.get())
    i = 0
    file_name = os.listdir(path.get())
    file_dirs = [os.path.join(path.get(), x) for x in file_name]

    while i < len(file_dirs):
        ext = os.path.splitext(file_dirs[i])[1]
        if ext.startswith('.x'):
            # excel
            xlApp = win32com.client.Dispatch('Excel.Application')
            xlApp.Visible = 0
            xlApp.EnableEvents = False
            xlApp.DisplayAlerts = False
            xlBook = xlApp.Workbooks.Open(file_dirs[i])
            xlApp.ActiveWorkbook.Sheets(1).PageSetup.Zoom = False
            xlApp.ActiveWorkbook.Sheets(1).PageSetup.FitToPagesWide = 1
            xlApp.ActiveWorkbook.Sheets(1).PageSetup.FitToPagesTall = 1
            xlBook.PrintOut(1, 99, )
            xlApp.quit()
        else:
            # word pdf txt
            win32api.ShellExecute(
                0,
                "print",
                file_dirs[i],
                '/d:"%s"' % win32print.GetDefaultPrinter(),
                ".",
                0
            )

        print(file_dirs[i])
        time.sleep(1)
        i = i + 1


root = Tk()

root.title('Print Tool')

width = 400
height = 280
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=True, height=True)

count = StringVar()

path = StringVar()
Label(root, text="").pack()
Label(root, text="Support batch printing of PDF, excel and word ", fg='red').pack(pady=10)
Label(root, text="path:").pack()
Entry(root, textvariable=path).pack()
Button(root, text="choose...", command=selectPath).pack(pady=10)
countLable = Label(root, textvariable=count, fg='blue').pack()
Button(root, text="start", command=startPritn).pack()

Label(root, text="author:zmy").pack()
Label(root, text="wechat:y551592").pack()

root.mainloop()

# ================================================================================================================
# import PySimpleGUI as sg
#
# sg.ChangeLookAndFeel('GreenTan')
#
# form = sg.FlexForm('Everything bagel', default_element_size=(40, 1))
#
# column1 = [[sg.Text('Column 1', background_color='#d3dfda',     justification='center', size=(10,1))],
#            [sg.Spin(values=('Spin Box 1', '2', '3'), initial_value='Spin Box 1')],
#            [sg.Spin(values=('Spin Box 1', '2', '3'), initial_value='Spin Box 2')],
#            [sg.Spin(values=('Spin Box 1', '2', '3'), initial_value='Spin Box 3')]]
# layout = [
#     [sg.Text('All graphic widgets in one form!', size=(30, 1), font=("Helvetica", 25))],
#     [sg.Text('Here is some text.... and a place to enter text')],
#     [sg.InputText('This is my text')],
#     [sg.Checkbox('My first checkbox!'), sg.Checkbox('My second checkbox!',     default=True)],
#     [sg.Radio('My first Radio!     ', "RADIO1", default=True), sg.Radio('My second Radio!', "RADIO1")],
#     [sg.Multiline(default_text='This is the default Text should you decide not to type anything', size=(35, 3)),
#  sg.Multiline(default_text='A second multi-line', size=(35, 3))],
#     [sg.InputCombo(('Combobox 1', 'Combobox 2'), size=(20, 3)),
#  sg.Slider(range=(1, 100), orientation='h', size=(34, 20), default_value=85)],
#     [sg.Listbox(values=('Listbox 1', 'Listbox 2', 'Listbox 3'), size=(30, 3)),
#  sg.Slider(range=(1, 100), orientation='v', size=(5, 20), default_value=25),
#  sg.Slider(range=(1, 100), orientation='v', size=(5, 20), default_value=75),
#  sg.Slider(range=(1, 100), orientation='v', size=(5, 20), default_value=10),
#  sg.Column(column1, background_color='#d3dfda')],
#     [sg.Text('_'  * 80)],
#     [sg.Text('Choose A Folder', size=(35, 1))],
#     [sg.Text('Your Folder', size=(15, 1), auto_size_text=False, justification='right'),
#  sg.InputText('Default Folder'), sg.FolderBrowse()],
#     [sg.Submit(), sg.Cancel()]
#  ]
#
# button, values = form.Layout(layout).Read()
# sg.Popup(button, values)

# import subprocess
# import sys
# import PySimpleGUI as sg
#
# """
#     Demo Program - Realtime output of a shell command in the window
#         Shows how you can run a long-running subprocess and have the output
#         be displayed in realtime in the window.
# """
#
# def main():
#     layout = [  [sg.Text('Enter the command you wish to run')],
#                 [sg.Input(key='_IN_')],
#                 [sg.Output(size=(60,15))],
#                 [sg.Button('Run'), sg.Button('Exit')] ]
#
#     window = sg.Window('Realtime Shell Command Output', layout)
#
#     while True:             # Event Loop
#         event, values = window.Read()
#         # print(event, values)
#         if event in (None, 'Exit'):
#             break
#         elif event == 'Run':
#             runCommand(cmd=values['_IN_'], window=window)
#     window.Close()
#
#
# def runCommand(cmd, timeout=None, window=None):
#     """ run shell command
#     @param cmd: command to execute
#     @param timeout: timeout for command execution
#     @param window: the PySimpleGUI window that the output is going to (needed to do refresh on)
#     @return: (return code from command, command output)
#     """
#     p = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
#     output = ''
#     for line in p.stdout:
#         line = line.decode(errors='replace' if (sys.version_info) < (3, 5) else 'backslashreplace').rstrip()
#         output += line
#         print(line)
#         window.Refresh() if window else None        # yes, a 1-line if, so shoot me
#
#     retval = p.wait(timeout)
#     return (retval, output)
#
#
# main()