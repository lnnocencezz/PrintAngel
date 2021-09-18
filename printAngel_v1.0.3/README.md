这里整理下win32print的API介绍，官网地址http://timgolden.me.uk/pywin32-docs/win32print.html

OpenPrinter　　打开指定的打印机，并获取打印机的句柄

GetPrinter　　取得与指定打印机有关的信息

SetPrinter　　对一台打印机的状态进行控制

ClosePrinter　　关闭一个打开的打印机对象

AddPrinterConnection　　连接指定的打印机

DeletePrinterConnection　　删除与指定打印机的连接

EnumPrinters　　枚举系统中安装的打印机

GetDefaultPrinter　　取得默认打印机名称 <type 'str'>

GetDefaultPrinterW　　取得默认打印机名称 <type 'unicode'>

SetDefaultPrinter　　对一台打印机名称 <type 'str'> 设置成默认打印机

SetDefaultPrinterW　　对一台打印机名称 <type 'unicode'> 设置成默认打印机

StartDocPrinter　　在后台打印的级别启动一个新文档

EndDocPrinter　　在后台打印程序的级别指定一个文档的结束

AbortPrinter　　删除与一台打印机关联在一起的缓冲文件

StartPagePrinter　　在打印作业中指定一个新页的开始

EndPagePrinter　　指定一个页在打印作业中的结尾

StartDoc　　开始一个打印作业

EndDoc　　结束一个成功的打印作业

AbortDoc　　取消一份文档的打印

StartPage　　打印一个新页前要先调用这个函数

EndPage　　用这个函数完成一个页面的打印，并准备设备场景，以便打印下一个页

WritePrinter　　将发送目录中的数据写入打印机

EnumJobs　　枚举打印队列中的作业

GetJob　　获取与指定作业有关的信息

SetJob　　对一个打印作业的状态进行控制

DocumentProperties　　打印机配置控制函数

EnumPrintProcessors　　枚举系统中可用的打印处理器

EnumPrintProcessorDatatypes　　枚举由一个打印处理器支持的数据类型

EnumPrinterDrivers　　枚举指定系统中已安装的打印机驱动程序

EnumForms　　枚举一台打印机可用的表单

AddForm　　为打印机的表单列表添加一个新表单

DeleteForm　　从打印机可用表单列表中删除一个表单

GetForm　　取得与指定表单有关的信息

SetForm 为指定的表单设置信息

AddJob　　用于获取一个有效的路径名，以便用它为作业创建一个后台打印文件。它也会为作业分配一个作业编号

ScheduleJob　　提交一个要打印的作业

DeviceCapabilities　　利用这个函数可获得与一个设备的能力有关的信息

GetDeviceCaps　　获取指定设备的参数设置

EnumMonitors　　枚举可用的打印监视器

EnumPorts　　枚举一个系统可用的端口

GetPrintProcessorDirectory　　判断指定系统中包含了打印机处理器驱动程序及文件的目录

GetPrinterDriverDirectory　　判断指定系统中包含了打印机驱动程序的目录是什么

AddPrinter　　在系统中添加一台新打印机

DeletePrinter　　将指定的打印机标志为从系统中删除

DeletePrinterDriver　　从系统删除一个打印机驱动程序

DeletePrinterDriverEx　　从系统删除一个打印机驱动程序和相关的文件

FlushPrinter　　更新打印机，清楚错误状态的打印机
