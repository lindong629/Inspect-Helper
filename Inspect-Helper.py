# -*- coding: utf-8 -*-
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from functools import partial
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from configparser import ConfigParser
from datetime import datetime
import sys
import telnetlib
import time
import re
import paramiko
import xlrd
import os
import threading

import Graphical_interface
import Validate_IP_Module

# ********************************************************************************
# * 
# * Inspect Helper GUI v3.0.0          [ December 19, 2022 ] 
# * Written by Vincent
# * 
# * Function: 
# *   * 从Excel表格中批量读取网络设备的IP地址、用户名密码等信息，自动连接到对应网络设备，
# *   批量执行所需的命令行并获取返回的执行结果；
# *   * 命令执行结果保存在本程序所在目录下的OutputFiles目录中，以IP地址、主机名和时间命名；
# *   * 连接设备的方式可以选择SSH协议或Telnet协议。
# * 
# ********************************************************************************

ProgramVersion = "v3.0.0"

class dateTimeDisplay():
    def showtime():
        time=QDateTime.currentDateTime() # 获取当前时间。
        timedisplay=time.toString("yyyy-MM-dd hh:mm:ss dddd") # 调整时间的显示格式。
        Gui.dayTime_Display.setText(timedisplay)

    def starttimer():
        Gui.timer=QTimer()
        Gui.timer.timeout.connect(dateTimeDisplay.showtime) # 通过调用槽函数来刷新时间。
        Gui.timer.start(1000) # 每隔一秒刷新一次, 这里设置为1000ms。

def Select_InfoFile_Path(Gui):
    GetInfoFilePath = filedialog.askopenfilename()        # 弹出"选择文件"的窗口。
    Gui.File_Path_Edit.setText(str(GetInfoFilePath))

def MainTask_SSH(Gui):
    Lock_Buttons(Gui)

    Gui.Message_Display_Edit.insertPlainText("正在读取设备信息文件……\n")    # 在"运行状态信息"下的文本框中追加显示状态信息。
    Message_Display_Edit_Scroll(Gui)                      # 执行下面定义好的函数, 让状态信息文本框始终滚动到最新提示的文字处。
    InfoFilePath = Gui.File_Path_Edit.text()
    is_xls = ".xls" in InfoFilePath
    if InfoFilePath == '' :                               # 判断"设备信息文件"是否没有选择。
        Gui.Message_Display_Edit.insertPlainText("找不到设备信息文件……\n")
        Message_Display_Edit_Scroll(Gui)
        messagebox.showwarning('错误','请选择设备信息文件!')
        Unlock_Buttons(Gui)
        return
    elif is_xls == False :                                # 判断"设备信息文件"是否是.xls格式, 如果不是, 则提示并停止运行。
        Gui.Message_Display_Edit.insertPlainText("设备信息文件只能选择 .xls 格式的文件……\n")
        Message_Display_Edit_Scroll(Gui)
        messagebox.showwarning('错误','设备信息文件只能是 .xls 格式!')
        Unlock_Buttons(Gui)
        return
    WorkBook = xlrd.open_workbook(InfoFilePath)
    WorkSheet = WorkBook.sheet_by_name('Devices')         # 读取Excel表格的内容, 根据工作簿底部的名称"Devices"查找工作表。

    Gui.Message_Display_Edit.insertPlainText("正在读取需要执行的命令……\n")
    Message_Display_Edit_Scroll(Gui)
    ScriptsCache = Gui.Scripts_Edit.toPlainText()
    ScriptsCacheFile = open(".\\Cache\\" + "Scripts-Cache.txt","a+")    # 将文本框中输入的命令保存到缓存文本文件中, 为了能够逐行读取。
    Gui.Message_Display_Edit.insertPlainText("正在清理缓存文件内容……\n")
    Message_Display_Edit_Scroll(Gui)
    ScriptsCacheFile.seek(0)                              # 在每次将命令写入缓存文件之前, 先清空该文件中原有的文本内容。
    ScriptsCacheFile.truncate()
    Gui.Message_Display_Edit.insertPlainText("将需要执行的命令写入缓存……\n")
    Message_Display_Edit_Scroll(Gui)
    ScriptsCacheFile.write(ScriptsCache)
    ScriptsCacheFile.close()

    Gui.Message_Display_Edit.insertPlainText("准备连接设备……\n")
    Message_Display_Edit_Scroll(Gui)
    for i in range(1,WorkSheet.nrows):                    # 循环读取Excel表中的设备信息, 每读取一行执行一遍。
        Gui.Progress_Bar.setValue(0)
        DeviceInfo = WorkSheet.row_values(i)
        DeviceName = str(DeviceInfo[0])                   # 数字代表Excel表的第几列。
        IPAddress = str(DeviceInfo[1])
        Username = str(DeviceInfo[2])
        Password = str(DeviceInfo[3])
        Gui.Progress_Bar.setValue(5)

        isIP = Validate_IP_Module.Validate_IP(IPAddress)  # 判断IP地址的格式是否正确。
        if isIP == False :
            Gui.Message_Display_Edit.insertPlainText("IP地址: "+ IPAddress +" 格式不正确……\n")
            Message_Display_Edit_Scroll(Gui)
            messagebox.showwarning('输入错误','IP地址: '+ IPAddress +' 格式不正确!')
            Gui.Progress_Bar.setValue(0)
            break
        Gui.Progress_Bar.setValue(10)

        Gui.Message_Display_Edit.insertPlainText("正在连接设备: " + IPAddress + " ……\n")
        Message_Display_Edit_Scroll(Gui)
        try:
            SSH_Client = paramiko.SSHClient()
            SSH_Client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            SSH_Client.connect(hostname=IPAddress, username=Username, password=Password, timeout=10)
        except:
            Gui.Message_Display_Edit.insertPlainText("设备: "+ IPAddress +" 连接失败……\n\n")
            Message_Display_Edit_Scroll(Gui)
            messagebox.showwarning('连接失败','连接失败, 请检查: '+ IPAddress +' 的IP地址、用户名和密码是否正确!')
            Gui.Progress_Bar.setValue(0)
            break
        runCommand = SSH_Client.invoke_shell()
        Gui.Progress_Bar.setValue(15)

        Scripts = open(".\\Cache\\Scripts-Cache.txt")     # 从缓存文件中读取所需执行的命令。
        Script = Scripts.readline()                       # 逐行读取。
        Gui.Progress_Bar.setValue(20)
        
        runCommand.send("\n")
        Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 已连接……\n")
        Message_Display_Edit_Scroll(Gui)
        time.sleep(2)
        Gui.Message_Display_Edit.insertPlainText("正在执行命令……\n")
        Message_Display_Edit_Scroll(Gui)
        while Script:                                     # 执行命令, 每一行执行一次。
            runCommand.send(Script + "\n")
            time.sleep(5)                                 # 每次执行后将程序暂停5秒钟, 给目标一个反应时间。
            Script = Scripts.readline()
        Scripts.close()
        Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 执行完毕, 正在保存结果……\n")
        Message_Display_Edit_Scroll(Gui)
        CmdResult = runCommand.recv(65535).decode("utf8","ignore")
        Gui.Progress_Bar.setValue(80)
        with open(".\\Cache\\CmdResult-Cache.txt", 'w') as CmdResultCache:    # 将设备返回的显示信息保存到缓存文件中;
            CmdResultCache.seek(0)                        # 先保存到缓存文件中, 是为了后面能够逐行读取, 清除空白行。
            CmdResultCache.truncate()
            CmdResultCache.write(CmdResult)
            CmdResultCache.close()
        Gui.Progress_Bar.setValue(90)

        now = datetime.now()
        nowDateTime = now.strftime("%Y-%m-%d_%H_%M_%S")
        OutputFile = open(".\\OutputFiles\\"+ IPAddress +"-"+ DeviceName +"-"+str(nowDateTime)+".txt","a+")
        with open(".\\Cache\\CmdResult-Cache.txt", 'r') as CmdResultCache:
            for line in CmdResultCache:
                if line.strip():                          # 逐行读取, 清除空白行;
                    OutputFile.write(line)                # 将处理后的文本保存到最终的输出文件中。
        OutputFile.close()
        Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 已保存结果……\n")
        Message_Display_Edit_Scroll(Gui)
        Gui.Progress_Bar.setValue(100)

        SSH_Client.close()
        Gui.Message_Display_Edit.insertPlainText("已退出设备: " + IPAddress + " ……\n\n")
        Message_Display_Edit_Scroll(Gui)
    
    # 在完全执行完毕后, 弹出一个窗口让用户确认。如果此处没有这个弹窗, 程序执行期间如果点击了隐藏的按钮的位置,
    # 即便是按钮已经隐藏并设为不可用了, 但还会在函数执行完毕之后的一瞬间触发点击动作, 弹出窗口是为了停顿一下,
    # 让用户点击"确认"之后再使按钮转为正常显示、可点击状态:
    if messagebox.showinfo('完成','运行结束!') == True :
        Unlock_Buttons(Gui)
    else:
        Unlock_Buttons(Gui)

def MainTask_Telnet(Gui):
    Lock_Buttons(Gui)

    Gui.Message_Display_Edit.insertPlainText("正在读取设备信息文件……\n")
    Message_Display_Edit_Scroll(Gui)
    InfoFilePath = Gui.File_Path_Edit.text()
    is_xls = ".xls" in InfoFilePath
    if InfoFilePath == '' :
        Gui.Message_Display_Edit.insertPlainText("找不到设备信息文件……\n")
        Message_Display_Edit_Scroll(Gui)
        messagebox.showwarning('错误','请选择设备信息文件!')
        Unlock_Buttons(Gui)
        return
    elif is_xls == False :
        Gui.Message_Display_Edit.insertPlainText("设备信息文件只能选择 .xls 格式的文件……\n")
        Message_Display_Edit_Scroll(Gui)
        messagebox.showwarning('错误','设备信息文件只能是 .xls 格式!')
        Unlock_Buttons(Gui)
        return
    WorkBook = xlrd.open_workbook(InfoFilePath)
    WorkSheet = WorkBook.sheet_by_name('Devices')

    Gui.Message_Display_Edit.insertPlainText("正在读取需要执行的命令……\n")
    Message_Display_Edit_Scroll(Gui)
    ScriptsCache = Gui.Scripts_Edit.toPlainText()
    ScriptsCacheFile = open(".\\Cache\\" + "Scripts-Cache.txt","a+")
    Gui.Message_Display_Edit.insertPlainText("正在清理缓存文件内容……\n")
    Message_Display_Edit_Scroll(Gui)
    ScriptsCacheFile.seek(0)
    ScriptsCacheFile.truncate()
    Gui.Message_Display_Edit.insertPlainText("将需要执行的命令写入缓存……\n")
    Message_Display_Edit_Scroll(Gui)
    ScriptsCacheFile.write(ScriptsCache)
    ScriptsCacheFile.close()

    Gui.Message_Display_Edit.insertPlainText("准备连接设备……\n")
    Message_Display_Edit_Scroll(Gui)
    for i in range(1,WorkSheet.nrows):
        Gui.Progress_Bar.setValue(0)
        DeviceInfo = WorkSheet.row_values(i)
        DeviceName = str(DeviceInfo[0])
        IPAddress = str(DeviceInfo[1])
        Username = str(DeviceInfo[2])
        Password = str(DeviceInfo[3])
        Gui.Progress_Bar.setValue(5)

        isIP = Validate_IP_Module.Validate_IP(IPAddress)
        if isIP == False :
            Gui.Message_Display_Edit.insertPlainText("IP地址: "+ IPAddress +" 格式不正确……\n")
            Message_Display_Edit_Scroll(Gui)
            messagebox.showwarning('输入错误','IP地址: '+ IPAddress +' 格式不正确!')
            Gui.Progress_Bar.setValue(0)
            break
        Gui.Progress_Bar.setValue(10)

        Gui.Message_Display_Edit.insertPlainText("正在连接设备: " + IPAddress + " ……\n")
        Message_Display_Edit_Scroll(Gui)
        try:
            Telnet_Client = telnetlib.Telnet(IPAddress, timeout=10)
        except:
            Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 连接失败……\n\n")
            Message_Display_Edit_Scroll(Gui)
            messagebox.showwarning('连接失败','连接失败, 请检查: '+ IPAddress +' 的IP地址是否正确!')
            Gui.Progress_Bar.setValue(0)
            break
        ReceivedMessage = Telnet_Client.read_some().decode('ascii')
        if ('sername' in ReceivedMessage) or ('ogin' in ReceivedMessage):
            Telnet_Client.write(Username.encode('ascii') + b'\n')
        if Telnet_Client.read_until(b'assword'):
            Telnet_Client.write(Password.encode('ascii') + b'\n')
        time.sleep(2)
        LoginResult = Telnet_Client.read_very_eager().decode('ascii')
        if ('nvalid' in LoginResult) or ('ailed' in LoginResult):
            Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 的用户名或密码错误……\n\n")
            Message_Display_Edit_Scroll(Gui)
            messagebox.showwarning('连接失败','连接失败, 设备: ' + IPAddress + ' 的用户名或密码错误!')
            Gui.Progress_Bar.setValue(0)
            break

        Gui.Progress_Bar.setValue(15)

        Scripts = open(".\\Cache\\Scripts-Cache.txt")
        Script = Scripts.readline()
        Gui.Progress_Bar.setValue(20)
        
        Telnet_Client.write(b'\n')
        Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 已连接……\n")
        Message_Display_Edit_Scroll(Gui)
        time.sleep(2)
        Gui.Message_Display_Edit.insertPlainText("正在执行命令……\n")
        Message_Display_Edit_Scroll(Gui)
        while Script:
            Telnet_Client.write(Script.encode('ascii') + b'\n')
            time.sleep(5)
            Script = Scripts.readline()
        Scripts.close()
        Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 执行完毕, 正在保存结果……\n")
        Message_Display_Edit_Scroll(Gui)
        CmdResult = Telnet_Client.read_very_eager().decode('ascii')

        regr = '\r'
        regn = '\n'
        reg08 = '\x08'
        CmdResult = re.sub(r'\\r',regr,CmdResult)
        CmdResult = re.sub(r'\\n',regn,CmdResult)
        CmdResult = re.sub(r'\\x08',reg08,CmdResult)

        Gui.Progress_Bar.setValue(80)
        with open(".\\Cache\\CmdResult-Cache.txt", 'w') as CmdResultCache:
            CmdResultCache.seek(0)
            CmdResultCache.truncate()
            CmdResultCache.write(CmdResult)
            CmdResultCache.close()
        Gui.Progress_Bar.setValue(90)

        now = datetime.now()
        nowDateTime = now.strftime("%Y-%m-%d_%H_%M_%S")
        OutputFile = open(".\\OutputFiles\\"+ IPAddress +"-"+ DeviceName +"-"+str(nowDateTime)+".txt","a+")
        with open(".\\Cache\\CmdResult-Cache.txt", 'r') as CmdResultCache:
            for line in CmdResultCache:
                if line.strip():
                    OutputFile.write(line)
        OutputFile.close()
        Gui.Message_Display_Edit.insertPlainText("设备: " + IPAddress + " 已保存结果……\n")
        Message_Display_Edit_Scroll(Gui)
        Gui.Progress_Bar.setValue(100)

        Telnet_Client.close()
        Gui.Message_Display_Edit.insertPlainText("已退出设备: " + IPAddress + " ……\n\n")
        Message_Display_Edit_Scroll(Gui)

    if messagebox.showinfo('完成','运行结束!') == True :
        Unlock_Buttons(Gui)
    else:
        Unlock_Buttons(Gui)

def MainTask_SSH_Thread(Gui):
    TaskThread = threading.Thread(target=MainTask_SSH(Gui))
    TaskThread.daemon=True
    TaskThread.start()

def MainTask_Telnet_Thread(Gui):
    TaskThread = threading.Thread(target=MainTask_Telnet(Gui))
    TaskThread.daemon=True
    TaskThread.start()

def Start_Main_Task(Gui):                                     # 在点击"开始执行"按钮时, 根据"单选按钮"的选择状态来判断使用哪种协议。
    if Gui.Connect_By_SSH_Button.isChecked():
        Gui.Message_Display_Edit.insertPlainText("准备使用SSH协议连接设备……\n")
        Message_Display_Edit_Scroll(Gui)
        MainTask_SSH_Thread(Gui)
    elif Gui.Connect_By_Telnet_Button.isChecked():
        Gui.Message_Display_Edit.insertPlainText("准备使用Telnet协议连接设备……\n")
        Message_Display_Edit_Scroll(Gui)
        MainTask_Telnet_Thread(Gui)

def Open_Saved_Dir():                                         # 打开输出文件所在的文件夹, 此函数由"浏览保存目录"按钮来连接。
    os.startfile('.\\OutputFiles\\')

def Message_Display_Edit_Scroll(Gui):
    Text_Cursor = Gui.Message_Display_Edit.textCursor()
    Text_Cursor.movePosition(Text_Cursor.End)
    Gui.Message_Display_Edit.setTextCursor(Text_Cursor)

def Message_Display_Clear(Gui):                               # 清空"运行状态信息"文本框的内容, 此函数由"清空状态信息"按钮来连接。
    Gui.Message_Display_Edit.setPlainText("")

def Lock_Buttons(Gui):
    Gui.Start_Main_Task_Button.setEnabled(False)          # 当开始连接设备时, 调用此函数将"开始执行"按钮和
    Gui.Start_Main_Task_Button.setVisible(False)          # "清空状态信息"按钮变为不可用状态并隐藏。
    Gui.Message_Display_Clear_Button.setEnabled(False)    # 隐藏后, 显示的灰色按钮是虚假按钮"Start_Main_Task_Button_Fake"
    Gui.Message_Display_Clear_Button.setVisible(False)    # 和"Message_Display_Clear_Button_Fake"。

def Unlock_Buttons(Gui):
    Gui.Start_Main_Task_Button.setEnabled(True)           # 当任务执行结束时, 调用此函数重新将隐藏的两个按钮恢复显示;
    Gui.Start_Main_Task_Button.setVisible(True)           # 包括"开始执行"按钮和"清空状态信息"按钮。
    Gui.Message_Display_Clear_Button.setEnabled(True)
    Gui.Message_Display_Clear_Button.setVisible(True)


if __name__ == '__main__':
    App = QApplication(sys.argv)
    MainWindow = QMainWindow()
    Gui = Graphical_interface.Ui_MainWindow()
    Gui.setupUi(MainWindow)
    MainWindow.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
    MainWindow.setFixedSize(MainWindow.width(), MainWindow.height())
    dateTimeDisplay.starttimer()
    Gui.Version_Display_Label.setText(str(ProgramVersion))    # 在窗口左下角显示版本信息。

    MainWindow.show()

    Gui.Select_File_Button.clicked.connect(partial(Select_InfoFile_Path,Gui))
    Gui.Start_Main_Task_Button.clicked.connect(partial(Start_Main_Task,Gui))
    Gui.Message_Display_Clear_Button.clicked.connect(partial(Message_Display_Clear,Gui))
    Gui.Open_Saved_Dir_Button.clicked.connect(partial(Open_Saved_Dir))

    sys.exit(App.exec_())
