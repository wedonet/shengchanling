#!/usr/bin/env python
# -*- coding: utf-8 -*-
import wx
import openpyxl
import xlrd

import sys
import os
import threading
import time
import datetime

import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.application import MIMEApplication
import random
import json

sourceExcelDir = r"D:\alirpa\奥的斯生生令项目\excel"
recycleExcelDir = r"D:\alirpa\奥的斯生生令项目\excelRecycle"
fileDir = r"D:\alirpa\奥的斯生生令项目\模拟文件"  # 图纸存放路径
sleetsecond = 2 # 发邮件间隔时间（秒）

sender = "support@tjhairong.com"
receiver1 = ["16216077@qq.com"]
mailAlias = "图纸"
smtpAddress = "smtp.qiye.aliyun.com"
smtpPass = "Hr_123456"


class HelloFrame(wx.Frame):
    def __init__(self, *args, **kw):
        super(HelloFrame, self).__init__(*args, **kw)

        # 创建一个Panel实例
        pn1 = wx.Panel(self)

        self.button = wx.Button(pn1, -1, 'Run', pos=(10, 20)) 
        
        self.contents = wx.TextCtrl(pn1, pos=(5, 50), size=(
            370, 150), style=wx.TE_MULTILINE | wx.HSCROLL | wx.TE_READONLY)

        self.Bind(wx.EVT_BUTTON, self.OnButtonClick, self.button)  

        # 提取参数设置
        #setting =  os.path.exists(r'./setting.json') 
        # 读取json文件内容,返回字典格式
        with open('./setting.json','r',encoding='utf8') as fp:
            json_data = json.load(fp)

        #print(json_data['a'])

        # 创建菜单栏
        # self.makeMenuBar()
        # 创建状态栏
        self.CreateStatusBar()
        # 设置状态栏要显示的文本内容
        self.SetStatusText("Status")

    def loopExcel(self):
        listExcel = os.listdir(sourceExcelDir)

        # 循环每个Excel
        for v in listExcel:
            self.handlerExcel(v)

        self.button.Enable(True)

    # 处理excel
    def handlerExcel(self, excelName):
        excelPath = os.path.join(sourceExcelDir, excelName)  
        excelPathRecycle = os.path.join(recycleExcelDir, excelName) 
        
        #wb=openpyxl.load_workbook(excelPath)
        data =  xlrd.open_workbook(excelPath)#文件名以及路径，如果路径或者文件名有中文给前面加一个r拜师原生字符。
        table = data.sheets()[0]          #通过索引顺序获取

        ## 循环Excel
        for num in range(1, table.nrows):
            #print(self.CharToNum('A'))

            #合同名称
            contactName = table.cell(num,  self.CharToNum('E')).value

            self.handlerData(contactName)

            # 发太快了，容易被垃圾邮件
            time.sleep(sleetsecond)

            #self.contents.AppendText(str(num)+"\n")
            #self.contents.ShowPosition(self.contents.GetLastPosition())


        #for num in range(0, 100):
        #    self.contents.AppendText(str(num)+"\n")
        #    self.contents.ShowPosition(self.contents.GetLastPosition())

    def makeMenuBar(self):
        # 创建菜单对象fileMenu(菜单栏主选项1)
        fileMenu = wx.Menu()

        # 在fileMenu中添加子项createItem
        # +item表示子项
        # +helpString表示对子项的说明，当鼠标移动到子项上时，会在状态栏显示
        # \t...语法允许用户键盘操作触发子项
        createItem = fileMenu.Append(
            wx.ID_ANY, item=u"新建文件(N)...\tCtrl-H", helpString="创建一个新的文件")

        # 在各子项中添加起分隔作用的横线
        fileMenu.AppendSeparator()

        # 在fileMenu中添加子项exitItem
        exitItem = fileMenu.Append(wx.ID_EXIT, item=u"退出")

        # 创建菜单对象helpMenu(菜单栏主选项2)
        helpMenu = wx.Menu()
        # 在fileMenu中添加子项aboutItem
        aboutItem = helpMenu.Append(wx.ID_ABOUT, item=u"关于")

        # 创建菜单栏
        menuBar = wx.MenuBar()
        # 添加各个菜单栏主选项到菜单栏中
        # "&"后的首字母+"alt"键触发菜单选项。该首字母会以下划线着重显示，按住alt键即能看见。
        menuBar.Append(fileMenu, u"文件(&F)")
        menuBar.Append(helpMenu, u"帮助(&H)")
        # 添加菜单栏到窗口
        self.SetMenuBar(menuBar)

        # 将主菜单的所有子项绑定动作
        self.Bind(wx.EVT_MENU, self.OnCreate, source=createItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, source=aboutItem)
        self.Bind(wx.EVT_MENU, self.OnExit, source=exitItem)

    def OnExit(self, event):
        # 关闭窗口
        self.Close(True)

    def OnCreate(self, event):
        wx.MessageBox(u"创建文件成功")

    def OnAbout(self, event):
        # MessageBox(message, caption=MessageBoxCaptionStr, style=OK|CENTRE, parent=None, x=DefaultCoord, y=DefaultCoord)
        # +调用message()方法将会弹出一个对话窗口
        # +message表示对话窗口显示的正文信息
        # +caption表示对话窗口的标题
        # +style表示对话窗口的按钮和图标样式
        wx.MessageBox(u"生产令机器人1.0.0",
                      wx.OK | wx.ICON_INFORMATION)

    def OnButtonClick(self, enent):
        # 开新线程
        t1 = threading.Thread(target=self.loopExcel,args=())
        t1.start()

        self.button.Enable(False)
        #wx.MessageBox(u"执行")
        pass

    # 发送图纸
    def handlerData(self, contactName):       
        fileName = contactName + '.jpg'
        
        filePath = os.path.join(fileDir, contactName + '.jpg')   

        #print(filePath)

        # 找到了就发送
        if( os.path.exists(filePath)):     
            print('找到了')
            
            #发送邮件
            self.test(fileName, filePath)

            self.contents.AppendText(contactName+"发送成功\n")
            self.contents.ShowPosition(self.contents.GetLastPosition())
            pass
           
        else:
            #print('没找到', filePath)
            pass
        pass

    # excel列名转数字    
    def CharToNum(self, str):
        switch = {'A': 0,                # 注意此处不要加括号
          'B': 1,
          'C': 2,
          'D': 3,
          'E': 4,
          'F': 5
          }
        return switch.get(str)  

    # 发送邮件
    def test(self, fileName, filePath):
        #sender = setting['sendermail']      # "support@tjhairong.com"
        #receiver1 = setting['receivers']    # "16216077@qq.com"
        #mailAlias = setting['receiverAlias']
        #smtpAddress = setting['smtpAddress']
        #smtpPass = setting['smtpPass'] 


        #receiver2 = "收件人邮箱2@xx.com"
        # 构造邮件类，输入关键的发件人和收件人信息
        test = Mail(smtpAddress, sender, receiver1)                             # Mail("smtp.qiye.aliyun.com",sender,[receiver1])

        # 为邮件实例设定“SMTP”授权码，对应授权码可以在邮箱提供方对应设置页面获取
        test.password =  smtpPass                                                   

        # 创建邮件正文文本，邮件正文使用Html进行编辑，方便插入表格及字体样式
        #text = test.create_html_table([["标题1","标题2","标题3"],[1,2,3]])
        #text = "<h1>邮件发送测试</h1><p>此邮件发送自python_1</p>"+text 

        text = '邮件内容'

        # 指定邮件标题
        title = "邮件标题" +  datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # 指定附件绝对路径
        file_path1 = filePath

        file_paths = {fileName:file_path1}
        
        # 将邮件标题、正文、附件路径写入到邮件类中
        msg =  test.write_mail(mailAlias, title,text,file_paths)
        
        # 发送邮件
        test.send_mail(msg)        


        
# 邮件处理类
class Mail(object):
    """
    邮件类工具，主要用于处理发邮件等操作，支持向正文中插入表格、文本、图片等形式
    1.构造方法中，需要输入smtp服务器名，发件人地址，收件人地址列表
    2.生成实例后，需要给实例的password赋值（即发件人smtp密码）
    3.write_mail方法需要传入参数收件人别名（多个收件人会用同一个名字）,邮件主题,html文本,附件路径(字典，附件名：附件绝对路径)
    4.send_mail方法需要传入参数msg，即write_mail方法的返回值。
    5.提供create_html_table工具方法，传入二维列表，自动生成html表格，可以作为html文本写到邮件正文中
    """
    # 初始化内容，smtp服务器，发件人地址，发件人smtp密码（初始为空），收件人地址列表
    def __init__(self, smtp_server:str, sender:str, receivers:list)->None:
        self.smtp_server = smtp_server
        self.sender = sender
        self.password = ""
        self.receivers = receivers
        pass

    
    # 根据二维列表生成对应的HTML表格
    def create_html_table(self,data):
        """
        根据二维列表生成对应的html文本\t
        Args:\t
            data(list<list>):二维列表\t
        Returns:\t
            html_table(str):html字符串\t
        """   
        html_table = "<table border=1>"
        for row in data:
            tr = "<tr>"
            for col in row:
                td = "<td>"+str(col)+"</td>"
                tr += td
            tr += "</tr>"
            html_table += tr
        html_table += "</table>"
        return html_table

    # 填写邮件内容
    def write_mail(self,receivers_alias,title,html_text,file_paths):
        """
        编写邮件内容，传入文件名\t
        Args:\t
            receivers_alias(str):收件人称谓\t
            title(str):邮件主题\t
            html_text(str):HTML文本，可以以此生成表格、链接、图片等信息\t
            file_paths(dict<str:str>):需要发送的附件的完整路径字典，形如文件名（需带后缀）:文件绝对路径\t
        Returns:\t
            msg:email模块的MIMEMultipart对象，可以带附件，带HTML文本\t
        """
        msg = MIMEMultipart()
        # 头信息
        msg['From'] = Header(self.sender)
        # 这里的收件人是指头信息中的收件人称呼，多个收件邮箱的情况下，收件方只能看到统一的称呼，看不到发送给了其他的谁
        msg['To'] = Header(receivers_alias)
        msg['Subject'] = Header(title)
    
        html_message = MIMEText(html_text, 'html', 'utf-8')
        msg.attach(html_message)
    
        # 添加多个附件
        if len(file_paths)>0:
            for key,value in file_paths.items():
                with open(value,'rb') as attach_file:
                    file_content = attach_file.read()
                file_part = MIMEApplication(file_content)
                file_part.add_header('Content-Disposition', 'attachment', filename=key)
                msg.attach(file_part)
        
        return msg

    # 发送邮件
    def send_mail(self,msg):
        """
        将给定的msg作为邮件内容向指定的收件人地址列表发送\t
        Args:\t
            msg(MIMEMultipart):邮件正文对象\t
        Returns:pass\t
        """
        try:
            # 开启发信服务，这里使用的是加密传输
            server = smtplib.SMTP_SSL(self.smtp_server)
            server.connect(self.smtp_server,465)
            # 登录发信邮箱
            server.login(self.sender, self.password)
            # 发送邮件
            server.sendmail(self.sender, self.receivers, msg.as_string())
            # 关闭服务器
            server.quit()
            print('success')
        except smtplib.SMTPException as e:
            print('error',e)
        pass



if __name__ == "__main__":
    app = wx.App()
    frame = HelloFrame(None, title="生产令机器人1.0.0")
    frame.Show()
    app.MainLoop()
