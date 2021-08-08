# 参考 https://blog.csdn.net/lishan132/article/details/108949975
 
from subprocess import Popen
from pywinauto import Desktop
from pywinauto.application import Application

import time
import os.path

path = r'c:/windows/system32/calc.exe'


# 设置txt文件保存路径
txt_path = './TXT/'
if not os.path.exists(txt_path):
    os.mkdir(txt_path)
 

# 设置控件信息的txt文件名
control_txt_name = txt_path + 'control.txt'

#Popen('calc.exe', shell=True)

#dlg = Desktop(backend="uia").Calculator
#dlg.wait('visible')

#print(dlg)

#time.sleep(3)


#app = Application().connect(class_name="计算器")

#app = Application().connect(path = path)

app = Application(backend="uia").connect(title=u"计算器")

time.sleep(5)

mainwin = app['计算器']

mainwin.set_focus()

 
print(mainwin)


# 打印控件信息到指定路径
#mainwin.print_control_identifiers(filename=control_txt_name)

btn0 = mainwin.window(auto_id = 'num0Button', control_type="Button")
time.sleep(0.2)

mainwin.window(auto_id = 'num0Button', control_type="Button").draw_outline(colour = 'red')
time.sleep(0.2)

mainwin.window(auto_id = 'num0Button', control_type="Button").click()
time.sleep(0.2)

btn1 = mainwin.window(auto_id = 'num1Button', control_type="Button")
time.sleep(0.2)
btn1.click()
btn1.draw_outline(colour = 'blue')

#btn0.draw_outline(colour = 'blue')

#mainwin.draw_outline()


