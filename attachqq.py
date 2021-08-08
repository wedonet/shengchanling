from pywinauto.application import Application

#app = Application(backend="uia").connect(process=16000)

#print(app)

import pywinauto

# 通过窗口打开
app = pywinauto.Desktop()
win = app['QQ']

win.draw_outline()


print(win)