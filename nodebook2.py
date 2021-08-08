# 打开关闭记事本

# 参考

# https://www.dazhuanlan.com/xiangjf/topics/1163516

from pywinauto.application import Application

app = Application(backend="uia").start("notepad.exe")
# app.UntitledNotepad.type_keys("%FX") 

# type_keys("%FX") 就是按下 “Alt+F,X”，从记事本菜单中可以看到，这个组合键对应于 “文件 | 退出”。这样，记事本退出也就不奇怪了。
app['无标题 - 记事本'].type_keys("%FX")   