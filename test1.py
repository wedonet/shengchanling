from pywinauto import Desktop, Application

app = Application(backend="uia").connect(path="explorer.exe", title="Program F")    


# 找到主窗口-title模糊匹配
wnd = app['Program']  

# 打印所有控件
# wnd.print_control_identifiers()  

common_files=wnd['common files']  


common_files.right_click_input()    