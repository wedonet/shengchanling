## 初始化

python3 -m venv ./venv

## 运行虚拟环境

./venv/Scripts/activate.bat





## 安装包

pip3 install wxPython

pip3 install openpyxl

*** 读xls格式的excel ***

pip3 install  xlrd  

https://www.cnblogs.com/lnd-blog/p/12535423.html



打包 pyinstaller -F -w index.py

## 问题

邮件中的excel文件，最好统一成一种格式 xlsx。

具体文件规则，从excel哪列，按什么规则去哪里找图纸。

it部提供一个，smtp信息用于发送邮件。

为避免被当作垃圾邮件拦截，标题加了发送时间可以接受吗

## 风险

smtp发送大文件，网络带宽
