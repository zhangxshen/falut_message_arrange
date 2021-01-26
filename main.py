# -*- coding = utf-8 -*-

import openpyxl
import os
from PyQt5.QtWidgets import QApplication
import warnings
import sys
from models import ui_control

warnings.filterwarnings("ignore")  # 屏蔽控制台中的非重要警告

"""---------------------------------
主函数~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
负责界面程序控制及按键绑定
---------------------------------"""


def main():
    app = QApplication(sys.argv)
    MainWindow = ui_control.tool_ctrl()
    MainWindow.pushButton.clicked.connect(
        MainWindow.file_import)  # 文件导入功能与按键的绑定
    MainWindow.pushButton_2.clicked.connect(MainWindow.run)  # 运行功能与按键的绑定
    MainWindow.pushButton_3.clicked.connect(MainWindow.close)  # 关闭按钮绑定
    MainWindow.pushButton_4.clicked.connect(
        MainWindow.showMinimized)  # 最小化按钮绑定
    MainWindow.show()

    print('输出文件保存在%s' % os.getcwd())
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
