import os.path
import sys
import time

from PyQt5.QtWidgets import QApplication, QMainWindow
from GUI import *
from compress import *
from convertor import Convertor


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.convertor = Convertor()

    def debugPrint(self, msg):
        self.plainTextEdit_debug.appendPlainText(str(msg))
        QApplication.processEvents()
        # 文本框显示到底部
        self.plainTextEdit_debug.moveCursor(self.plainTextEdit_debug.textCursor().End)
        # 睡眠
        time.sleep(0.1)

    def pathFilter(self, path):
        if path.startswith('file://'):
            path = path[8:]
        return path

    def uncompress_start(self):
        self.pushButton_uncompress_start.setEnabled(False)
        self.debugPrint('>> 开始解压')
        filePath = self.pathFilter(self.textEdit_uncompress_filePath.toPlainText().strip())
        # file:///C:/Users/SXF/Desktop/Converted files.zip
        if os.path.isfile(filePath):
            self.debugPrint('>> 是压缩包，即将进入解压模式！')
            if filePath.endswith('rar'):
                self.debugPrint('>> rar格式！')
                try:
                    rar_uncompress(filePath, debugPrint=self.debugPrint)
                except Exception as e:
                    self.debugPrint(">> 解压出错: " + str(e))
            elif filePath.endswith('zip'):
                self.debugPrint('>> zip格式！')
                try:
                    zip_uncompress(filePath, debugPrint=self.debugPrint)
                except Exception as e:
                    self.debugPrint(">> 解压出错: " + str(e))
            else:
                self.debugPrint('>> 格式不支持！')
        else:
            self.debugPrint('>> 不是压缩包！')
        self.pushButton_uncompress_start.setEnabled(True)

    def compress_start(self):
        self.pushButton_compress_start.setEnabled(False)
        self.debugPrint('>> 开始压缩')
        dirPath = self.pathFilter(self.textEdit_compress_filePath.toPlainText().strip())
        if os.path.isdir(dirPath):
            self.debugPrint('>> 是文件夹，即将进入压缩模式！')
            zip_compress(dirPath, debugPrint=self.debugPrint)
        else:
            self.debugPrint('>> 不是文件夹！')
        self.pushButton_compress_start.setEnabled(True)
    
    def wps2image_start(self):  # file:///C:/Users/SXF/Desktop/学术成果一览表.xlsx
        self.pushButton_wps2image_start.setEnabled(False)
        self.debugPrint('>> 开始Excel/Word/()转图片')
        filePath = self.pathFilter(self.textEdit_wps2image_filePath.toPlainText().strip())
        self.debugPrint(">> 待转换文件为: " + filePath)
        if os.path.exists(filePath):
            try:
                if filePath.endswith('xls') or filePath.endswith('xlsx'):
                    self.convertor.excel2image(filePath, debugPrint=self.debugPrint)
                elif filePath.endswith('doc'):
                    # self.debugPrint('>> 先将doc转为docx!(更建议手动转换，自动转可能格式出错)')
                    # self.convertor.doc2docx(filePath, debugPrint=self.debugPrint)
                    # filePath = filePath + 'x'
                    self.convertor.doc2image(filePath, debugPrint=self.debugPrint)
                elif filePath.endswith('docx'):
                    self.convertor.word2image(filePath, debugPrint=self.debugPrint)

                else:
                    self.debugPrint('>> 转换类型不支持!')
            except Exception as e:
                self.debugPrint(">> 转换出错: "+str(e))
        else:
            self.debugPrint(">> 文件不存在")
        self.pushButton_wps2image_start.setEnabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyWindow()
    myWin.show()
    sys.exit(app.exec_())



