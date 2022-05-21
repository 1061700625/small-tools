import os.path
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from GUI import *
from excel2image import *
from compress import *


class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)

    def debugPrint(self, msg):
        previous = self.plainTextEdit_debug.toPlainText().strip()
        self.plainTextEdit_debug.setPlainText(previous+'\n'+str(msg))
        # print(previous+'\n'+str(msg))
        QApplication.processEvents()

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
    
    def excel2image_start(self):  # file:///C:/Users/SXF/Desktop/学术成果一览表.xlsx
        self.pushButton_excel2image_start.setEnabled(False)
        self.debugPrint('>> 开始Excel转图片')
        filePath = self.pathFilter(self.textEdit_excel2image_filePath.toPlainText().strip())
        self.debugPrint(">> 待转换文件为: " + filePath)
        if os.path.exists(filePath):
            try:
                res = excel2image(filePath, debugPrint=self.debugPrint)
                if res:
                    downloadFile(res['FileName'], res['FolderName'], debugPrint=self.debugPrint)
            except Exception as e:
                self.debugPrint(">> 转换出错: "+str(e))
        else:
            self.debugPrint(">> 文件不存在")
        self.pushButton_excel2image_start.setEnabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyWindow()
    myWin.show()
    sys.exit(app.exec_())



