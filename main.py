import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from GUI import *

class MyWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)

    def uncompress_start(self):
      print('uncompress_start')
      print(self.textEdit_uncompress_filePath.toPlainText())
      pass

    def compress_start(self):
      print('compress_start')
      pass
    
    def excel2image_start(self):
      print('excel2image_start')
      pass




if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyWindow()
    myWin.show()
    sys.exit(app.exec_())



