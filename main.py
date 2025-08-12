import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtCore import pyqtSignal, QThread

from qrcodeUI import Ui_MainWindow
from PyQt5.QtWidgets import QFileDialog, QMainWindow, QMessageBox
from gen import imageSet
import os, shutil
from generate import Generator


class CalTheard(QThread):
    _run = pyqtSignal(str)  # 信号类型 str

    def __init__(self, inp, outp, mode):
        super().__init__()
        self.inp = inp
        self.outp = outp
        self.mode = mode

    def run(self):
        # s, k = 0, 0
        # while k <= 50000000:
        #     s += k
        #     k += 1
        GeneratorTemp = Generator()
        df = pd.read_excel(self.inp)
        df.to_csv('temp.csv', index=False)
        df = pd.read_csv('temp.csv')
        titles = list(df.head(0))
        # print(titles)
        for index, row in df.iterrows():
            strs = ''
            for i in range(0, len(titles)):
                strs += titles[i] + ':' + str(row[i]) + '\n'
            GeneratorTemp.gen_code(strs, index, self.outp)
        self._run.emit(str(1))  # 计算结果完成后，发送结果


class UI(Ui_MainWindow, QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.progressBar.hide()
        self.pushButton_2.clicked.connect(self.clicker)
        self.pushButton.clicked.connect(self.runall)
        self.outp = ""
        self.inp = ""
        self.mode = 0

    def clicker(self):
        fname, _ = QFileDialog.getOpenFileNames(self, "打开文件", "", "工作簿xlsx(推荐)(*.xlsx);;工作簿xls(*.xls);;所有文件(*)")
        if fname:
            self.plainTextEdit.setPlainText(fname[0])

    def messageDialog_1(self):
        msg_box = QMessageBox(QMessageBox.Warning, '警告', '输入路径不得为空！')
        msg_box.exec_()

    def messageDialog_2(self):
        msg_box = QMessageBox(QMessageBox.Warning, '警告', '输出文件夹名不得为空！')
        msg_box.exec_()

    def messageDialog_3(self):
        msg_box = QMessageBox(QMessageBox.Warning, '警告', '输入路径不存在！')
        msg_box.exec_()

    def messageDialog_4(self):
        msg_box = QMessageBox(QMessageBox.Information, '成功', '生成完毕！')
        msg_box.exec_()

    def runall(self):
        path_in = self.plainTextEdit.toPlainText()
        path_out = self.plainTextEdit_2.toPlainText()
        flag = 0
        if path_in == "":
            self.messageDialog_1()
        elif not os.path.exists(path_in):
            self.messageDialog_3()
        else:
            flag = 1
            pass
        if path_out == "":
            self.messageDialog_2()
        else:
            if flag == 1:
                prefix = os.path.dirname(path_in)
                if not os.path.exists(os.path.join(prefix, path_out)):
                    os.makedirs(os.path.join(prefix, path_out))
                else:
                    shutil.rmtree(os.path.join(prefix, path_out))
                    os.makedirs(os.path.join(prefix, path_out))
                flag += 1
            else:
                flag = 0

        if flag == 2:
            mode = 0
            self.outp = path_out
            self.inp = path_in
            self.mode = mode
            if not self.checkBox.isChecked():
                self.mode = 1
            self.pushButton.setEnabled(False)  # 将按钮设置为不可点击
            self.progressBar.show()
            self.progressBar.setRange(0, 0)
            prefix = os.path.dirname(self.inp)
            path_out = os.path.join(prefix, path_out)
            self.outp = path_out
            self.cal = CalTheard(path_in, path_out, self.mode)  # 创建一个线程
            self.cal._run.connect(self.update_run)  # 线程发过来的信号挂接到槽函数update_sum
            self.cal.start()  # 线程启动
            # self.final_run(path_in, path_out, mode)

    def update_run(self, r):
        # self.show_label.setText('  ' + r + '  ')  # 信号发过来时，更新QLabel内容
        # self.btn.setText(' 点击计算1+2+...+50000000 ')  # 更新按钮
        self.pushButton.setEnabled(True)  # 让按钮恢复可点击状态
        self.progressBar.setRange(0, 100)
        self.progressBar.setValue(100)
        SetImage = imageSet(self.outp, self.mode, self.inp)
        SetImage.run()
        self.messageDialog_4()
        self.progressBar.hide()
        self.pushButton.setEnabled(True)

    # def final_run(self, inp, outp, mode):
    #     GeneratorTemp = Generator()
    #     df = pd.read_excel(inp)
    #     df.to_csv('temp.csv', index=False)
    #     df = pd.read_csv('temp.csv')
    #     titles = list(df.head(0))[:-1]
    #     # print(titles)
    #     self.progressBar.show()
    #     self.progressBar.setRange(0, 0)
    #     for index, row in df.iterrows():
    #         strs = ''
    #         for i in range(0, len(titles)):
    #             strs += titles[i] + ':' + str(row[i]) + '\n'
    #         GeneratorTemp.gen_code(strs, index, outp)
    #     self.progressBar.setRange(0, 100)
    #     self.progressBar.setValue(100)
    #     SetImage = imageSet(outp, mode, inp)
    #     SetImage.run()
    #     self.messageDialog_4()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    window = UI()
    window.show()
    sys.exit(app.exec_())
