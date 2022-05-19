from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        #Dialog.resize(491, 781)
        self.bgwidget = QtWidgets.QWidget(Dialog)
        self.bgwidget.setGeometry(QtCore.QRect(0, 0, 491, 781))
        self.bgwidget.setMinimumSize(QtCore.QSize(0, 0))
        self.bgwidget.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.bgwidget.setStyleSheet("QWidget#bgwidget{\n"
"background-color:qlineargradient(spread:pad, x1:0.353, y1:0.210409, x2:0.930348, y2:1, stop:0.0746269 rgba(60, 0, 28, 255), stop:0.850746 rgba(123, 82, 156, 255))}")
        self.bgwidget.setObjectName("bgwidget")
        self.pushButton = QtWidgets.QPushButton(self.bgwidget)
        self.pushButton.setGeometry(QtCore.QRect(210, 630, 80, 80))
        self.pushButton.setStyleSheet("QPushButton {\n"
"    border: 1px solid #8f8f91;\n"
"    border-radius: 40px;\n"
"    background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,\n"
"                                      stop: 0 #f6f7fa, stop: 1 #dadbde);\n"
"    min-width: 80px;\n"
"}\n"
"\n"
"QPushButton:pressed {\n"
"    background-color: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,\n"
"                                      stop: 0 #dadbde, stop: 1 #f6f7fa);\n"
"}\n"
"\n"
"QPushButton:flat {\n"
"    border: none; \n"
"}\n"
"\n"
"QPushButton:default {\n"
"    border-color: navy;\n"
"}")
        self.pushButton.setText("")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setIcon(QtGui.QIcon('C:\\Users\\OUTLIER\\Desktop\\voiceAssistant\\icon\\mic.ico'))
        self.pushButton.setIconSize(QtCore.QSize(75,75))
        self.label = QtWidgets.QLabel(self.bgwidget)
        self.label.setGeometry(QtCore.QRect(140, 70, 241, 81))
        self.label.setStyleSheet("font: 36pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.bgwidget)
        self.label_2.setGeometry(QtCore.QRect(80, 160, 361, 31))
        self.label_2.setStyleSheet("font: 16pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);")
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(self.bgwidget)
        self.line.setGeometry(QtCore.QRect(0, 660, 491, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.lineEdit = QtWidgets.QLineEdit(self.bgwidget)
        self.lineEdit.setGeometry(QtCore.QRect(90, 230, 335, 35))
        self.lineEdit.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);\n"
"background-color:rgba(0,0,0,0);\n"
"")
        self.lineEdit.setInputMethodHints(QtCore.Qt.ImhNone)
        self.lineEdit.setText("")
        self.lineEdit.setFrame(True)
        self.lineEdit.setCursorPosition(0)
        self.lineEdit.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit.setClearButtonEnabled(True)
        self.lineEdit.setObjectName("lineEdit")
        self.label.raise_()
        self.label_2.raise_()
        self.line.raise_()
        self.pushButton.raise_()
        self.lineEdit.raise_()

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Lora", "Lora"))
        Dialog.setFixedSize(491, 781)
        self.label.setText(_translate("Dialog", "Merhaba"))
        self.label_2.setText(_translate("Dialog", "Sana nasıl yardımcı olabilirim?"))
        self.lineEdit.setPlaceholderText(_translate("Dialog", "Lora\'ya Yaz"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    Dialog.setWindowIcon(QtGui.QIcon('C:\\Users\\OUTLIER\\Desktop\\voiceAssistant\\icon\\lora.ico'))
    ui = Ui_Dialog()    
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
