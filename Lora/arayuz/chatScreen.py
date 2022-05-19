from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(475, 772)
        self.bgwidget = QtWidgets.QWidget(Dialog)
        self.bgwidget.setGeometry(QtCore.QRect(-10, 0, 491, 781))
        self.bgwidget.setMinimumSize(QtCore.QSize(0, 0))
        self.bgwidget.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.bgwidget.setStyleSheet("QWidget#bgwidget{\n"
"background-color:qlineargradient(spread:pad, x1:0.353, y1:0.210409, x2:0.930348, y2:1, stop:0.0746269 rgba(60, 0, 28, 255), stop:0.850746 rgba(123, 82, 156, 255))}")
        self.bgwidget.setObjectName("bgwidget")
        self.label = QtWidgets.QLabel(self.bgwidget)
        self.label.setGeometry(QtCore.QRect(190, 30, 121, 81))
        self.label.setStyleSheet("font: 36pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.bgwidget)
        self.label_2.setGeometry(QtCore.QRect(140, 120, 221, 31))
        self.label_2.setStyleSheet("font: 16pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);")
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(self.bgwidget)
        self.line.setGeometry(QtCore.QRect(0, 620, 481, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.line_2 = QtWidgets.QFrame(self.bgwidget)
        self.line_2.setGeometry(QtCore.QRect(10, 160, 481, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.textEdit = QtWidgets.QTextEdit(self.bgwidget)
        self.textEdit.setGeometry(QtCore.QRect(40, 660, 361, 91))
        self.textEdit.setStyleSheet("background-color:rgba(0,0,0,0);")
        self.textEdit.setObjectName("textEdit")
        self.textBrowser = QtWidgets.QTextBrowser(self.bgwidget)
        self.textBrowser.setGeometry(QtCore.QRect(30, 190, 441, 421))
        self.textBrowser.setStyleSheet("background-color:rgba(0,0,0,0);")
        self.textBrowser.setObjectName("textBrowser")
        self.pushButton = QtWidgets.QPushButton(self.bgwidget)
        self.pushButton.setGeometry(QtCore.QRect(420, 660, 61, 91))
        self.pushButton.setStyleSheet("font:70 12pt \"MS Shell Dlg 2\";")
        self.pushButton.setObjectName("pushButton")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "Lora"))
        self.label_2.setText(_translate("Dialog", "Ne bilmek istersin?"))
        self.pushButton.setText(_translate("Dialog", "↑"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
