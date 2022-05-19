from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(475, 773)
        self.bgwidget = QtWidgets.QWidget(Dialog)
        self.bgwidget.setGeometry(QtCore.QRect(-10, 0, 491, 781))
        self.bgwidget.setMinimumSize(QtCore.QSize(0, 0))
        self.bgwidget.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.bgwidget.setStyleSheet("QWidget#bgwidget{\n"
"background-color:qlineargradient(spread:pad, x1:0.353, y1:0.210409, x2:0.930348, y2:1, stop:0.0746269 rgba(60, 0, 28, 255), stop:0.850746 rgba(123, 82, 156, 255))}")
        self.bgwidget.setObjectName("bgwidget")
        self.pushButton = QtWidgets.QPushButton(self.bgwidget)
        self.pushButton.setGeometry(QtCore.QRect(80, 540, 341, 51))
        self.pushButton.setStyleSheet("QPushButton {\n"
"border-radius:20px;\n"
"background-color: rgb(170, 255, 255);\n"
"color:black;\n"
"    font: 75 15pt \"MS Shell Dlg 2\";\n"
"}\n"
"\n"
"QPushButton:hover {\n"
"background-color: rgb(136, 235, 229)\n"
"}")
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(self.bgwidget)
        self.label.setGeometry(QtCore.QRect(130, 70, 241, 81))
        self.label.setStyleSheet("font: 36pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.bgwidget)
        self.label_2.setGeometry(QtCore.QRect(120, 180, 261, 31))
        self.label_2.setStyleSheet("font: 16pt \"MS Shell Dlg 2\";\n"
"color: rgb(255, 255, 255);")
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(self.bgwidget)
        self.line.setGeometry(QtCore.QRect(10, 660, 481, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.lineEdit = QtWidgets.QLineEdit(self.bgwidget)
        self.lineEdit.setGeometry(QtCore.QRect(70, 310, 351, 51))
        self.lineEdit.setStyleSheet("background-color: rgba(0,0,0,0);\n"
"font: 12pt \"MS Shell Dlg 2\";\n"
"")
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.bgwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(70, 420, 351, 51))
        self.lineEdit_2.setStyleSheet("background-color:rgba(0,0,0,0);\n"
"font: 12pt \"MS Shell Dlg 2\";")
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_3 = QtWidgets.QLabel(self.bgwidget)
        self.label_3.setGeometry(QtCore.QRect(70, 270, 101, 31))
        self.label_3.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color: white;")
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.bgwidget)
        self.label_4.setGeometry(QtCore.QRect(70, 390, 101, 21))
        self.label_4.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color: white;")
        self.label_4.setObjectName("label_4")
        self.label.raise_()
        self.label_2.raise_()
        self.line.raise_()
        self.pushButton.raise_()
        self.lineEdit.raise_()
        self.lineEdit_2.raise_()
        self.label_3.raise_()
        self.label_4.raise_()

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.pushButton.setText(_translate("Dialog", "Giriş Yap"))
        self.label.setText(_translate("Dialog", "Giriş Yap"))
        self.label_2.setText(_translate("Dialog", "Hesabınıza giriş yapın"))
        self.label_3.setText(_translate("Dialog", "E-posta"))
        self.label_4.setText(_translate("Dialog", "Şifre"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
