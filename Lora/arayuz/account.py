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
        self.pushButton.setGeometry(QtCore.QRect(70, 580, 341, 51))
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
        self.line = QtWidgets.QFrame(self.bgwidget)
        self.line.setGeometry(QtCore.QRect(10, 660, 481, 20))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.lineEdit = QtWidgets.QLineEdit(self.bgwidget)
        self.lineEdit.setGeometry(QtCore.QRect(70, 200, 351, 51))
        self.lineEdit.setStyleSheet("background-color: rgba(0,0,0,0);\n"
"font: 12pt \"MS Shell Dlg 2\";\n"
"")
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.bgwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(70, 300, 351, 51))
        self.lineEdit_2.setStyleSheet("background-color:rgba(0,0,0,0);\n"
"font: 12pt \"MS Shell Dlg 2\";")
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_3 = QtWidgets.QLabel(self.bgwidget)
        self.label_3.setGeometry(QtCore.QRect(70, 170, 101, 31))
        self.label_3.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color: white;")
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.bgwidget)
        self.label_4.setGeometry(QtCore.QRect(70, 270, 101, 21))
        self.label_4.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color: white;")
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.bgwidget)
        self.label_5.setGeometry(QtCore.QRect(70, 360, 131, 31))
        self.label_5.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color: white;")
        self.label_5.setObjectName("label_5")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.bgwidget)
        self.lineEdit_3.setGeometry(QtCore.QRect(70, 400, 351, 51))
        self.lineEdit_3.setStyleSheet("background-color:rgba(0,0,0,0);\n"
"font: 12pt \"MS Shell Dlg 2\";")
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.radioButton = QtWidgets.QRadioButton(self.bgwidget)
        self.radioButton.setGeometry(QtCore.QRect(130, 500, 101, 21))
        self.radioButton.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color:white;")
        self.radioButton.setObjectName("radioButton")
        self.radioButton_2 = QtWidgets.QRadioButton(self.bgwidget)
        self.radioButton_2.setGeometry(QtCore.QRect(260, 500, 101, 21))
        self.radioButton_2.setStyleSheet("font: 12pt \"MS Shell Dlg 2\";\n"
"color:white;")
        self.radioButton_2.setObjectName("radioButton_2")
        self.label.raise_()
        self.line.raise_()
        self.pushButton.raise_()
        self.lineEdit.raise_()
        self.lineEdit_2.raise_()
        self.label_3.raise_()
        self.label_4.raise_()
        self.label_5.raise_()
        self.lineEdit_3.raise_()
        self.radioButton.raise_()
        self.radioButton_2.raise_()

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.pushButton.setText(_translate("Dialog", "Kayıt Ol"))
        self.label.setText(_translate("Dialog", "Hesabın"))
        self.label_3.setText(_translate("Dialog", "İsim"))
        self.label_4.setText(_translate("Dialog", "Soyisim"))
        self.label_5.setText(_translate("Dialog", "Doğum Tarihi"))
        self.radioButton.setText(_translate("Dialog", "Kadın"))
        self.radioButton_2.setText(_translate("Dialog", "Erkek"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
