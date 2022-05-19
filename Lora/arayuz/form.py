from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_guigui(object):
    def setupUi(self, guigui):
        guigui.setObjectName("guigui")
        guigui.resize(560, 827)
        self.grad = QtWidgets.QLabel(guigui)
        self.grad.setGeometry(QtCore.QRect(-30, -110, 601, 941))
        self.grad.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0.959909, stop:0 rgba(255, 185, 0, 255), stop:1 rgba(37, 255, 255, 255));")
        self.grad.setText("")
        self.grad.setObjectName("grad")
        self.sendtext = QtWidgets.QLineEdit(guigui)
        self.sendtext.setGeometry(QtCore.QRect(70, 170, 271, 51))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        self.sendtext.setFont(font)
        self.sendtext.setStyleSheet("")
        self.sendtext.setObjectName("sendtext")
        self.pushButton = QtWidgets.QPushButton(guigui)
        self.pushButton.setGeometry(QtCore.QRect(380, 170, 111, 51))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(14)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.pushButton.setFont(font)
        self.pushButton.setStyleSheet("color: rgb(255, 255, 255);\n"
"background-color: rgba(47, 47, 47,200);\n"
"border: none;\n"
"border-radius:25px;\n"
"font: 14pt \"Arial\";")
        self.pushButton.setObjectName("pushButton")
        self.sendandrec = QtWidgets.QTextEdit(guigui)
        self.sendandrec.setGeometry(QtCore.QRect(60, 250, 461, 561))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        self.sendandrec.setFont(font)
        self.sendandrec.setObjectName("sendandrec")
        self.label = QtWidgets.QLabel(guigui)
        self.label.setGeometry(QtCore.QRect(16, 10, 501, 61))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(28)
        self.label.setFont(font)
        self.label.setStyleSheet("color: rgb(255, 255, 255);")
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(guigui)
        self.label_2.setGeometry(QtCore.QRect(180, 90, 171, 21))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(11)
        font.setItalic(True)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.retranslateUi(guigui)
        QtCore.QMetaObject.connectSlotsByName(guigui)

    def retranslateUi(self, guigui):
        _translate = QtCore.QCoreApplication.translate
        guigui.setWindowTitle(_translate("guigui", "guigui"))
        self.pushButton.setText(_translate("guigui", "Send"))
        self.label.setText(_translate("guigui", "Chatting Application"))
        self.label_2.setText(_translate("guigui", "By Arvendra Chhonkar"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    guigui = QtWidgets.QWidget()
    ui = Ui_guigui()
    ui.setupUi(guigui)
    guigui.show()
    sys.exit(app.exec_())
