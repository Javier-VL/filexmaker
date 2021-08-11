


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_SecondWindow(object):
    def setupUi(self, SecondWindow):
        SecondWindow.setObjectName("SecondWindow")
        SecondWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(SecondWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.le_ExcelDirectory = QtWidgets.QLineEdit(self.centralwidget)
        self.le_ExcelDirectory.setGeometry(QtCore.QRect(20, 30, 221, 21))
        self.le_ExcelDirectory.setObjectName("le_ExcelDirectory")
        self.pb_BrowseFile = QtWidgets.QPushButton(self.centralwidget)
        self.pb_BrowseFile.setGeometry(QtCore.QRect(260, 30, 75, 23))
        self.pb_BrowseFile.setObjectName("pb_BrowseFile")
        SecondWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(SecondWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        SecondWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(SecondWindow)
        self.statusbar.setObjectName("statusbar")
        SecondWindow.setStatusBar(self.statusbar)

        self.retranslateUi(SecondWindow)
        QtCore.QMetaObject.connectSlotsByName(SecondWindow)

        #
        #BROWSE"
        self.pb_BrowseFile.clicked.connect(self.browseFile)

    def browseFile(self):
        fname = QtWidgets.QFileDialog.getExistingDirectory(None,"abre")
        #fname = QtWidgets.QFileDialog.getOpenFileName(None,'Open File', '.','(*.xlsx)')

        self.le_ExcelDirectory.setText(fname[0])#POSITION CERO PARA ARCHIVOS
        print(fname)

    def retranslateUi(self, SecondWindow):  
        _translate = QtCore.QCoreApplication.translate
        SecondWindow.setWindowTitle(_translate("SecondWindow", "MainWindow"))
        self.pb_BrowseFile.setText(_translate("SecondWindow", "PushButton"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    SecondWindow = QtWidgets.QMainWindow()
    ui = Ui_SecondWindow()
    ui.setupUi(SecondWindow)
    SecondWindow.show()
    sys.exit(app.exec_())
