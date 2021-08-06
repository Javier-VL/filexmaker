from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(800, 600)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pb_OpenExcel = QtWidgets.QPushButton(self.centralwidget)
        self.pb_OpenExcel.setEnabled(True)
        self.pb_OpenExcel.setMinimumSize(QtCore.QSize(70, 70))
        self.pb_OpenExcel.setMaximumSize(QtCore.QSize(90, 90))
        self.pb_OpenExcel.setObjectName("pb_OpenExcel")
        self.horizontalLayout.addWidget(self.pb_OpenExcel)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        self.menuFilex = QtWidgets.QMenu(self.menubar)
        self.menuFilex.setObjectName("menuFilex")
        self.menuReimpresiones = QtWidgets.QMenu(self.menubar)
        self.menuReimpresiones.setObjectName("menuReimpresiones")
        self.menuJobs = QtWidgets.QMenu(self.menubar)
        self.menuJobs.setObjectName("menuJobs")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionCrear_Constancia = QtWidgets.QAction(MainWindow)
        self.actionCrear_Constancia.setObjectName("actionCrear_Constancia")
        self.actionConsultar_Base = QtWidgets.QAction(MainWindow)
        self.actionConsultar_Base.setObjectName("actionConsultar_Base")
        self.actionCrear_Constancia_Jobs = QtWidgets.QAction(MainWindow)
        self.actionCrear_Constancia_Jobs.setObjectName("actionCrear_Constancia_Jobs")
        self.actionCrear_Constancia_Filex = QtWidgets.QAction(MainWindow)
        self.actionCrear_Constancia_Filex.setObjectName("actionCrear_Constancia_Filex")
        self.menuFilex.addAction(self.actionCrear_Constancia_Filex)
        self.menuJobs.addAction(self.actionCrear_Constancia_Jobs)
        self.menubar.addAction(self.menuFilex.menuAction())
        self.menubar.addAction(self.menuReimpresiones.menuAction())
        self.menubar.addAction(self.menuJobs.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


        #
        self.pb_OpenExcel.clicked.connect(importExcel)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pb_OpenExcel.setText(_translate("MainWindow", "Importar Excel"))
        self.menuFilex.setTitle(_translate("MainWindow", "Filex"))
        self.menuReimpresiones.setTitle(_translate("MainWindow", "Reimpresiones"))
        self.menuJobs.setTitle(_translate("MainWindow", "Jobs"))
        self.actionCrear_Constancia.setText(_translate("MainWindow", "Crear Constancia"))
        self.actionConsultar_Base.setText(_translate("MainWindow", "Consultar Base"))
        self.actionCrear_Constancia_Jobs.setText(_translate("MainWindow", "Crear Constancia Jobs"))
        self.actionCrear_Constancia_Jobs.setStatusTip(_translate("MainWindow", "Crear Constancia Jobs"))
        self.actionCrear_Constancia_Jobs.setShortcut(_translate("MainWindow", "Ctrl+J"))
        self.actionCrear_Constancia_Filex.setText(_translate("MainWindow", "Crear Constancia Filex"))
        self.actionCrear_Constancia_Filex.setStatusTip(_translate("MainWindow", "Nueva Constancia Filex"))
        self.actionCrear_Constancia_Filex.setShortcut(_translate("MainWindow", "Ctrl+F"))


#
def importExcel(self):
    print("hehe")

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

