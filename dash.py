import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QLabel
from PyQt5.QtGui import QColor
from openpyxl import load_workbook

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1920, 1200)
        MainWindow.setStyleSheet("#main {\n"
"  background-color: white;\n"
"}\n"
"\n"
"#left {\n"
"  background-image: url(\'C:/coding/bg.jpg\');\n"
"  /* You may want to add more background properties like background-size, background-repeat, etc. for better styling. */\n"
"\n"
"}\n"
"\n"
"#table {\n"
"  border: 1px solid black;\n"
"}\n"
"\n"
"#geninfo {\n"
"  background-color: transparent;\n"
"  /* You can add more styles based on your design requirements. */\n"
"}\n"
"#main {\n"
"  background-color: white;\n"
"}\n"
"\n"
"#left {\n"
"  background-image: url(\'C:/coding/bg.jpg\');\n"
"  /* You may want to add more background properties like background-size, background-repeat, etc. for better styling. */\n"
"\n"
"}\n"
"\n"
"#table {\n"
"  border: 1px solid black;\n"
"}\n"
"\n"
"#geninfo {\n"
"  background-color: transparent;\n"
"  font-family: \"Arial\", sans-serif;\n"
"  color: #fff; /* Set the color according to your design */\n"
"  font-size: 30px;\n"
"\n"
"\n"
" \n"
"}\n"
"#Dash{\n"
"  background-color: transparent;\n"
"  font-family: \"Arial\", sans-serif;\n"
"  color: #333; /* Set the color according to your design */\n"
"  font-size: 50px;\n"
"\n"
"\n"
"}\n"
"#actual,#annual{\n"
"  background-color: transparent;\n"
"  font-family: \"Arial\", sans-serif;\n"
"  color: #fff; /* Set the color according to your design */\n"
"  font-size: 30px;\n"
"  text-align: left;\n"
"\n"
"}\n"
"\n"
"#actual:hover {\n"
"  background-color: gray;\n"
"}\n"
"\n"
"#annual:hover {\n"
"  background-color: gray;\n"
"}\n"
"\n"
"\n"
"")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        self.main = QtWidgets.QWidget(self.centralwidget)
        self.main.setGeometry(QtCore.QRect(10, 10, 1920, 1080))
        self.main.setObjectName("main")

        
        self.left = QtWidgets.QWidget(self.main)
        self.left.setGeometry(QtCore.QRect(0, 0, 384, 1080))
        self.left.setObjectName("left")
        
        self.geninfo = QtWidgets.QComboBox(self.left)
        self.geninfo.setGeometry(QtCore.QRect(3, 200, 384, 41))
        self.geninfo.setObjectName("geninfo")
        self.geninfo.addItem("")
        self.geninfo.addItem("")
        self.geninfo.addItem("")
        
        self.Dash = QtWidgets.QPushButton(self.left)
        self.Dash.setGeometry(QtCore.QRect(0, 27, 384, 51))
        self.Dash.setObjectName("Dash")
        
        self.actual = QtWidgets.QPushButton(self.left)
        self.actual.setGeometry(QtCore.QRect(3, 300, 384, 41))
        self.actual.setObjectName("actual")
        
        self.annual = QtWidgets.QPushButton(self.left)
        self.annual.setGeometry(QtCore.QRect(3, 400, 384, 41))
        self.annual.setObjectName("annual")
        
        self.table = QtWidgets.QTableView(self.main)
        self.table.setGeometry(QtCore.QRect(450, 161, 1380, 800))
        self.table.setObjectName("table")






        
        
        self.pro = QtWidgets.QComboBox(self.main)
        self.pro.setGeometry(QtCore.QRect(450, 90, 201, 41))
        self.pro.setObjectName("pro")
        self.pro.addItem("North")
        self.pro.addItem("Western")
        self.pro.addItem("Eastern")
        self.pro.addItem("")
        self.pro.currentIndexChanged.connect(self.pro_changed)
        
        self.tla = QtWidgets.QComboBox(self.main)
        self.tla.setGeometry(QtCore.QRect(700, 90, 201, 41))
        self.tla.setObjectName("tla")
        self.tla.addItem("")
        

        
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1561, 26))
        self.menubar.setObjectName("menubar")
        
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.geninfo.setItemText(0, _translate("MainWindow", "Gentral Info A"))
        self.geninfo.setItemText(1, _translate("MainWindow", "Gentral Info B"))
        self.geninfo.setItemText(2, _translate("MainWindow", "Gentral Info C"))
        self.tla.setItemText(0, _translate("MainWindow", "DISTRICT"))

        
        self.Dash.setText(_translate("MainWindow", "Dash Board"))
        self.actual.setText(_translate("MainWindow", "Actual"))
        self.annual.setText(_translate("MainWindow", "Annual"))








    def pro_changed(self):
        self.tla.clear()

        selected_pro = self.pro.currentText()
        if selected_pro == "North":
            self.tla.addItems(["Jaffna", "Killinochi","Mannar","Vavuniya","Mulaithiwu"])
        elif selected_pro == "Western":
            self.tla.addItems(["Colombo", "Kaluthura","Galle"])


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
