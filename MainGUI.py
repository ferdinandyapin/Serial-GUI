#Import Library
from PyQt5 import QtCore, QtGui, QtWidgets
from serialThread import myThread
from PyQt5.QtCore import pyqtSignal,QTimer
from serial.tools.list_ports import comports
import pyqtgraph as pg
import xlsxwriter
import pandas as pd


class Ui_SimpleModbus(object):
    def setupUi(self, SimpleModbus):
    #Initialize Window
        SimpleModbus.setObjectName("SimpleModbus")
        SimpleModbus.resize(259, 416)
    #Initialize list
        self.msg =[]
        self.slave1 =[]
        self.slave2 =[]
        self.PS=False
    #Setting the content of the window
        self.centralwidget = QtWidgets.QWidget(SimpleModbus)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 20, 231, 51))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(self.gridLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.portcombo = QtWidgets.QComboBox(self.gridLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.portcombo.sizePolicy().hasHeightForWidth())
        self.portcombo.setSizePolicy(sizePolicy)
        self.portcombo.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.portcombo.setObjectName("portcombo")
        self.gridLayout.addWidget(self.portcombo, 0, 1, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.gridLayoutWidget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 1, 0, 1, 1)
        self.Baudcombo = QtWidgets.QComboBox(self.gridLayoutWidget)
        self.Baudcombo.setObjectName("Baudcombo")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.Baudcombo.addItem("")
        self.gridLayout.addWidget(self.Baudcombo, 1, 1, 1, 1)
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(10, 70, 241, 271))
        self.groupBox.setObjectName("groupBox")
        self.listWidget = QtWidgets.QListWidget(self.groupBox)
        self.listWidget.setGeometry(QtCore.QRect(0, 20, 231, 241))
        self.listWidget.setObjectName("listWidget")
        self.layoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.layoutWidget.setGeometry(QtCore.QRect(80, 350, 158, 25))
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.b_connect = QtWidgets.QPushButton(self.layoutWidget)
        self.b_connect.setObjectName("b_connect")
        self.horizontalLayout.addWidget(self.b_connect)
        SimpleModbus.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(SimpleModbus)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 259, 21))
        self.menubar.setObjectName("menubar")
        self.menuData = QtWidgets.QMenu(self.menubar)
        self.menuData.setObjectName("menuData")
        SimpleModbus.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(SimpleModbus)
        self.statusbar.setObjectName("statusbar")
        SimpleModbus.setStatusBar(self.statusbar)
        self.PlotData = QtWidgets.QAction(SimpleModbus)
        self.PlotData.setObjectName("PlotData")
        self.SaveData = QtWidgets.QAction(SimpleModbus)
        self.SaveData.setObjectName("SaveData")
        self.menuData.addAction(self.PlotData)
        self.menuData.addAction(self.SaveData)
        self.menubar.addAction(self.menuData.menuAction())
        self.listWidget.addItem("")
        self.listWidget.addItem("")
    #Initialize timer for plotting
        self.timer = QtCore.QTimer()
        
    #Call class function
        self.setPortCombo()
        self.retranslateUi(SimpleModbus)
        self.connection()
        QtCore.QMetaObject.connectSlotsByName(SimpleModbus)

        


    def retranslateUi(self, SimpleModbus):
    # Label the option and stuff
        _translate = QtCore.QCoreApplication.translate
        SimpleModbus.setWindowTitle(_translate("SimpleModbus", "SimpleModbus"))
        self.label.setText(_translate("SimpleModbus", "Master Port :"))
        self.label_2.setText(_translate("SimpleModbus", "Baud Rate :"))
        self.Baudcombo.setItemText(0, _translate("SimpleModbus", "300 baud"))
        self.Baudcombo.setItemText(1, _translate("SimpleModbus", "1200 baud"))
        self.Baudcombo.setItemText(2, _translate("SimpleModbus", "2400 baud"))
        self.Baudcombo.setItemText(3, _translate("SimpleModbus", "4800 baud"))
        self.Baudcombo.setItemText(4, _translate("SimpleModbus", "9600 baud"))
        self.Baudcombo.setItemText(5, _translate("SimpleModbus", "19200 baud"))
        self.Baudcombo.setItemText(6, _translate("SimpleModbus", "38400 baud"))
        self.Baudcombo.setItemText(7, _translate("SimpleModbus", "57600 baud"))
        self.Baudcombo.setItemText(8, _translate("SimpleModbus", "74880 baud"))
        self.Baudcombo.setItemText(9, _translate("SimpleModbus", "115200 baud"))
        self.Baudcombo.setItemText(10, _translate("SimpleModbus", "230400 baud"))
        self.Baudcombo.setItemText(11, _translate("SimpleModbus", "250000 baud"))
        self.Baudcombo.setItemText(12, _translate("SimpleModbus", "500000 baud"))
        self.Baudcombo.setItemText(13, _translate("SimpleModbus", "1000000 baud"))
        self.Baudcombo.setItemText(14, _translate("SimpleModbus", "2000000 baud"))
        self.groupBox.setTitle(_translate("SimpleModbus", "Data Read"))
        self.b_connect.setText(_translate("SimpleModbus", "Connect"))
        self.menuData.setTitle(_translate("SimpleModbus", "Data"))
        self.PlotData.setText(_translate("SimpleModbus", "Plot Data"))
        self.PlotData.setShortcut(_translate("SimpleModbus", "Ctrl+P"))
        self.SaveData.setText(_translate("SimpleModbus","Save Data"))
        self.SaveData.setShortcut(_translate("SimpleModbus", "Ctrl+S"))

    def connection(self):
    #window event
        self.b_connect.clicked.connect(self.p_conClick)
        self.PlotData.triggered.connect(self.clickPlotData)
        self.SaveData.triggered.connect(self.saveData)
        

    def setPortCombo(self):
    #searching for available port
        a=comports()
        for port in a:
            self.portcombo.addItem(port.device)


    def p_conClick(self):
    #when connect button clicked get user port and baudrate
        port_name=str(self.portcombo.currentText())
        baud_rate=self.Baudcombo.currentText().split()
        self.state = True
        for a in baud_rate:
            if a.isdigit():
                baud_rate=int(a)
    #set threading to read data
        self.Thread=myThread(port_name,baud_rate,self.state)
        self.Thread.msg.connect(self.dispData)
        self.Thread.start()

    def dispData(self,listX):
    #getting slave data from port reading
        self.datList=listX
        self.createData()

    def createData(self):
    #string manipulation to get data
        R_data = list(dict.fromkeys(self.datList))
        R_data = R_data[-1]
        self.data=R_data.strip().split('\t')
        for i,dat in enumerate(self.data):
            if ((len(dat)==1) and (dat==':')):
                del self.data[i]
        for i,dat in enumerate(self.data):
            self.data[i]=self.data[i].split(':')
        for i,dat in enumerate(self.data):
            self.data[i]=dat[:-1]
        for i,dat in enumerate(self.data):
    #display data on window
            if i==0:
                self.slave1.append(int(dat[1]))
                self.slavename1="Slave "+ dat[0]
                s= self.slavename1 + " Data Receive : " + dat[1]
            elif i==1:
                self.slave2.append(int(dat[1]))
                self.slavename2="Slave "+ dat[0]
                s= self.slavename2 + " Data Receive : " + dat[1]
            self.listWidget.item(i).setText(s)

    def clickPlotData(self):
    #when plot data menu triggered, plot the list data
        p=pg.plot()
        p.setWindowTitle('Live Plot from Serial')
        p.setInteractive(True)
        self.curve = p.plot(pen=(255,0,0), name="Slave 1")
        self.curve2 = p.plot(pen=(0,255,0), name="Slave 2")
    
        self.timer.timeout.connect(self.update)
        self.timer.start(0)


    def update(self):
    #updating the plot everytime the data changes
        self.curve.setData(self.slave1)
        self.curve2.setData(self.slave2)
        QtWidgets.QApplication.processEvents()

    def saveData(self):
    #when save data menu triggered, save the data to microsoft excel file
        df = pd.DataFrame({self.slavename1:self.slave1,self.slavename2:self.slave2})
        writer = pd.ExcelWriter('Slave Data.xlsx', engine='xlsxwriter')
        df.to_excel(writer,sheet_name='Data')
        writer.save()

                



        
        
                
if __name__ == "__main__":
    #Main program to start the window
    import sys
    app = QtWidgets.QApplication(sys.argv)
    SimpleModbus = QtWidgets.QMainWindow()
    ui = Ui_SimpleModbus()
    ui.setupUi(SimpleModbus)
    SimpleModbus.show()
    sys.exit(app.exec_())
