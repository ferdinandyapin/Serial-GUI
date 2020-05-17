from PyQt5.QtCore import Qt, QThread, pyqtSignal
import serial
import time

class myThread(QThread):
    def __init__(self, port, baud,state):
    #getiing parameter to serial read
        super().__init__()
        self.port=port
        self.baud=baud
        self.state=state
        self.ser=serial.Serial(self.port,self.baud)
    msg=pyqtSignal(list)
    def run(self):
    #run threading to serial read the port and emit the list
        a=[]
        while True:
            try:
                text=str(self.ser.readline().decode())
            except:
                continue
            a.append(text)
            self.msg.emit(a)


