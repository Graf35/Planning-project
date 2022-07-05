#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog
import Script
import shutil
#Этот модуль позволяет использовать многопоточность
import threading



#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Monthly_volume_planning.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class Monthly_volume_planning(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.btnClicked1)
        self.pushButton_2.clicked.connect(self.btnClicked2)
        self.pushButton_3.clicked.connect(self.btnClicked3)

    def btnClicked1(self):
        # Объявляем новый поток
        self.demanm1 = threading.Thread(target=Script.monthly_plan(self))
        # Запускаем новый поток
        self.deman1.start()


    def btnClicked2(self):
        pass

    def btnClicked3(self):
        pass