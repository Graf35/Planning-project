#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
import Script
#Этот модуль позволяет использовать многопоточность
import threading
import logging




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
        logging.info("Окно месячного планирования объёмов создано")


    def btnClicked1(self):
        # Объявляем новый поток
        self.deman31 = threading.Thread(target=Script.monthly_plan(self))
        # Запускаем новый поток
        self.deman31.start()


    def btnClicked2(self):
        self.deman32 = threading.Thread(target=Script.release_zvr(self))
        # Запускаем новый поток
        self.deman32.start()

