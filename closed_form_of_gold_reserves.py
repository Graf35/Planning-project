#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
import Script
#Этот модуль позволяет использовать многопоточность
import threading
import logging




#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Closed_form_of_gold_reserves.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class Closed_form_of_gold_reserves(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.btnClicked1)
        logging.info("Окно подтверждения закрытия ЗВР")
        # Объявляем новый поток
        self.deman43 = threading.Thread(target=Script.database_read(self))
        # Запускаем новый поток
        self.deman43.start()


    def btnClicked1(self):
        # Объявляем новый поток
        self.deman41 = threading.Thread(target=Script.closed_ZVR(self))
        # Запускаем новый поток
        self.deman41.start()


