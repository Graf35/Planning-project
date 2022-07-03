#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
from PyQt5.QtWidgets import QFileDialog
from Annual_volume_planning import Annual_volume_planning
from monthly_volume_planning import Monthly_volume_planning

#Этот модуль позволяет использовать многопоточность
import threading

#Определяем имя и путь до файлас формой окна.
ui=uic.loadUiType("interface/Main_window.ui")[0]

#Этот класс определяет параметры окна и взаимодействие с ним.
class MaimWindow(QtWidgets.QMainWindow, ui):
    def __init__(self):
        #Инициализируем окно
        super().__init__()
        self.setupUi(self)
        # Прописываем действие на нажатие кнопки
        self.pushButton.clicked.connect(self.btnClicked1)
        self.pushButton_2.clicked.connect(self.btnClicked2)
        #self.pushButton_3.clicked.connect(self.btnClicked3)

    def btnClicked1(self):
        self.Annual_volume_planning_window = Annual_volume_planning()  # Создаём объект класса`
        # Объявляем новый поток
        self.deman1 = threading.Thread(target=self.Annual_volume_planning_window.show())
        # Запускаем новый поток
        self.deman1.start()

    def btnClicked2(self):
        self.monthly_volume_planning_window = Monthly_volume_planning()  # Создаём объект класса`
        # Объявляем новый поток
        self.deman1 = threading.Thread(target=self.monthly_volume_planning_window.show())
        # Запускаем новый поток
        self.deman1.start()