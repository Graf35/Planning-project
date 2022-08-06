#Эти библиотеки позволяют работать с графикой.
from PyQt5 import QtWidgets
from PyQt5 import  uic
import Script
#Этот модуль позволяет использовать многопоточность
import threading
import logging
import pandas as pd
from datetime import date
from PyQt5.QtWidgets import QTableWidgetItem
import numpy as np




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
        self.deman43 = threading.Thread(target=self.database_read())
        # Запускаем новый поток
        self.deman43.start()

    def btnClicked1(self):
        # Объявляем новый поток
        self.deman41 = threading.Thread(target=self.closed_ZVR())
        # Запускаем новый поток
        self.deman41.start()

    def database_read(self):
        str_index = 0
        month_daes = {"Январь": 31, "Февраль": 28, "Март": 31, "Апрель": 30,
                      "Май": 31, "Июнь": 30, "Июль": 31, "Август": 31,
                      "Сентябрь": 30, "Октябрь": 31, "Ноябрь": 30, "Декабрь": 31}
        month_nom = {"Январь": "01", "Февраль": "02", "Март": "03", "Апрель": "04",
                     "Май": "05", "Июнь": "06", "Июль": "07", "Август": "08",
                     "Сентябрь": "09", "Октябрь": "10", "Ноябрь": "11", "Декабрь": "12"}
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        for i in range((assets.shape[0])):
            if str(assets["Дата окончания"][i]) != "NaT" and str(assets["Дата принятия"][i]) == "NaT" or "NaN" and str(
                    assets["Статус ЗВР"][i]) == "Выпущено":
                self.tableWidget.insertRow(str_index)
                self.tableWidget.setItem(str_index, 0, QTableWidgetItem(assets["Номер ЗВР"][i]))
                self.tableWidget.setItem(str_index, 1, QTableWidgetItem(assets["Номер родительского актива"][i]))
                self.tableWidget.setItem(str_index, 2, QTableWidgetItem(assets["Номер актива"][i]))
                self.tableWidget.setItem(str_index, 3, QTableWidgetItem(assets["Операция с активом"][i]))
                self.tableWidget.setItem(str_index, 4, QTableWidgetItem(assets["Статус ЗВР"][i]))
                self.tableWidget.setItem(str_index, 5,
                                           QTableWidgetItem(
                                               "01-" + str(month_nom[assets["Месяц ремонта"][i]]) + "-2023"))
                self.tableWidget.setItem(str_index, 6, QTableWidgetItem(
                    str(month_daes[assets["Месяц ремонта"][i]]) + "-" + str(
                        month_nom[assets["Месяц ремонта"][i]]) + "-2023"))
                self.tableWidget.setItem(str_index, 7, QTableWidgetItem(str(assets["Дата начала"][i])))
                self.tableWidget.setItem(str_index, 8, QTableWidgetItem(str(assets["Дата начала"][i])))
                self.tableWidget.setItem(str_index, 9, QTableWidgetItem(str(assets["Дата окончания"][i])))
                self.tableWidget.setItem(str_index, 10, QTableWidgetItem(str(assets["Дата принятия"][i])))

    def closed_ZVR(self):
        logging.info("Синхронизируюсь с базой")
        assets = pd.read_excel('Database/assets.xlsx')
        for i in range(self.tableWidget.rowCount()):
            try:
                index = assets[assets["Номер актива"] == self.tableWidget.item(i, 2).text()].index[0]
            except:
                continue
            if self.tableWidget.item(i, 10).text() == "ДА" or self.tableWidget.item(i,
                                                                                        10).text() == "Да" or self.tableWidget.item(
                    i, 10).text() == "дА" or self.tableWidget.item(i, 10).text() == "да":
                assets["Статус ЗВР"][index] = "Завершено"
                assets["Дата принятия"][index] = date.today()
            elif self.tableWidget.item(i, 10).text() == "НЕТ" or self.tableWidget.item(i,
                                                                                           10).text() == "Нет" or self.tableWidget.item(
                    i, 10).text() == "нет":
                assets["Дата окончания"][index] = np.nan
            else:
                logging.error(
                    "У актива " + str(self.tableWidget.item(i, 2).text()) + " не принято решение о завершении")
        Script.base_update(assets)