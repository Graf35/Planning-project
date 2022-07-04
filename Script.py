#Импортируем модуль math. Он необходим для сложных вычислений.
import time
#Импортируем настройки конфигурации записи логирования
#import Log
#Импортируем модуль logging. Он необходим для возможности логирования.
import logging
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QTableWidgetItem, QTableWidget


#Эта функция ожидает и преобразует данные, полученные от пользователя через интерфейс
def dialog(parent):
    #Задаём локальную переменную, которая ожидает ввода данных
    variable = 0
    #Задаём бесконечный цикл проверки ввода
    while variable == 0:
        #Если ввода переменной не произошло то останавливаем поток на 1 сек
        time.sleep(1)
        #присваиваем переменной ожидания текущие значение на табло
        variable = parent.inpat
    #Выводим пользователю сообщение
    parent.label_2.setText("Продолжаю расчёт")
    #Стираем введенное значение из окна
    parent.inpat=0
    return (variable)

def loading_assets(parent, filename):
    str_index=0
    assets = pd.read_excel(filename)
    EAM = pd.read_csv('Database\EAM607.csv', ";", encoding='windows-1251')
    assets["Номер родительского актива"] = np.nan
    for i in range((EAM.shape[0])):
        if EAM["Номер актива"][i] in assets["Номер актива"].tolist():
            assets["Номер родительского актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = EAM["Номер родительского актива"][i]
            # Выводим пользователю сообщение
            parent.tableWidget.insertRow(str_index)
            parent.tableWidget.setItem(str_index, 0, QTableWidgetItem(str(assets["Организация "][str_index])))
            parent.tableWidget.setItem(str_index, 1, QTableWidgetItem(str(assets["Номер актива"][str_index])))
            parent.tableWidget.setItem(str_index, 2, QTableWidgetItem(str(assets["Номер родительского актива"][str_index])))
            str_index +=1
    assets.to_excel("Database/assets.xlsx")


