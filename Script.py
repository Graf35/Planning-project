#Импортируем модуль math. Он необходим для сложных вычислений.
import time
#Импортируем настройки конфигурации записи логирования
#import Log
#Импортируем модуль logging. Он необходим для возможности логирования.
import logging
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QTableWidgetItem, QTableWidget

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

def creature_zvr(parent):
    assets = pd.read_excel('Database/assets.xlsx')
    assets[["Месяц ремонта", "Операция с активом", "Номер ЗВР", "Статус ЗВР"]] = np.nan
    for i in range(parent.tableWidget.rowCount()):
        assets["Месяц ремонта"][i]= parent.tableWidget.item(i, 3).text()
        assets["Операция с активом"][i] = parent.tableWidget.item(i, 4).text()
        config=filereader("config.config")
        zvrnomber=int(config['zvrnomber'])
        assets["Номер ЗВР"][i]=int(zvrnomber+1)
        configupdate(int(assets["Номер ЗВР"][i]))
        parent.tableWidget.setItem(i, 5, QTableWidgetItem(str(int(assets["Номер ЗВР"][i]))))
        assets["Статус ЗВР"][i] = "Проект"
        parent.tableWidget.setItem(i, 6, QTableWidgetItem(str(assets["Статус ЗВР"][i])))
    assets.to_excel("Database/assets.xlsx")


def filereader(fail_name):
    sting=[]
    file=open(fail_name, 'r')
    for line in file:
        sting.append(line.strip().split(";"))
    file.close()
    filetext={}
    for i in range(len(sting)):
        filetext[sting[i][0]]=(sting[i][1])
    return (filetext)

def configupdate(zvr):
    lines = ["zvrnomber;" + str(zvr) + "\n"]
    with open("config.config", "w") as file:
        for line in lines:
            if line != "":
                file.write(line)


