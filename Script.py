#Импортируем модуль math. Он необходим для сложных вычислений.
import time
#Импортируем модуль logging. Он необходим для возможности логирования.
import logging
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QTableWidgetItem, QTableWidget
import threading
from data_for_ZVR import data_for_zvr

def loading_assets(parent, filename):
    logging.info("Запущена проверка загрузки активов")
    str_index=0
    assets = pd.read_excel(filename)
    assets["Номер родительского актива"]= np.nan
    assets["Описание"]= np.nan
    EAM = pd.read_csv('Database\EAM607.csv', ";", encoding='windows-1251')
    logging.info("Открыт ЕАМ607")
    for i in range((EAM.shape[0])):
        if EAM["Номер актива"][i] in assets["Номер актива"].tolist():
            assets["Номер родительского актива"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = EAM["Номер родительского актива"][i]
            assets["Описание"][assets["Номер актива"].tolist().index(EAM["Номер актива"][i])] = EAM["Описание актива"][i]
            # Выводим пользователю сообщение
            parent.tableWidget.insertRow(str_index)
            parent.tableWidget.setItem(str_index, 0, QTableWidgetItem(str(assets["Организация "][str_index])))
            parent.tableWidget.setItem(str_index, 1, QTableWidgetItem(str(assets["Номер актива"][str_index])))
            parent.tableWidget.setItem(str_index, 2, QTableWidgetItem(str(assets["Номер родительского актива"][str_index])))
            parent.tableWidget.setItem(str_index, 3,QTableWidgetItem(str(assets["Описание"][str_index])))
            str_index +=1
    logging.info("Сверка с базой закончена")
    assets.to_excel("Database/assets.xlsx")
    logging.info("Создана база годового объёма ремонта")

def date_for_zvr(parent,cfo, project, task, department, basis_operation):
    logging.info("Заполняю ЗВР")
    logging.info("Синхронизируюсь с базой")
    assets = pd.read_excel('Database/assets.xlsx')
    assets["ЦФО"] = np.nan
    assets["Проект"] = np.nan
    assets["Отдел"] = np.nan
    assets["Описание отдела"] = np.nan
    assets["Задача"] = np.nan
    assets["Описание задачи"] = np.nan
    assets["Основание операции"] = np.nan
    assets["Операция с активом"] = np.nan
    assets["Описание операции"] = np.nan
    assets["Месяц выполнения работ"] = np.nan
    assets["Трудоёмкость"] = np.nan
    assets["Номер родительского актива"] = np.nan
    print(1)
    for i in range(parent.tableWidget.rowCount()):
        assets["ЦФО"][i]=cfo
        assets["Проект"]=project
        assets["Отдел"] = department
        assets["Описание отдела"] = np.nan
        assets["Задача"] = task
        assets["Описание задачи"] = np.nan
        assets["Основание операции"] = basis_operation
        print(1)
        parent.tableWidget.setItem(i, 5, QTableWidgetItem(str(assets["ЦФО"][i])))
        parent.tableWidget.setItem(i, 6, QTableWidgetItem(str(assets["Проект"][i])))
        parent.tableWidget.setItem(i, 7, QTableWidgetItem(str(assets["Отдел"][i])))
        parent.tableWidget.setItem(i, 8, QTableWidgetItem(str(assets["Описание отдела"][i])))
        parent.tableWidget.setItem(i, 9, QTableWidgetItem(str(assets["Задача"][i])))
        parent.tableWidget.setItem(i, 10, QTableWidgetItem(str(assets["Описание задачи"][i])))
        parent.tableWidget.setItem(i, 11, QTableWidgetItem(str(assets["Основание операции"][i])))
    assets.to_excel("Database/assets.xlsx")


def creature_zvr(parent):
    logging.info("Создаю ЗВР")
    logging.info("Синхронизируюсь с базой")
    assets = pd.read_excel('Database/assets.xlsx')
    assets["Номер ЗВР"] = np.nan
    assets["Статус ЗВР"] = np.nan
    for i in range(parent.tableWidget.rowCount()):
        try:
            assets["Месяц ремонта"][i]= parent.tableWidget.item(i, 12).text()
        except:
            logging.error("У актива "+str(assets["Номер актива"][i])+"не указана дата ремонта")
            assets["Месяц ремонта"][i]=" "
            continue
        try:
            assets["Операция с активом"][i] = parent.tableWidget.item(i, 14).text()
        except:
            logging.error("У актива " + str(assets["Номер актива"][i]) + "не указана операция с активом")
            assets["Операция с активом"][i]=" "
            continue
        config=filereader("config.config")
        zvrnomber=int(config['zvrnomber'])
        assets["Номер ЗВР"][i] = 'АВ-' + str(int(zvrnomber) + 1)
        configupdate(int(zvrnomber) + 1)
        parent.tableWidget.setItem(i, 16, QTableWidgetItem(assets["Номер ЗВР"][i]))
        assets["Статус ЗВР"][i] = "проект"
        parent.tableWidget.setItem(i, 17, QTableWidgetItem(str(assets["Статус ЗВР"][i])))
        logging.info("Создан ЗВР " + str(assets["Номер ЗВР"][i]))
    logging.info("Создание ЗВР завершено")
    assets.to_excel("Database/assets.xlsx")
    logging.info("База годового объёма ремонта обновлена")

def monthly_plan(parent):
    month=parent.comboBox.currentText()


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


