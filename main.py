# Импортируем модуль Window. Он необходим для создания графического окна
import Main_window
#Импортируем PyQt5 для работы с графикой
from PyQt5 import QtWidgets
#Модуль sys обеспечивает доступ к некоторым переменным и функциям, взаимодействующим с интерпретатором python.
import sys

def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = Main_window.MaimWindow()  # Создаём объект класса
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()