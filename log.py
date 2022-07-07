#Импортируем модуль logging. Он необходим для работы функций логирования.
import logging

#Создаём класс демона логирования
class Deman_log(object):
    def __init__(self):
        #Создаём файл журнала логирования
        logfile = open('log.txt', "w")
        logfile.close()
        #Прописываем формат вывода сообщения в журнал
        logging.basicConfig(format = u'%(filename)s %(funcName)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s', level = logging.DEBUG, filename="log.txt")
        #Тестовая запись в журнал
        logging.info("Деман лога призван!")