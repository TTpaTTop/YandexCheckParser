#!python3.7

# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import pyqtSignal, QStringListModel, pyqtSlot
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import time
import email
import email.utils
import base64
import imaplib
import re
import os
import threading
import configparser
import time
import sys

class Ui_MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

    text_signal = pyqtSignal(str)

    def setupUi(self, MainWindow):
        #Основное окно
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(600, 230)
        MainWindow.setDocumentMode(False)

        #Центральный виджет
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        #Основной заголовок (Название)
        font = QtGui.QFont()
        font.setFamily("Courier New")
        font.setPointSize(28)

        self.labelTitleMain = QtWidgets.QLabel(self.centralwidget)
        self.labelTitleMain.setGeometry(QtCore.QRect(60, 0, 481, 40))
        self.labelTitleMain.setFont(font)
        self.labelTitleMain.setAlignment(QtCore.Qt.AlignCenter)
        self.labelTitleMain.setObjectName("labelTitleMain")

        #Текст версии
        font.setPointSize(8)
        self.labelVersion = QtWidgets.QLabel(self.centralwidget)
        self.labelVersion.setGeometry(QtCore.QRect(503, 212, 120, 20))
        self.labelVersion.setFont(font)
        self.labelVersion.setObjectName("labelVersion")

        #Подзаголовок (Разработчики)
        font.setPointSize(9)
        font.setUnderline(True)
        self.labelTitleAuthors = QtWidgets.QLabel(self.centralwidget)
        self.labelTitleAuthors.setGeometry(QtCore.QRect(290, 30, 150, 20))
        self.labelTitleAuthors.setFont(font)
        self.labelTitleAuthors.setAlignment(QtCore.Qt.AlignHCenter|QtCore.Qt.AlignTop)
        self.labelTitleAuthors.setObjectName("labelTitleAuthors")

        #(Условно) Рабочее пространство программы (groupBox)
        self.groupBoxWorkSpace = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBoxWorkSpace.setGeometry(QtCore.QRect(10, 50, 580, 165))
        self.groupBoxWorkSpace.setAutoFillBackground(False)
        self.groupBoxWorkSpace.setTitle("")
        self.groupBoxWorkSpace.setFlat(False)
        self.groupBoxWorkSpace.setCheckable(False)
        self.groupBoxWorkSpace.setObjectName("groupBoxWorkSpace")

        #Индикатор загрузки
        self.progressBarWork = QtWidgets.QProgressBar(self.groupBoxWorkSpace)
        self.progressBarWork.setGeometry(QtCore.QRect(110, 135, 381, 20))
        self.progressBarWork.setProperty("value", 0)
        self.progressBarWork.setInvertedAppearance(False)
        self.progressBarWork.setObjectName("progressBarWork")
        self.progressBarWork.setMaximum(0)

        #Журнал работы программы
        self.listViewWorkLogs = QtWidgets.QListView(self.groupBoxWorkSpace)
        self.listViewWorkLogs.setGeometry(QtCore.QRect(330, 19, 240, 100))
        self.listViewWorkLogs.setObjectName("listViewWorkLogs")
        self.listViewWorkLogs.wordWrap()
        # self.listViewWorkLogs.setReadOnly(True)

        # Настройка модели для QlistView
        self.model = QStringListModel()
        self.listViewWorkLogs.setModel(self.model)

        # Подключите сигнал с методом обновления
        self.text_signal.connect(self.update_list)

        #(Условно) Рабочее пространство селекторов CheckBox
        self.groupBoxCheckBoxesSpace = QtWidgets.QGroupBox(self.groupBoxWorkSpace)
        self.groupBoxCheckBoxesSpace.setEnabled(True)
        self.groupBoxCheckBoxesSpace.setGeometry(QtCore.QRect(170, 10, 150, 110))
        self.groupBoxCheckBoxesSpace.setTitle("")
        self.groupBoxCheckBoxesSpace.setObjectName("groupBoxCheckBoxesSpace")

        #Селектор включения записи логов в файл
        self.checkBoxWriteLogFile = QtWidgets.QCheckBox(self.groupBoxCheckBoxesSpace)
        self.checkBoxWriteLogFile.setGeometry(QtCore.QRect(10, 70, 15, 15))
        self.checkBoxWriteLogFile.setText("")
        self.checkBoxWriteLogFile.setChecked(False)
        self.checkBoxWriteLogFile.setObjectName("checkBoxWriteLogFile")
        self.checkBoxWriteLogFile.setEnabled(False)

        #Надпись checkBoxWriteLogFile
        self.labelWriteLogFile = QtWidgets.QLabel(self.groupBoxCheckBoxesSpace)
        self.labelWriteLogFile.setGeometry(QtCore.QRect(30, 67, 110, 20))
        self.labelWriteLogFile.setWordWrap(True)
        self.labelWriteLogFile.setObjectName("labelWriteLogFile")

        #Название groupBoxWorkSpace
        self.labelOptions = QtWidgets.QLabel(self.groupBoxCheckBoxesSpace)
        self.labelOptions.setGeometry(QtCore.QRect(55, 0, 40, 15))
        self.labelOptions.setObjectName("labelOptions")

        #Селектор открытия папки с файлами после работы
        self.checkBoxOpenExplorer = QtWidgets.QCheckBox(self.groupBoxCheckBoxesSpace)
        self.checkBoxOpenExplorer.setGeometry(QtCore.QRect(10, 30, 20, 20))
        self.checkBoxOpenExplorer.setText("")
        self.checkBoxOpenExplorer.setObjectName("checkBoxOpenExplorer")

        #Надпись чекбокса checkBoxOpenExplorer
        self.labelOpenExplorer = QtWidgets.QLabel(self.groupBoxCheckBoxesSpace)
        self.labelOpenExplorer.setGeometry(QtCore.QRect(30, 23, 110, 30))
        self.labelOpenExplorer.setWordWrap(True)
        self.labelOpenExplorer.setObjectName("labelOpenExplorer")
        
        #Надпись рабочего пространства listViewWorkLogs
        self.labelLogStatus = QtWidgets.QLabel(self.groupBoxWorkSpace)
        self.labelLogStatus.setGeometry(QtCore.QRect(430, 0, 47, 20))
        self.labelLogStatus.setAlignment(QtCore.Qt.AlignCenter)
        self.labelLogStatus.setObjectName("labelLogStatus")

        #(Условно) Рабочее пространство параметров поиска (DateEdit)
        self.groupBoxSearchSpace = QtWidgets.QGroupBox(self.groupBoxWorkSpace)
        self.groupBoxSearchSpace.setEnabled(True)
        self.groupBoxSearchSpace.setGeometry(QtCore.QRect(10, 10, 150, 110))
        self.groupBoxSearchSpace.setTitle("")
        self.groupBoxSearchSpace.setObjectName("groupBoxSearchSpace")

        #Название рабочего пространства параметров поиска
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(12)
        self.labelTitleDateEditTitle = QtWidgets.QLabel(self.groupBoxSearchSpace)
        self.labelTitleDateEditTitle.setGeometry(QtCore.QRect(0, 0, 151, 21))
        self.labelTitleDateEditTitle.setFont(font)
        self.labelTitleDateEditTitle.setAlignment(QtCore.Qt.AlignCenter)
        self.labelTitleDateEditTitle.setObjectName("labelTitleDateEditTitle")

        #Текстовый указатель DateEditEnd
        self.labelTitleDateEditEnd = QtWidgets.QLabel(self.groupBoxSearchSpace)
        self.labelTitleDateEditEnd.setGeometry(QtCore.QRect(0, 50, 31, 21))
        self.labelTitleDateEditEnd.setFont(font)
        self.labelTitleDateEditEnd.setAlignment(QtCore.Qt.AlignCenter)
        self.labelTitleDateEditEnd.setObjectName("labelTitleDateEditEnd")

        #Кнопка старта поиска
        self.pushButtonSearch = QtWidgets.QPushButton(self.groupBoxSearchSpace)
        self.pushButtonSearch.setGeometry(QtCore.QRect(10, 80, 131, 21))
        self.pushButtonSearch.setFlat(False)
        self.pushButtonSearch.setObjectName("pushButtonSearch")
        self.pushButtonSearch.clicked.connect(self.onClick)

        #Текстовый указатель DateEditStart
        self.labelTitleDateEditStart = QtWidgets.QLabel(self.groupBoxSearchSpace)
        self.labelTitleDateEditStart.setGeometry(QtCore.QRect(10, 20, 16, 21))
        self.labelTitleDateEditStart.setFont(font)
        self.labelTitleDateEditStart.setAlignment(QtCore.Qt.AlignCenter)
        self.labelTitleDateEditStart.setObjectName("labelTitleDateEditStart")

        #Селектор dateEdit начала даты поиска
        self.dateEditStartDate = QtWidgets.QDateEdit(self.groupBoxSearchSpace)
        self.dateEditStartDate.setGeometry(QtCore.QRect(30, 20, 110, 22))
        self.dateEditStartDate.setObjectName("dateEditStartDate")
        self.dateEditStartDate.setDateTime(QtCore.QDateTime(QtCore.QDate(2024, 1, 1)))

        #Селектор dateEdit конца даты поиска
        self.dateEditEndDate = QtWidgets.QDateEdit(self.groupBoxSearchSpace)
        self.dateEditEndDate.setGeometry(QtCore.QRect(30, 50, 110, 22))
        self.dateEditEndDate.setObjectName("dateEditEndDate")
        self.dateEditEndDate.setDateTime(QtCore.QDateTime(QtCore.QDate(2024, 1, 1)))
        #Кнопка вызова справки
        self.pushButtonHelp = QtWidgets.QPushButton(self.centralwidget)
        self.pushButtonHelp.setGeometry(QtCore.QRect(544, 0, 61, 21))
        self.pushButtonHelp.setCursor(QtGui.QCursor(QtCore.Qt.WhatsThisCursor))
        self.pushButtonHelp.setFlat(True)
        self.pushButtonHelp.setObjectName("pushButtonHelp")

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def update_list(self, text):
        # Получите текущий список строк
        current_data = self.model.stringList()
        # Добавить новый текст в список
        current_data.append(text)
        # Обновите модель с новым списком
        self.model.setStringList(current_data)
        self.listViewWorkLogs.scrollToBottom()

    def onClick(self):
        self.text_signal.emit(f'Инициализация. Возможно краткое зависание')
        thread = Work(self.text_signal)
        thread = threading.Thread(target = self.onClickWork)
        thread.start()

    def onClickWork(self):
        self.text_signal.emit(f'Поиск чеков...')
        checkBoxOpenExplorer = self.checkBoxOpenExplorer.isChecked()
        checkBoxWriteLogFile = self.checkBoxWriteLogFile.isChecked()
        work = Work(self.text_signal)
        work.progress_signal.connect(self.update_progress)
        self.progressBarWork.setMaximum(100)
        gg = work.buttonPressed(self.dateEditStartDate, self.dateEditEndDate, checkBoxOpenExplorer, checkBoxWriteLogFile)

    @pyqtSlot(int)
    def update_progress(self, progress):
        self.progressBarWork.setValue(progress)
        print(f'Текущее кол-во процентов: {self.progressBarWork.value()}')

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MailReader"))
        self.labelTitleMain.setText(_translate("MainWindow", "MailReader"))
        self.labelVersion.setText(_translate("MainWindow", "Ver. 2.0 (Fix)"))
        self.labelTitleAuthors.setText(_translate("MainWindow", "by Cal1pso and TR3M0R"))
        self.progressBarWork.setFormat(_translate("MainWindow", "%p%"))
        self.labelWriteLogFile.setText(_translate("MainWindow", "Запись логов в файл"))
        self.labelOptions.setText(_translate("MainWindow", "Опции:"))
        self.labelOpenExplorer.setText(_translate("MainWindow", "Открыть папку по завершению работы"))
        self.labelLogStatus.setText(_translate("MainWindow", "Статус:"))
        self.labelTitleDateEditTitle.setText(_translate("MainWindow", "Интервал поиска"))
        self.labelTitleDateEditEnd.setText(_translate("MainWindow", "По:"))
        self.pushButtonSearch.setText(_translate("MainWindow", "Поиск"))
        self.labelTitleDateEditStart.setText(_translate("MainWindow", "С:"))
        self.pushButtonHelp.setText(_translate("MainWindow", "Справка"))


class UiFunc:
    def dateEditFormat(self, dateEditStartDate, dateEditEndDate):
        dateEditStartFormat = dateEditStartDate.date().toPyDate().strftime('%d-%b-%Y')
        dateEditEndFormat = dateEditEndDate.date().toPyDate().strftime('%d-%b-%Y')
        return[dateEditStartFormat, dateEditEndFormat]


class Work(UiFunc, Ui_MainWindow):
    progress_signal = pyqtSignal(int)
    def __init__(self, signal):
        super().__init__()
        self.text_signal = signal
        config = configparser.ConfigParser()
        config.read('config.ini')
     # Подключение к почтовому серверу по протоколу IMAP
        imap_server = "imap.mail.ru" #IMAP-сервер
        self.username = config['Authentication']['Mail'] #Почта для парсинга
        self.password = config['Authentication']['Password'] #Пароль от почты (Пароль приложения)
        self.imap = imaplib.IMAP4_SSL(imap_server)
        self.imap.login(self.username, self.password)
        self.imap.select("INBOX")
        # Подключение к образу GoogleChrome
        self.options = Options()
        self.options.add_argument("--headless")
        self.options.add_experimental_option(
            "excludeSwitches", 
            ["enable-automation"]
        )
        self.options.add_experimental_option('useAutomationExtension', False)
        self.options.add_argument("--disable-blink-features=AutomationControlled")

        self.driver = webdriver.Chrome(options=self.options)
        self.driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument", 
            {
                'source': 
                    '''
                        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
                        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
                        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
                    '''
            }
        )


    def searchMessageLink(self, dateEditFormat):
        print(self.text_signal)
        MessageArrayID = re.findall('([-+]?\d+)', str(self.imap.search(None, f'(SINCE "{dateEditFormat[0]}" BEFORE "{dateEditFormat[1]}")')))
        MessageArrayLink = []
        MessageArrayDate = []
        MessageArrayError = []
        for i in MessageArrayID:
            _, msg = self.imap.fetch(str(i), '(RFC822)')
            msg = email.message_from_bytes(msg[0][1])
            MessageDate = email.utils.parsedate_to_datetime(msg["Date"])
            for part in msg.walk():
                if part.get_content_maintype() == 'text' and part.get_content_subtype() == 'html':
                    content = base64.b64decode(part.get_payload()).decode()
                    content = str(content)
                    findStart = content.find("https://receipts-renderer" )
                    findFinish = content.find("mode=mobile")
                    if len(str(content[findStart:findFinish + 11])) == 100:
                        MessageArrayLink.append(str(content[findStart:findFinish + 11]))
                        MessageArrayDate.append(str(MessageDate))
                    else:
                        MessageArrayError.append(f'Ошибка в письме #{i}, дата и время письма: {str(MessageDate)}')
                        print(f'Ошибка в письме #{i}, дата и время письма: {str(MessageDate)}')
                        self.text_signal.emit(f'Ошибка в письме #{i}')
                        self.text_signal.emit(f'Дата и время письма: {str(MessageDate)}')
        return[MessageArrayLink, MessageArrayDate, MessageArrayError, self.progress_signal.emit(100), self.text_signal.emit(f'В заданном промужутке найдено:\nВсего {len(MessageArrayLink) + len(MessageArrayError)} чеков\n{len(MessageArrayError)} с ошибками')]

                

    def getMessageCheck(self, MessageArrayLink, MessageArrayDate, OpenExplorerDir = False):
        time.sleep(0.8)
        print(len(MessageArrayDate))
        os.makedirs(r'ScreenShots', exist_ok=True)
        for i in range(len(MessageArrayLink)):
            print(i)
            progress_step = len(MessageArrayLink)
            link = MessageArrayLink[i]
            datePath = MessageArrayDate[i].replace('-', '.')
            datePath = datePath.replace(':', '_')
            datePath = datePath.replace('+03-00', '')
            self.driver.get(link)
            s = lambda x: self.driver.execute_script(
                'return document.body.parentNode.scroll' + x
            )
            self.driver.set_window_size(s('Width'), s('Height'))
            print(f'{datePath}.png')
            self.driver.find_element(By.TAG_NAME, 'body').screenshot(f'ScreenShots/{datePath}.png')
            self.text_signal.emit(f'Чек №{i + 1} сохранен')
            progress = int((i + 1) / progress_step * 100)
            self.progress_signal.emit(progress)
        if OpenExplorerDir == True:
            return(os.system('start ScreenShots'),self.text_signal.emit(f'Сохранение завершено\nПапка открыта'))
        else: 
            return(print("Работа завершена!"), self.text_signal.emit(f'Сохранение завершено'), self.progress_signal.emit(100))
            


    def buttonPressed(self, dateEditStartDate, dateEditEndDate, checkBoxExplorer = False, checkBoxWriteLogFile = False):
        dateEditFormat = self.dateEditFormat(dateEditStartDate, dateEditEndDate)
        searchMessage = self.searchMessageLink(dateEditFormat)
        getMessageCheck = self.getMessageCheck(searchMessage[0], searchMessage[1], checkBoxExplorer)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = Ui_MainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())