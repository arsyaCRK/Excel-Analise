import sys
import os
from pathlib import Path
import openpyxl.utils.exceptions
import openpyxl_dictreader
from PyQt6 import QtCore, QtWidgets
from PyQt6.QtWidgets import QMessageBox
from openpyxl import Workbook


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(351, 414)
        MainWindow.setFixedSize(351, 414)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.widget = QtWidgets.QWidget(self.centralwidget)
        self.widget.setGeometry(QtCore.QRect(0, 3, 351, 411))
        self.widget.setObjectName("widget")
        self.gridLayout = QtWidgets.QGridLayout(self.widget)
        self.gridLayout.setContentsMargins(7, 0, 7, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.widget)
        self.lineEdit.setObjectName("lineEdit")
        self.gridLayout.addWidget(self.lineEdit, 1, 0, 1, 1)
        self.label_5 = QtWidgets.QLabel(self.widget)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 9, 0, 1, 1)
        self.pushButton = QtWidgets.QPushButton(self.widget)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 11, 0, 1, 1)
        self.v_filter = QtWidgets.QLineEdit(self.widget)
        self.v_filter.setObjectName("v_filter")
        self.v_filter.setPlaceholderText('Введите значение фильтра...')
        self.gridLayout.addWidget(self.v_filter, 6, 0, 1, 1)
        self.column_name = QtWidgets.QLineEdit(self.widget)
        self.column_name.setObjectName("column_name")
        self.column_name.setPlaceholderText('Введите название столбца...')
        self.gridLayout.addWidget(self.column_name, 4, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.widget)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 5, 0, 1, 1)
        self.label = QtWidgets.QLabel(self.widget)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.date_name = QtWidgets.QLineEdit(self.widget)
        self.date_name.setObjectName("date_name")
        self.date_name.setPlaceholderText('Введите значение столбца дата...')
        self.gridLayout.addWidget(self.date_name, 8, 0, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.widget)
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 7, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.widget)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 3, 0, 1, 1)
        self.f_date = QtWidgets.QLineEdit(self.widget)
        self.f_date.setObjectName("f_date")
        self.f_date.setPlaceholderText('Введите значение даты...')
        self.gridLayout.addWidget(self.f_date, 10, 0, 1, 1)
        self.toolButton = QtWidgets.QToolButton(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Policy.Preferred, QtWidgets.QSizePolicy.Policy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.toolButton.sizePolicy().hasHeightForWidth())
        self.toolButton.setSizePolicy(sizePolicy)
        self.toolButton.setObjectName("toolButton")
        self.gridLayout.addWidget(self.toolButton, 2, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.add_functions()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Для Алёны v1.0"))
        self.label.setText(_translate("MainWindow", "Выберите файл:"))
        self.toolButton.setText(_translate("MainWindow", "..."))
        self.label_2.setText(_translate("MainWindow", "Введите название столбца для фильра:"))
        self.label_3.setText(_translate("MainWindow", "Введите значение фильра:"))
        self.label_4.setText(_translate("MainWindow", "Название столбца дата:"))
        self.label_5.setText(_translate("MainWindow", "Введите дату:"))
        self.pushButton.setText(_translate("MainWindow", "Выполнить"))

    def add_functions(self):
        self.pushButton.clicked.connect(lambda: self.execute_filter(self.column_name.text(), self.v_filter.text(),
                                                                    self.date_name.text(), self.f_date.text()))
        self.toolButton.clicked.connect(lambda: self.showFileDialog())

    def execute_filter(self, column_name, filter_value, col_date_name, f_date):
        wb = Workbook()
        ws = wb.active
        file_path = os.path.abspath(self.lineEdit.text())
        try:
            reader = openpyxl_dictreader.DictReader(file_path)
            if column_name != '' and filter_value == '':
                self.showErrorMessage('Вы не заполнили поле фильтр!')
                wb.remove(wb.active)
                wb.close()
            else:
                for row in reader:
                    if f_date != '':
                        date_value = row[col_date_name].find(f_date)
                        if date_value >= 0:
                            if row[column_name] == filter_value:
                                cells = tuple(row.values())
                                ws.append(cells)
                    else:
                        if row[column_name] == filter_value:
                            cells = tuple(row.values())
                            ws.append(cells)
                filename = self.showSaveFileDialog(f'{filter_value}')
                if filename != '':
                    for column_cells in ws.columns:
                        length = max(len(str(cell.value)) for cell in column_cells)
                        ws.column_dimensions[column_cells[0].column_letter].width = length * 1.3
                    wb.save(filename)
                    wb.remove(wb.active)
                    wb.close()
                    self.showInfoMessage('Файл успешно сохранён!')
                else:
                    wb.remove(wb.active)
                    wb.close()
        except KeyError:
            wb.remove(wb.active)
            wb.close()
            self.showErrorMessage('Вы не верно ввели названия столбцов \n"Дата" или "Фильтр". \nПопробуйте снова.')
        except openpyxl.utils.exceptions.InvalidFileException:
            wb.remove(wb.active)
            wb.close()
            self.showErrorMessage('Вы не выбрали файл или \nне верно указано имя файла!')

    def showFileDialog(self):
        home_dir = str(Path.home())
        dialog = QtWidgets.QFileDialog()
        f_name = dialog.getOpenFileName(MainWindow, 'Выберите файл в в формате Excel', home_dir, filter='*.xlsx')
        if type(f_name) == tuple:
            self.lineEdit.setText(f_name[0])
        else:
            self.lineEdit.setText(str(f_name))

    def showSaveFileDialog(self, filename):
        dialog = QtWidgets.QFileDialog()
        f_name = dialog.getSaveFileName(MainWindow, 'Сохранить...', directory=f'{filename}.xlsx', filter='*.xlsx')
        if type(f_name) == tuple:
            return f_name[0]
        else:
            return str(f_name)

    def showErrorMessage(self, w_text):
        msg_box = QMessageBox()
        msg_box.critical(MainWindow, 'Ошибка!!!', w_text)

    def showInfoMessage(self, w_text):
        msg_box = QMessageBox()
        msg_box.information(MainWindow, 'Информация...', w_text)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
