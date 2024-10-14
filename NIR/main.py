from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QDialog, QLineEdit, QLabel, QVBoxLayout, QWidget, QHBoxLayout, QSpinBox
from PyQt5 import QtGui
from functions import initial_data, creating_excel_lab3, creating_excel_lab1, all, mean, strings, mastostr, mastostr2, find_plateau_time
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import openpyxl
from openpyxl.drawing.image import Image
import sys
from decimal import Decimal


class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()

        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

        self.setWindowTitle("Программа для выполнения ЛР")
        self.setGeometry(500, 250, 750, 250)

        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        self.button1 = QtWidgets.QPushButton(self)
        self.button1.setText("Ввести начальные данные для лабораторной работы по рассеиванию")
        self.button1.adjustSize()
        self.button1.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        self.button1.clicked.connect(self.open_lab1)

        self.button2 = QtWidgets.QPushButton(self)
        self.button2.setText("Ввести начальные данные для лабораторной работы ФДЭ")
        self.button2.adjustSize()
        self.button2.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        self.button2.clicked.connect(self.open_lab3)

        text_greatings1 = ("Вас приветствует программа для составления отчетов по лабораторным работам \n"
                          "№1 «Изучение особенностей рассеяния лазерного излучения в модельных биосредах»\n"
                          "и №3 «Изучение фотодинамического действия света на клеточные структуры»\n\n"
                          "Цель ЛР1: овладение навыками экспериментального исследования процесса рассеяния лазерного излучения\n"
                          "для оценки размеров рассеивающих частиц, характерных для биологических сред.")
        self.main_text1 = QtWidgets.QLabel(self)
        self.main_text1.setText(text_greatings1)
        self.main_text1.adjustSize()


        text_greatings2 = ("Цель ЛР3: приобретение навыков оценки эффективности метода фотодинамического действия света на\n"
                           "клеточные структуры на примере наблюдения фотодинамического эффекта (ФДЭ)\n"
                           "на предложенных преподавателем образца.")
        self.main_text2 = QtWidgets.QLabel(self)
        self.main_text2.setText(text_greatings2)
        self.main_text2.adjustSize()
        # self.main_text2.move(10, 165)

        layout.addWidget(self.main_text1)
        layout.addWidget(self.button1)
        layout.addWidget(self.main_text2)
        layout.addWidget(self.button2)

    def open_lab3(self):
        self.open_lab3 = WindowLab3()
        self.open_lab3.show()
        self.close()

    def open_lab1(self):
        self.open_lab1 = WindowLab1()
        self.open_lab1.show()
        self.close()


class WindowLab3(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Форма ввода данных')
        self.initUI()
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

    def initUI(self):
        # Создание основного виджета и главного макета
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # Метка и поле для ввода времени
        label_time = QLabel('Полное время облучения в секундах, через пробел:', self)
        self.edit_time = QLineEdit(self)
        layout.addWidget(label_time)
        layout.addWidget(self.edit_time)

        # Метка и поле для ввода значения пропускания
        label_transmittance = QLabel('Пропускание в процентах, через пробел:', self)
        self.edit_transmittance = QLineEdit(self)
        layout.addWidget(label_transmittance)
        layout.addWidget(self.edit_transmittance)

        # Метка и поле для ввода длины волны
        label_wavelength = QLabel('Длина волны в нанометрах:', self)
        self.edit_wavelength = QLineEdit(self)
        layout.addWidget(label_wavelength)
        layout.addWidget(self.edit_wavelength)

        label_power = QLabel('Плотность мощности матрицы, Вт/м^2:', self)
        self.edit_power = QLineEdit(self)
        layout.addWidget(label_power)
        layout.addWidget(self.edit_power)

        label_width = QLabel('Ширина , м:', self)
        self.edit_width = QLineEdit(self)
        layout.addWidget(label_width)
        layout.addWidget(self.edit_width)

        label_heigh = QLabel('Длина , м:', self)
        self.edit_heigh = QLineEdit(self)
        layout.addWidget(label_heigh)
        layout.addWidget(self.edit_heigh)

        # Кнопка для отправки данных
        submit_button = QPushButton('Отправить данные', self)
        submit_button.clicked.connect(self.submitData)
        layout.addWidget(submit_button)

        back_button = QPushButton("Назад", self)
        back_button.clicked.connect(self.go_back)
        layout.addWidget(back_button)

    def go_back(self):
        self.main_window = Window()
        self.main_window.show()
        self.close()

    def submitData(self):
        # Пример обработчика события, который выводит данные в консоль
        time = self.edit_time.text()
        transmittance = self.edit_transmittance.text()
        wavelength = self.edit_wavelength.text()
        power = self.edit_power.text()
        width = self.edit_width.text()
        heigh = self.edit_heigh.text()

        transmittance = transmittance.strip()
        time = time.strip()
        power = power.strip()
        width = width.strip()
        heigh = heigh.strip()
        wavelength = wavelength.strip()

        try:
            time_array = [float(num) for num in time.split(' ')]
            transmittance_array = [float(num) for num in transmittance.split(' ')]
            wavelength1 = float(wavelength)
            power1 = float(power)
            width1 = float(width)
            heigh1 = float(heigh)
            dose = []

            for i in time_array:
                dose.insert(len(dose),round(i * width1 * heigh1 * power1, 8))
            creating_excel_lab3(transmittance_array, time_array, wavelength1, width1, heigh1, power1)

            # Отображаем окно с графиком и данными
            plot_filename = "images/graph_lab3.png"
            plot_filename2 = "images/graph2_lab3.png"
            plot_window = PlotWindowLab3(plot_filename,plot_filename2, wavelength1, width1, heigh1, power1, time, transmittance, dose, time_array, transmittance_array)
            plot_window.exec_()
            self.close()

        except ValueError:
            error = ErrorWindow()
            error.exec_()


class PlotWindowLab3(QDialog):
    def __init__(self, plot_filename, plot_filename2, wave_len, width, heigh, power, array_time, array_trans, dose, time_array, trans_array, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Отчет')
        # self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

        h_layout = QHBoxLayout()

        plateau_time = find_plateau_time(time_array, trans_array)

        # Добавляем текстовую информацию
        info_text = f"""
        Длина волны: {wave_len} нм
        Ширина пятна: {width} м
        Высота пятна: {heigh} м
        Плотность мощности матрицы: {power} Вт/м^2
        Время облучения, c: {array_time}
        Пропускание, %: {array_trans}
        Общая доза облучения, Дж: {mastostr2(dose)}
        Время выхода на плато ~ {plateau_time} c
        """
        info_label = QLabel(info_text, self)
        layout.addWidget(info_label)

        # Добавляем график в окно
        # label = QLabel(self)
        # pixmap = QtGui.QPixmap(plot_filename)
        # label.setPixmap(pixmap)
        # h_layout.addWidget(label)

        label1 = QLabel(self)
        pixmap2 = QtGui.QPixmap(plot_filename2)
        label1.setPixmap(pixmap2)
        h_layout.addWidget(label1)

        layout.addLayout(h_layout)

        self.another_window = AnotherMesureLab3(self)
        self.another_window.show()
        self.close()


class WindowLab1(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Форма ввода данных')
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

        # Создание основного виджета и главного макета
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        # Метка и поле для ввода
        number_of_mes = QLabel('Введите количество серий измерений\n'
                               'для изучаемой среды', self)
        layout.addWidget(number_of_mes)

        self.spin_box = QSpinBox(self)
        self.spin_box.setMinimum(1)
        self.spin_box.setMaximum(7)
        layout.addWidget(self.spin_box)

        submit_button = QPushButton('Далее', self)
        submit_button.clicked.connect(self.showlab1)
        layout.addWidget(submit_button)

        back_button = QPushButton("Назад", self)
        back_button.clicked.connect(self.go_back)
        layout.addWidget(back_button)

    def go_back(self):
        self.main_window = Window()
        self.main_window.show()
        self.close()

    def showlab1(self):
        measurement_count = self.spin_box.value()
        self.main_window = Lab1(measurement_count)
        self.main_window.show()
        self.close()


class Lab1(QMainWindow):
    def __init__(self, measurement_count):
        super().__init__()
        self.setWindowTitle('Форма ввода данных')
        self.initUI(measurement_count)
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

    def initUI(self, measurement_count):
        # Создание основного виджета и главного макета
        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        h_layout = QHBoxLayout()

        hh_layout = QHBoxLayout()

        name_sub = QLabel('Название среды', self)
        self.nameQ = QLineEdit(self)
        main_layout.addWidget(name_sub)
        main_layout.addWidget(self.nameQ)

        self.inputs_exp = []
        self.inputs_deg = []
        self.inputs_amp = []
        self.inputs_wl = []

        label_wavelenght = QLabel('Длина волны:   ', self)
        self.edit_wavelenght = QLineEdit(self)
        self.inputs_wl.append(self.edit_wavelenght)
        main_layout.addWidget(label_wavelenght)
        main_layout.addWidget(self.edit_wavelenght)

        label_degrees = QLabel('Градусы, через пробел:     ', self)
        degrees = QLineEdit(self)

        self.inputs_deg.append(degrees)

        main_layout.addWidget(label_degrees)
        main_layout.addWidget(degrees)

        for i in range(measurement_count):
            v_layout = QVBoxLayout()
            experiment = QLabel(f'Измерение № {i+1}', self)
            v_layout.addWidget(experiment)

            # Метка и поле для ввода
            label_amp = QLabel('Значение измерений, через пробел:', self)
            amplitude = QLineEdit(self)

            v_layout.addWidget(label_amp)
            v_layout.addWidget(amplitude)

            self.inputs_exp.append(experiment)
            self.inputs_amp.append(amplitude)

            h_layout.addLayout(v_layout)
            main_layout.addLayout(h_layout)

        submit_button = QPushButton('Отправить данные', self)
        hh_layout.addWidget(submit_button)
        submit_button.clicked.connect(self.submitData)

        back_button = QPushButton("Назад", self)
        back_button.clicked.connect(self.go_back)
        hh_layout.addWidget(back_button)

        main_layout.addLayout(hh_layout)

    def go_back(self):
        self.main_window = WindowLab1()
        self.main_window.show()
        self.close()

    def submitData(self):
        try:
            data_amp = []
            for i in self.inputs_amp:
                i = i.text()
                i = i.strip()
                data_amp.append(i)

            data_deg = []
            for i in self.inputs_deg:
                i = i.text()
                i = i.strip()
                data_deg.append(i)

            data_wl = []
            data_wl.append(self.edit_wavelenght.text())

            data_name = []
            data_name.append(self.nameQ.text())

            creating_excel_lab1(data_deg, data_amp, data_name, data_wl)
            self.close()

            # Отображаем окно с графиком и данными
            plot_filename = "images/graph_lab1.png"
            plot_window = PlotWindowLab1(plot_filename, data_deg, data_amp, data_name, data_wl)
            plot_window.exec_()


        except ValueError:
            error = ErrorWindow()
            error.exec_()


class PlotWindowLab1(QDialog):
    def __init__(self, plot_filename, data_deg, data_amp, data_name, wavelenght, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Отчет')
        self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

        data1 = strings(data_amp)
        data2 = mean(all(data_amp))

        # Добавляем текстовую информацию
        info_text = f"""
        Вещество: {data_name[0]}
        Длина волны: {wavelenght[0]}
        Градусы: {mastostr(data_deg)}
        {data1}
        Среднее значение: {mastostr2(data2)}
        """
        info_label = QLabel(info_text, self)
        layout.addWidget(info_label)

        # Добавляем график в окно
        label = QLabel(self)
        pixmap = QtGui.QPixmap(plot_filename)
        label.setPixmap(pixmap)
        layout.addWidget(label)

        self.another_window = AnotherMesureLab1(self)
        self.another_window.show()
        self.close()


class ErrorWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Ошибка')
        layout = QVBoxLayout()
        self.setGeometry(500,500,250,80)
        self.setLayout(layout)
        self.setWindowIcon(QtGui.QIcon('images/error.jpg'))

        info_text = f"Ошибка ввода данных"

        self.main_text = QtWidgets.QLabel(self)
        self.main_text.setText(info_text)
        self.main_text.adjustSize()
        self.main_text.setAlignment(QtCore.Qt.AlignCenter)
        self.main_text.move(10, 10)

        self.button1 = QtWidgets.QPushButton(self)
        self.button1.setText("Ok")
        self.button1.move(65, 70)
        self.button1.adjustSize()
        self.button1.clicked.connect(self.close)

        layout.addWidget(self.main_text)
        layout.addWidget(self.button1)



class AnotherMesureLab1(QDialog):
    def __init__(self, parent_window, parent=None):
        super().__init__(parent)
        self.parent_window = parent_window
        self.setWindowTitle('Измерение')
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        h_layout = QHBoxLayout()

        info_text = f"Желаете сделать ещё измерение?"

        self.main_text = QtWidgets.QLabel(self)
        self.main_text.setText(info_text)
        self.main_text.adjustSize()
        main_layout.addWidget(self.main_text)

        self.button1 = QtWidgets.QPushButton(self)
        self.button1.setText("Да")
        self.button1.adjustSize()
        self.button1.clicked.connect(self.open_lab1)
        h_layout.addWidget(self.button1)

        self.button2 = QtWidgets.QPushButton(self)
        self.button2.setText("Нет")
        self.button2.adjustSize()
        self.button2.clicked.connect(self.go_back)
        h_layout.addWidget(self.button2)

        main_layout.addLayout(h_layout)

    def open_lab1(self):
        self.open_lab1 = WindowLab1()
        self.open_lab1.show()
        self.close()

    def go_back(self):
        self.main_window = Window()
        self.main_window.show()
        self.close()

class AnotherMesureLab3(QDialog):
    def __init__(self, parent_window, parent=None):
        super().__init__(parent)
        self.parent_window = parent_window
        self.setWindowTitle('Измерение')
        self.setWindowIcon(QtGui.QIcon('images/icon.jpg'))

        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        h_layout = QHBoxLayout()

        info_text = f"Желаете сделать ещё измерение?"

        self.main_text = QtWidgets.QLabel(self)
        self.main_text.setText(info_text)
        self.main_text.adjustSize()
        main_layout.addWidget(self.main_text)

        self.button1 = QtWidgets.QPushButton(self)
        self.button1.setText("Да")
        self.button1.move(15, 70)
        self.button1.adjustSize()
        self.button1.clicked.connect(self.open_lab3)
        h_layout.addWidget(self.button1)

        self.button2 = QtWidgets.QPushButton(self)
        self.button2.setText("Нет")
        self.button2.move(155, 70)
        self.button2.adjustSize()
        self.button2.clicked.connect(self.go_back)
        h_layout.addWidget(self.button2)

        main_layout.addLayout(h_layout)

    def open_lab3(self):
        self.open_lab3 = WindowLab3()
        self.open_lab3.show()
        self.close()

    def go_back(self):
        self.main_window = Window()
        self.main_window.show()
        self.close()

def application():
    app = QApplication(sys.argv)
    window = Window()

    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    application()