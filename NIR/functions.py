from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from openpyxl.drawing.image import Image
from openpyxl import Workbook
from openpyxl.chart import LineChart,Reference
import numpy as np

import matplotlib.pyplot as plt
import openpyxl
import sys
from decimal import Decimal


def initial_data():
    # Ввод начальных данных

    global wave_len, width, heigh, array_time, array_trans

    wave_len = float(input("Введите длину волны в нанометрах:"))
    print("Введите размер пятна облучения")

    width = float(input("Ширина:"))
    heigh = float(input("Высота:"))

    time = input("Введите время облучения в секундах через пробел:")
    array_time = time.split()
    array_time = list(map(int, array_time))

    trans = input("Пропускание вещества в процентах% через пробел:")
    array_trans = trans.split()
    array_trans = list(map(int, array_trans))


def creating_excel_lab3(array_trans, array_time, wave_len, width, heigh, power):
    # Создаем новый файл Excel
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['G1'], sheet['G2'] = "Длина волны", wave_len
    sheet['D1'], sheet['D2'] = "Ширина пятна", width
    sheet['E1'], sheet['E2'] = "Высота пятна", heigh
    sheet['F1'], sheet['F2'] = "Плотность мощности матрицы", power
    sheet['A1'] = "Полное время облучения, с"
    sheet['B1'] = "Пропускание, %"
    sheet['C1'] = "Полная доза облучение, Дж"

    row = 2
    for i in array_trans:
        sheet[f'A{row}'] = i
        i += 1
        row += 1

    row = 2
    for i in array_time:
        sheet[f'B{row}'] = i
        i += 1
        row += 1

    row = 2
    array_power = []
    for i in array_time:
        sheet[f'C{row}'] = i * width * heigh * power
        array_power.append(i * width * heigh * power)
        i += 1
        row += 1

    plateau_time = find_plateau_time(array_time, array_trans)
    sheet['H1'] = "Время выхода на плато"
    sheet['H2'] = plateau_time

    plt.plot(array_power, array_trans, color='g')
    plt.title('График зависимости пропускания от дозы облучения')
    plt.scatter(array_power, array_trans, color='g')
    plt.xlabel('Доза, Дж')
    plt.ylabel('Пропускание, %')
    plot_filename2 = 'images/graph2_lab3.png'
    plt.savefig(plot_filename2)
    plt.close()

    plt.scatter(array_time, array_trans, color='b', label='Измерения')
    plt.plot(array_time, array_trans, color='b')
    plt.axvline(x=plateau_time, color='r', linestyle='--', label=f'Время выхода на плато: {plateau_time}')
    plt.title('График зависимости пропускания от времени')
    plt.xlabel('Время, с')
    plt.ylabel('Пропускание, %')
    plot_filename = 'images/graph_lab3.png'
    plt.legend()
    plt.savefig(plot_filename)
    plt.close()

    img = Image(plot_filename)
    img.anchor = 'T2'
    sheet.add_image(img)

    img2 = Image(plot_filename2)
    img2.anchor = 'I2'
    sheet.add_image(img2)

    wb.save("excel/Отчет_ЛР_3.xlsx")


def creating_excel_lab1(data_deg, data_amp, data_name, wavelenght):
    # Создаем новый файл Excel
    wb = openpyxl.Workbook()
    sheet = wb.active
    name = data_name[0]
    data_deg_num = [float(num) for num in data_deg[0].split(' ')]

    sheet['A1'] = f"{name}"

    sheet['A2'] = "Градусы"
    row = 3
    for i in data_deg_num:
        sheet[f'A{row}'] = i
        row += 1

    code = ord("B")
    code_new = code

    for k in data_amp:
        data_amp_num = [float(num) for num in k.split(' ')]
        j = 3
        sheet[f'{chr(code_new)}2'] = f"Измерение {code_new+1-code}"
        for l in data_amp_num:
            sheet[f'{chr(code_new)}{j}'] = f'{l}'
            j += 1
        code_new += 1

    # all_data = []
    # for k in data_amp:
    #     all_data.append([float(num) for num in k.split(' ')])

    all_data = all(data_amp)

    # mean_mas = []
    # for i in range(len(all_data[0])):
    #     mas = []
    #     for k in all_data:
    #         mas.append(k[i])
    #     mean_mas.append(sum(mas)/len(mas))

    mean_mas = mean(all_data)

    sheet[f'{chr(code_new)}2'] = "Среднее значение"
    j = 3
    for i in mean_mas:
        sheet[f"{chr(code_new)}{j}"] = f"{i}"
        j += 1
    code_new += 1

    sheet[f"{chr(code_new)}2"] = "Длина волны"
    sheet[f"{chr(code_new)}3"] = f"{wavelenght[0]}"
    code_new += 1

    # Преобразование градусов в радианы
    radians = np.deg2rad(data_deg_num)

    # Построение графика
    plt.polar(radians, mean_mas)
    plt.title(f'Индикатриса рассеяния для вещества "{name}"')
    plot_filename = 'images/graph_lab1.png'
    plt.savefig(plot_filename)
    plt.close()

    img = Image(plot_filename)
    img.anchor = f'{chr(code_new)}3'
    sheet.add_image(img)

    wb.save("excel/Отчет_ЛР_1.xlsx")

def mean(all_data):
    mean_mas = []
    for i in range(len(all_data[0])):
        mas = []
        for k in all_data:
            mas.append(k[i])
        mean_mas.append(sum(mas)/len(mas))

    return mean_mas

def all(data_amp):
    all_data = []
    for k in data_amp:
        all_data.append([float(num) for num in k.split(' ')])

    return all_data

def strings(data_amp):
    i = 1
    mas = []
    for k in data_amp:
        string = f"Измерение №{i}: " + f"{k}" + "\n"
        mas.append(string)
        i += 1

    stingus = "\n"
    for k in mas:
        stingus += k

    return stingus

def mastostr(mas):
    str = ""
    for i in mas:
        str += i + " "
    return str

def mastostr2(mas):
    str = ""
    for i in mas:
        str += f"{i}" + " "
    return str

def find_plateau_time(array_time, array_trans, threshold=2):
    n = len(array_trans)

    mid = n // 2  # Начинаем с середины массива

    # Проверка центрального элемента и элементов вокруг него
    for i in range(mid, n - 1):
        if abs(array_trans[i] - array_trans[i - 1]) <= threshold and abs(array_trans[i] - array_trans[i + 1]) <= threshold:
            return array_time[i]

    for i in range(mid - 1, 0, -1):
        if abs(array_trans[i] - array_trans[i - 1]) <= threshold and abs(array_trans[i] - array_trans[i + 1]) <= threshold:
            return array_time[i]

    return array_time[-1]