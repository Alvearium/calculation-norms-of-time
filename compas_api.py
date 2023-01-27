import os, re

import pythoncom
import subprocess
from win32com.client import Dispatch, gencache

k0 = 0
k1 = 0
k_tp = 0
k_scale = 0

table_norm_time_for_A4 = {
    '5': 0.32,
    '6':  0.4,
    '7-8': 0.48,
    '9-10': 0.54,
    '11-13': 0.62,
    '14-17': 0.7,
    '18-21': 0.82,
    '22-27': 0.96,
    '28-34': 1.05,
    '35-44': 1.28,
    '45-56': 1.44
}
table_norm_time_for_A3 = {
    '7-8': 0.58,
    '9-10': 0.65,
    '11-13': 0.73,
    '14-17': 0.84,
    '18-21': 0.98,
    '22-27': 1.2,
    '28-34': 1.32,
    '35-44': 1.46,
    '45-56': 1.7,
    '52-71': 1.96,
    '72-91': 2.4
}
table_norm_time_for_A2 = {
    '11-13': 1.10,
    '14-17': 1.24,
    '18-21': 1.48,
    '22-27': 1.56,
    '28-34': 1.8,
    '35-44': 2.24,
    '45-56': 2.7,
    '52-71': 2.96,
    '72-91': 3.5,
    '92-115': 4.0,
    '116-147': 4.6
}
table_norm_time_for_A1 = {
    '18-21': 2.1,
    '22-27': 2.5,
    '28-34': 3.0,
    '35-44': 3.6,
    '45-56': 4.1,
    '52-71': 4.6,
    '72-91': 5.1,
    '92-115': 6.0,
    '116-147': 6.9,
    '148-187': 8.2,
    '188-238': 9.6
}
table_norm_time_for_A0 = {
    '28-34': 4.2,
    '35-44': 4.6,
    '45-56': 5.0,
    '52-71': 5.4,
    '72-91': 7.4,
    '92-115': 8.4,
    '116-147': 9.4,
    '148-187': 10.4,
    '188-238': 11.4,
    '239-300': 12.4,
    '300-10000': 14
}

general_view_drawing = {
    '7': 15,
    '8-12': 18,
    '13-21': 20,
    '22-35': 24,
    '36-60': 28,
    '61-103': 31,
    '104-10000': 36
}

dict_k1 = {
    'A4': 0.1,
    'A3': 0.2,
    'A2': 0.4,
    'A1': 1,
    'A0': 1.6
}

drawing_scale = {
    '1:1': 1,
    '1:2': 1.05,
    '1:10': 1.05,
    '1:20': 1.05,
    '1:100': 1.05,
    '1:1000': 1.05,
    '1:2,5': 1.1,
    '1:4': 1.1,
    '1:5': 1.1,
    '1:40': 1.1,
    '1:50': 1.1,
    '1:200': 1.1,
    '1:400': 1.1,
    '1:500': 1.1,
    '1:800': 1.1,
    '2:1': 1.1,
    '4:1': 1.1,
    '5:1': 1.1,
    '1:15': 1.15,
    '1:25': 1.15,
    '1:75': 1.15,
}

def calculation(format, number_of_size, obj_window):
    list_coefficients = []
    time = 0
    if obj_window.radioButton_6.isChecked():
        global k0
        if format == 'A4':
            keys = [*table_norm_time_for_A4]
            for key in keys:
                if len(key) < 2:
                    if number_of_size <= int(key) < 6:
                        time = table_norm_time_for_A4[key]
                    elif number_of_size == int(key):
                        time = table_norm_time_for_A4[key]
                else:
                    iter_keys = key.split('-')
                    if int(iter_keys[0]) <= number_of_size <= int(iter_keys[1]):
                        time = table_norm_time_for_A4[key]
        elif format == 'A3':
            keys = [*table_norm_time_for_A3]
            for key in keys:
                iter_keys = key.split('-')
                if int(iter_keys[0]) <= number_of_size <= int(iter_keys[1]):
                    time = table_norm_time_for_A3[key]
        elif format == 'A2':
            keys = [*table_norm_time_for_A2]
            for key in keys:
                iter_keys = key.split('-')
                if int(iter_keys[0]) <= number_of_size <= int(iter_keys[1]):
                    time = table_norm_time_for_A2[key]
        elif format == 'A1':
            keys = [*table_norm_time_for_A1]
            for key in keys:
                iter_keys = key.split('-')
                if int(iter_keys[0]) <= number_of_size <= int(iter_keys[1]):
                    time = table_norm_time_for_A1[key]
        elif format == 'A0':
            keys = [*table_norm_time_for_A0]
            for key in keys:
                iter_keys = key.split('-')
                if int(iter_keys[0]) <= number_of_size <= int(iter_keys[1]):
                    time = table_norm_time_for_A0[key]
        k0 = time

    # Чертеж общего вида
    elif obj_window.radioButton_4.isChecked():
        global k1
        keys = [*general_view_drawing]
        for key in keys:
            if len(key) < 2:
                if number_of_size <= int(key):
                    time = general_view_drawing[key]
                    time *= dict_k1[format]
                    k1 = dict_k1[format]
                    k0 = general_view_drawing[key]

            else:
                iter_keys = key.split('-')
                if int(iter_keys[0]) <= number_of_size <= int(iter_keys[1]):
                    time = general_view_drawing[key]
                    k0 = general_view_drawing[key]
                    time *= dict_k1[format]

                    k1 = dict_k1[format]



    # Поправочный коэффициент по Типу производства
    global k_tp
    if obj_window.radioButton.isChecked():
        time *= 1
        k_tp = 1
    elif obj_window.radioButton_2.isChecked():
        time *= 1.2
        k_tp = 1.2
    elif obj_window.radioButton_3.isChecked():
        time *= 1.3
        k_tp = 1.3

    # Поправочный коэффициент по масштабу
    global k_scale
    if obj_window.label_scale.text() != None or obj_window.label_scale.text() != '***':
        time *= drawing_scale[obj_window.label_scale.text()]
        k_scale = drawing_scale[obj_window.label_scale.text()]


    return time

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const


def is_running():
    proc_list = subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"',
                                 shell=False,
                                 stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False


def amount_sheet(doc7, obj_window):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
        if format.FormatMultiplicity != 0:
            obj_window.label_format.setText("A" + str(format.Format))
    return sheets


def stamp_scale(doc7):
    stamp = doc7.LayoutSheets.Item(0).Stamp  # Item(0) указывает на штамп первого листа
    return stamp.Text(32).Str


def stamp(doc7, obj_window):
    for sheet in range(doc7.LayoutSheets.Count):
        style_filename = os.path.basename(doc7.LayoutSheets.Item(sheet).LayoutLibraryFileName)
        style_number = int(doc7.LayoutSheets.Item(sheet).LayoutStyleNumber)

        if style_filename in ['graphic.lyt', 'Graphic.lyt'] and style_number == 1:
            stamp = doc7.LayoutSheets.Item(sheet).Stamp

            if stamp.Text(110).Str != None:
                obj_window.label_des.setText(stamp.Text(110).Str)
            else:
                obj_window.label_des.setText('Не указано')
            obj_window.label_scale.setText(str(re.search(r"\d+:\d+", stamp.Text(6).Str).group()))

            if stamp.Text(40) or stamp.Text(41) or stamp.Text(42):
                obj_window.label_lit.setText(stamp.Text(40).Str + stamp.Text(41).Str + stamp.Text(42).Str)

            return {"Scale": re.search(r"\d+:\d+", stamp.Text(6).Str).group(),
                    "Designer": stamp.Text(110).Str}

    return {"Scale": 'Неопределенный стиль оформления',
            "Designer": 'Неопределенный стиль оформления'}


def parse_stamp(doc7, number_sheet):
    stamp = doc7.LayoutSheets.Item(number_sheet).Stamp
    for i in range(10000):
        if stamp.Text(i).Str:
            print('Номер ячейки = %-5d Значение = %s' % (i, stamp.Text(i).Str))


def count_TT(doc7, module7, obj_window):
    doc2D_s = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'],
                                           pythoncom.IID_IDispatch)
    doc2D = module7.IDrawingDocument(doc2D_s)
    text_TT = doc2D.TechnicalDemand.Text

    count_tt = 0  # Количество пунктов технических требований
    for i in range(text_TT.Count):  # Проходим по каждой строчке технических требований
        if text_TT.TextLines[i].Numbering == 1:  # и проверяем, есть ли у строки нумерация
            count_tt += 1

    # Если нет нумерации, но есть текст
    if not count_tt and text_TT.TextLines[0]:
        count_tt += 1

    obj_window.label_tt.setText(str(count_tt))

    return count_tt


def count_dimension(doc7, module7, obj_window):
    IKompasDocument2D = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'],
                                                     pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(IKompasDocument2D)
    views = doc2D.ViewsAndLayersManager.Views

    count_dim = 0
    for i in range(views.Count):
        ISymbols2DContainer = views.View(i)._oleobj_.QueryInterface(module7.NamesToIIDMap['ISymbols2DContainer'],
                                                                    pythoncom.IID_IDispatch)
        dimensions = module7.ISymbols2DContainer(ISymbols2DContainer)

        # Складываем все необходимые размеры
        count_dim += dimensions.AngleDimensions.Count + \
                     dimensions.ArcDimensions.Count + \
                     dimensions.Bases.Count + \
                     dimensions.BreakLineDimensions.Count + \
                     dimensions.BreakRadialDimensions.Count + \
                     dimensions.DiametralDimensions.Count + \
                     dimensions.Leaders.Count + \
                     dimensions.LineDimensions.Count + \
                     dimensions.RadialDimensions.Count + \
                     dimensions.RemoteElements.Count + \
                     dimensions.Roughs.Count + \
                     dimensions.Tolerances.Count

    obj_window.label_dimension.setText(str(count_dim))

    return count_dim


def parse_design_documents(paths, obj_window):
    is_run = is_running()  # Установим флаг, который нам говорит,
    # запущена ли программа до запуска нашего скрипта

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы
    table = []  # Создаём таблицу параметров
    for path in paths:
        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=True,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять

        obj_window.label_name.setText(doc7.Name)
        row = amount_sheet(doc7, obj_window)  # Посчитаем кол-во листов каждого формат
        row.update(stamp(doc7, obj_window))
        count_dimension_val = count_dimension(doc7, module7, obj_window)# Читаем основную надпись
        row.update({
            "Filename": doc7.Name,  # Имя файла
            "CountTT": count_TT(doc7, module7, obj_window),  # Количество пунктов технических требований
            "CountDim": count_dimension_val,  # Количество размеров на чертеже
        })
        table.append(row)  # Добавляем строку параметров в таблицу

        obj_window.label_nt.setText(str(calculation(stamp_scale(doc7), count_dimension_val, obj_window)))

        doc7.Close(const7.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run: app7.Quit()  # Выходим из программы
    return table


def print_to_excel(result, obj_window):
    excel = Dispatch("Excel.Application")  # Подключаемся к Excel
    excel.Visible = True  # Делаем окно видимым
    wb = excel.Workbooks.Add()  # Добавляем новую книгу
    sheet = wb.ActiveSheet  # Получаем ссылку на активный лист

    # Создаём заголовок таблицы
    sheet.Range("A1:J1").value = ["Имя файла", "Разработчик",
                                  "Кол-во размеров", "Кол-во пунктов ТТ",
                                  "А0", "А1", "А2", "А3", "А4", "Масштаб"]

    # Заполняем таблицу
    for i, row in enumerate(result):
        sheet.Range("A2:J2").value = [row['Filename'],
                                      row['Designer'],
                                      row['CountDim'],
                                      row['CountTT'],
                                      row['A0'],
                                      row['A1'],
                                      row['A2'],
                                      row['A3'],
                                      row['A4'],
                                      "".join(('="', row['Scale'], '"'))]

    sheet.Cells(1, 11).value = 'Норма времени'
    sheet.Cells(2, 11).value = obj_window.label_nt.text()

    global k0, k1, k_tp, k_scale
    sheet.Cells(4, 1).value = 'Норма времени без коэффицентов'
    sheet.Cells(5, 1).value = k0
    sheet.Cells(4, 2).value = 'Коэффицент формата'
    sheet.Cells(5, 2).value = k1
    sheet.Cells(4, 3).value = 'Коэффициент Типа производтсва'
    sheet.Cells(5, 3).value = k_tp
    sheet.Cells(4, 4).value = 'Коэффициент Масштаба'
    sheet.Cells(5, 4).value = k_scale

    wb.Save()



    # wb.save(filename)
    # send_file(wb)
