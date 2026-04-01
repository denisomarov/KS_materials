import streamlit as st
import io
import pandas as pd
import math
import numpy as np
from fontTools.cu2qu.cu2qu import NAN
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side, Alignment

def process_data(FileName, FileName_Sheet, Voc_FileName):
    # Инииализация переменных

    sign_of_existing_cable  = 'Существующий кабель'
    opt_c_name = ['6XV8100', '6XV8100 LC-LC', '6XV8100 ST-ST', '6XV8100 ST-LC', '6XV8100 LC-ST',
                  'ОВК-Б-нг(А) HF – 1Г – 0,5 кН LC-LC', 'ОВК-Б-нг(А) HF – 1Г – 0,5 кН ST-ST',
                  'ОВК-Б-нг(А) HF – 1Г – 0,5 кН ST-LC', 'ОВК-Б-нг(А) HF – 1Г – 0,5 кН LC-ST',
                  'ОВК-Б-нг(А) HF – 2Г – 2,7 кН LC-LC', 'ОВК-Б-нг(А) HF – 2Г – 2,7 кН ST-ST',
                  'ОВК-Б-нг(А) HF – 2Г – 2,7 кН ST-LC', 'ОВК-Б-нг(А) HF – 2Г – 2,7 кН LC-ST',
                  'ОВК-Б-нг(А)-HF-1Г LC-LC', 'ОВК-Б-нг(А)-HF-1Г ST-ST',
                  'ОВК-Б-нг(А)-HF-1Г LC-ST', 'ОВК-Б-нг(А)-HF-1Г ST-LC',
                  'ОВК-Б-нг(А)-HF-2Г LC-LC', 'ОВК-Б-нг(А)-HF-2Г ST-ST',
                  'ОВК-Б-нг(А)-HF-2Г LC-ST', 'ОВК-Б-нг(А)-HF-2Г ST-LC'
                  ]

    # Инициализация структуры для сохранения кабельной продукции
    columns = ['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']
    df_cable = pd.DataFrame(columns=columns)
    df_cable.loc[0] = ['Кабельная продукция', '', '', '', np.nan]

    # фиксируем формат таблицы кабеля
    df_cable['Наименование'] = df_cable['Наименование'].astype(str)
    df_cable['Марка'] = df_cable['Марка'].astype(str)
    df_cable['Код'] = df_cable['Код'].astype(str)
    df_cable['Ед. изм.'] = df_cable['Ед. изм.'].astype(str)
    df_cable['Количество'] = df_cable['Количество'].astype(float)

    # считываем кабельный журнал
    df_a = pd.read_excel(FileName, FileName_Sheet, header=1)
    df_a = df_a.iloc[1:, 1:10]
    df_a.columns = ['№ кабеля', 'Марка кабеля', 'Жильность x сечение', 'Кол-во использ. Жил', 'Откуда', 'Куда',
                    'Длина проект, м', 'Длина факт, м', 'Примечание']
    df_a = df_a.dropna(how='all')

    # считываем словарь Кабеля
    df_vc = pd.read_excel(Voc_FileName, 'Кабель')
    df_vc = df_vc.dropna(how='all')

    # считываем словарь Кабельных материалов
    df_vm = pd.read_excel(Voc_FileName, 'Словарь')
    df_vm = df_vm.dropna(how='all')

    # проверяем, что данные были считаны
    if ((not df_a.empty) and (not df_vc.empty) and (not df_vm.empty)):

        # Очистка полученных данных и формирование базы
        # Убираем существующий кабель
        df_a = df_a.loc[df_a['Примечание'] != sign_of_existing_cable]
        # Убираем пустые строки в Номере кабеля
        df_a = df_a.loc[df_a['№ кабеля'].str.replace(' ','') != '']
        # Убираем непечатные символы
        df_a = df_a.apply(lambda x: x.str.replace(r'[\r\n\t]', '', regex=True) if x.dtype == "str" else x)
        # Проверяем значение в столбце Длина проект, м, если не цифра - заменяем на 0
        df_a['Длина проект, м'] = df_a['Длина проект, м'].astype(str)
        df_a['Длина проект, м'] = df_a['Длина проект, м'].str.replace(',', '.')
        df_a['Длина проект, м'] = pd.to_numeric(df_a['Длина проект, м'], errors = 'coerce', downcast='float')
        df_a['Длина проект, м'] = df_a['Длина проект, м'].fillna(0)
        # присваиваем тип столбцам таблицы
        df_a['№ кабеля'] = df_a['№ кабеля'].astype(str)
        df_a['Марка кабеля'] = df_a['Марка кабеля'].astype(str)
        df_a['Жильность x сечение'] = df_a['Жильность x сечение'].astype(str)
        df_a['Кол-во использ. Жил'] = df_a['Кол-во использ. Жил'].astype(str)
        df_a['Откуда'] = df_a['Откуда'].astype(str)
        df_a['Куда'] = df_a['Куда'].astype(str)
        df_a['Длина факт, м'] = df_a['Длина факт, м'].astype(str)
        df_a['Примечание'] = df_a['Примечание'].astype(str)
        # Меняем точки в сечении кабеля на запятые
        df_a['Жильность x сечение'] = df_a['Жильность x сечение'].str.replace('.', ',')

        # обработка оптического кабеля

        for patch_name in opt_c_name:

            df_optic = df_a.loc[df_a['Марка кабеля'] == patch_name]
            df_optic['Марка кабеля'] = df_optic['Марка кабеля'] + ' ' + df_optic['Длина проект, м'].astype(str) + ' м'
            df_optic_sorted = df_optic.groupby(['Марка кабеля', 'Длина проект, м']).agg('count').iloc[:, 0:1]
            df_optic_sorted.reset_index(inplace=True)

            df_optic_sorted.columns = ['Наименование', 'Длина, м', 'Количество']
            df_optic_sorted['Длина, м'] = df_optic_sorted['Длина, м'].astype(float)
            df_optic_sorted['Количество'] = df_optic_sorted['Количество'].astype(float)
            df_optic_sorted.sort_values(['Длина, м'], inplace=True)
            df_optic_sorted = df_optic_sorted[['Наименование', 'Количество']]
            df_optic_sorted['Наименование'] = df_optic_sorted['Наименование'].astype(str)

            # дополняем полученную таблицу из словаря Кабель
            df_vc_oc = df_vc[['Полная марка', 'Марка', 'Код', 'Ед. изм.']]
            df_vc_oc['Полная марка'] = df_vc_oc['Полная марка'].astype(str)
            df_vc_oc.columns = ['Наименование', 'Марка', 'Код', 'Ед. изм.']

            # переведем все значения столбца Наименование в заглавные буквы
            df_optic_sorted['Наименование'] = df_optic_sorted['Наименование'].str.upper()
            df_vc_oc['Наименование'] = df_vc_oc['Наименование'].str.upper()

            df_optic_sorted = pd.merge(df_optic_sorted, df_vc_oc, on='Наименование', how='left')
            df_optic_sorted = df_optic_sorted[['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']]
            df_optic_sorted.fillna('', inplace=True)

            # формируем Наименование согласно словаря Кабель
            df_vc_oc = df_vc[['Марка', 'Наименование']]
            df_optic_sorted = df_optic_sorted[['Марка', 'Код', 'Ед. изм.', 'Количество']]
            df_optic_sorted = pd.merge(df_optic_sorted, df_vc_oc, on='Марка', how='left')
            df_optic_sorted = df_optic_sorted[['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']]

            # увеличиваем в два раза число оптических пачкордов с 1 жилой
            df_optic_sorted.loc[df_optic_sorted['Наименование'].str.contains('– 1Г –', na=False), 'Количество'] *= 2

            # заносим данные по оптическому кабелю в общую таблицу кабеля
            if len(df_optic_sorted) > 0:
                df_cable = pd.concat([df_cable, df_optic_sorted], ignore_index=True)

        # фиксируем формат финальной таблицы кабеля
        df_cable['Наименование'] = df_cable['Наименование'].astype(str)
        df_cable['Марка'] = df_cable['Марка'].astype(str)
        df_cable['Код'] = df_cable['Код'].astype(str)
        df_cable['Ед. изм.'] = df_cable['Ед. изм.'].astype(str)
        df_cable['Количество'] = df_cable['Количество'].astype(float)

        # обработка медного кабеля
        # выделяем медный кабель

        df_cupper = df_a.loc[~df_a['Марка кабеля'].isin(opt_c_name)]
        df_cupper['Марка кабеля'] = df_cupper['Марка кабеля'] + ' ' + df_cupper['Жильность x сечение']
        df_cupper['Марка кабеля'] = df_cupper['Марка кабеля'].str.replace('+SH', '')
        df_cupper['Марка кабеля'] = df_cupper['Марка кабеля'].str.replace('+Sh', '')
        df_cupper['Марка кабеля'] = df_cupper['Марка кабеля'].str.replace('+sh', '')

        # обработка медного кабеля общего (включая внутри шкафов)
        df_cupper = df_cupper[['Марка кабеля', 'Длина проект, м']]
        df_cupper.columns = ['Марка', 'Количество']

        # переведем все значения столбца Марка таблицы в заглавные буквы
        df_cupper['Марка'] = df_cupper['Марка'].str.upper()

        df_cupper['Количество'] = df_cupper['Количество'].astype(float)
        df_cupper = df_cupper.groupby('Марка').agg('sum')
        df_cupper.sort_index()
        df_cupper.reset_index(inplace=True)

        # оформляем таблицу медного кабеля общего (включая внутри шкафов)
        df_cupper = df_cupper[['Марка', 'Количество']]

        # оформляем таблицу словаря медного кабеля
        df_vc_oc = df_vc[['Полная марка', 'Марка', 'Наименование', 'Код', 'Ед. изм.']]
        df_vc_oc.columns = ['Полная марка', 'Новая марка', 'Наименование', 'Код', 'Ед. изм.']
        df_vc_oc['Полная марка'] = df_vc_oc['Полная марка'].str.rstrip()

        # переведем все значения столбца Марка в заглавные буквы
        df_vc_oc['Полная марка'] = df_vc_oc['Полная марка'].str.upper()

        # формируем Наименование согласно словаря Кабель
        df_cupper = pd.merge(df_cupper, df_vc_oc, left_on='Марка', right_on='Полная марка', how='left')
        df_cupper = df_cupper[['Наименование', 'Новая марка', 'Код', 'Ед. изм.', 'Количество']]
        df_cupper.columns = ['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']
        df_cupper.fillna('', inplace=True)

        # корректируем Марку кабеля, заменим X на x
        df_cupper['Марка'] = df_cupper['Марка'].str.replace('X', "x")

        # заполняем полную таблицу кабеля
        if len(df_cupper) > 1:
            df_cable = pd.concat([df_cable, df_cupper], axis=0, ignore_index=True)

        # фиксируем формат финальной таблицы кабеля
        df_cable['Наименование'] = df_cable['Наименование'].astype(str)
        df_cable['Марка'] = df_cable['Марка'].astype(str)
        df_cable['Код'] = df_cable['Код'].astype(str)
        df_cable['Ед. изм.'] = df_cable['Ед. изм.'].astype(str)
        df_cable['Количество'] = df_cable['Количество'].astype(float)

        # РАСЧЕТ КАБЕЛЬНЫХ МАТЕРИАЛОВ

        # расчет числа маркировок кабеля
        cable_len = df_a['Длина проект, м'].sum()
        cable_cnt = df_a['Длина проект, м'].count()

        cable_tag = math.ceil((cable_cnt*(2 + 3) + cable_len/50)*1.1)

        # выделяем медный кабель
        df_cupper = df_a.loc[~df_a['Марка кабеля'].isin(opt_c_name)][['Марка кабеля', 'Жильность x сечение']]

        #определяем количество кабелей STP
        RJ45_cnt = math.ceil(df_cupper['Марка кабеля'].loc[df_cupper['Марка кабеля'] == 'STP Cord RJ45'].count()*1.2)

        #определяем количество стяжек кабеля
        cable_tie = math.ceil(cable_tag / 1000) * 1000

        # обработка медного кабеля
        # удаляем кабель STP
        df_cupper = df_cupper.loc[df_cupper['Марка кабеля'] != 'STP Cord RJ45']
        df_cupper.reset_index(inplace=True, drop=True)

        # делаем все буквы столбца Марка кабеля заглавными
        df_cupper['Марка кабеля'] = df_cupper['Марка кабеля'].str.upper()

        df_cupper['Жильность'] = df_cupper['Жильность x сечение'].str.replace('+SH','')
        df_cupper['Жильность'] = df_cupper['Жильность'].str.replace('.',',')

        # формируем Полную марку кабеля
        df_cupper['Марка кабеля'] = df_cupper['Марка кабеля'].str.replace(' ', '')
        df_cupper['Полная марка'] = df_cupper['Марка кабеля'] + ' ' + df_cupper['Жильность']

        # формируем таблицу по Полной марке с указанием количества кабеля
        df_cupper_gr = df_cupper.groupby('Полная марка').count()
        df_cupper_gr.reset_index(inplace=True)
        df_cupper_gr.columns = ['Полная марка', 'Кол-во, шт', 'Жильность, шт', 'Всего жил, шт']

        # определяем количество жил в кабеле
        df_cupper_gr['Жильность, шт'] = df_cupper_gr['Полная марка'].str.split(' ').str[1].str.split('x').str[0].str.replace(',','.').astype(float)
        df_cupper_gr['Всего жил, шт'] = df_cupper_gr['Жильность, шт'] * df_cupper_gr['Кол-во, шт']


        # готовим таблицу словаря и кабеля
        df_vc['Полная марка'] = df_vc['Полная марка'].str.upper()
        df_vc['Полная марка'] = df_vc['Полная марка'].str.rstrip()

        df_cupper_gr['Полная марка'] = df_cupper_gr['Полная марка'].str.upper()

        # объединяем таблицы кабеля и словаря
        df_result = df_cupper_gr.set_index('Полная марка').join(df_vc.set_index('Полная марка'))

        # группируем наконечники и рассчитываем количество с запасом 20%
        df_cc_lug = pd.DataFrame(df_result.groupby('Наконечник')['Всего жил, шт'].sum())
        df_cc_lug.reset_index(inplace=True)
        df_cc_lug.columns = ['Марка', 'Количество']
        df_cc_lug['Количество'] = np.ceil(df_cc_lug['Количество'] * 2 * 1.2 / 100) * 100
        df_cc_lug = df_cc_lug.join(df_vm.set_index('Марка'), on='Марка')
        df_cc_lug = df_cc_lug[['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']]
        df_cc_lug.sort_values('Марка', inplace=True)

        # фиксируем формат таблицы наконечников
        df_cc_lug['Наименование'] = df_cc_lug['Наименование'].astype(str)
        df_cc_lug['Марка'] = df_cc_lug['Марка'].astype(str)
        df_cc_lug['Код'] = df_cc_lug['Код'].astype(str)
        df_cc_lug['Ед. изм.'] = df_cc_lug['Ед. изм.'].astype(str)
        df_cc_lug['Количество'] = df_cc_lug['Количество'].astype(float)

        # группируем термоусадочную трубку и рассчитываем длину из расчета 30 см на один кабель
        df_cc_tut = pd.DataFrame(df_result.groupby('Термоусадочная трубка')['Кол-во, шт'].sum())
        df_cc_tut.reset_index(inplace=True)
        df_cc_tut.columns = ['Марка', 'Количество']
        df_cc_tut['Количество'] = df_cc_tut['Количество'] * 0.3
        df_cc_tut = df_cc_tut.join(df_vm.set_index('Марка'), on='Марка')
        df_cc_tut = df_cc_tut[['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']]
        df_cc_tut.sort_values('Марка', inplace=True)

        # фиксируем формат таблицы термоусадочной трубки
        df_cc_tut['Наименование'] = df_cc_tut['Наименование'].astype(str)
        df_cc_tut['Марка'] = df_cc_tut['Марка'].astype(str)
        df_cc_tut['Код'] = df_cc_tut['Код'].astype(str)
        df_cc_tut['Ед. изм.'] = df_cc_tut['Ед. изм.'].astype(str)
        df_cc_tut['Количество'] = df_cc_tut['Количество'].astype(float)

        # группируем маркировочную трубку и рассчитываем длину из рассчета 5 см на одну жилу
        df_cc_tube = pd.DataFrame(df_result.groupby('Трубка')['Всего жил, шт'].sum())
        df_cc_tube.reset_index(inplace=True)
        df_cc_tube.columns = ['Марка', 'Количество']
        df_cc_tube['Количество'] = df_cc_tube['Количество'] * 2 * 0.05
        df_cc_tube = df_cc_tube.join(df_vm.set_index('Марка'), on='Марка')
        df_cc_tube = df_cc_tube[['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']]
        df_cc_tube.sort_values('Марка', inplace=True)

        # фиксируем формат таблицы маркировочной трубки
        df_cc_tube['Наименование'] = df_cc_tube['Наименование'].astype(str)
        df_cc_tube['Марка'] = df_cc_tube['Марка'].astype(str)
        df_cc_tube['Код'] = df_cc_tube['Код'].astype(str)
        df_cc_tube['Ед. изм.'] = df_cc_tube['Ед. изм.'].astype(str)
        df_cc_tube['Количество'] = df_cc_tube['Количество'].astype(float)

        #собираем общую таблицу полученных данных
        columns = ['Наименование', 'Марка', 'Код', 'Ед. изм.', 'Количество']
        df_final = pd.DataFrame(columns=columns)
        df_final.loc[1] = ['Монтажные изделия и материалы', '', '', '', np.nan]

        # фиксируем формат общей таблицы
        df_final['Наименование'] = df_final['Наименование'].astype(str)
        df_final['Марка'] = df_final['Марка'].astype(str)
        df_final['Код'] = df_final['Код'].astype(str)
        df_final['Ед. изм.'] = df_final['Ед. изм.'].astype(str)
        df_final['Количество'] = df_final['Количество'].astype(float)

        df_final = pd.concat([df_cable, df_final, df_cc_lug, df_cc_tut, df_cc_tube])
        df_final = df_final.dropna(how='all')
        df_final.reset_index(inplace=True, drop=True)
        df_final.loc[len(df_final)] = ['Бирки маркировочные (Маркировка кабеля (21,5х43,5) 9,8 мм', '', '261-107-0961', 'шт', cable_tag]
        df_final.loc[len(df_final)] = ['Разъем RJ45 категория 6 экранированный (FTP) (для корпусов DTG-2RM и DTS-2RM)', 'RJ45', '274-702-1505 -0104', 'шт', RJ45_cnt]
        df_final.loc[len(df_final)] = ['Стяжка кабельная из синтетического материала', '9х180', '252-207-0101', 'шт', cable_tie]
        df_final.loc[len(df_final)] = ['Стяжка кабельная из синтетического материала', '9х265', '252-207-0102', 'шт', cable_tie]
        df_final.loc[len(df_final)] = ['Изолента', 'ПВХ', '247-216-1102', 'кг', 1]
        df_final.loc[len(df_final)] = ['Труба из полипропилена гибкая со структурированной стенкой диаметром 32 мм', '', '241-207-0104', 'м', NAN]
        df_final.fillna('', inplace=True)

        # выгружаем Групповую спецификацию изделий кабельной продукции в файл Excel
        wb = Workbook()
        ws = wb.active

        # заполняем ячейки из таблицы
        for r in dataframe_to_rows(df_final, index=True, header=True):
            ws.append(r)

        # зададим ширину столбцов 20 единиц
        ws.column_dimensions['A'].width  = 5
        ws.column_dimensions['B'].width  = 90
        ws.column_dimensions['C'].width  = 25
        ws.column_dimensions['D'].width  = 20
        ws.column_dimensions['E'].width  = 20
        ws.column_dimensions['F'].width  = 20

        # зададим стиль первого ряда - заголовка таблицы
        # Создание стиля шрифта - жирный
        bold_font = Font(bold=True, color="000000") # Жирный, черный

        # Создание стиля выравнивания ячеек
        center_center = Alignment(horizontal='center', vertical='center') # Центрирование значения ячейки
        right_center  = Alignment(horizontal='right', vertical='center') # Выравнивание по горизонтали - правое, во вертикали - центр
        left_center  = Alignment(horizontal='left', vertical='center') # Выравнивание по горизонтали - левое, во вертикали - центр

        # Создание стиля границы ячейки
        thin_side = Side(style='thin', color="000000")  # Черная тонкая граница
        cell_border = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side) # Черная тонкая граница по всему периметру ячейки

        # Создание стиля границы ячейки
        thin_side = Side(style='thin', color="000000")  # Черная тонкая граница
        cell_border = Border(top=thin_side, left=thin_side, right=thin_side,
                             bottom=thin_side)  # Черная тонкая граница по всему периметру ячейки

        # применение стиля к столбцам

        for cell in ws['A']:
            cell.alignment = center_center
            cell.border = cell_border
            cell.alignment = center_center

        for cell in ws['B']:
            cell.alignment = left_center
            cell.border = cell_border
            cell.alignment = Alignment(wrap_text=True)

        for cell in ws['C']:
            cell.alignment = center_center
            cell.border = cell_border
            cell.alignment = center_center

        for cell in ws['D']:
            cell.alignment = center_center
            cell.border = cell_border
            cell.alignment = center_center

        for cell in ws['E']:
            cell.alignment = center_center
            cell.border = cell_border
            cell.alignment = center_center

        for cell in ws['F']:
            cell.alignment = center_center
            cell.border = cell_border
            cell.alignment = center_center

        # применение стилей к ячейкам заголовка

        col_A1 = ws['A1']
        col_A1.font = bold_font
        col_A1.alignment = center_center

        cell_B1 = ws['B1']
        cell_B1.font = bold_font
        cell_B1.alignment = center_center

        cell_C1 = ws['C1']
        cell_C1.font = bold_font
        cell_C1.alignment = center_center

        cell_D1 = ws['D1']
        cell_D1.font = bold_font
        cell_D1.alignment = center_center

        cell_D1 = ws['E1']
        cell_D1.font = bold_font
        cell_D1.alignment = center_center

        cell_D1 = ws['F1']
        cell_D1.font = bold_font
        cell_D1.alignment = center_center

        # запись результата в буфер
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()


# --- Интерфейс Streamlit ---
st.title("📊 Выделение материалов из кабельного журнала")

# Используем колонки для красоты интерфейса
col1, col2 = st.columns(2)

# Переменные для хранения выбранных листов
sheet_a = None
sheet_b = None

with col1:
    st.subheader("📁 Файл кабельного журнала")
    file_a = st.file_uploader("Загрузите файл кабельного журнала", type=['xlsx'], key="u1")
    if file_a:
        # Получаем список листов первого файла
        sheets1 = pd.ExcelFile(file_a).sheet_names
        sheet_a = st.selectbox("Выберите лист в файле №1:", sheets1, key="s1")


with col2:
    st.subheader("📁 Файл словаря материалов")
    file_b = st.file_uploader("Загрузите файл словаря материалов", type=['xlsx'], key="u2")

st.divider()

# Проверяем, что оба файла загружены и листы выбраны
if file_a and file_b and sheet_a:
    st.success(f"Готово к работе: Лист '{sheet_a}'")

    if st.button("🚀 Начать обработку"):
        with st.spinner('Обработка кабельного журнала...'):
            try:
                # Сбрасываем указатели файлов
                file_a.seek(0)
                file_b.seek(0)

                final_result = process_data(file_a, sheet_a, file_b)

                st.download_button(
                    label="📥 Скачать результат",
                    data=final_result,
                    file_name="Materials.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Ошибка: {e}")
else:
    st.info("Пожалуйста, загрузите оба файла и выберите нужные вкладки.")