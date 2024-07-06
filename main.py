import os
import sys
from docxtpl import DocxTemplate
import xlwings as xw
from collections import OrderedDict

# Параметры для подстановки в шаблоны
nomergruppy = 3253
mestopraktiki = 'Университет "Дубна"'
datanachala = '24.06.2024'
dataokonchaniya = "07.07.2024"
god = 2024
dataotcheta = "06.07.2024"
zavkafedroy = "Кореньков Владимир Васильевич"
weeks = 2
направлен_value = None
кафедра_value = None
форма_обучения_value = None
column_name = None
row_number = None
search_string_name = "Наименование"
search_string_block2 = "Блок 2.Практика"
search_string_block3 = "Преддипломная практика"

# Установка рабочей директории на директорию скрипта
os.chdir(sys.path[0])

# Создание необходимых директорий
os.mkdir('титул_дневник')
os.mkdir('шаблон_характеристика')
os.mkdir('аттестационный_лист')

# Открытие книги Excel
wb = xw.Book('Очный_план.xlsx')

# Лист с титулом
titulSheet = wb.sheets['Титул']

# Поиск ячейки с текстом 'направлен'
for cell in titulSheet.used_range:
    cell_value = cell.value
    if 'направлен' in str(cell_value):
        направлен_value = cell_value
        break

# Поиск ячейки с текстом 'Кафедра'
for cell in titulSheet.used_range:
    cell_value = cell.value
    if 'Кафедра ' in str(cell_value):
        кафедра_value = cell_value
        break

# Поиск ячейки с текстом 'Форма обучения'
for cell in titulSheet.used_range:
    cell_value = cell.value
    if 'Форма обучения' in str(cell_value):
        форма_обучения_value = cell_value
        break

# Извлечение значений из найденных ячеек
направлен_value = направлен_value.split("подготовки ")[1]
форма_обучения_value = форма_обучения_value.split("Форма обучения: ")[1]

# Лист с планом практик
sheet = wb.sheets['ПланСвод']

# Поиск колонки с 'Наименование'
used_range = sheet.used_range
for cell in used_range:
    if search_string_name in str(cell.value):
        column_name = cell.column
        break

# Поиск строки с 'Блок 2.Практика'
for cell in used_range:
    if search_string_block2 in str(cell.value):
        row_number = cell.row
        break

# Проверка, что оба заголовка найдены
if column_name is None or row_number is None:
    print("Не удалось найти указанные заголовки в листе.")
    sys.exit(1)

# Инициализация массива для хранения данных
practices = []

# Сбор данных ниже указанной ячейки до 'Блок 3'
current_row = row_number + 1
while True:
    cell_value = sheet.range((current_row, column_name)).value
    if cell_value and search_string_block3 in str(cell_value):
        practices.append([cell_value, sheet.range((current_row, column_name)).address])
        break
    if cell_value and not sheet.range((current_row, column_name)).api.MergeCells:
        practices.append([cell_value, sheet.range((current_row, column_name)).address])
    current_row += 1

# Поиск колонки с 'По плану'
for cell in used_range:
    if 'По плану' in str(cell.value):
        column_hour = cell.column
        break

# Счетчик для имен файлов
name_counter = 1

# Обработка каждой практики
for practice in practices:
    practice_name = practice[0]
    cell_range = practice[1]
    
    # Получаем букву столбца из диапазона ячейки
    column_letter = cell_range[1]
    
    # Вычисляем номер строки из диапазона ячейки
    row_number = int(cell_range[3:])
    
    # Вычисляем адрес ячейки слева
    index_cell = f"{chr(ord(column_letter) - 1)}{row_number}"

    # Получаем значение из левой ячейки
    index = sheet.range(index_cell).value

    # Лист с компетенциями
    kompSheet = wb.sheets('Компетенции(2)')
    competencies_cell = None
    for cell in kompSheet.used_range:
        if cell.value == index:
            competencies_cell = cell
            break

    if competencies_cell:
        if competencies_cell.column + 3:
            target_cell = competencies_cell.offset(column_offset=3)
            competencien = target_cell.value
    competencien_list = [comp.strip() for comp in competencien.split(';')]
    unique_competencies = list(OrderedDict.fromkeys([item[:-2] for item in competencien_list]))

    # Данные для заполнения в таблице Word
    table_data = []

    compSheet = wb.sheets('Компетенции')
    for comp_number in unique_competencies:
        comp_cell = None
        for cell in compSheet.used_range:
            if cell.value == comp_number:
                comp_cell = cell
                break

        if comp_cell:
            if comp_cell.column + 3:
                target_cell = comp_cell.offset(column_offset=3)
                comp_value = target_cell.value
                table_data.append({"col1": comp_number, "col2": comp_value})

    # Определение типа практики
    if 'У' in str(index):
        practice_type = 'учебной'
    else:
        practice_type = 'производственной'

    # Получение семестра
    right_cell = f"{chr(ord(column_letter) + 1)}{row_number}"
    semester = sheet.range(right_cell).value 

    while semester is None:
        column_letter = chr(ord(column_letter) + 1)
        right_cell = f"{column_letter}{row_number}"
        semester = sheet.range(right_cell).value

    curse = (int(semester) // 2 + int(semester) % 2)

    # Получение значения часов
    hour_cell = f"{chr(column_hour+64)}{row_number}"
    hour_value = sheet.range(hour_cell).value

    # Заполнение шаблона дневника и титула
    context = {'naimenovanie':practice_name, 'vidpraktiki': practice_type, 'nomerkursa': curse, 'kodnaprapleniya':направлен_value, 
               'formaobucheniya':форма_обучения_value, 'kafedra': кафедра_value, 'itogochasov': hour_value, 'nomergruppy': nomergruppy, 
               'mestopraktiki':mestopraktiki, 'datanachala':datanachala, 'dataokonchaniya':dataokonchaniya, 'god': god, 
               'dataotcheta':dataotcheta, 'zavkafedroy':zavkafedroy}
    doc = DocxTemplate('Шаблон_дневника_и_титула_enwords.docx')
    doc.render(context)

    # Сохранение документа дневника и титула
    docName = 'титул_дневник/Шаблон_дневника_и_титула ' + str(name_counter) + ' практики.docx'
    doc.save(docName)

    # Изменение формы обучения для характеристики
    форма_обучения_value = форма_обучения_value[:-2] + 'ой'

    # Заполнение шаблона характеристики
    context2 = {'rows': table_data, 'naimenovanie':practice_name, 'nomerkursa': curse, 'formaobucheniya':форма_обучения_value, 
                'kodnaprapleniya':направлен_value, 'itogochasov': hour_value, 'weeks': weeks}

    tpl = DocxTemplate('Шаблон_Характеристика.docx')
    tpl.render(context2)

    table = tpl.tables[0]

    for row_data in table_data:
        new_row = table.add_row()
        new_row.cells[0].text = str(row_data["col1"])
        new_row.cells[1].text = str(row_data["col2"])

    # Сохранение документа характеристики
    tpl.save('шаблон_характеристика/Шаблон_характеристика '+ str(name_counter) + ' практики.docx')

    # Заполнение шаблона аттестационного листа
    context3 = {'rows': table_data, 'naimenovanie':practice_name, 'nomerkursa': curse, 'formaobucheniya':форма_обучения_value, 
                'kodnaprapleniya':направлен_value, 'itogochasov': hour_value, 'weeks': weeks,'nomergruppy': nomergruppy, 
                'kafedra': кафедра_value }
    tpl_attestation = DocxTemplate('Шаблон_Аттестиционный_лист.docx')
    tpl_attestation.render(context3)
    
    table1 = tpl_attestation.tables[0]
    table2 = tpl_attestation.tables[1]

    for row_data in table_data:
        new_row = table1.add_row()
        new_row.cells[0].text = str(row_data["col1"]) + " " + str(row_data["col2"])

    for row_data in table_data:
        empty_cell_found = False
        for row in table2.rows:
            if row.cells[0].text == "":
                row.cells[0].text = str(row_data["col1"]) + " " + str(row_data["col2"])
                empty_cell_found = True
                break
        if not empty_cell_found:
            new_row = table2.add_row()
            new_row.cells[0].text = str(row_data["col1"]) + " " + str(row_data["col2"])

    # Сохранение документа аттестационного листа
    tpl_attestation.save('аттестационный_лист/Шаблон_аттестации '+ str(name_counter) + ' практики.docx')

    name_counter += 1

# Закрытие рабочей книги Excel
wb.close()
