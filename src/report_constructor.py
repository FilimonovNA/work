import os
from tkinter import filedialog
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT


# Открытие диалогового окна и выбор юзером папки
def select_path():
    user_path = ''
    while user_path == '':
        user_path = filedialog.askdirectory(title="Select a File")
        # Возвращает строку содержащую путь к выбранной папке
    return user_path


# Принимает на вход название документа и путь сохраняет документ если это возможно, иначе возвращает 1
def save_report(_doc, _path):
    try:
        _doc.save(_path + '/test.docx')
    except PermissionError:
        return 1


# На основании полученного на вход пути возвращает список строк содержащих названия картинок в папке
def get_pictures_list(_path):
    all_files = os.listdir(_path)
    list_of_pictures = []
    for file in all_files:
        if file[-4:] == '.jpg' or file[-4:] == '.png':
            list_of_pictures.append(file)
    return list_of_pictures


# Настройка полей для документа, можно добавить доп настройки: размер документа, ориентация и т.д.
def set_margin(doc):
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1)


# Возвращает номер этажа, на котором будет располагаться данная картинка
def get_picture_number_of_floor(picture):
    if picture[1].isdigit():
        floor_num = picture[:2]
    else:
        floor_num = picture[0]
    return floor_num


# Возвращает порядковый номер скрипта картинки для функции add_picture_caption
def get_script_number_of_picture(picture):
    script_num = picture[picture.find("fl")+3:picture.find("fl")+5]
    return script_num


# Возвращает отсортированный список этажей
def get_floor_list(pictures_list):
    floor_list = []
    for elem in pictures_list:
        floor_num = get_picture_number_of_floor(elem)
        if floor_num not in floor_list:
            floor_list.append(floor_num)
    return sorted(floor_list)


# Возвращает список картинок для конкретного этажа
def get_list_of_floor_pictures(floor, _all_pictures):
    pictures_list = []
    for elem in _all_pictures:
        pic_num = get_picture_number_of_floor(elem)
        if str(floor) == pic_num:
            pictures_list.append(elem)
    return pictures_list


# Добавление нижнего колонтитула в файл, на вход принимает документ
def add_footer_in_doc(doc):
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]

    # Должен вызывать функцию для ввода информации пользователем {site_ID} {location} {full_address}
    site_id = '25654'
    location = 'БЦ Юникон'
    full_address = 'Москва, Плеханова, д.4А'
    footer_text = footer_para.add_run(f'Indoor, SiteID - {site_id}\n{location}\n{full_address}')
    footer_text.font.name = 'Times new roman'
    footer_text.font.size = Pt(9)


# Добавляет верхний колонтитул
def add_header_in_doc(doc):
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0]
    header_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    header_text = header_para.add_run(f'Отчет по результатам проведения\n'
                                      f'Indoor-измерений в сети ПАО «МегаФон»')

    # Нужно добавить лого слева
    header_text.font.name = 'Times new roman'
    header_text.font.size = Pt(11)


# Добавляет заголовок этажа и форматирует его
def add_floor_title_in_file(floor, doc):
    floor_title = doc.add_heading().add_run(f'{floor} этаж')
    floor_title.font.name = 'Times new roman'
    floor_title.font.size = Pt(16)
    floor_title.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph()


# Добавляет таблицу со значениями для каждого этажа
# Необходима доработка для вычисляемых значений из немо(чтение из файла)
def add_table_with_values(doc):
    rows = 4
    cols = 4
    table = doc.add_table(rows=rows, cols=cols)
    table.rows[0].cells[1].text = '2G'
    table.rows[0].cells[2].text = '3G'
    table.rows[0].cells[3].text = '4G'
    table.rows[1].cells[0].text = 'Средний уровень сигнала, дБм'
    table.rows[2].cells[0].text = 'Cредняя скорость DL, Мб/c'
    table.rows[3].cells[0].text = 'Cредняя скорость UL, Мб/c'
    table.rows[1].cells[1].text = 'Rx Level'                # Rx level
    table.rows[2].cells[1].text = '-'                       # DL 2G
    table.rows[3].cells[1].text = '-'                       # UL 2G
    table.rows[1].cells[2].text = 'RSCP'                    # RSCP
    table.rows[2].cells[2].text = 'DL 3G'                   # DL 3G
    table.rows[3].cells[2].text = 'UL 3G'                   # UL 3G
    table.rows[1].cells[3].text = 'RSRP'                    # RSRP
    table.rows[2].cells[3].text = 'DL 4G'                   # DL 2G
    table.rows[3].cells[3].text = 'UL 4G'                   # UL 2G
    doc.add_paragraph()
    adjust_table_with_values(table, rows, cols)


# Добавляет 1 картинку файл
def add_picture_in_file(_path, doc, picture):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = paragraph.add_run("")
    run.add_picture(_path + '/' + picture, width=Pt(480))


# Форматирование таблицы со значениями
def adjust_table_with_values(tab, rows, cols):
    tab.alignment = WD_TAB_ALIGNMENT.CENTER
    for cell in tab.columns[0].cells:
        cell.width = Cm(7.5)
    for i in range(rows):
        for j in range(1, cols):
            tab.cell(i, j).paragraphs[0].paragraph_format.alignment = WD_TAB_ALIGNMENT.CENTER


# Добавление названия к картинке
def add_picture_caption(picture, doc, serial_number):
    picture_captions = {'01': 'Качественные показатели покрытия 2G', '02': 'Качественные показатели покрытия 3G',
                        '03': 'Качественные показатели покрытия 4G', '04': 'Скорость ППД DL/UL 3G',
                        '05': 'Скорость ППД DL/UL 4G', '06': 'Функциональные показатели CSFB',
                        '07': 'Функциональные показатели LTE Carrier Aggregation',
                        '08': 'Функциональные показатели LTE MIMO', '09': 'Функциональные показатели 2G indoor scanner',
                        '10': 'Функциональные показатели 3G indoor scanner',
                        '11': 'Функциональные показатели LTE indoor scanner'
                        }
    script_number_of_picture = get_script_number_of_picture(picture)
    picture_caption = picture_captions.get(script_number_of_picture)
    paragraph = doc.add_paragraph('')
    text = paragraph.add_run(f'Рисунок {serial_number} - {picture_caption}')

    # Форматирование подписи
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    text.font.name = 'Times new roman'
    text.font.size = Pt(14)


# Добавление подписи для раздела измерений со сканера
def add_scanner_title(doc):
    doc.add_page_break()
    floor_title = doc.add_heading().add_run('Функциональные показатели со сканера')
    floor_title.font.name = 'Times new roman'
    floor_title.font.size = Pt(14)
    floor_title.font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph()


# Полная "Сборка" 1 этажа для отчета
def add_floor_in_report(_path, doc, floor, floor_pictures_list, picture_number_in_file):
    add_floor_title_in_file(floor, doc)
    add_table_with_values(doc)
    is_scanner_title_was_add = 0
    for picture in floor_pictures_list:
        script_number = get_script_number_of_picture(picture)
        if script_number in ['09', '10', '11'] and is_scanner_title_was_add == 0:
            add_scanner_title(doc)
            is_scanner_title_was_add = 1
        add_picture_in_file(_path, doc, picture)
        add_picture_caption(picture, doc, picture_number_in_file)
        picture_number_in_file += 1
    doc.add_page_break()
    return picture_number_in_file


# Main, в котором происходит сборка всего отчета и вызов функций
def main():
    path = 'C:/Users/PC/Desktop/Work/'  # legacy for save time
    report_doc = Document()
    set_margin(report_doc)
    path_with_pictures = path + '/Pictures'
    # path_with_pictures = select_path()
    all_pictures = get_pictures_list(path_with_pictures)
    all_floors = get_floor_list(all_pictures)
    # report_path = select_path()
    report_path = path

    add_footer_in_doc(report_doc)
    add_header_in_doc(report_doc)

    if save_report(report_doc, report_path) != 1:
        absolut_picture_number_in_file = 1
        for current_floor in all_floors:
            current_floor_pictures_list = get_list_of_floor_pictures(current_floor, all_pictures)
            absolut_picture_number_in_file = add_floor_in_report(path_with_pictures, report_doc, current_floor,
                                                                 current_floor_pictures_list,
                                                                 absolut_picture_number_in_file)
        save_report(report_doc, report_path)
        print("SUCCESS")
    else:
        print("CLOSE FILE")


# Точка входа в программу
main()
