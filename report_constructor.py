import os
from tkinter import filedialog
from docx import Document
from docx.shared import Pt, Cm

'''
дописать функции floor_pictures, add_table, construct_floor
'''


def select_path():
    user_path = ''
    while user_path == '':
        user_path = filedialog.askdirectory(title="Select a File")
    return user_path


def get_pictures_list(user_path):
    all_files = os.listdir(user_path)
    list_of_pictures = []
    for file in all_files:
        if file[-4:] == '.jpg' or file[-4:] == '.png':
            list_of_pictures.append(file)
    return list_of_pictures


def add_picture_in_file(picture, doc):
    doc.add_picture(picture + '/' + picture, width=Pt(500))


def margin_set(doc):
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1)


def get_floor_list(pictures_list):
    floor_list = []
    for elem in pictures_list:
        if elem[1].isdigit():
            floor_num = elem[:2]
        else:
            floor_num = elem[0]
        if floor_num not in floor_list:
            floor_list.append(floor_num)
    return sorted(floor_list)


def add_floor_title_in_file(floor, doc):
    floor_title = doc.add_paragraph().add_run(f'{floor} этаж')
    floor_title.font.name = 'Times new roman'
    floor_title.font.size = Pt(16)
    floor_title.bold = True
    doc.add_paragraph()


def floor_pictures(floor, pictures_list):
    # если Начало совпадает с номеро этажа то добавляем
    # возвращаем список картинок
    return floor_list


def construct_floor(report_doc, floor, floor_pictures_list):
    add_floor_title_in_file(floor, report_doc)
    add_table()
    for picture in floor_pictures_list:
        add_picture_in_file(picture, report_doc)


def main_report_constructor():
    path = 'C:/Users/PC/Desktop/Work/'  # legacy for save time
    report_doc = Document()
    margin_set(report_doc)
    path_with_pictures = path + '/Pictures'
    # path_with_pictures = select_path()
    all_pictures = get_pictures_list(path_with_pictures)
    all_floors = get_floor_list(all_pictures)
    for floor in all_floors:
        floor_pictures_list = floor_pictures(floor, all_pictures)  # получаем список картинок для конкретного этажа
        construct_floor(report_doc, floor, floor_pictures_list)  # создаем конструкцию первого этажа
    # report_path = select_path()
    report_path = path

    try:
        report_doc.save(report_path + '/test.docx')
        print('SUCCESS')
    except PermissionError:
        print('Close the file pls')


main_report_constructor()
