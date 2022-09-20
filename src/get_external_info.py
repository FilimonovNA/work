import os
from tkinter import filedialog


# Открытие диалогового окна и выбор юзером папки
def select_path():
    user_path = ''
    while user_path == '':
        user_path = filedialog.askdirectory(title="Select a File")
        # Возвращает строку содержащую путь к выбранной папке
    return user_path


# Принимает название документа и путь сохраняет документ если это возможно, иначе возвращает 1
def save_report(_doc, _path):
    try:
        _doc.save(_path + '/report.docx')
    except PermissionError:
        return 1


def remove_report(_doc, _path):
    try:
        _doc.delete(_path + '/report.docx')
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


# Получаем данные для таблиц перед этажами
def get_data(_path):
    filename = _path+'/data.txt'
    if os.path.isfile(filename):
        with open(filename) as data_file:
            data = data_file.read().splitlines()
        data_file.close()
        return data
    else:
        return -1
