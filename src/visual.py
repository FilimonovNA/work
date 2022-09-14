from tkinter import Tk, Label, Button, filedialog
from main import generate_file


window = Tk()
window.title("Report Generator")
window.geometry('700x400')


def clicked():
    user_path = filedialog.askdirectory(title="Select a File")
    label_path_with_pictures.config(text=user_path)
    pic_path = user_path
    return pic_path


def clicked_1():
    user_path = filedialog.askdirectory(title="Select a File")
    label_path_for_report.config(text=user_path)
    rep_path = user_path
    return rep_path


def generate():
    if label_path_with_pictures['text'].find('/') == -1 or label_path_for_report['text'].find('/') == -1:
        print("ERROR")
    else:
        generate_file(label_path_with_pictures['text'], label_path_for_report['text'])


def config():
    print(window.title)


label_select_path_with_pictures = Label(window, text="Select path with pictures")
label_select_path_with_pictures.grid(column=1, row=0)

btn_select_path_with_pictures = Button(window, text="Select", command=clicked)
btn_select_path_with_pictures.grid(column=1, row=1)

label_path_with_pictures = Label(window, text="")
label_path_with_pictures.grid(column=1, row=2)

label_select_path_for_report = Label(window, text="Select path to save report")
label_select_path_for_report.grid(column=3, row=0)

btn_select_path_for_report = Button(window, text="Select", command=clicked_1)
btn_select_path_for_report.grid(column=3, row=1)

label_path_for_report = Label(window, text="")
label_path_for_report.grid(column=3, row=2)

btn_configure = Button(window, text="Configure", command=generate)
btn_configure.grid(column=5, row=4)

btn_quit = Button(window, text="Quit", command=window.destroy)
btn_quit.grid(column=10, row=10)

window.mainloop()
