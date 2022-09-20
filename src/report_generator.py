from tkinter import Tk, Label, Button, filedialog
from word_file_generator import generate_file


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
        btn_configure.config(text="Select both paths")
    else:
        btn_configure.config(text=generate_file(label_path_with_pictures['text'], label_path_for_report['text']))


def config():
    print(window.title)


label_select_path_with_pictures = Label(window, text="Select path with pictures", font=13)
label_select_path_with_pictures.place(x=10, y=10, width=200, height=20)

btn_select_path_with_pictures = Button(window, text="Select", command=clicked)
btn_select_path_with_pictures.place(x=65, y=40, width=80, height=20)

label_path_with_pictures = Label(window, text="Your path will be here")
label_path_with_pictures.place(x=10, y=70, width=200, height=20)

label_select_path_for_report = Label(window, text="Select path to save report", font=13)
label_select_path_for_report.place(x=260, y=10, width=200, height=20)

btn_select_path_for_report = Button(window, text="Select", command=clicked_1)
btn_select_path_for_report.place(x=310, y=40, width=80, height=20)

label_path_for_report = Label(window, text="Your path will be here")
label_path_for_report.place(x=260, y=70, width=200, height=20)

btn_configure = Button(window, text="GENERATE\nREPORT", command=generate, font=('Times', 35), fg='#f5f5f5', background="#66b0ab")
btn_configure.place(x=10, y=100, width=680, height=200)

btn_quit = Button(window, text="Quit", command=window.destroy)
btn_quit.place(x=620, y=360, width=70, height=30)

window.mainloop()
