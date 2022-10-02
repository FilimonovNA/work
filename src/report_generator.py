from tkinter import Tk, Label, Button, filedialog
from word_file_generator import generate_file

NO_SYMBOL_ID = -1
WINDOW_WIDTH = 700
WINDOW_HEIGHT = 400

window = Tk()
window.title("Report Generator")
window.geometry(f'{WINDOW_WIDTH}x{WINDOW_HEIGHT}')


def select_pictures_path():
    user_path = filedialog.askdirectory(title="Select a File")
    label_path_with_pictures.config(text=user_path)
    pic_path = user_path
    return pic_path


def select_report_path():
    user_path = filedialog.askdirectory(title="Select a File")
    label_report_path.config(text=user_path)
    rep_path = user_path
    return rep_path


def generate():
    if label_path_with_pictures['text'].find('/') == NO_SYMBOL_ID or \
            label_report_path['text'].find('/') == NO_SYMBOL_ID:
        btn_configure.config(text="Select both paths")
    else:
        btn_configure.config(
            text=generate_file(label_path_with_pictures['text'], label_report_path['text']))


def config():
    print(window.title)


label_select_pictures_path = Label(window, text="Select path with pictures", font=13)
label_select_pictures_path.place(x=10, y=10, width=200, height=20)

btn_select_pictures_path = Button(window, text="Select", command=select_pictures_path)
btn_select_pictures_path.place(x=65, y=40, width=80, height=20)

label_path_with_pictures = Label(window, text="Your path will be here")
label_path_with_pictures.place(x=10, y=70, width=200, height=20)

label_select_report_path = Label(window, text="Select path to save report", font=13)
label_select_report_path.place(x=260, y=10, width=200, height=20)

btn_select_report_path = Button(window, text="Select", command=select_report_path)
btn_select_report_path.place(x=310, y=40, width=80, height=20)

label_report_path = Label(window, text="Your path will be here")
label_report_path.place(x=260, y=70, width=200, height=20)

btn_configure = Button(window, text="GENERATE\nREPORT", command=generate, font=('Times', 15), fg='#f5f5f5', background="#66b0ab")
btn_configure.place(x=520, y=270, width=130, height=60)



btn_quit = Button(window, text="Quit", command=window.destroy)
btn_quit.place(x=620, y=360, width=70, height=30)

window.mainloop()
