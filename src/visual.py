from tkinter import *
from tkinter import filedialog

window = Tk()
window.title("Report Generator")
window.geometry('700x400')


def clicked():
    user_path = filedialog.askdirectory(title="Select a File")
    #label_path.configure(show=user_path)
    label_path.config(text=user_path)
    window.title(user_path)


def config():
    print(window.title)


user_path = ""
label_select_path = Label(window, text = "Select path with pictures")
label_select_path.grid(column=1, row=0)

btn_select_path = Button(window, text="Click", command=clicked)
btn_select_path.grid(column=1, row=1)

label_path = Label(window, text = "path")
label_path.grid(column=1, row=2)

btn_configure = Button(window, text="Configure", command=config)
btn_configure.grid(column=2, row=1)

window.mainloop()
