from tkinter import *
from tkinter import filedialog

window = Tk()
window.title("Report Generator")
window.geometry('700x400')

def clicked():
    user_path = filedialog.askdirectory(title="Select a File")
    label_path.configure(text = user_path)
    print(label_path['text'])

label_select_path = Label(window, text = "Select path with pictures")
label_select_path.grid(column=1, row=0)

btn_select_path = Button(window, text="Select", command=clicked)
btn_select_path.grid(column=1, row=2)

label_path = Label(window, text="Yours path will be here")
label_path.grid(column=1, row=4)


window.mainloop()
