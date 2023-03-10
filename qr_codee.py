import qrcode
from tkinter import filedialog
import pandas as pd
from tkinter import *
from tkinter import messagebox
from fpdf import FPDF
import os
root = Tk()
root.title('Display a Text File')
root.geometry('500x250')
root.title('ergo company')
root.config(background='black')

def openFile():
    global df
    filepath = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                          title="Open file okay?",
                                          filetypes=(
                                              ('All files', '*.*'),
                                              ('excel files', '*.xlsx'),
                                              ('csv files', '*.csv')))


    df = pd.read_excel(filepath, header=0)
    print(df)
    messagebox.showinfo('Info', 'Done')


excel_file = pd.ExcelWriter("C:\\Users\\tarek\\OneDrive\\Desktop\\Python\\" + filename2 + " حضور المشرف .xlsx")


def format():

    for data in df.iloc[:,0]:
        img = qrcode.make(data)
        #img.save("C:\\Users\\hp\\Desktop\\pics\\"+data+'.png')
        img.save("C:\\Users\\tarek\\OneDrive\\Desktop\\python\\pics\\"+data+'.png')

    messagebox.showinfo('Info', 'Done')


def pdf():
    pdf = FPDF()
    pdf.set_auto_page_break(0)

    #file_path = "C:\\Users\\hp\\Desktop\\pics\\"
    file_path = "C:\\Users\\tarek\\OneDrive\\Desktop\\python\\pics\\"


    img_list = [x for x in os.listdir(file_path)]
    print(img_list)

    for img in img_list:
        pdf.add_page()
        images = "{}\\".format(file_path) + img
        pdf.image(images, w=180, h=260)

    pdf.output("{}pics.pdf".format(file_path))
    messagebox.showinfo('Info', 'Done')

logo_label = Label(root, font = ('Times', 30, 'bold'), bg = 'black', fg = 'white', text = "QRCODE")
logo_label.place(x=150, y=5)

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
open_btn = Button(root,width=10, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='black', fg='white', padx=4, pady=4, command=lambda: [openFile()])
open_btn.place(x=50, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Format', bg='black', fg='white', padx=4, pady=4, command=format)
append_btn.place(x=180, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='PDF', bg='black', fg='white', padx=4, pady=4, command=pdf)
append_btn.place(x=310, y=150)

root.mainloop()

