import pandas as pd
import datetime
import numpy as np
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
import os
import datetime


root2=Tk()
root2.title('Display a Text File')
root2.geometry('500x250')
root2.title('Ergo Company')
root2.config(background='#FFC300')
data=[]

def open_file():
    global df
    global final
    filepath2=filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                          title="Open file okay?",
                                          filetypes=(
                                              ('All files','*.*'),
                                              ('excel files','*.xlsx'),
                                              ('csv files','*.csv'),
                                              ('All files','*.*')))

    df=pd.read_excel(filepath2,sheet_name=None)
    for x in df.keys():
            df1=pd.read_excel(filepath2,sheet_name=x)
            df1['Sheet Name'] = x
            data.append(df1)
    final = pd.concat(data, ignore_index=True)
    messagebox.showinfo('Info','Done')



def export():
    curr=datetime.datetime.now().strftime("%Y-%m-%d")
    excel_file=pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\"+ curr +" Ergo All.xlsx")
       final.to_excel(excel_file,sheet_name="All",index=False,startrow=2)

    excel_file.save()
    messagebox.showinfo('Info','Export succesfully')
    root2.destroy()


medical=Button(root2,width=10,font=('Times',12,'bold'),bd=4,text='Open File',bg='#FFC300',fg='black',padx=4,pady=4,command=open_file)
medical.place(x=50,y=170)

social=Button(root2,width=10,font=('Times',12,'bold'),bd=4,text='Export',bg='#FFC300',fg='black',padx=4,pady=4,command=export)
social.place(x=350,y=170)
logo_label=Label(root2,font=('Times',25,'bold'),bg='#FFC300',fg='black',text='Append Sheets')
logo_label.place(x=270,y=20)

root2.mainloop()




