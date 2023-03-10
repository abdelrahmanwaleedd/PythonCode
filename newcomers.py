from tkinter import filedialog
import pandas as pd
from tkinter import *
from tkinter import messagebox
import os
import datetime
import numpy as np
from ttkwidgets.autocomplete import AutocompleteCombobox
import warnings
from datetime import date


root = Tk()
root.title('Display a Text File')
root.geometry('600x400')
root.title('Ergo Company')
root.config(background='#AED6F1')
y = IntVar()
var1 = StringVar()
data = []
listt=[]

def openFile():
        global filepathm, filename2
        warnings.filterwarnings("ignore")

        filepath = filedialog.askopenfilenames(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?", filetypes=(
                ('All files', '*.*'), ('excel files', '*.xlsx'), ('csv files', '*.csv'), ('All files', '*.*')))

        for file in filepath:
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            try:
                df = pd.read_excel(file, header=None, sheet_name="Attendance")
            except:
                messagebox.showerror('Python Error', 'Attendance sheet doesnt exist in ' + filename2)

            try:
                head = df[df.apply(lambda row: row.astype(str).str.contains('قوم|بطاق').any(), axis=1)].index.values
                df3 = pd.read_excel(file, header=head, sheet_name="Attendance")
            except:
                # head = df[df.apply(lambda row: row.astype(str).str.contains('قوم|بطاق').any(), axis = 1)].index.values
                df3 = pd.read_excel(file, header=0, sheet_name="Attendance")

            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            print(filename2)
            df3["المصنع"] = filename2


            data.append(df3)

        messagebox.showinfo('Info', len(filepath))

def append_data():
    global final
    final = pd.concat(data, ignore_index=True)
    final = final[final.columns.drop(list(final.filter(regex='Unna')))]
    final.drop(final[final['القومى'] == '$'].index, inplace=True)
    final.drop(final[final['القومى'] == 'S'].index, inplace=True)
    final.drop(final[final['القومى'] == '2'].index, inplace=True)
    final.drop(final[final['الاسم'] == 4].index, inplace=True)
    final.drop(final[final['القومى'] == '-'].index, inplace=True)
    final.dropna(subset=['القومى'], inplace=True)
    final["القومى"] = pd.to_numeric(final["القومى"])

    final = final[["القومى","تاريخ التعيين","الاسم","المصنع","الوردية"]]

    final['char_length'] = final['المصنع'].str.len()

    conditions = [final['char_length'] < 20, (final['char_length'] > 20) & (final['char_length'] < 38), final['char_length'] > 38]

    values = [final['المصنع'].str[2:], final['المصنع'].str[7:], final['المصنع'].str[17:]]

    final["المصنع"] = np.select(conditions, values)

    messagebox.showinfo('Info', 'Done')

def new():
    global final
    mon = date.today().month

    final['year'] = pd.DatetimeIndex(final["تاريخ التعيين"]).year
    final['month'] = pd.DatetimeIndex(final["تاريخ التعيين"]).month

    final["year"] = final["year"].fillna(0)
    final["year"] = final.year.astype(int)
    final["month"] = final["month"].fillna(0)
    final["month"] = final.month.astype(int)
    final=final[final["year"] ==2022]
    final= final[(final["month"] == mon) | (final["month"] == mon-1)]

    final.drop(labels=['char_length', 'year', 'month'], axis=1,inplace=True)


    messagebox.showinfo('Info', 'Done')

def export():
    #excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\Newcomers.xlsx")
    #excel_file = pd.ExcelWriter("C:\\Users\\hp\\Desktop\\Newcomers.xlsx")
    excel_file = pd.ExcelWriter("C:\\Users\\BOTA\\Desktop\\Newcomers.xlsx")

    for x in final["المصنع"]:
        final2 = final[final["المصنع"] == x]
        final2.to_excel(excel_file, sheet_name=x, index=False)

    excel_file.save()
    messagebox.showinfo('Info', 'export sucessfully')
    root.destroy()


# ------------------------------------------------------------------------------------------------------------------------------------------------------------
f1 = Frame(root, bg='#30475E', width=800, height=150)
f1.place(x=0, y=0)
f2 = Frame(root, bg='#30475E', width=800, height=200)
f2.place(x=0, y=250)
logo_label = Label(f1, font=('Times', 30, 'bold'), bg='#30475E', fg='black', text='Detect Newcomers')
logo_label.place(x=150, y=5)

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
openf = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='#222831', fg='white',
               padx=4, pady=4, command=lambda: [openFile()])
openf.place(x=20, y=300)
append_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Append files', bg='#222831',
                    fg='white', padx=4, pady=4, command=append_data)
append_btn.place(x=160, y=300)

neww = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Report', bg='#222831',
                    fg='white', padx=4, pady=4, command=new)
neww.place(x=300, y=300)

export_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Export to xlsx', bg='#222831',
                    fg='white', padx=4, pady=4, command=export)
export_btn.place(x=440, y=300)


root.mainloop()
