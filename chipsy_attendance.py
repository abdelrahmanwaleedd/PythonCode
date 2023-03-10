from tkinter import filedialog
import pandas as pd
from tkinter import *
from tkinter import messagebox
import numpy as np
import tkinter as tk
import os
import datetime
import warnings
pd.options.mode.chained_assignment = None
root = Tk()
root.title('Display a Text File')
root.geometry('700x250')
root.title('ergo company')
root.config(background='black')
new = []
date2 = []
var2 = tk.StringVar()

def openFile():
        global df1, filename2, tasgel
        filepath = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                              title="Open file okay?",
                                              filetypes=(
                                                  ('All files', '*.*'),
                                                  ('excel files', '*.xlsx'),
                                                  ('csv files', '*.csv')))

        df5 = pd.read_excel(filepath, header=None, sheet_name='تسجيل الحضور')
        y = df5[df5.apply(lambda row: row.astype(str).str.contains('تاري').any(), axis=1)].index.values
        df = pd.read_excel(filepath, header=y, sheet_name='تسجيل الحضور')

        tasgel = pd.read_excel(filepath, sheet_name='الأسماء')
        tasgel.dropna(subset=['الرقم القومى'],inplace=True)

        tasgel_twakof = pd.read_excel(filepath, sheet_name='تسجيل التوقف')
        df = df.append(tasgel_twakof)
        print(df)
        #tasgel["الرقم القومى"] = pd.to_numeric(tasgel["الرقم القومى"])
        df['الإسم'] = df['الإسم'].apply(int)
        filename = os.path.basename(filepath)
        filename2, extension = os.path.splitext(filename)
        print(filename2)

        for x in df.iloc[:, 0].unique():
            ts = pd.to_datetime(str(x))
            d = ts.strftime('%m/%d/%y')
            date2.append(d)

        print(tasgel)

        def choose_one(event):
            global df1
            if var2.get() != 0:
                df1 = df[df.iloc[:, 0] >= var2.get()]
                print(df1)
                messagebox.showinfo('Info', 'Done')

        sheets_droplist = OptionMenu(root, var2, *(date2), command=choose_one)
        sheets_droplist.place(x=20, y=200)
        sheets_droplist.config(width=10)

def format():
        global df1
        global df2
        global df3
        df2 = pd.DataFrame()
        df2["القومي"] = df1.iloc[:, 1]
        df2.drop_duplicates(subset="القومي", inplace=True)
        for x in df1.iloc[:, 0].unique():
            new.append(x)
            print(new)

        df3 = pd.DataFrame()
        df3["القومي"] = df1.iloc[:, 1]

        def choose_one():
            global df3, df2
            for day in range(len(new)):

                try:
                    conditions = [(df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'A12'),
                                  (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'A8'),
                                  (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'B12'),

                                  (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'B8'),
                                  (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'C8'),
                                    (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'S1'),
                                    (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'S2'),
                                  (df1["التاريخ"] == new[day]) & (df1['الوردية'] == 'C12')]

                    values = ["A12", "A8", "B12", "B8", "C8","S1","S2","C12"]
                    df3[new[day]] = np.select(conditions, values, default='')
                except:
                    try:
                        conditions = [(df1["التاريخ"] == new[day]) & (df1['الوظيفة'] == 'مساعد بيع'),
                                      (df1["التاريخ"] == new[day]) & (df1['الوظيفة'] == 'عامل نظافة'),
                                      (df1["التاريخ"] == new[day]) & (df1['الوظيفة'] == 'فنى كهرباء'),
                                      (df1["التاريخ"] == new[day]) & (df1['الوظيفة'] == 'الجملة'),
                                      (df1["التاريخ"] == new[day]) & (df1['الوظيفة'] == 'المباشر'), ]
                        values = ["B", "C", "A", "A", "A"]
                        df3[new[day]] = np.select(conditions, values, default='')
                    except:

                        conditions = [df1["التاريخ"] == new[day]]
                        values = ["A"]
                        df3[new[day]] = np.select(conditions, values, default='')

            df3 = df3.drop_duplicates().groupby('القومي', sort=False, as_index=False).sum()

            # df3=df3.groupby('القومي').sum()
            A8 = df3.apply(lambda row: sum(row[0:] == 'A8'), axis=1)
            A12 = df3.apply(lambda row: sum(row[0:] == 'A12'), axis=1)


            B8 = df3.apply(lambda row: sum(row[0:] == 'B8'), axis=1)
            B12 = df3.apply(lambda row: sum(row[0:] == 'B12'), axis=1)

            C = df3.apply(lambda row: sum(row[0:] == 'C8'), axis=1)
            C12 = df3.apply(lambda row: sum(row[0:] == 'C12'), axis=1)
            S1 = df3.apply(lambda row: sum(row[0:] == 'S1'), axis=1)
            S2 = df3.apply(lambda row: sum(row[0:] == 'S2'), axis=1)

            df3["total days"] = C + B8 + B12 + A12 + A8 + C12
            df3["total stoppage"] = S1+S2

            df2 = pd.merge(df2, df3, on="القومي")

        choose_one()

        messagebox.showinfo('Info', 'Done')

def current():
        global filepath2, dtt, df10, df11, df66
        filepath2 = filedialog.askopenfilenames(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                                title="Open file okay?",
                                                filetypes=(
                                                    ('All files', '*.*'),
                                                    ('excel files', '*.xlsx'),
                                                    ('csv files', '*.csv'),
                                                    ('All files', '*.*')))

        for file in filepath2:
            dtt = pd.read_excel(file, header=None)
            y = dtt[dtt.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
            dtt = pd.read_excel(file, header=y)
            dtt.dropna(subset=['م'], inplace=True)
            dtt.drop(dtt[dtt['م'] == '$'].index, inplace=True)
            dtt.drop(dtt[dtt['م'] == 'S'].index, inplace=True)

            dtt.drop(dtt[dtt['م'] == '&'].index, inplace=True)
            dtt.drop(dtt[dtt['م'] == 'م'].index, inplace=True)
            dtt.drop(dtt[dtt['م'] == 0].index, inplace=True)
            dtt["القومى"] = pd.to_numeric(dtt["القومى"])
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            dtt["المصنع"] = filename2
            dtt = dtt.iloc[:, [0, 1, 2, 3, 4]]

        print(tasgel)
        df10 = pd.merge(dtt, df2, left_on='القومى', right_on='القومي', how='left')
        df11 = pd.merge(dtt, df2, left_on='القومى', right_on='القومي', how='right')
        df11 = df11[df11['القومى'].isna()]

        df11["القومي"] = pd.to_numeric(df11["القومي"])

        df66 = pd.merge(df11, tasgel, left_on='القومي', right_on='الرقم القومى', how='left')


        messagebox.showinfo('Info', "Done")

def sections():
    global section, final
    section = df1.copy()
    conditions=[(df1['القسم']=="مخزن المنتج التام"),(df1['القسم']=="فرز البطاطس"),(df1['القسم']=="تعبئة خط 5 TC"),
                (df1['القسم']=="المخازن المبردة"),(df1['القسم']=="تعبئة خط 6"),(df1['القسم']=="تعبئة خط 5 PC"),
                (df1['القسم']=="تصنيع TC"),(df1['القسم']=="تعبئة خط 9"),(df1['القسم']=="تعبئة خط 8"),(df1['القسم']=="تعبئة خط 4"),
                (df1['القسم']=="تعبئة خط 7"),(df1['القسم']=="تصنيع خط 3"),(df1['القسم']=="تعبئة خط 2"),(df1['القسم']=="الشئون الادارية"),
                (df1['القسم']=="المخلفات"),(df1['القسم']=="تصنيع خط 1"),(df1['القسم']=="تعبئة خط 1"),(df1['القسم']=="تعبئة خط 3"),
                (df1['القسم']=="تصنيع خط 2")]
    values=["Non Production labor","Production Labor Cairo","Production Labor Cairo","Non Production labor","Production Labor Cairo",
            "Production Labor Cairo","Production Labor Cairo","Production Labor Cairo","Production Labor Cairo","Production Labor Cairo",
            "Production Labor Cairo","Production Labor Cairo","Production Labor Cairo","Non Production labor","Non Production labor",
            "Production Labor Cairo","Production Labor Cairo","Production Labor Cairo","Production Labor Cairo"]
    section['الوظيفة'] = np.select(conditions,values,default='')
    section = section.rename(columns={'الإسم': 'القومي'})
    #------------------------------------------------------------------------------------------------------------
    prod = section[["القومي","الوظيفة"]]
    prod['Production Labor Cairo'] = np.where(section['الوظيفة'] == "Production Labor Cairo", 1, 0)
    prod = prod.groupby(['القومي'], as_index=False)['Production Labor Cairo'].sum()

    non = section[["القومي","الوظيفة"]]
    non['Non Production labor'] = np.where(section['الوظيفة'] == "Non Production labor", 1, 0)
    non = non.groupby(['القومي'], as_index=False)['Non Production labor'].sum()

    empty = section[["القومي","الوظيفة"]]
    empty['Empty'] = np.where((((section['الوظيفة'] != "Non Production labor") & (section['الوظيفة'] != "Production Labor Cairo"))), 1, 0)
    empty = empty.groupby(['القومي'], as_index=False)['Empty'].sum()

    final = pd.concat([prod,non,empty], ignore_index=True)
    final = final.groupby('القومي').agg({'Production Labor Cairo': 'sum', 'Non Production labor': 'sum','Empty':'sum'}).reset_index()




    conditions5 = [((final['Production Labor Cairo'] > final['Non Production labor']) & (final['Production Labor Cairo'] > final['Empty'])),
                   ((final['Non Production labor'] > final['Production Labor Cairo']) & (final['Non Production labor'] > final['Empty'])),
                   ((final['Empty'] > final['Production Labor Cairo']) & (final['Empty'] > final['Non Production labor'])),
                   ((final['Empty'] == final['Production Labor Cairo']) | (final['Empty'] == final['Non Production labor']) |(final['Production Labor Cairo'] == final['Non Production labor']))]
    values5 = ['Production Labor Cairo','Non Production labor','Empty',"Balanced"]
    final["Jop"] = np.select(conditions5,values5,default="")

    messagebox.showinfo('Info', 'Done')


def export():
        #excel_file=pd.ExcelWriter("E:\\Khaled\\Application Output\\" + filename2 +  " حضور المشرف .xlsx")
        #excel_file=pd.ExcelWriter("C:\\Users\\hp\\Desktop\\" + filename2 +  " حضور المشرف .xlsx")
        #excel_file = pd.ExcelWriter("D:\\Abdel Rahman\\" + filename2 + " حضور المشرف .xlsx")
        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\Desktop\\" + filename2 + " حضور المشرف .xlsx")
        #excel_file = pd.ExcelWriter("C:\\Users\\tarek\\OneDrive\\Desktop\\Python\\"+ filename2 +  " حضور المشرف .xlsx")

        df2.to_excel(excel_file, sheet_name="Application Att", index=False)
        df10.to_excel(excel_file, sheet_name="Current Month", index=False)
        df66.to_excel(excel_file, sheet_name="Not exist in cm", index=False)
        section.to_excel(excel_file, sheet_name="Sections", index=False)
        final.to_excel(excel_file, sheet_name="Statistics", index=False)


        excel_file.save()


        messagebox.showinfo('Info', 'Done')
        root.destroy()
logo_label = Label(root, font=('Times', 30, 'bold'), bg='black', fg='white', text='Chipsy Attendance')
logo_label.place(x=150, y=5)
# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
open_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='black', fg='white',
                      padx=4, pady=4, command=lambda: [openFile()])
open_btn.place(x=20, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Format', bg='black', fg='white', padx=4,
                        pady=4, command=format)
append_btn.place(x=150, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='CM', bg='black', fg='white', padx=4,
                    pady=4, command=current)
append_btn.place(x=280, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Statistics', bg='black', fg='white', padx=4,
                        pady=4, command=sections)
append_btn.place(x=410, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Export', bg='black', fg='white', padx=4,
                        pady=4, command=export)
append_btn.place(x=540, y=150)
root.mainloop()


