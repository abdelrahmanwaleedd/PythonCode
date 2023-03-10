from tkinter import filedialog
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import numpy as np
import os
import datetime
from tkinter import ttk
from dateutil.relativedelta import relativedelta

root5 = tk.Tk()
root5.title('Display a Text File')
root5.geometry('400x150')
root5.title('ergo company')
root5.config(background='black')


def fact():
    root = tk.Tk()
    root.title('Display a Text File')
    root.geometry('800x300')
    root.title('ergo company')
    root.config(background='black')
    s = ttk.Style()
    y = tk.IntVar()
    c = tk.IntVar()
    data = []
    fridays_col = []
    cola = []
    ir = 0
    it = 0
    fact = []

    def openFile():
        global filepath
        filepath = filedialog.askopenfilenames(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?",
                                               filetypes=(
                                                   ('All files', '*.*'),
                                                   ('excel files', '*.xlsx'),
                                                   ('csv files', '*.csv'),
                                                   ('All files', '*.*')))

        for file in filepath:
            df1 = pd.read_excel(file, header=None)
            y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
            df1 = pd.read_excel(file, header=y)
            df1.dropna(subset=['م'], inplace=True)
            df1.drop(df1[df1['م'] == '$'].index, inplace=True)
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            df1["المصنع"] = filename2
            data.append(df1)
        messagebox.showinfo('Info', len(filepath))

    def append_data():
        global final
        final = pd.concat(data, ignore_index=True)
        final = final[final.columns.drop(list(final.filter(regex='Unna')))]

        messagebox.showinfo('Info', 'Done')

    def weekend():
        global final2, holiday_sab
        for days in final.columns[9:40]:

            temp = pd.Timestamp(days)
            if temp.dayofweek == 4:
                fridays_col.append(final[temp])
                final2 = pd.concat(fridays_col, axis=1)

    def holiday():
        global holiday_sab
        global H
        global temp2

        for days in final.columns[9:40]:
            temp2 = pd.Timestamp(days)

            if temp2.dayofyear == 6 or temp2.dayofyear == 27 or temp2.dayofyear == 115 or temp2.dayofyear == 121 or temp2.dayofyear == 123 or temp2.dayofyear == 133 or temp2.dayofyear == 134 or temp2.dayofyear == 135 or temp2.dayofyear == 182 or temp2.dayofyear == 199 or temp2.dayofyear == 200 or temp2.dayofyear == 201 or temp2.dayofyear == 202 or temp2.dayofyear == 203 or temp2.dayofyear == 224 or temp2.dayofyear == 280 or temp2.dayofyear == 294:
                cola.append(final[temp2])
                holiday_sab = pd.concat(cola, axis=1)
                H = holiday_sab.apply(lambda row: sum(row[0:] == 'A') or sum(row[0:] == 'B'), axis=1)
        if len(cola) == 0:
            H = 0

    def calc():
        weekend()
        holiday()
        month = datetime.date.today().month
        today = datetime.date.today()
        future_day = today.day
        future_month = (today.month + 4) % 12
        future_year = today.year + ((today.month - 6) // 12)
        six_months_later = datetime.date(future_year, future_month, future_day)

        final.iloc[:, 2] = pd.to_datetime(final.iloc[:, 2]).dt.date
        A = final.apply(lambda row: sum(row[0:] == 'A'), axis=1)
        B = final.apply(lambda row: sum(row[0:] == 'B'), axis=1)
        B8 = final.apply(lambda row: sum(row[0:] == 'B8'), axis=1)
        C = final.apply(lambda row: sum(row[0:] == 'C'), axis=1)
        S1 = final.apply(lambda row: sum(row[0:] == 'S1'), axis=1)
        S2 = final.apply(lambda row: sum(row[0:] == 'S2'), axis=1)
        F1 = final2.apply(lambda row: sum(row[0:] == 'A'), axis=1)
        F2 = final2.apply(lambda row: sum(row[0:] == 'B'), axis=1)
        F3 = final2.apply(lambda row: sum(row[0:] == 'C'), axis=1)
        F4 = final2.apply(lambda row: sum(row[0:] == 'S1'), axis=1)
        F5 = final2.apply(lambda row: sum(row[0:] == 'S2'), axis=1)
        S = final["مرضى بيبسى"]
        # -------------------------------------------------------------------------------------------------------------------------------------------------
        final["total days"] = z = A + B + C + S1 + S2 + S
        final["actual days"] = (A + B + C) - (H) - (F1 + F2 + F3 + F4 + F5)
        final["stoppage"] = S1 + S2
        final["sick leave"] = S
        final["fridays"] = F1 + F2 + F3 + F4 + F5
        final["official holidays"] = H
        final["day overtime"] = final["actual days"] * 2
        final["night overtime"] = final["actual days"] * 2
        final["fridays/h"] = (F1 + F2 + F3 + F4 + F5) * 12
        final["official holidays\h"] = H * 12
        final["Type of account"] = final["Type of account"]
        final["month"] = pd.to_datetime(final["تاريخ التعيين"]).dt.month
        yy=datetime.datetime.now() - relativedelta(years=1)
        # -----------------------------------------------------------------------------------------------------------------------------------------------
        df = pd.read_excel("C:\\Users\\abdoo\\OneDrive\\Desktop\\qoutation2.xlsx", sheet_name=None)
        # df=pd.read_excel("C:\\Users\\OSAMA\\Desktop\\qoutation2.xlsx",sheet_name=None)
        #df = pd.read_excel("D:\\Abdel Rahman\\qoutation2.xlsx", sheet_name=None)

        global df1
        sheets = []
        var3 = tk.StringVar()
        sheet_name = df.keys()
        for i in sheet_name:
            sheets.append(i)

        filepath = "C:\\Users\\abdoo\\OneDrive\\Desktop\\qoutation2.xlsx"

        # filepath="C:\\Users\\OSAMA\\Desktop\\qoutation2.xlsx"
        #filepath="D:\\Abdel Rahman\\qoutation2.xlsx"

        def sheet(event):
            global df
            for i in range(len(sheet_name)):
                if var3.get() == sheets[i]:
                    df = pd.read_excel(filepath, sheet_name=sheets[i])

            final["basic salary"] = df.iloc[0, 1] / 26 * (
                    final["actual days"] + final["stoppage"] + final["sick leave"] + final["official holidays"])

            conditions2 = [
                final["actual days"] + final["official holidays"] + final["stoppage"] + final["sick leave"] >= 1 | (
                        final.iloc[:, 2] <= yy)]
            values2 = [df.iloc[0, 2] / 26 * (final["actual days"] + final["sick leave"] + final["official holidays"])]
            final["variable salary"] = np.select(conditions2, values2, 0)

            final["daily attendance"] = final["basic salary"] + final["variable salary"]
            final["extra hours"] = (final["day overtime"] * 1479.63 / 26 / 8 * 1.35) + (
                    final["night overtime"] * 1479.63 / 26 / 8 * 1.7)
            final["fridays/h/salary"] = df.iloc[0, 7] * final["fridays/h"]
            final["holiday/h/salary"] = df.iloc[0, 7] * final["official holidays\h"]
            final["meal allowance"] = df.iloc[0, 12] * (A + B + C)
            final["for change"] = 0
            final["sick leave cost"] = df.iloc[0, 17] * S

            conditions = [(final["actual days"] + S + S1 + S2 + H >= 20) & (final.iloc[:, 2] <= six_months_later)]
            values = [df.iloc[0, 8]]
            final["annual leave cost"] = np.select(conditions, values, default=0)

            final["transportation"] = 0
            # --------------------------------------------------------------------------------------------------------------------------------------------------------------
            final["حافز اضافى"] = 0
            final["total commission"] = 0

            final["total"] = final["daily attendance"] + final["extra hours"] + final["fridays/h/salary"] + final[
                "holiday/h/salary"] + final["meal allowance"] + final["for change"] + final["annual leave cost"] + \
                             final[
                                 "transportation"] + final["حافز اضافى"] + final["total commission"]

            final[" "] = " "
            final["total loans"] = final["اجمالى السلف"]
            final["permission"] = 0
            final["factory penalties"] = 0
            final["Ergo penalties"] = 0
            final["Total"] = final["total loans"] + final["permission"] + final["factory penalties"] + final[
                "Ergo penalties"]
            final["salary for slipes"] = final["total"] - final["Total"]
            final.drop(final[final['salary for slipes'] == 0].index, inplace=True)

            final["خصم اتلافات"] = 0
            final["سيفتي"] = 0
            final["تيشيرت"] = 0
            final["المرحل"] = 0
            final["final salary slipes"] = final["salary for slipes"]
            final["CIB Account"] = final["CIB Account"]
            final["Bank Misr"] = final["Bank Misr"]
            final[""] = 0
            final[""] = 0

            global net_salary
            net_salary = final[
                ["م", "القومى", "تاريخ التعيين", "الاسم", "المصنع", "الوردية", "basic salary", "variable salary",
                 "daily attendance", "extra hours", "fridays/h/salary",
                 "holiday/h/salary", "meal allowance", "for change", "transportation", "annual leave cost",
                 "حافز اضافى", "total commission", "total",
                 " ", " ", "permission", "total loans", "factory penalties",
                 "Ergo penalties", "Total", " ", " ", "salary for slipes",
                 "خصم اتلافات", "سيفتي", "تيشيرت", "المرحل", "final salary slipes", "Type of account",
                 "CIB Account",
                 "Bank Misr", " ", " "]]
            # ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            messagebox.showinfo('Info', 'Done')
            return final

        sheets_droplist = tk.OptionMenu(root, var3, *(sheets), command=sheet)
        sheets_droplist.place(x=325, y=255)
        sheets_droplist.config(width=15)

    def export_data():
        curr = datetime.datetime.now().strftime("%Y-%m-%d")
        #excel_file = pd.ExcelWriter("D:\\Abdel Rahman\\" + curr + " payroll factory.xlsx")
        # excel_file=pd.ExcelWriter("C:\\Users\\OSAMA\\Desktop\\" + curr + " payroll factory.xlsx")
        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\" + curr + " payroll factory.xlsx")
        final.to_excel(excel_file, sheet_name="attendance", index=False)
        net_salary.to_excel(excel_file, sheet_name='salaries', index=False)
        excel_file.save()
        messagebox.showinfo('Info', 'export sucessfully')
        root.destroy()

    # ------------------------------------------------------------------------------------------------------------------------------------------------------------
    logo_label = tk.Label(root, font=('Times', 30, 'bold'), bg='black', fg='white', text='Factories Payroll')
    logo_label.place(x=400, y=5)

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------
    x = tk.Button(root, width=7, font=('Times', 12, 'bold'), bd=4, text='-> Open', bg='black', fg='white', padx=6,
                  pady=6,
                  command=openFile)
    x.place(x=20, y=200)
    append_btn = tk.Button(root, width=13, font=('Times', 15, 'bold'), bd=4, text='^ Append files', bg='black',
                           fg='white', padx=4,
                           pady=4, command=append_data)
    append_btn.place(x=200, y=200)
    labor_view_btn = tk.Button(root, width=13, font=('Times', 15, 'bold'), bd=4, text='+ Calculations', bg='black',
                               fg='white',
                               padx=4, pady=4, command=calc)
    labor_view_btn.place(x=400, y=200)

    export_btn = tk.Button(root, width=7, font=('Times', 12, 'bold'), bd=4, text='Export ->', bg='black', fg='white',
                           padx=6,
                           pady=6, command=export_data)
    export_btn.place(x=700, y=200)

    root.mainloop()

def sales():
    root = tk.Tk()
    root.title('Display a Text File')
    root.geometry('800x300')
    root.title('ergo company')
    root.config(background='black')
    s = ttk.Style()
    y = tk.IntVar()
    c = tk.IntVar()
    data = []
    fridays_col = []
    cola = []
    ir = 0
    it = 0
    fact = []

    def openFile():
        global filepath
        filepath = filedialog.askopenfilenames(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?",
                                               filetypes=(
                                                   ('All files', '*.*'),
                                                   ('excel files', '*.xlsx'),
                                                   ('csv files', '*.csv'),
                                                   ('All files', '*.*')))

        for file in filepath:
            df1 = pd.read_excel(file, header=None)
            y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
            df1 = pd.read_excel(file, header=y)
            df1.drop(df1.index[[0, 1]], axis=0, inplace=True)
            df1.drop(df1[df1['م'] == '$'].index, inplace=True)
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            df1["المصنع"] = filename2
            data.append(df1)
        messagebox.showinfo('Info', len(filepath))

    def append_data():
        global final
        final = pd.concat(data, ignore_index=True)
        final = final[final.columns.drop(list(final.filter(regex='Unna')))]
        final.drop("الوظيفة", axis=1, inplace=True)

        messagebox.showinfo('Info', 'Done')

    def weekend():
        global final2, holiday_sab
        for days in final.columns[9:40]:

            temp = pd.Timestamp(days)
            if temp.dayofweek == 4:
                fridays_col.append(final[temp])
                final2 = pd.concat(fridays_col, axis=1)

    def holiday():
        global holiday_sab
        global H
        global temp2

        for days in final.columns[9:40]:
            temp2 = pd.Timestamp(days)
            if temp2.dayofyear == 6 or temp2.dayofyear == 27 or temp2.dayofyear == 115 or temp2.dayofyear == 121 or temp2.dayofyear == 123 or temp2.dayofyear == 133 or temp2.dayofyear == 134 or temp2.dayofyear == 135 or temp2.dayofyear == 182 or temp2.dayofyear == 199 or temp2.dayofyear == 200 or temp2.dayofyear == 201 or temp2.dayofyear == 202 or temp2.dayofyear == 203 or temp2.dayofyear == 224 or temp2.dayofyear == 280 or temp2.dayofyear == 294:
                cola.append(final[temp2])
                holiday_sab = pd.concat(cola, axis=1)
                H = holiday_sab.apply(lambda row: sum(row[0:] == 'A') or sum(row[0:] == 'B'), axis=1)
        if len(cola) == 0:
            H = 0

    def calc():
        weekend()
        holiday()
        today = datetime.date.today()
        future_day = today.day
        future_month = (today.month + 6) % 12
        future_year = today.year + ((today.month - 6) // 12)
        six_months_later = datetime.date(future_year, future_month, future_day)
        final.iloc[:, 2] = pd.to_datetime(final.iloc[:, 2]).dt.date

        A = final.apply(lambda row: sum(row[0:] == 'A'), axis=1)
        B = final.apply(lambda row: sum(row[0:] == 'B'), axis=1)
        C = final.apply(lambda row: sum(row[0:] == 'C'), axis=1)
        S1 = final.apply(lambda row: sum(row[0:] == 'S1'), axis=1)
        S2 = final.apply(lambda row: sum(row[0:] == 'S2'), axis=1)
        F1 = final2.apply(lambda row: sum(row[0:] == 'A'), axis=1)
        F2 = final2.apply(lambda row: sum(row[0:] == 'B'), axis=1)
        F3 = final2.apply(lambda row: sum(row[0:] == 'C'), axis=1)
        F4 = final2.apply(lambda row: sum(row[0:] == 'S1'), axis=1)
        F5 = final2.apply(lambda row: sum(row[0:] == 'S2'), axis=1)
        S = final["مرضى بيبسى"]
        # -------------------------------------------------------------------------------------------------------------------------------------------------
        final["total days"] = A + B + C + S1 + S2 + S
        final.drop(final[final['total days'] == 0].index, inplace=True)

        final["actual days"] = (A + B + C) - (H) - (F1 + F2 + F3 + F4 + F5)
        final["stoppage"] = S1 + S2
        final["sick leave"] = S
        final["fridays"] = F1 + F2 + F3 + F4 + F5
        final["official holidays"] = H
        final["day overtime"] = 0
        final["night overtime"] = 0
        final["fridays/h"] = (F1 + F2 + F3 + F4 + F5) * 8
        final["official holidays\h"] = H * 8
        final["Type of account"] = final["Type of account"]
        # -----------------------------------------------------------------------------------------------------------------------------------------------
        df = pd.read_excel("C:\\Users\\abdoo\\OneDrive\\Desktop\\qoutation2.xlsx", sheet_name=None)
        #df = pd.read_excel("D:\\Abdel Rahman\\qoutation2.xlsx", sheet_name=None)

        global df1
        sheets = []
        var3 = tk.StringVar()
        sheet_name = df.keys()
        for i in sheet_name:
            sheets.append(i)

        filepath = "C:\\Users\\abdoo\\OneDrive\\Desktop\\qoutation2.xlsx"
        #filepath = "D:\\Abdel Rahman\\qoutation2.xlsx"
        def sheet(event):
            global df
            for i in range(len(sheet_name)):
                if var3.get() == sheets[i]:
                    df = pd.read_excel(filepath, sheet_name=sheets[i])

            final["basic salary"] = df.iloc[0, 1] / 26 * (
                        final["actual days"] + final["stoppage"] + final["sick leave"] + final["official holidays"])
            final['basic salary'] = final['basic salary'].fillna(0)

            final["variable salary"] = 0
            final["daily attendance"] = final["basic salary"]
            final["extra hours"] = 0
            final["fridays/h/salary"] = df.iloc[0, 7] * final["fridays/h"]
            final["holiday/h/salary"] = df.iloc[0, 7] * final["official holidays\h"]
            final["meal allowance"] = df.iloc[0, 12] / 26 * A + B + C
            final["for change"] = 0
            final["sick leave cost"] = df.iloc[0, 17] * S

            conditions = [(final["total days"] >= 20) & (final.iloc[:, 2] <= six_months_later)]
            values = [df.iloc[0, 1] / 26 * 1.75]

            final["annual leave cost"] = np.select(conditions, values, default=0)
            final["transportation"] = df.iloc[0, 4] / 26 * (
                        final["actual days"] + final["official holidays"] + final["fridays"])

            # ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            messagebox.showinfo('Info', 'Done')
            return final

        sheets_droplist = tk.OptionMenu(root, var3, *(sheets), command=sheet)
        sheets_droplist.place(x=325, y=255)
        sheets_droplist.config(width=15)

    def commession_sheet():
        global filepath2, new_df
        filepath2 = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?",
                                               filetypes=(
                                                   ('All files', '*.*'),
                                                   ('excel files', '*.xlsx'),
                                                   ('csv files', '*.csv'),
                                                   ('All files', '*.*')))

        print(final)
        print("-----------------------------------------------")
        df1 = pd.read_excel(filepath2, header=None)
        y = df1[df1.apply(lambda row: row.astype(str).str.contains('Employee').any(), axis=1)].index.values
        df1 = pd.read_excel(filepath2, header=y)
        df1.dropna(subset=['Employee Id'], inplace=True)
        print(df1)
        new_df = pd.merge(final, df1, left_on='رقم العامل', right_on='Employee Id', how='left')
        # new_df["حافز اضافى"]=new_df["حافز استثنائي"]
        new_df['حافز استثنائي'] = new_df['حافز استثنائي'].fillna(0)
        # new_df["total commission"]=new_df["Total Commission"]
        new_df['Total Commission'] = new_df['Total Commission'].fillna(0)
        new_df["total"] = new_df["daily attendance"] + new_df["extra hours"] + new_df["fridays/h/salary"] + new_df[
            "holiday/h/salary"] + new_df["meal allowance"] + new_df["for change"] + new_df["annual leave cost"] + \
                          new_df["transportation"] + new_df["حافز استثنائي"] + new_df["Total Commission"]

        new_df[" "] = " "
        new_df["total loans"] = new_df["اجمالى السلف"]
        new_df["permission"] = 0
        new_df["factory penalties"] = 0
        new_df["Ergo penalties"] = 0
        new_df["Total"] = new_df["total loans"] + new_df["permission"] + new_df["factory penalties"] + new_df[
            "Ergo penalties"]
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        new_df["salary for slipes"] = new_df["total"] - new_df["Total"]
        new_df.drop(new_df[new_df['salary for slipes'] == 0].index, inplace=True)

        new_df["خصم اتلافات"] = 0
        new_df["سيفتي"] = 0
        new_df["تيشيرت"] = 0
        new_df["المرحل"] = 0
        new_df["final salary slipes"] = new_df["salary for slipes"]
        new_df["CIB Account"] = new_df["CIB Account"]
        new_df["Bank Misr"] = new_df["Bank Misr"]
        new_df[""] = 0
        new_df[""] = 0

        global net_salary
        net_salary = new_df[
            ["م", "القومى", "تاريخ التعيين", "الاسم", "المصنع", "الوردية", "basic salary", "variable salary",
             "daily attendance", "extra hours", "fridays/h/salary",
             "holiday/h/salary", "meal allowance", "for change", "transportation", "annual leave cost", "حافز استثنائي",
             "Total Commission", "total",
             " ", " ", "permission", "total loans", "factory penalties",
             "Ergo penalties", "Total", " ", " ", "salary for slipes",
             "خصم اتلافات", "سيفتي", "تيشيرت", "المرحل", "final salary slipes", "Type of account", "CIB Account",
             "Bank Misr", " ", " "]]
        messagebox.showinfo('Info', 'Done')

    def export_data():
        curr = datetime.datetime.now().strftime("%Y-%m-%d")
        #excel_file = pd.ExcelWriter("D:\\Abdel Rahman\\" + curr + " payroll sales.xlsx")

        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\" + curr + "  payroll sales.xlsx")
        new_df.to_excel(excel_file, sheet_name="attendance", index=False)
        net_salary.to_excel(excel_file, sheet_name='salaries', index=False)
        excel_file.save()
        messagebox.showinfo('Info', 'export sucessfully')
        root.destroy()

    # ------------------------------------------------------------------------------------------------------------------------------------------------------------
    logo_label = tk.Label(root, font=('Times', 30, 'bold'), bg='black', fg='white', text='Sales Payroll')
    logo_label.place(x=500, y=5)

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------
    x = tk.Button(root, width=7, font=('Times', 12, 'bold'), bd=4, text='-> Open', bg='black', fg='white', padx=6,
                  pady=6, command=openFile)
    x.place(x=20, y=200)
    append_btn = tk.Button(root, width=13, font=('Times', 15, 'bold'), bd=4, text='^ Append files', bg='black',
                           fg='white', padx=4, pady=4, command=append_data)
    append_btn.place(x=130, y=200)
    labor_view_btn = tk.Button(root, width=13, font=('Times', 15, 'bold'), bd=4, text='+ Calculations', bg='black',
                               fg='white', padx=4, pady=4, command=calc)
    labor_view_btn.place(x=315, y=200)
    comm_btn = tk.Button(root, width=13, font=('Times', 15, 'bold'), bd=4, text='* Commession', bg='black', fg='white',
                         padx=4, pady=4, command=commession_sheet)
    comm_btn.place(x=500, y=200)
    export_btn = tk.Button(root, width=7, font=('Times', 12, 'bold'), bd=4, text='Export ->', bg='black', fg='white',
                           padx=6, pady=6, command=export_data)
    export_btn.place(x=700, y=200)

    root.mainloop()


logo_label = tk.Label(root5, font=('Times', 25, 'bold'), bg='black', fg='white', text='Pepsi payroll')
logo_label.place(x=70, y=5)

export_btn = tk.Button(root5, width=12, font=('Times', 15, 'bold'), bd=4, text='Factories', bg='black', fg='white',
                       padx=6, pady=6, command=fact)
export_btn.place(x=20, y=80)

export_btn = tk.Button(root5, width=12, font=('Times', 15, 'bold'), bd=4, text='Sales', bg='black', fg='white', padx=6,
                       pady=6, command=sales)
export_btn.place(x=200, y=80)
root5.mainloop()