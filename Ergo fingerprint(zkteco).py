from tkinter import filedialog
import pandas as pd
from tkinter import *
from tkinter import messagebox
import numpy as np
import calendar
import datetime
from datetime import date
root = Tk()
root.title('Display a Text File')
root.geometry('500x250')
root.title('ergo company')
root.config(background='black')
Date = []


now = datetime.datetime.now()
days = calendar.monthrange(now.year, now.month)[1]

year, month = now.year, now.month

First, Last = calendar.monthrange(year, month)

first = datetime.date(year, month, 1)
last = datetime.date(year, month, Last)


count_frid = 0
for d_ord in range(first.toordinal(), last.toordinal()):
    d = date.fromordinal(d_ord)
    if (d.weekday() == 4):
        count_frid += 1
def openFile():
        global att_df
        filepath = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                            title="Open file okay?",
                                            filetypes=(
                                                ('All files', '*.*'),
                                                ('excel files', '*.xlsx'),
                                                ('csv files', '*.csv')))

        att_df = pd.read_excel(filepath, header=0)
        print(att_df)

        att_df["Time"] = att_df["Time"].astype(str)

        att_df[['Date', 'Time']] = att_df.Time.str.split(expand=True)
        att_df = att_df[['no', 'Name', 'Date', 'Time']]
        print(att_df)
        messagebox.showinfo('Info', 'Done')
def format():
        global format_df, test_format_df, late_col, late_col_test, overtime_col, report

        time_leaving = pd.to_timedelta('17:30:59')
        format_df = pd.DataFrame()
        format_df["Name"] = att_df['Name']
        format_df.drop_duplicates(subset="Name",inplace=True)

        for x in att_df['Date'].unique():
            Date.append(x)

        test_format_df = pd.DataFrame()
        test_format_df["Name"] = att_df['Name']


        for day in range(len(Date)):
                    test_format_df[Date[day]] = np.where(att_df["Date"] == Date[day], "ab", '')

        test_format_df = test_format_df.groupby('Name').sum()
        test_format_df.replace(['abab', 'ababab'], 'ab', inplace=True)
        test_format_df["Attandance"] = test_format_df.apply(lambda row: sum(row[0:] == 'ab'), axis=1)
        format_df = pd.merge(format_df, test_format_df, on="Name")

        #---------------------------------------------------------------------------------
        late_col_test = att_df[(att_df['Time'] > '08:45:59') & (att_df['Time'] < '13:59:00')]
        late_col = pd.DataFrame()
        late_col['Late/days'] = late_col_test.groupby('Name')['Time'].count()
        late_col = late_col.reset_index()
        format_df = pd.merge(format_df, late_col, on="Name", how='left')
        #--------------------------------------------------------------------------------------

        att_df['overtime/h'] = np.where(att_df['Time'] > '18:00:00', pd.to_timedelta(att_df['Time'] - time_leaving), 0)
        #att_df['overtime/h'] = np.where(att_df['overtime/h'] < '00:59:00', 0, att_df['overtime/h'])

        overtime_col = pd.DataFrame()
        overtime_col['Overtime/h'] = att_df.groupby('Name')['overtime/h'].sum()
        overtime_col = overtime_col.reset_index()
        format_df = pd.merge(format_df, overtime_col, on="Name", how='left')
        att_df.drop('overtime/h', axis=1, inplace=True)
        #---------------------------------------------------------------------
        report = format_df.loc[:, ['Name', 'Attandance', 'Late/days', 'Overtime/h']]
        report['Month Days'] = days
        report['Fridays/mon'] = count_frid
        report['Total Days'] =report['Attandance'] + report['Fridays/mon']
        report = report.loc[:, ['Name', 'Attandance', 'Fridays/mon', 'Total Days', 'Month Days', 'Late/days', 'Overtime/h']]
        format_df.drop(['Late/days','Overtime/h'],axis=1,inplace=True)
        att_df.sort_values(ascending=True,by=['no','Date'],inplace=True)

        messagebox.showinfo('Info', 'Done')
def export():
        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\Desktop\\Ergo Attendance.xlsx")
        att_df.to_excel(excel_file, sheet_name="DataBase", index=False)
        format_df.to_excel(excel_file, sheet_name="Attendance", index=False)
        report.to_excel(excel_file, sheet_name="Report", index=False)
        excel_file.save()
        messagebox.showinfo('Info', 'Done')
        root.destroy()

logo_label = Label(root, font = ('Times', 30, 'bold'), bg = 'black', fg = 'white', text = "Ergo Attendance")
logo_label.place(x=150, y=5)

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
open_btn = Button(root,width=10, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='black', fg='white', padx=4, pady=4, command=lambda: [openFile()])
open_btn.place(x=50, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='format', bg='black', fg='white', padx=4, pady=4, command=format)
append_btn.place(x=180, y=150)
append_btn = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Export', bg='black', fg='white', padx=4, pady=4, command=export)
append_btn.place(x=310, y=150)

root.mainloop()
