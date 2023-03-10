from tkinter import filedialog
import pandas as pd
from tkinter import *
from tkinter import messagebox
import datetime
import numpy as np
import os

root2 = Tk()
root2.title('Display a Text File')
root2.geometry('500x150')
root2.title('Ergo Company')
root2.config(background='black')


def social():
    root = Tk()
    root.title('Display a Text File')
    root.geometry('700x400')
    root.title('Ergo Company')
    root.config(background='white')
    y = IntVar()
    data = []

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
            try:
                df1 = pd.read_excel(file, header=None, sheet_name='Attendance For Slips')
                y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
                df1 = pd.read_excel(file, header=y, sheet_name='Attendance For Slips')
            except:
                df1 = pd.read_excel(file, header=None, sheet_name='Attendance')
                y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
                df1 = pd.read_excel(file, header=y, sheet_name='Attendance')
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            df1["sheet name"] = filename2
            print(filename2)
            df1.drop(df1[df1['رقم البطاقة'] == '$'].index, inplace=True)

            df1.dropna(subset=['الإســــــــــــــــــــــــــــم'],inplace=True)
            data.append(df1)

        messagebox.showinfo('Info', len(filepath))

    def append_data():
        global final
        final = pd.concat(data, ignore_index=True)
        final = final[final.columns.drop(list(final.filter(regex='Unna')))]
        messagebox.showinfo('Info', 'Done')

    def select_col():
        root5 = Tk()
        root5.geometry("200x600")
        root5.config(background='white')
        root5.title("Select Columns")
        data1 = []
        var = StringVar()

        def test(event):
            global df1
            for n in range(len(final.columns)):
                if var.get() == final.columns[n]:
                    data1.append(final.iloc[:, n])

        y = OptionMenu(root5, var, *(final.columns), command=test)
        y.place(x=20, y=30)
        y.config(width=20)

        def printt():
            global final2
            global final
            final2 = pd.DataFrame(data1).T
            final = final2
            final["رقم البطاقة"] = pd.to_numeric(final["رقم البطاقة"])

            final["delete"] = np.where(
                (final["الوردية"] == "ازرق تصفية أولى") | (final["الوردية"] == "ازرق تصفية ثانية") | (
                            final["الوردية"] == "تصفية أولى") | (final["الوردية"] == "تصفية") | (
                            final["الوردية"] == "ابيض تصفية أولى") | (final["الوردية"] == "تصفية ثانية") | (
                    final["الوردية"].isna()), "B", "A")
            final.sort_values(by='delete', ascending=True, inplace=True)
            final.drop_duplicates("رقم البطاقة", keep='first', inplace=True)
            final.drop('delete', axis='columns', inplace=True)
            messagebox.showinfo('Info', 'Done')
            root5.destroy()

        done_btn = Button(root5, text="done", width=10, bg='blue', fg='white', padx=4, pady=4,
                          font=('Times', 10, 'bold'),
                          command=printt)
        done_btn.place(x=70, y=500)
        root5.mainloop()

    def check():
        global medical_df, newdf2, newdf, inn, out, addition, deletion, notincm
        filepath2 = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?",
                                               filetypes=(
                                                   ('All files', '*.*'),
                                                   ('excel files', '*.xlsx'),
                                                   ('csv files', '*.csv'),
                                                   ('All files', '*.*')))
        df6 = pd.read_excel(filepath2, header=None)
        y = df6[df6.apply(lambda row: row.astype(str).str.contains('ocia').any(), axis=1)].index.values
        df = pd.read_excel(filepath2, header=y)

        medical_df = df.iloc[:, [8, 7, 9, 1, 6, 20, 21, 22, 23, 24]]
        newdf = pd.merge(final, medical_df, left_on='رقم البطاقة', right_on='ID No.', how='left')
        inn = newdf[newdf['Social No.'].notnull()]
        out = newdf[newdf['Social No.'].isna()]
        addition = out[
            (out['الوردية'] == "ابيض أولى") | (out['الوردية'] == "ابيض ثانية") | (out['الوردية'] == "ازرق أولى")
            | (out['الوردية'] == "ازرق ثانية") | (out['الوردية'] == "أولى") | (out['الوردية'] == "توقف") | (
                        out['الوردية'] == "ثالثة") | (out['الوردية'] == "ثانية")]

        deletion = inn[
            (inn['الوردية'] == "تصفية أولى") | (inn['الوردية'] == "تصفية") | (inn['الوردية'] == "ابيض تصفية أولى") |
            (inn['الوردية'] == "ازرق تصفية ثانية") | (inn['الوردية'] == "ازرق تصفية أولى") | (
                        inn['الوردية'] == "تصفية ثانية")]
        messagebox.showinfo('Info', 'Done')

        notincm = pd.merge(final, medical_df, left_on='رقم البطاقة', right_on='ID No.', how='right')
        notincm = notincm[notincm["رقم البطاقة"].isna()]

    def export():
        curr = datetime.datetime.now().strftime("%Y-%m-%d")
        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\" + curr + "Social sheet.xlsx")
        final.to_excel(excel_file, sheet_name="cm", index=False)
        medical_df.to_excel(excel_file,sheet_name="med",index=False)
        inn.to_excel(excel_file,sheet_name="inn",index=False)
        notincm.to_excel(excel_file,sheet_name="not in cm",index=False)
        out.to_excel(excel_file,sheet_name="out",index=False)
        addition.to_excel(excel_file,sheet_name="addition",index=False)
        deletion.to_excel(excel_file,sheet_name="deletion",index=False)

        excel_file.save()
        messagebox.showinfo('Info', 'export sucessfully')
        root.destroy()

        # ------------------------------------------------------------------------------------------------------------------------------------------------------------

    f1 = Frame(root, bg='#C9002B', width=700, height=150)
    f1.place(x=0, y=0)
    f2 = Frame(root, bg='#004B93', width=700, height=200)
    f2.place(x=0, y=250)
    logo_label = Label(f1, font=('Times', 30, 'bold'), bg='#C9002B', fg='white', text='Social Insurance')
    logo_label.place(x=350, y=5)

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------
    open = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='#004B93', fg='white', padx=4,
                  pady=4, command=lambda: [openFile()])
    open.place(x=15, y=300)
    append_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Append files', bg='#004B93', fg='white',
                        padx=4, pady=4, command=append_data)
    append_btn.place(x=160, y=300)
    dele = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Check', bg='#004B93', fg='white', padx=4,
                  pady=4, command=check)
    dele.place(x=450, y=300)
    select = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='select col', bg='#004B93', fg='white',
                    padx=4, pady=4, command=select_col)
    select.place(x=305, y=300)

    dele = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Export', bg='#004B93', fg='white', padx=4,
                  pady=4, command=export)
    dele.place(x=595, y=300)
    root.mainloop()


def life():
    root = Tk()
    root.title('Display a Text File')
    root.geometry('700x400')
    root.title('Ergo Company')
    root.config(background='white')
    y = IntVar()
    data = []
    data2 = []

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
            df1 = pd.read_excel(file, header=None, sheet_name='Attendance For Slips')
            y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
            df1 = pd.read_excel(file, header=y, sheet_name='Attendance For Slips')
            df1.drop(df1[df1['رقم البطاقة'] == '$'].index, inplace=True)
            df1.drop(df1[df1['رقم البطاقة'] == 'S'].index, inplace=True)

            df1.dropna(subset=['الإســــــــــــــــــــــــــــم'], inplace=True)

            data.append(df1)
        messagebox.showinfo('Info', len(filepath))

    def append_data():
        global final
        final = pd.concat(data, ignore_index=True)
        final = final[final.columns.drop(list(final.filter(regex='Unna')))]
        messagebox.showinfo('Info', 'Done')

    def select_col():
        root5 = Tk()
        root5.geometry("200x600")
        root5.config(background='white')
        root5.title("Select Columns")
        data1 = []
        var = StringVar()

        def test(event):
            global df1
            for n in range(len(final.columns)):
                if var.get() == final.columns[n]:
                    data1.append(final.iloc[:, n])

        y = OptionMenu(root5, var, *(final.columns), command=test)
        y.place(x=20, y=30)
        y.config(width=20)

        def printt():
            global final2
            global final
            final2 = pd.DataFrame(data1).T
            final = final2
            final["رقم البطاقة"] = pd.to_numeric(final["رقم البطاقة"])

            final["delete"] = np.where(
                (final["الوردية"] == "ازرق تصفية أولى") | (final["الوردية"] == "ازرق تصفية ثانية") | (
                            final["الوردية"] == "تصفية أولى") | (final["الوردية"] == "تصفية") | (
                            final["الوردية"] == "ابيض تصفية أولى") | (final["الوردية"] == "تصفية ثانية") | (
                    final["الوردية"].isna()), "B", "A")
            final.sort_values(by='delete', ascending=True, inplace=True)
            final.drop_duplicates("رقم البطاقة", keep='first', inplace=True)
            final.drop('delete', axis='columns', inplace=True)
            messagebox.showinfo('Info', 'Done')
            root5.destroy()

        done_btn = Button(root5, text="done", width=10, bg='blue', fg='white', padx=4, pady=4,
                          font=('Times', 10, 'bold'),
                          command=printt)
        done_btn.place(x=70, y=500)
        root5.mainloop()

    def check():
        global medical_df, newdf2, newdf, inn, out, addition, deletion, notincm2
        filepath2 = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?",
                                               filetypes=(
                                                   ('All files', '*.*'),
                                                   ('excel files', '*.xlsx'),
                                                   ('csv files', '*.csv'),
                                                   ('All files', '*.*')))
        df6 = pd.read_excel(filepath2, header=None)
        y = df6[df6.apply(lambda row: row.astype(str).str.contains('ationa').any(), axis=1)].index.values
        df = pd.read_excel(filepath2, header=y)
        medical_df = df.iloc[:, [12, 2, 3, 1, 5, 6, 7, 15]]
        medical_df.dropna(subset=["National ID / Passport#"],inplace=True)

        newdf = pd.merge(final, medical_df, left_on='رقم البطاقة', right_on='National ID / Passport#', how='left')
        inn = newdf[newdf['National ID / Passport#'].notnull()]
        out = newdf[newdf['National ID / Passport#'].isna()]
        addition = out[
            (out['الوردية'] == "ابيض أولى") | (out['الوردية'] == "ابيض ثانية") | (out['الوردية'] == "ازرق أولى")
            | (out['الوردية'] == "ازرق ثانية") | (out['الوردية'] == "أولى") | (out['الوردية'] == "توقف") | (
                        out['الوردية'] == "ثالثة") | (out['الوردية'] == "ثانية")]

        deletion = inn[
            (inn['الوردية'] == "تصفية أولى") | (inn['الوردية'] == "تصفية") | (inn['الوردية'] == "ابيض تصفية أولى") |
            (inn['الوردية'] == "ازرق تصفية ثانية") | (inn['الوردية'] == "ازرق تصفية أولى") | (
                        inn['الوردية'] == "تصفية ثانية")]

        notincm2 = pd.merge(final, medical_df, left_on='رقم البطاقة', right_on='National ID / Passport#', how='right')
        notincm2 = notincm2[notincm2["رقم البطاقة"].isna()]
        messagebox.showinfo('Info', 'Done')

    def export():
        curr = datetime.datetime.now().strftime("%Y-%m-%d")
        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\" + curr + "  Life sheet.xlsx")
        final.to_excel(excel_file, sheet_name="cm", index=False)
        medical_df.to_excel(excel_file, sheet_name="med", index=False)
        inn.to_excel(excel_file, sheet_name="inn", index=False)
        notincm2.to_excel(excel_file, sheet_name="not in cm", index=False)
        out.to_excel(excel_file, sheet_name="out", index=False)
        addition.to_excel(excel_file, sheet_name="addition", index=False)
        deletion.to_excel(excel_file, sheet_name="deletion", index=False)

        excel_file.save()
        messagebox.showinfo('Info', 'export sucessfully')
        root.destroy()

        # ------------------------------------------------------------------------------------------------------------------------------------------------------------

    f1 = Frame(root, bg='#C9002B', width=700, height=150)
    f1.place(x=0, y=0)
    f2 = Frame(root, bg='#004B93', width=700, height=200)
    f2.place(x=0, y=250)
    logo_label = Label(f1, font=('Times', 30, 'bold'), bg='#C9002B', fg='white', text='Life Insurance')
    logo_label.place(x=350, y=5)

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------
    open = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='#004B93', fg='white', padx=4,
                  pady=4, command=lambda: [openFile()])
    open.place(x=15, y=300)
    append_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Append files', bg='#004B93', fg='white',
                        padx=4, pady=4, command=append_data)
    append_btn.place(x=160, y=300)
    dele = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Check', bg='#004B93', fg='white', padx=4,
                  pady=4, command=check)
    dele.place(x=450, y=300)
    select = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='select col', bg='#004B93', fg='white',
                    padx=4, pady=4, command=select_col)
    select.place(x=305, y=300)

    dele = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Export', bg='#004B93', fg='white', padx=4,
                  pady=4, command=export)
    dele.place(x=595, y=300)
    root.mainloop()


def medical():
    root = Tk()
    root.title('Display a Text File')
    root.geometry('700x400')
    root.title('Ergo Company')
    root.config(background='white')
    y = IntVar()
    data = []
    data2 = []

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
            try:
                df1 = pd.read_excel(file, header=None, sheet_name='Attendance For Slips')
                y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
                df1 = pd.read_excel(file, header=y, sheet_name='Attendance For Slips')
            except:
                df1 = pd.read_excel(file, header=None, sheet_name='Attendance')
                y = df1[df1.apply(lambda row: row.astype(str).str.contains('بطاق|قوم').any(), axis=1)].index.values
                df1 = pd.read_excel(file, header=y, sheet_name='Attendance')
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            df1["sheet name"] = filename2
            print(filename2)
            df1.drop(df1[df1['رقم البطاقة'] == '$'].index, inplace=True)

            df1.dropna(subset=['الإســــــــــــــــــــــــــــم'],inplace=True)
            data.append(df1)

        messagebox.showinfo('Info', len(filepath))

    def append_data():
        global final
        final = pd.concat(data, ignore_index=True)
        final = final[final.columns.drop(list(final.filter(regex='Unna')))]
        messagebox.showinfo('Info', 'Done')

    def select_col():
        root5 = Tk()
        root5.geometry("200x600")
        root5.config(background='white')
        root5.title("Select Columns")
        data1 = []
        var = StringVar()

        def test(event):
            global df1
            for n in range(len(final.columns)):
                if var.get() == final.columns[n]:
                    data1.append(final.iloc[:, n])

        y = OptionMenu(root5, var, *(final.columns), command=test)
        y.place(x=20, y=30)
        y.config(width=20)

        def printt():
            global final2
            global final
            final2 = pd.DataFrame(data1).T
            final = final2
            final["رقم البطاقة"] = pd.to_numeric(final["رقم البطاقة"])

            final["delete"] = np.where(
                (final["الوردية"] == "ازرق تصفية أولى") | (final["الوردية"] == "ازرق تصفية ثانية") | (
                            final["الوردية"] == "تصفية أولى") | (final["الوردية"] == "تصفية") | (
                            final["الوردية"] == "ابيض تصفية أولى") | (final["الوردية"] == "تصفية ثانية") | (
                    final["الوردية"].isna()), "B", "A")
            final.sort_values(by='delete', ascending=True, inplace=True)
            final.drop_duplicates("رقم البطاقة", keep='first', inplace=True)
            final.drop('delete', axis='columns', inplace=True)
            messagebox.showinfo('Info', 'Done')
            root5.destroy()

        done_btn = Button(root5, text="done", width=10, bg='blue', fg='white', padx=4, pady=4,
                          font=('Times', 10, 'bold'),
                          command=printt)
        done_btn.place(x=70, y=500)
        root5.mainloop()

    def check():
        global medical_df, newdf2, newdf, inn, out, addition, deletion, notincm2

        filepath2 = filedialog.askopenfilename(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?",
                                               filetypes=(
                                                   ('All files', '*.*'),
                                                   ('excel files', '*.xlsx'),
                                                   ('csv files', '*.csv'),
                                                   ('All files', '*.*')))
        df6 = pd.read_excel(filepath2, header=None)
        y = df6[df6.apply(lambda row: row.astype(str).str.contains('ationa').any(), axis=1)].index.values
        df = pd.read_excel(filepath2, header=y)
        medical_df = df.iloc[:, [35, 3, 10, 7, 6, 11, 14, 12]]
        newdf = pd.merge(final, medical_df, left_on='رقم البطاقة', right_on='National ID', how='left')
        inn = newdf[newdf['National ID'].notnull()]
        out = newdf[newdf['National ID'].isna()]
        addition = out[
            (out['الوردية'] == "ابيض أولى") | (out['الوردية'] == "ابيض ثانية") | (out['الوردية'] == "ازرق أولى")
            | (out['الوردية'] == "ازرق ثانية") | (out['الوردية'] == "أولى") | (out['الوردية'] == "توقف") | (
                        out['الوردية'] == "ثالثة") | (out['الوردية'] == "ثانية")]

        deletion = inn[
            (inn['الوردية'] == "تصفية أولى") | (inn['الوردية'] == "تصفية") | (inn['الوردية'] == "ابيض تصفية أولى") |
            (inn['الوردية'] == "ازرق تصفية ثانية") | (inn['الوردية'] == "ازرق تصفية أولى") | (
                        inn['الوردية'] == "تصفية ثانية")]

        notincm2 = pd.merge(final, medical_df, left_on='رقم البطاقة', right_on='National ID', how='right')
        notincm2 = notincm2[notincm2["رقم البطاقة"].isna()]

        messagebox.showinfo('Info', 'Done')

    def export():
        curr = datetime.datetime.now().strftime("%Y-%m-%d")
        excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\" + curr + "  Medical sheet.xlsx")
        final.to_excel(excel_file, sheet_name="cm", index=False)
        medical_df.to_excel(excel_file, sheet_name="med", index=False)
        inn.to_excel(excel_file, sheet_name="inn", index=False)
        notincm2.to_excel(excel_file, sheet_name="not in cm", index=False)
        out.to_excel(excel_file, sheet_name="out", index=False)
        addition.to_excel(excel_file, sheet_name="addition", index=False)
        deletion.to_excel(excel_file, sheet_name="deletion", index=False)

        excel_file.save()
        messagebox.showinfo('Info', 'export sucessfully')
        root.destroy()

        # ------------------------------------------------------------------------------------------------------------------------------------------------------------

    f1 = Frame(root, bg='#C9002B', width=700, height=150)
    f1.place(x=0, y=0)
    f2 = Frame(root, bg='#004B93', width=700, height=200)
    f2.place(x=0, y=250)
    logo_label = Label(f1, font=('Times', 30, 'bold'), bg='#C9002B', fg='white', text='Medical Insurance')
    logo_label.place(x=350, y=5)

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------
    open = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='#004B93', fg='white', padx=4,
                  pady=4, command=lambda: [openFile()])
    open.place(x=15, y=300)
    append_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Append files', bg='#004B93', fg='white',
                        padx=4, pady=4, command=append_data)
    append_btn.place(x=160, y=300)
    dele = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Check', bg='#004B93', fg='white', padx=4,
                  pady=4, command=check)
    dele.place(x=450, y=300)
    select = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='select col', bg='#004B93', fg='white',
                    padx=4, pady=4, command=select_col)
    select.place(x=305, y=300)

    dele = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Export', bg='#004B93', fg='white', padx=4,
                  pady=4, command=export)
    dele.place(x=595, y=300)
    root.mainloop()


#logo_label = Label(root2, font=('Times', 25, 'bold'), bg='black', fg='white', text="🏥")
#logo_label.place(x=230, y=15)

social = Button(root2, width=10, font=('Times', 12, 'bold'), bd=4, text='Social', bg='black', fg='white', padx=4,
                pady=4, command=social)
social.place(x=350, y=75)

social2 = Button(root2, width=10, font=('Times', 12, 'bold'), bd=4, text='Life', bg='black', fg='white', padx=4, pady=4,
                 command=life)
social2.place(x=50, y=75)

social3 = Button(root2, width=10, font=('Times', 12, 'bold'), bd=4, text='Medical', bg='black', fg='white', padx=4,
                 pady=4, command=medical)
social3.place(x=200, y=75)
root2.mainloop()
