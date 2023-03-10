from tkinter import filedialog
import pandas as pd
from tkinter import *
from tkinter import messagebox
import os
import datetime
import numpy as np
from ttkwidgets.autocomplete import AutocompleteCombobox
import warnings

root = Tk()
root.title('Display a Text File')
root.geometry('750x400')
root.title('Ergo Company')
root.config(background='#AED6F1')
y = IntVar()
var1 = StringVar()
data = []
sheet_names = ["Attendance", "Invoice", "Quotation", "Attendance For Slips", "New Salary", "Sheet1", "Sheet2", "Sheet3"]


def openFile():
    global filepathm, filename2
    warnings.filterwarnings("ignore")

    if len(var1.get()) == 0:
        messagebox.showerror('Python Error', 'please write a file name')


    else:
        filepath = filedialog.askopenfilenames(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                               title="Open file okay?", filetypes=(
                ('All files', '*.*'), ('excel files', '*.xlsx'), ('csv files', '*.csv'), ('All files', '*.*')))

        for file in filepath:
            filename = os.path.basename(file)
            filename2, extension = os.path.splitext(filename)
            try:
                df = pd.read_excel(file, header=None, sheet_name=var1.get())
            except:
                messagebox.showerror('Python Error', 'this sheet doesnt exist in ' + filename2)

            try:
                head = df[df.apply(lambda row: row.astype(str).str.contains('قوم|بطاق').any(), axis=1)].index.values
                df3 = pd.read_excel(file, header=head, sheet_name=var1.get())
            except:
                # head = df[df.apply(lambda row: row.astype(str).str.contains('قوم|بطاق').any(), axis = 1)].index.values
                df3 = pd.read_excel(file, header=0, sheet_name=var1.get())

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
        global final

        final2 = pd.DataFrame(data1).T
        final = final2

        messagebox.showinfo('Info', 'Done')
        root5.destroy()

    done_btn = Button(root5, text="done", width=10, bg='blue', fg='white', padx=4, pady=4, font=('Times', 10, 'bold'),
                      command=printt)
    done_btn.place(x=70, y=500)
    root5.mainloop()


def remove_duplicates():
    try:
        try:
            final.dropna(subset=['القومى'], inplace=True)
            final.drop(final[final['القومى'] == 2].index, inplace=True)
            final.drop(final[final['القومى'] == "S"].index, inplace=True)
            final.drop(final[final['الوردية'] == '$'].index, inplace=True)
            # final.dropna(subset = ['الوردية'], inplace = True)
            # final["القومى"] = pd.to_numeric(final["القومى"])
            #final['القومى'] = final['القومى'].apply(int)

            final["delete"] = np.where(
                (final["الوردية"] == "ازرق تصفية أولى") | (final["الوردية"] == "ازرق تصفية ثانية") | (
                        final["الوردية"] == "تصفية أولى") | (final["الوردية"] == "تصفية") | (
                            final["الوردية"] == "ابيض تصفية أولى") | (final["الوردية"] == "تصفية ثانية") | (
                    final["الوردية"].isna()), "B", "A")
            final.sort_values(by='delete', ascending=True, inplace=True)
            final.drop_duplicates("القومى", keep='first', inplace=True)
            final.drop('delete', axis='columns', inplace=True)
            messagebox.showinfo('Info', 'Done')
        except:
            final.dropna(subset=['رقم البطاقة'], inplace=True)
            final.drop(final[final['رقم البطاقة'] == 2].index, inplace=True)
            final.drop(final[final['رقم البطاقة'] == "S"].index, inplace=True)
            final.drop(final[final['الوردية'] == '$'].index, inplace=True)
            # final.dropna(subset = ['الوردية'], inplace = True)
            # final["رقم البطاقة"] = pd.to_numeric(final["رقم البطاقة"])

            final["delete"] = np.where(
                (final["الوردية"] == "ازرق تصفية أولى") | (final["الوردية"] == "ازرق تصفية ثانية") | (
                        final["الوردية"] == "تصفية أولى") | (final["الوردية"] == "تصفية") | (
                        final["الوردية"] == "ابيض تصفية أولى") | (final["الوردية"] == "تصفية ثانية") | (
                    final["الوردية"].isna()), "B", "A")
            final.sort_values(by='delete', ascending=True, inplace=True)
            final.drop_duplicates("رقم البطاقة", keep='first', inplace=True)
            final.drop('delete', axis='columns', inplace=True)
            messagebox.showinfo('Info', 'Done')

    except KeyError:
        messagebox.showerror('Python Error', 'you cant remove duplicates')


def export():
    curr = datetime.datetime.now().strftime("%Y-%m-%d")
    #excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\" + curr + var1.get() + " All companies.xlsx")
    excel_file = pd.ExcelWriter("C:\\Users\\tarek\\OneDrive\\Desktop\\Python\\" + curr + var1.get() + " All companies.xlsx")

    final.to_excel(excel_file, sheet_name="كل المصانع", index=False)
    excel_file.save()
    messagebox.showinfo('Info', 'export sucessfully')
    root.destroy()


# ------------------------------------------------------------------------------------------------------------------------------------------------------------
f1 = Frame(root, bg='#AED6F1', width=800, height=150)
f1.place(x=0, y=0)
f2 = Frame(root, bg='#AED6F1', width=800, height=200)
f2.place(x=0, y=250)
logo_label = Label(f1, font=('Times', 30, 'bold'), bg='#AED6F1', fg='black', text='Ergo company')
logo_label.place(x=400, y=5)

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
openf = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Open file', bg='#AED6F1', fg='black',
               padx=4, pady=4, command=lambda: [openFile()])
openf.place(x=20, y=300)
append_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Append files', bg='#AED6F1',
                    fg='black', padx=4, pady=4, command=append_data)
append_btn.place(x=160, y=300)
select_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Select columns', bg='#AED6F1',
                    fg='black', padx=4, pady=4, command=select_col)
select_btn.place(x=300, y=300)
dup = Button(root, width=14, font=('Times', 12, 'bold'), bd=4, text='Remove Duplicates', bg='#AED6F1',
             fg='black', padx=4, pady=4, command=remove_duplicates)
dup.place(x=440, y=300)

xx = Label(root, font=('Times', 20, 'bold'), bg='#AED6F1', fg='black', text='optional')
xx.place(x=450, y=250)


export_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Export to xlsx', bg='#AED6F1',
                    fg='black', padx=4, pady=4, command=export)


export_btn.place(x=600, y=300)
shet = Label(f1, font=('Times', 20, 'bold'), bg='#AED6F1', fg='black', text='Sheet Name')
shet.place(x=5, y=110)

name_ent = AutocompleteCombobox(root, width=20, textvariable=var1, font=('Arial Greek', 12), completevalues=sheet_names)
name_ent.place(x=160, y=115, height=30)

root.mainloop()
