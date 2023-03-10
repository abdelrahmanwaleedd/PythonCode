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
root.geometry('500x300')
root.title('Ergo Company')
root.config(background='#AED6F1')

data = []



def openFile():
    global filepath, filename2

    filepath = filedialog.askopenfilenames(initialdir="C:\\Users\\Cakow\\PycharmProjects\\Main",
                                           title="Open file okay?",
                                           filetypes=(
                                               ('All files', '*.*'),
                                               ('excel files', '*.xlsx'),
                                               ('csv files', '*.csv'),
                                               ('All files', '*.*')))

    for file in filepath:
        df1 = pd.read_excel(file, header=None, sheet_name="Invoice")
        y = df1[df1.apply(lambda row: row.astype(str).str.contains('Loc').any(), axis=1)].index.values
        df1 = pd.read_excel(file, header=y, sheet_name="Invoice")

        filename = os.path.basename(file)
        filename2, extension = os.path.splitext(filename)
        print(filename2)
        df1["المصنع"] = filename2
        data.append(df1)

    messagebox.showinfo('Info', len(filepath))


def append_data():
    global final, final2
    final = pd.concat(data, ignore_index=True)
    final = final[final.columns.drop(list(final.filter(regex='Unna')))]
    final.dropna(subset=["Job"], inplace=True)

    final['Highest Actual Attend'] = final['Highest Actual Attend'].fillna(0)

    conditions = [final["Job"] == "Ambulance driver", final["Job"] == "Assistant Storekeeper",
                  final["Job"] == "Data entry", final["Job"] == "Technician",
                  final["Job"] == "Tractor Driver", final["Job"] == "Yard Coordinator",
                  final["Job"] == "Forklift driver", final["Job"] == "Lab Technician",
                  final["Job"] == "Operator", final["Job"] == "Production Labor Cairo",
                  final["Job"] == "Non Production Labor Cairo", final["Job"] == "Administrator",
                  final["Job"] == "Supervisor"
        , final["Job"] == "License Coordinator", final["Job"] == "Driver", final["Job"] == "Agro Lab Technician",
                  final["Job"] == "Non Production Labor Cairo", final["Job"] == "Non Production Labor Cairo "
        , final["Job"] == "Non Production labor", final["Job"] == "Data Entry", final["Job"] == "Retail Helper",
                  final["Job"] == "Whole Sale Helper", final["Job"] == "Store Keeper", final["Job"] == "Scale Operator"
                  , final["Job"] == "Helper", final["Job"] == "Service Labor", final["Job"] == "Electrical Technician",final["Job"] == "Hepler"
                  ,final["Job"] == "Skilled Labor"]

    values = [final["Highest Actual Attend"] * 258.7975966, final["Highest Actual Attend"] * 291.3561296,
              final["Highest Actual Attend"] * 290.895, final["Highest Actual Attend"] * 276.1970093,
              final["Highest Actual Attend"] * 369.1222985, final["Highest Actual Attend"] * 251.2744687,
              final["Highest Actual Attend"] * 358.2390581, final["Highest Actual Attend"] * 228.0551217,
              final["Highest Actual Attend"] * 252.610524, final["Highest Actual Attend"] * 255.0801233,
              final["Highest Actual Attend"] * 232.5787612, final["Highest Actual Attend"] * 245.2622195,
              final["Highest Actual Attend"] * 251.2744687, final["Highest Actual Attend"] * 301.3688969,
              final["Highest Actual Attend"] * 411.2390228, final["Highest Actual Attend"] * 228.0551217,
              final["Highest Actual Attend"] * 232.5787612, final["Highest Actual Attend"] * 232.5787612
        , final["Highest Actual Attend"] * 232.5787612, final["Highest Actual Attend"] * 290.895,
              final["Highest Actual Attend"] * 228.374581221875, final["Highest Actual Attend"] * 249.62,
              final["Highest Actual Attend"] * 301.38, final["Highest Actual Attend"] * 252.61
              , final["Highest Actual Attend"] * 256.491487924375, final["Highest Actual Attend"] *251.42301454325
              , final["Highest Actual Attend"] * 417.878725, final["Highest Actual Attend"] * 256.491487924375
              , final["Highest Actual Attend"] * 282.974674202]

    final["Emp share"] = np.select(conditions, values, default=0)

    final.duplicated(subset='المصنع', keep='first').sum()
    final2 = final.groupby(['المصنع', 'Location'], as_index=False)['Emp share'].sum()

    messagebox.showinfo('Info', 'Done')


def export():
    curr = datetime.datetime.now().strftime("%Y-%m-%d")
    excel_file = pd.ExcelWriter("C:\\Users\\abdoo\\OneDrive\\Desktop\\"+curr  +" Invoices.xlsx")
    #excel_file = pd.ExcelWriter("C:\\Users\\Khaled Saqr\\Desktop\\" + curr + " All companies.xlsx")
    #excel_file = pd.ExcelWriter("D:\\Abdel Rahman\\" + curr + " All companies.xlsx")

    final.to_excel(excel_file, sheet_name="كل المصانع", index=False)
    final2.to_excel(excel_file, sheet_name="كل ", index=False)

    excel_file.save()
    messagebox.showinfo('Info', 'export sucessfully')
    root.destroy()


# ------------------------------------------------------------------------------------------------------------------------------------------------------------

logo_label = Label(root, font=('Times', 30, 'bold'), bg='#AED6F1', fg='black', text='Ergo Invoices')
logo_label.place(x=100, y=5)

# -----------------------------------------------------------------------------------------------------------------------------------------------------------------
openf = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Open files', bg='#AED6F1', fg='black',
               padx=4, pady=4, command=lambda: [openFile()])
openf.place(x=20, y=200)
append_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Format', bg='#AED6F1',
                    fg='black', padx=4, pady=4, command=append_data)
append_btn.place(x=160, y=200)

export_btn = Button(root, width=12, font=('Times', 12, 'bold'), bd=4, text='Export to xlsx', bg='#AED6F1',
                    fg='black', padx=4, pady=4, command=export)
export_btn.place(x=300, y=200)



root.mainloop()