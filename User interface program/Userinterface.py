from tkinter import *
import datetime
import pandas as pd
from tkinter import messagebox
from ttkwidgets.autocomplete import AutocompleteCombobox

root2 = Tk()
root2.geometry('900x600')
root2.title('ادخال بيانات ارجو')
root2.configure(background='#30475E')
root2.resizable(width=True, height=True)
var = StringVar()
var1 = StringVar()
var2 = StringVar()
var3 = StringVar()
var4 = IntVar()
var5 = IntVar()
var6 = IntVar()
var7 = IntVar()
var8 = IntVar()
counter1 = 0
counter2 = 0
counter3 = 0
counter4 = 0
counter5 = 0
namee = []
excel_file = "your DB path\\Uniform DataBase.xlsx"

df = pd.read_excel(excel_file)

fact = ['شيبسي اكتوبر', 'شيبسي العبور', 'شيبسي طناش', 'شيبسي سيلز عين شمس', 'شيبسي سيلز اكتوبر', 'شيبسي سيلز طناش',
        'شيبسي سيلز قطاميه', 'شيبسي مخزن عين شمس', 'شيبسي مخزن طناش',
        '-----------------------', 'بيبسي اكتوبر', 'بيبسي طنطا', 'بيبسي العمريه', 'بيبسي عين شمس', 'بيبسي امبابه',
        'بيبسي المعادي', 'بيبسي المنصوره', 'بيبسي مسترد', 'بيبسي مدينة نصر', 'بيبسي الزقازيق']


def choose_fact(event):
    for one in range(len(fact)):
        if var.get() == fact[one]:
            print(var.get())


def add_items():
    global counter1, counter2, counter3, counter4, counter5, df

    values = {'id': var1.get(), 'name': var2.get(), 'factory': var.get(), 'date': var3.get(), 'safety': var4.get(),
              'tshirt': var5.get(), 'jacket': var6.get(), 'pantalon': var7.get(), 'vest': var8.get()}
    df = df.append(values, ignore_index=True)
    df.to_excel(excel_file, index=False)
    if var4.get() == 1:
        counter1 += 1
    if var5.get() == 1:
        counter2 += 1
    if var6.get() == 1:
        counter3 += 1
    if var7.get() == 1:
        counter4 += 1
    if var8.get() == 1:
        counter5 += 1
    id_ent.delete(0, END)
    name_ent.delete(0, END)
    date_ent.delete(0, END)
    saf.deselect()
    ts.deselect()
    pa.deselect()
    vs.deselect()
    ja.deselect()
    counter11 = Label(f5, text=counter1, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
    counter11.place(x=40, y=50)
    counter21 = Label(f4, text=counter2, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
    counter21.place(x=40, y=50)
    counter31 = Label(f3, text=counter3, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
    counter31.place(x=40, y=50)
    counter41 = Label(f2, text=counter4, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
    counter41.place(x=40, y=50)
    counter51 = Label(f, text=counter5, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
    counter51.place(x=40, y=50)
    messagebox.showinfo('Info', 'تمت الاضافه بنجاح')


fact1 = Label(root2, font=('Times', 25, 'bold'), bg='#30475E', fg='white', text='المصنع')
fact1.place(x=800, y=160)
fact_droplist = OptionMenu(root2, var, *(fact), command=choose_fact)
fact_droplist.place(x=350, y=170)
fact_droplist.config(width=60)

#---------------------------------labels-------------------------------------------------
id = Label(root2, font=('Times', 25, 'bold'), bg='#30475E', fg='white', text='القومى')
id.place(x=800, y=230)
name = Label(root2, font=('Times', 25, 'bold'), bg='#30475E', fg='white', text='  الاسم')
name.place(x=802, y=300)
date = Label(root2, font=('Times', 25, 'bold'), bg='#30475E', fg='white', text='التاريخ')
date.place(x=800, y=370)
item = Label(root2, font=('Times', 25, 'bold'), bg='#30475E', fg='white', text='الصنف')
item.place(x=793, y=440)
saf = Checkbutton(root2, text="سيفتي", font=('Times', 20, 'bold'), bg='white', variable=var4)
saf.place(x=700, y=480)
ts = Checkbutton(root2, text="تيشرت", font=('Times', 20, 'bold'), bg='white', variable=var5)
ts.place(x=590, y=480)
ja = Checkbutton(root2, text="جاكت", font=('Times', 20, 'bold'), bg='white', variable=var6)
ja.place(x=490, y=480)
pa = Checkbutton(root2, text="بنطلون", font=('Times', 20, 'bold'), bg='white', variable=var7)
pa.place(x=380, y=480)
vs = Checkbutton(root2, text="فيست", font=('Times', 20, 'bold'), bg='white', variable=var8)
vs.place(x=280, y=480)

# ------------------------------------entries-----------------------------------------------------------------
id_ent = Entry(root2, width=45, textvariable=var1, font=('Arial Greek', 12), bg='white', fg='black')
id_ent.place(x=350, y=240, height=30)
name_ent = AutocompleteCombobox(root2, width=43, textvariable=var2, font=('Arial Greek', 12), completevalues=namee)
name_ent.place(x=350, y=310, height=30)
date_ent = Entry(root2, width=45, textvariable=var3, font=('Arial Greek', 12), bg='white', fg='black')
date_ent.place(x=350, y=380, height=30)
# -----------------------------frames-------------------------------------------------------
f = Frame(root2, bg='white', width=100, height=100)
f.place(x=50, y=10)
f2 = Frame(root2, bg='white', width=100, height=100)
f2.place(x=200, y=10)
f3 = Frame(root2, bg='white', width=100, height=100)
f3.place(x=350, y=10)
f4 = Frame(root2, bg='white', width=100, height=100)
f4.place(x=500, y=10)
f5 = Frame(root2, bg='white', width=100, height=100)
f5.place(x=650, y=10)
# ---------------------------labels------------------------------------------------------------------------------------------
safty = Label(f, text="فيست", bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
safty.place(x=25, y=10)
tshrt = Label(f2, text="بنطلون", bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
tshrt.place(x=18, y=10)
vesst = Label(f3, text="جاكت", bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
vesst.place(x=25, y=10)
jaket = Label(f4, text="تيشرت", bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
jaket.place(x=25, y=10)
pantal = Label(f5, text="سيفتي", bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
pantal.place(x=25, y=10)
counter11 = Label(f5, text=counter1, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
counter11.place(x=40, y=50)
counter21 = Label(f4, text=counter2, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
counter21.place(x=40, y=50)
counter31 = Label(f3, text=counter3, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
counter31.place(x=40, y=50)
counter41 = Label(f2, text=counter4, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
counter41.place(x=40, y=50)
counter51 = Label(f, text=counter5, bg='white', fg='#222831', font=('Arial Greek', 20, 'bold'))
counter51.place(x=40, y=50)
curr = datetime.datetime.now().strftime("%d/%m/%Y")
datee = Label(root2, text=curr, bg='#30475E', fg='#222831', font=('Arial Greek', 30, 'bold'))
datee.place(x=40, y=210)
# --------------------------------buttons---------------------------------------------------------------------------------------------

add = Button(root2, width=15, font=('Times', 15, 'bold'), bd=4, text='<-   ادخال', bg='#222831', fg='white', padx=4,
             pady=4, command=add_items)
add.place(x=25, y=480)

root2.mainloop()
