import glob
import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

root = Tk()
root.title('Folders')
root.geometry('500x250')
root.config(background='black')


def folder():
    global file_path, to_ex
    file_path = filedialog.askdirectory()
    to_ex = "{}\\".format(file_path)
    print(to_ex)
    messagebox.showinfo('Info', file_path)

def extract():
    files = glob.glob(file_path + '/**/*.jpg', recursive=True)
    files2 = glob.glob(file_path + '/**/*.jpeg', recursive=True)
    # For each image
    for file in files:
        # Get File name and extension
        filename = os.path.basename(file)
        # Copy the file with os.rename
        os.rename(
            file,
            to_ex + filename
        )

    for file in files2:
        # Get File name and extension
        filename = os.path.basename(file)
        # Copy the file with os.rename
        os.rename(
            file,
            to_ex + filename
        )
    for entry in os.scandir(file_path):
        if os.path.isdir(entry.path) and not os.listdir(entry.path):
            os.rmdir(entry.path)
    messagebox.showinfo('Info', 'Done')

def formaat():
    global test
    jpg = "{}/*-2.jpg".format(to_ex)
    jpg2 = "{}/*-3.jpg".format(to_ex)

    jpeg = "{}/*-2.jpeg".format(to_ex)
    jpeg2 = "{}/*-3.jpeg".format(to_ex)

    removingfiles = glob.glob(jpg)
    for i in removingfiles:
        os.remove(i)

    removingfiles = glob.glob(jpeg)
    for i in removingfiles:
        os.remove(i)
    removingfiles = glob.glob(jpg2)
    for i in removingfiles:
        os.remove(i)

    removingfiles = glob.glob(jpeg2)
    for i in removingfiles:
        os.remove(i)

    os.chdir(to_ex)
    for file in os.listdir():
        name, ext = os.path.splitext(file)

        splitted = name.split("-")
        new_name = f"{splitted[0]}{ext}"
        os.rename(file, new_name)

    for file in os.listdir():
        name, ext = os.path.splitext(file)
        splitted = name.split("_")
        new_name = f"{splitted[1]}{ext}"
        os.rename(file, new_name)

    messagebox.showinfo('Info', 'Done')
    root.destroy()


logo_label = Label(root, font=('Times', 30, 'bold'), bg='black', fg='white', text='Extract Pics')
logo_label.place(x=120, y=5)
sfold = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Select Folder', bg='black', fg='white', padx=4,pady=4,command=folder)
sfold.place(x=50, y=150)
exfold = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Extract', bg='black', fg='white', padx=4,pady=4,command=extract)
exfold.place(x=200, y=150)
forfold = Button(root, width=10, font=('Times', 12, 'bold'), bd=4, text='Format', bg='black', fg='white', padx=4,pady=4,command=formaat)
forfold.place(x=350, y=150)
root.mainloop()















""""
import glob
import os

# Location with subdirectories
my_path = "C:\\Users\\abdoo\\OneDrive\\Desktop\\New folder\\pics"
# Location to move images to
main_dir = "C:\\Users\\abdoo\\OneDrive\\Desktop\\New folder\\pics\\"

# Get List of all images
files = glob.glob(my_path + '/**/*.jpg', recursive=True)
files2 = glob.glob(my_path + '/**/*.jpeg', recursive=True)
# For each image
for file in files:
    # Get File name and extension
    filename = os.path.basename(file)
    # Copy the file with os.rename
    os.rename(
        file,
        main_dir + filename
    )

for file in files2:
    # Get File name and extension
    filename = os.path.basename(file)
    # Copy the file with os.rename
    os.rename(
        file,
        main_dir + filename
    )


"""""