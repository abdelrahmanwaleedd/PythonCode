from fpdf import FPDF
import os
from tkinter import filedialog
import glob
pdf = FPDF()
pdf.set_auto_page_break(0)
file_path = filedialog.askdirectory()
data=[]
print(file_path)
print("-------------------------------------------------")

folders_list = [x for x in os.listdir(file_path)]
print(folders_list)
print("-------------------------------------------------")

for x in folders_list:

        new_path = f"{file_path}/"+x
        print(new_path)
        print("-------------------------------------------------")

        img_list = [x for x in os.listdir(new_path)]
        print(img_list)
        print("-------------------------------------------------")
        pdf = FPDF()

        for img in img_list:
                print(img)
                print("-------------------------------------------------")

                pdf.add_page()
                images = "{}\\".format(new_path) + img
                print(images)
                print("-------------------------------------------------")

                pdf.image(images, w=180, h=260)
                #pdf.image(images, w=100, h=100)
        pdf.output("{}.pdf".format(new_path))
        print("complete")
        print("-------------------------------------------------")








