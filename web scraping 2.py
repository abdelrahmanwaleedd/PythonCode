import requests
import urllib.request
from bs4 import BeautifulSoup
import csv
from itertools import zip_longest

name = []
year = []

result = requests.get("https://www.imdb.com/chart/toptv/")


src = result.content
soup = BeautifulSoup(src, "lxml")

movies_name = soup.find_all('td', {"class":"titleColumn"})
year_lunched = soup.find_all('span', {"class":"secondaryInfo"})


for i in range(len(movies_name)):
    name.append(movies_name[i].text.replace("\n","").replace(" ",""))
    year.append(year_lunched[i].text.replace("-","").replace("()",""))

print(name)
print(year)
file_exp = [name, year]
exported = zip_longest(*file_exp)
print(exported)
with open("C:\\Users\\abdoo\\OneDrive\\Desktop\\tttttmov.csv","w") as myfile:
    wr = csv.writer(myfile)
    wr.writerow(["Name","Year"])
    wr.writerows(exported)

