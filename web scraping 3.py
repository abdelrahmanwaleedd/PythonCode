import requests
import csv
from bs4 import BeautifulSoup
from itertools import zip_longest
no = []
title = []
year = []
genre = []
gross = []
duration = []
votes = []

page_num = 1
while page_num <1001:
    site = requests.get(f"https://www.imdb.com/search/title/?groups=top_1000&sort=user_rating,desc&count=100&start={page_num}&ref_=adv_nxt")

    src = site.content

    soup =  BeautifulSoup(src,"lxml")

    number = soup.find_all('span', {'class': 'lister-item-index unbold text-primary'})
    movietitle = soup.find_all('h3', {'class': "lister-item-header"})
    year_lunch = soup.find_all('span', {'class': 'lister-item-year text-muted unbold'})
    movgenre = soup.find_all('span', {'class': 'genre'})
    #movgross = soup.find_all('span', {'class': 'text-muted'})
    movdur = soup.find_all('span', {'class': 'runtime'})
    #movotes = soup.find_all('', {'class': ''})


    for i in range(len(number)):
        no.append(number[i].text.replace('.',''))
        title.append(movietitle[i].text.replace('\n','').strip())
        year.append(year_lunch[i].text)
        genre.append(movgenre[i].text.replace('\n','').strip())
        duration.append(movdur[i].text.strip())

    page_num +=100

    print("page switched : ",page_num)


file_exp = zip_longest(*[no, title, year, genre, duration])

with open("C:\\Users\\abdoo\\OneDrive\\Desktop\\imdb22.csv","w") as myfile:
    wr = csv.writer(myfile)
    wr.writerow(["No","Title","Year","Genre","Duration"])
    wr.writerows(file_exp)