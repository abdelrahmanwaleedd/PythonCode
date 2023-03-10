import requests
import urllib.request
from bs4 import BeautifulSoup
import csv
from itertools import zip_longest

jop_title = []
comp_name = []
jop_loc = []
skills = []
page_num = 0
#https://wuzzuf.net/search/jobs/?a=hpb&filters%5Bcareer_level%5D%5B0%5D=Entry%20Level&filters%5Byears_of_experience_max%5D%5B0%5D=2&filters%5Byears_of_experience_min%5D%5B0%5D=0&q=analysis
#https://wuzzuf.net/search/jobs/?a=hpb&filters%5Bcareer_level%5D%5B0%5D=Entry%20Level&filters%5Byears_of_experience_max%5D%5B0%5D=2&filters%5Byears_of_experience_min%5D%5B0%5D=0&q=analysis&start=1
while page_num < 18:
    #result = requests.get(f"https://wuzzuf.net/search/jobs/?a=hpb&q=python&start={page_num}")
    result = requests.get(f"https://wuzzuf.net/search/jobs/?a=hpb&filters%5Bcareer_level%5D%5B0%5D=Entry%20Level&filters%5Byears_of_experience_max%5D%5B0%5D=2&filters%5Byears_of_experience_min%5D%5B0%5D=0&q=analysis&start={page_num}")

    src = result.content
    soup = BeautifulSoup(src, "lxml")

    jop_titles = soup.find_all('h2', {"class":"css-m604qf"})
    company_names = soup.find_all('a', {"class":"css-17s97q8"})
    jop_location = soup.find_all('span', {"class":"css-5wys0k"})
    jop_skills = soup.find_all('div', {"class":"css-y4udm8"})


    for i in range(len(jop_titles)):
        jop_title.append(jop_titles[i].text.strip())
        comp_name.append(company_names[i].text.strip())
        jop_loc.append(jop_location[i].text.strip())
        skills.append(jop_skills[i].text.strip())

    page_num +=1
    print("page switched=",page_num)

file_exp = [jop_title, comp_name, jop_loc, skills]
exported = zip_longest(*file_exp)

with open("C:\\Users\\abdoo\\OneDrive\\Desktop\\wuzuff_test.csv","w", encoding='utf-8') as myfile:
    wr = csv.writer(myfile)
    wr.writerow(["jop title","company name","jop location","skills"])
    wr.writerows(exported)

