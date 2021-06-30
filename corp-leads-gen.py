import requests
from bs4 import BeautifulSoup
import urllib
import pandas as pd
from time import sleep
import smtplib
from email.message import EmailMessage
import ratelim
import datetime
import csv
import time
import os.path
from os import path


keywords = ["M&A","Mergers and Acquisitions","Corporate Development","Corporate Strategy","CFO","cheif financial officer","Strategic initiatives","CEO","Cheif execution offficer"]
column_names = ["Name","keyword","Link 1","Link 2","Link 3","Link 4","Link 5"]
company_names = []
response_code = 200
resultCount = 5
filePath = r"D:\\python-workspace\files\\"
sourceFile = "tech_trail_10k.xlsx"

def save_csv(name, keyName, link_arr):
    listt = []
    newList = []
    if len(link_arr) > resultCount:
        newList = link_arr[:resultCount] + link_arr[len(link_arr) :]
    elif len(link_arr) < resultCount:
        for index in range(resultCount):
            print(index)
            if index < len(link_arr):
                newList.insert(index, link_arr[index]);
            else:
                newList.insert(index, "NA")
    else:
        newList = link_arr
        
    today = datetime.date.today()
    file_name = str(today) + '.csv'
    if(path.isfile(filePath + file_name) == False):
        listt.append(column_names)
        
    tmp = []
    tmp.append(name)
    tmp.append(keyName)
    tmp += newList
    listt.append(tmp)
    
    df = pd.DataFrame(listt)
    df.columns = column_names
    df.to_csv(filePath + file_name, mode='a', header=False)

    

@ratelim.patient(9,45)
def call_api(search_query, company):
    print(search_query)
    link_arr = []
    api_key = 'AIzaSyC8E_VqSDA7Dd8okXclmO3n70aglDuUAL0'
    page = requests.get(f"https://google.com/search?q={search_query}&num={resultCount}",auth = (api_key,'')) 
    global response_code
    response_code=page.status_code
    if page.status_code==200:
        
        soup = BeautifulSoup(page.content, "html.parser")   
        links = soup.findAll("a")     


        result1 = soup.findAll({"class": "eqAnXb D0ONmb"})
        result2 = soup.findAll({"class":"ULSxyf"})
        result3 = soup.findAll({"class":"v3jTId"})
        a = "NO RESULT FOUND"
        if len(result1)>0:
            link_arr.append(a)
            return link_arr
                
        elif len(result2)>0:
            link_arr.append(a)
            return link_arr
              
        elif len(result3)>0:
            link_arr.append(a)
            return link_arr
        
        for link in links:                      
            link_href = link.get('href')
            if "url?q=" in link_href and not "webache" in link_href:
                url = link.get('href').split("?q=")[1].split("&sa=U")[0]
                link_arr.append(url)
        return link_arr             
    
    else:
        return link_arr
    
def main_function():
    no_of_companies = 1
    df = pd.read_excel(sourceFile)
    if (len(df.columns)==1):
        df['status']=[0]*10000
    status = list(df['status'])
    i=0
    while(status[i]==1 and i<10000):
        i+=1
    for j in range(i,i+1):
        status[j]=1          
    df['status']= status
    df = df[['Company Name', 'status']]
    df.to_excel(sourceFile)
    df = df[df['status']==1]
    df = df.tail(no_of_companies)
    global company_names
    company_names = list(df['Company Name'])
    for company in company_names:
        for keyword in keywords:
            search_query = urllib.parse.quote(company) + "%20"+ urllib.parse.quote(keyword) + '%20site:linkedin.com/in%20-intitle:profiles' #taking keywords and company name from array
            if(response_code==200):
                print("call function for [ " + company + " ] [" + keyword +" ]")
                urlList = call_api(search_query,company)
                if len(urlList) > 0:
                    save_csv(company, keyword, urlList)
                time.sleep(5)
            else:
                print("call_api() function called off due to response status code: " + str(response_code))
                break
main_function()
