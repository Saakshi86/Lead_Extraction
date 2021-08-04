import json
import numpy as np
import pandas as pd
import datetime
from googleapiclient.discovery import build
import os.path
from os import path
import urllib


#keywords = ["M&A","Mergers & Acquisitions","Mergers and Acquisitions","Corporate Development","Corporate Strategy","CFO","cheif financial officer","Strategic initiatives","CEO","Cheif execution offficer"]
keywords = ["Partner", "Managing Partner", "Principle"]
column_names = ["Name","keyword","Link 1","Link 2","Link 3","Link 4","Link 5"]
company_names = []
resultCount = 5
filePath = r"D:\\python-workspace\files\\"
sourceFile = filePath + "tech_trail 10k.xlsx"
api_key = 'xxxxxxx'
search_engine_id = 'xxxxxxx'

def save_csv(name, keyName, link_arr):
    listt = []
    newList = []
    if len(link_arr) > resultCount:
        newList = link_arr[:resultCount] + link_arr[len(link_arr) :]
    elif len(link_arr) < resultCount:
        for index in range(resultCount):
            print(index)
            if index < len(link_arr):
                newList.insert(index, link_arr[index])
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

    

def call_api(search_query, company):
    print(search_query)
    link_arr = []
    urls = []
    service = build("customsearch", "v1", developerKey = api_key)
    response = service.cse().list(q = search_query, cx = search_engine_id, start = 0, num = resultCount).execute()

    #print(response)
    
    if "items" in response:
        print("resultes items found")
        for item in response['items']:
            urls.append(item['formattedUrl'])
            
        print(urls)
        link_arr = urls[:resultCount]
        
    else:
        print("items key doesn't exist in the response data")
        

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
        print("fetching for company: " + company)
        for keyword in keywords:
            print("fetching for keyword: " + keyword)
            #search_query = urllib.parse.quote(company) + "%20"+ urllib.parse.quote(keyword) + '%20site:linkedin.com/in%20-intitle:profiles' #taking keywords and company name from array
            search_query = company + " "+ keyword + ' site:linkedin.com/in -intitle:profiles' 
            print("call function for [ " + company + " ] [ " + keyword +" ]")
            urlList = call_api(search_query, company)
            if len(urlList) > 0 & len(urlList) <= 5:
                save_csv(company, keyword, urlList)
                break
            
main_function()

