from cmath import e
from collections import UserString
from ctypes import cdll
import http
from http.cookiejar import LWPCookieJar
from multiprocessing.sharedctypes import Value
from operator import index
from pyclbr import Class
from shutil import ignore_patterns
from sqlite3 import Date, Timestamp
from tokenize import Ignore
from typing import Text
from unicodedata import name
from cv2 import split
from matplotlib.pyplot import table, text
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from datetime import date
from datetime import timedelta
from datetime import datetime
from bs4 import BeautifulSoup as soup
from openpyxl import load_workbook
from xlwt.Workbook import *
from pandas import ExcelWriter
import xlsxwriter
import pandas as pd
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.styles import Font
import time 
import requests
from selenium.common.exceptions import NoSuchElementException
import cv2
import xlwt
import pytesseract 
import urllib3
import re
import os
import glob
import numpy as np
from array import *
import csv
import openpyxl as pxl
import random



workbook = xlsxwriter.Workbook('twitter_D1.xlsx')
head_format = workbook.add_format({'bold' : True, 'text_wrap' : True, 'valign' : 'top'})
head_format.set_text_wrap()
bold = workbook.add_format({'bold': True})
# wb = workbook()


tweet_data4xls = {"Time of crape" : [], "Name": [], "Username":[], "Date": [],"Tweet" : []}
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
driver = webdriver.Edge()
driver.get("https://twitter.com/explore")
list_hashtag = ['#NOTAM']
list_account = ['@TGSpaceStation','@CNDeepSpace','@SpaceTrackOrg','@AerospaceCorp','@SpaceForceDoD','@sharemyspacee','@JAXA_en','@LeoLabs_Space','@isro']                      
                # '@LeoLabs_Space','@isro'               
# print(list_hashtag)

data = []

tweet_ids = set()
count = 0


def get_tweet_data(card):
    overdate = "Fasle"
    name = card.find_element(By.XPATH,'.//span').text
    username = card.find_element(By.XPATH,'.//span[contains(text(),"@")]').text
    try:
        timenow = datetime.now()
        date = card.find_element(By.XPATH,'.//time').get_attribute('datetime')
        #print(type(date))
        datetime_object = datetime.strptime(date, '%Y-%m-%dT%H:%M:%S.000Z')  
        #print(type(datetime_object))
        #print(datetime_object)
        x = timenow-datetime_object
        #print(x.total_seconds()/60/60/24)
        day_diff = x.total_seconds()/60/60/24
        if day_diff > 3 :
            overdate = "True"
    except :
        return
    text = card.find_element(By.XPATH,'.//div[@data-testid="tweetText"]').text

    try:
        img_links = card.find_elements(By.XPATH,'.//div[@class="css-1dbjc4n"]//a[@class="css-4rbku5 css-18t94o4 css-1dbjc4n r-1loqt21 r-1pi2tsx r-1ny4l3l"]')
        for img_link in img_links:
            img_src = img_link.find_element(By.XPATH,'.//img[@alt="Image"]').get_attribute('src')
            image = requests.get(img_src)
            print("Image from Scraping : ")
            print(img_src,"\n")
            time.sleep(1)
            with open(r'C:/Users/Lenovo/Downloads/twitter/images/tmp.jpg','wb' ) as f:
                f.write(image.content)
            time.sleep(1)
            img = cv2.imread('C:\\Users\\Lenovo\\Downloads\\twitter\\images\\tmp.jpg') 
            time.sleep(1)
            kernel2 = np.array([[-2, -1, 0],
                                [-1, 3, 1],
                                [0, 1, 2]])
            img = cv2.filter2D(src=img, ddepth=-1, kernel=kernel2)
            # cv2.imshow("filter2d image", img)
            # cv2.imwrite("Filter2d Sharpened Image.jpg", img)
            string = pytesseract.image_to_string(img)
            print("Text from OCR_image :")
            print(string)
    except NoSuchElementException:
        pass
    
    tweet = (name, username, date,text,overdate)
    timeis = datetime.today()
    
    tweet_data4xls["Name"].append(name)
    tweet_data4xls["Username"].append(username)
    tweet_data4xls["Date"].append(date)
    tweet_data4xls["Tweet"].append(text)
    tweet_data4xls["Time of crape"].append(str(timeis))
    return tweet 

    # assert "Python" in driver.title


for hashtag in list_hashtag:
    isOverdate = False
    last_position = driver.execute_script('return window.pageYOffset;')
    time.sleep(2)
    search_input = driver.find_element(By.XPATH, '//input[@aria-label="Search query"]')
    search_input.send_keys(Keys.CONTROL + "a")
    search_input.send_keys(Keys.DELETE)
    time.sleep(1)
    search_input.send_keys(hashtag)
    search_input.send_keys(Keys.RETURN)
    time.sleep(2)
    driver.find_element(By.LINK_TEXT, 'Latest').click()
    print("\n")
    print("##### Data from Hashtag #####\n")
    print(hashtag)
    time.sleep(2)
    scollTo = 0
    
    while not isOverdate:
        cards = driver.find_elements(By.XPATH,'//article[@data-testid="tweet"]')
        scollTo +=500
        for card in cards:
            tweet = get_tweet_data(card)
            if tweet:
                tweet_id = ''.join(tweet)
                if tweet_id not in tweet_ids:
                    if tweet[4] == "True":
                        isOverdate = True
                        print(tweet)
                        break
                    tweet_ids.add(tweet_id)
                    data.append(tweet)
                    print("Data from Scraping :")
                    print(tweet)

        driver.execute_script("window.scrollTo({top: %s ,behavior: 'smooth'});" % str(scollTo))
        time.sleep(1)
        curr_position = driver.execute_script('return window.pageYOffset;')

        if last_position == curr_position:
            break

#print(data)

for account in list_account:
    time.sleep(1)
    print("----------------------------------------------------------\n\n")
    print("##### Data from Account #####\n")
    print('https://twitter.com/'+ account)
    driver.get('https://twitter.com/'+ account)
    isOverdate = False
    last_position = driver.execute_script('return window.pageYOffset;')
    time.sleep(2)

    while not isOverdate:
        cards = driver.find_elements(By.XPATH,'//article[@data-testid="tweet"]')
        scollTo +=200
        for card in cards:
            tweet = get_tweet_data(card)
            if tweet:
                tweet_id = ''.join(tweet)
                if tweet_id not in tweet_ids:
                    if tweet[4] == "True":
                        isOverdate = True
                        break
                    tweet_ids.add(tweet_id)
                    data.append(tweet)
                    print("Data from Scraping : \n")
                    print(tweet , "\n")

        driver.execute_script("window.scrollTo({top: %s ,behavior: 'smooth'});" % str(scollTo))
        time.sleep(1)
        curr_position = driver.execute_script('return window.pageYOffset;')

        if last_position == curr_position:
            break

class dataframe_1:
    Overdate = "Fasle"
    today = date.today()
    nimt = timedelta(days= 1)
    print(type(nimt))
    yesterday = today <= (today - timedelta(days= 7))
    print(yesterday)
    # ConvertToInt = (yesterday.strftime('%Y-%m-%d'))
    # print(type(ConvertToInt))
    df = pd.DataFrame(tweet_data4xls)
    # df['date'] = pd.to_datetime(df['Time']).dt.date.to_string().splitlines()
    # print(type(df['date']))
    df = df.drop_duplicates(subset=['Date'],keep='last')
    rows = df.shape[0]
    cols = df.shape[1]
    convertedStr = str(df)
    print(convertedStr)
    if yesterday is today :
            Overdate = "True"
            while not Overdate:
                if yesterday[2] == "True":
                                    Overdate = True
                                    break
        # print(f'type: {type(convertedStr)}')
    writer = pd.ExcelWriter(r'C:/Users/Lenovo/Desktop/Intern/MyData/twitter_19082022.xlsx')
    convertedStr = df.to_excel(writer,'',index=False)

    workbook = writer.book
    worksheet = writer.sheets['']

    
    for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                writer.sheets[''].set_column(col_idx, 3, column_width)

                cell_format = workbook.add_format()
                cell_format.set_align('center')

                cell_format2 = workbook.add_format()
                cell_format2.set_text_wrap()
                cell_format2.set_align('bottom')

                cell_format3 = workbook.add_format()
                cell_format3.set_text_wrap()
                cell_format3.set_align('center')
                cell_format3.set_align('vcenter')

                worksheet.set_row(0,25)
                worksheet.set_column('A:A', 20, cell_format3)
                worksheet.set_column('B:B', 20, cell_format3)
                worksheet.set_column('C:C', 25, cell_format3)
                worksheet.set_column('D:D', 50, cell_format3)
                worksheet.set_column('E:E', 50, cell_format3)
                worksheet.set_column('F:F', 20, cell_format3)
                worksheet.freeze_panes(1, 0)

                    
    writer.save()

class dataframe_2:
    Overdate = "Fasle"
    today = date.today()
    nimt = timedelta(days= 1)
    print(type(nimt))
    yesterday = today <= (today - timedelta(days= 7))
    print(yesterday)
    # ConvertToInt = (yesterday.strftime('%Y-%m-%d'))
    # print(type(ConvertToInt))
    df = pd.DataFrame(tweet_data4xls)
    # df['date'] = pd.to_datetime(df['Time']).dt.date.to_string().splitlines()
    # print(type(df['date']))
    df = df.drop_duplicates(subset=['Date'],keep='last')
    rows = df.shape[0]
    cols = df.shape[1]
    convertedStr = str(df)
    print(convertedStr)
    if yesterday is today :
            Overdate = "True"
            while not Overdate:
                if yesterday[2] == "True":
                                    Overdate = True
                                    break
        # print(f'type: {type(convertedStr)}')
    writer = pd.ExcelWriter(r'C:/Users/Lenovo/Desktop/Intern/MyData/twitter_20082022.xlsx')
    convertedStr = df.to_excel(writer,'',index=False)

    workbook = writer.book
    worksheet = writer.sheets['']

    
    for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                writer.sheets[''].set_column(col_idx, 3, column_width)

                cell_format = workbook.add_format()
                cell_format.set_align('center')

                cell_format2 = workbook.add_format()
                cell_format2.set_text_wrap()
                cell_format2.set_align('bottom')

                cell_format3 = workbook.add_format()
                cell_format3.set_text_wrap()
                cell_format3.set_align('center')
                cell_format3.set_align('vcenter')

                worksheet.set_row(0,25)
                worksheet.set_column('A:A', 20, cell_format3)
                worksheet.set_column('B:B', 20, cell_format3)
                worksheet.set_column('C:C', 25, cell_format3)
                worksheet.set_column('D:D', 50, cell_format3)
                worksheet.set_column('E:E', 50, cell_format3)
                worksheet.set_column('F:F', 20, cell_format3)
                worksheet.freeze_panes(1, 0)

                    
    writer.save()

class dataframe_3:
    Overdate = "Fasle"
    today = date.today()
    nimt = timedelta(days= 1)
    print(type(nimt))
    yesterday = today <= (today - timedelta(days= 7))
    print(yesterday)
    # ConvertToInt = (yesterday.strftime('%Y-%m-%d'))
    # print(type(ConvertToInt))
    df = pd.DataFrame(tweet_data4xls)
    # df['date'] = pd.to_datetime(df['Time']).dt.date.to_string().splitlines()
    # print(type(df['date']))
    df = df.drop_duplicates(subset=['Date'],keep='last')
    rows = df.shape[0]
    cols = df.shape[1]
    convertedStr = str(df)
    print(convertedStr)
    if yesterday is today :
            Overdate = "True"
            while not Overdate:
                if yesterday[2] == "True":
                                    Overdate = True
                                    break
        # print(f'type: {type(convertedStr)}')
    writer = pd.ExcelWriter(r'C:/Users/Lenovo/Desktop/Intern/MyData/twitter_21082022.xlsx')
    convertedStr = df.to_excel(writer,'',index=False)

    workbook = writer.book
    worksheet = writer.sheets['']

    
    for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                writer.sheets[''].set_column(col_idx, 3, column_width)

                cell_format = workbook.add_format()
                cell_format.set_align('center')

                cell_format2 = workbook.add_format()
                cell_format2.set_text_wrap()
                cell_format2.set_align('bottom')

                cell_format3 = workbook.add_format()
                cell_format3.set_text_wrap()
                cell_format3.set_align('center')
                cell_format3.set_align('vcenter')

                worksheet.set_row(0,25)
                worksheet.set_column('A:A', 20, cell_format3)
                worksheet.set_column('B:B', 20, cell_format3)
                worksheet.set_column('C:C', 25, cell_format3)
                worksheet.set_column('D:D', 50, cell_format3)
                worksheet.set_column('E:E', 50, cell_format3)
                worksheet.set_column('F:F', 20, cell_format3)
                worksheet.freeze_panes(1, 0)

                    
    writer.save()

class dataframe_4:
    Overdate = "Fasle"
    today = date.today()
    nimt = timedelta(days= 1)
    print(type(nimt))
    yesterday = today <= (today - timedelta(days= 7))
    print(yesterday)
    # ConvertToInt = (yesterday.strftime('%Y-%m-%d'))
    # print(type(ConvertToInt))
    df = pd.DataFrame(tweet_data4xls)
    # df['date'] = pd.to_datetime(df['Time']).dt.date.to_string().splitlines()
    # print(type(df['date']))
    df = df.drop_duplicates(subset=['Date'],keep='last')
    rows = df.shape[0]
    cols = df.shape[1]
    convertedStr = str(df)
    print(convertedStr)
    if yesterday is today :
            Overdate = "True"
            while not Overdate:
                if yesterday[2] == "True":
                                    Overdate = True
                                    break
        # print(f'type: {type(convertedStr)}')
    writer = pd.ExcelWriter(r'C:/Users/Lenovo/Desktop/Intern/MyData/twitter_22082022.xlsx')
    convertedStr = df.to_excel(writer,'',index=False)

    workbook = writer.book
    worksheet = writer.sheets['']

    
    for column in df:
                column_width = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                writer.sheets[''].set_column(col_idx, 3, column_width)

                cell_format = workbook.add_format()
                cell_format.set_align('center')

                cell_format2 = workbook.add_format()
                cell_format2.set_text_wrap()
                cell_format2.set_align('bottom')

                cell_format3 = workbook.add_format()
                cell_format3.set_text_wrap()
                cell_format3.set_align('center')
                cell_format3.set_align('vcenter')

                worksheet.set_row(0,25)
                worksheet.set_column('A:A', 20, cell_format3)
                worksheet.set_column('B:B', 20, cell_format3)
                worksheet.set_column('C:C', 25, cell_format3)
                worksheet.set_column('D:D', 50, cell_format3)
                worksheet.set_column('E:E', 50, cell_format3)
                worksheet.set_column('F:F', 20, cell_format3)
                worksheet.freeze_panes(1, 0)

                    
    writer.save()




