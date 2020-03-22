#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function
import io
import codecs
import html2text
import codecs
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os
from getpass import getpass
import pyautogui,time
import pandas as pd


BLOCKSIZE = 1048576
str1 = '['
str2 = '.'
str3 = '<SelectedValues>'
k = 0
j = 0
firstClient = 'randomword'
flag = False
temppath = os.getenv('TEMP')


#get html web-page
browser = webdriver.Chrome('chromedriver.exe')
r = browser.get('https://wiki.itfb.ru/pages/viewpage.action?pageId=13894474')
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt')
s_username = browser.find_element_by_id('os_username')
s_password = browser.find_element_by_id('os_password')
s_continue = browser.find_element_by_id('loginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
page = browser.page_source
file = codecs.open(temppath + '\clients.html', 'w', encoding = 'utf-8')
file.write(page)
file.close()
with open(temppath + '\clients.html', 'r', encoding = 'utf-8') as file:
    for line in file:
        if "Телефоны сотрудников" in line:
            start_table = line.find("Телефоны сотрудников")
            end_table = line.find("2. Доступ в Интернет/Телефония", start_table - 1)
            print(start_table)
            print(end_table)
            out = line[start_table:end_table]
            with open(temppath + '\clients1.html', 'w', encoding = 'utf-8') as file1:
                file1.write(out + "\n")           
tables = pd.read_html(temppath + '\clients1.html')
print(tables[0].iloc[:,1:6])
browser.close()

#deleting temp files
os.remove(temppath + '\clients.html')
os.remove(temppath + '\clients1.html')
