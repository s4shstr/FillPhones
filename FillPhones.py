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
br = webdriver.Chrome('chromedriver.exe')
r = br.get('https://wiki.itfb.ru/pages/viewpage.action?pageId=13894474')
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt')
s_username = br.find_element_by_id('os_username')
s_password = br.find_element_by_id('os_password')
s_continue = br.find_element_by_id('loginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
page = br.page_source
file = codecs.open(temppath + '/clients.html', 'w', encoding = 'utf-8')
file.write(page)
file.close()


tables = pd.read_html(page)

#print(page)
print(tables[1].iloc[:,1:6])

br.close()

"""
#get html web-page
br = webdriver.Chrome('chromedriver.exe')
br.get('https://wiki.itfb.ru/display/CLIEN/Clients')
pyautogui.keyDown('alt')
pyautogui.press('tab')
pyautogui.keyUp('alt')
s_username = br.find_element_by_id('os_username')
s_password = br.find_element_by_id('os_password')
s_continue = br.find_element_by_id('loginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
page = br.page_source
file = codecs.open(temppath + '/clients.html', 'w', encoding = 'utf-8')
file.write(page)
file.close()


#conversion html to txt
html = codecs.open(temppath + '/clients.html', 'r', 'utf-8')
f = html.read()
w = open(temppath + '/clients.txt', 'w')
w.write(html2text.html2text(f))
html.close()
w.close()


#conversion to utf-8
with codecs.open(temppath + '/clients.txt', 'r', 'cp1251') as sourceFile:
    with codecs.open(temppath + '/clients_new.txt', 'w', 'utf8') as targetFile:
        while True:
            contents = sourceFile.read(BLOCKSIZE)
            if not contents:
                break
            targetFile.write(contents)  


#finding clients on txt file
with io.open(temppath + '/clients_new.txt', 'r', encoding='utf-8') as infile, io.open(temppath + '/clients_new_out.txt', 'w',encoding='utf-8') as outfile:
    for line in infile:
        if flag == True:
            break
        if str1 in line:
            for i in range(len(line)):
                if str1 == line[i]:
                    if str2 == line[i + 4] and firstClient != line[i + 1 : i + 4]:
                        k = k + 1
                        if k == 1:
                            firstClient = line[i + 1 : i + 4]
                        print(line[i + 1 : i + 4])
                        outfile.write(line[i + 1 : i + 4] + '\n')
                    if str2 == line[i + 5] and firstClient != line[i + 1 : i + 5]:
                        k = k + 1
                        if k == 1:
                            firstClient = line[i + 1 : i + 5]
                        print(line[i + 1 : i + 5])
                        outfile.write(line[i + 1 : i + 5] + '\n')
                if (firstClient == line[i + 1 : i + 4] or firstClient == line[i + 1 : i + 5]) and k > 1:
                    flag = True
                    break
print("Колическо компаний: " + str(k))

#creating xml file
with open('./template.xml', 'r', encoding='utf-8') as infile_xml, open(temppath + '/By client, all types.xml', 'w', encoding='utf-8') as outfile_xml:
    for line in infile_xml:
        if str3 in line:
            j = j + 1
        else:
            outfile_xml.write(line)
        if j == 1:
            with io.open(temppath + '/clients_new_out.txt', 'r',encoding='utf-8') as infile:
                for line in infile:
                    outfile_xml.write('<SelectedValues>' + line[:-1] + '</SelectedValues>' + '\n')


#xml file upload
br.get('http://hd.itfb.ru/index.pl?Action=AgentStatistics;Subaction=Import')
s_username = br.find_element_by_id('User')
s_password = br.find_element_by_id('Password')
s_continue = br.find_element_by_id('LoginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
WebDriverWait(br, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@type="file"][@name="File"]'))).send_keys(temppath + '/By client, all types.xml')
s_continue = br.find_element_by_class_name('Primary.CallForAction')
s_continue.click()
s_continue = br.find_element_by_id('SaveAndFinish')
s_continue.click()
"""

#deleting temp files
os.remove(temppath + '/clients.html')
#os.remove(temppath + '/clients.txt')
#os.remove(temppath + '/clients_new.txt')
#os.remove(temppath + '/clients_new_out.txt')
#os.remove(temppath + '/By client, all types.xml')
