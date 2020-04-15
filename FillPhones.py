#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import print_function
import io
import codecs
import html2text
import codecs
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass
import pyautogui,time
import pandas as pd
import numpy as np


BLOCKSIZE = 1048576
str1 = '['
str2 = '.'
str3 = '<SelectedValues>'
k = 0
j = 0
firstClient = 'randomword'
flag = False
temppath = os.getenv('TEMP')
browser = webdriver.Chrome('chromedriver.exe')


# Confluence parsing and tables preparing
browser.get('https://wiki.itfb.ru/pages/viewpage.action?pageId=13894474')
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
try:
    os.mkdir(temppath + "\FillPhones");
except OSError:
    shutil.rmtree(temppath + "\FillPhones");
    os.mkdir(temppath + "\FillPhones");
file = codecs.open(temppath + '\FillPhones\ADT_conf_raw.html', 'w', encoding = 'utf-8')
file.write(page)
file.close()
with open(temppath + '\FillPhones\ADT_conf_raw.html', 'r', encoding = 'utf-8') as file:
    for line in file:
        if "Телефоны сотрудников" in line:
            start_table = line.find("Телефоны сотрудников")
            end_table = line.find("2. Доступ в Интернет/Телефония", start_table - 1)
            out = line[start_table:end_table]
            with open(temppath + '\FillPhones\ADT_conf_staff_raw.html', 'w', encoding = 'utf-8') as file1:
                file1.write(out + "\n")           
conf_staff_raw_html = pd.read_html(temppath + '\FillPhones\ADT_conf_staff_raw.html', encoding = 'utf-8')
conf_staff_raw_xlsx = conf_staff_raw_html[0].drop(conf_staff_raw_html[0].columns[[0, 4, 6, 7]], axis='columns')
conf_staff_raw_xlsx.to_excel(temppath + '\FillPhones\ADT_conf_staff_raw.xlsx', index=False)
conf_staff_split_xlsx = conf_staff_raw_xlsx['Ф.И.О.'].str.split(' ',expand=True)
conf_staff_split_xlsx.to_html(temppath + '\FillPhones\ADT_conf_staff_split.html', index=False)
conf_staff_split_html = pd.read_html(temppath + '\FillPhones\ADT_conf_staff_split.html', encoding = 'utf-8')
conf_staff_names = conf_staff_split_html[0].drop(conf_staff_split_html[0].columns[[2]], axis='columns')
conf_staff_names.to_excel(temppath + '\FillPhones\ADT_conf_staff_names.xlsx', index=False)
conf_staff_info = conf_staff_raw_html[0].drop(conf_staff_raw_html[0].columns[[0, 1, 4, 6, 7]], axis='columns')
conf_staff_info.to_excel(temppath + '\FillPhones\ADT_conf_staff_info.xlsx', index=False)
conf_staff_login = pd.read_excel(temppath + '\FillPhones\ADT_conf_staff_names.xlsx', encoding = 'utf-8')
conf_staff_login['Логин'] = ""
conf_staff = conf_staff_login.drop(conf_staff_login.columns[[0, 1]], axis='columns')
conf_staff = conf_staff.merge(conf_staff_names, left_index=True, right_index=True)
conf_staff = conf_staff.merge(conf_staff_info, left_index=True, right_index=True)
conf_staff.to_excel(temppath + '\FillPhones\ADT_conf_staff.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер"])

# OTRS parsing and tables preparing
browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Nav=adt')
s_username = browser.find_element_by_id('User')
s_password = browser.find_element_by_id('Password')
s_continue = browser.find_element_by_id('LoginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
company_name = browser.find_element_by_xpath('//*[@id="Search"]')
company_search = browser.find_element_by_xpath('//form[contains(@class, "SearchBox")]//button[contains(@title,"Поиск")]')
browser.execute_script("arguments[0].click();", company_name)
browser.execute_script("arguments[0].value = 'ADT';", company_name);
browser.execute_script("arguments[0].click();", company_search)
page = browser.page_source
otrs_staff_raw_html = pd.read_html(browser.find_element_by_id("CustomerTable").get_attribute('outerHTML'))
otrs_staff_raw_xlsx = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[3, 4, 5]], axis='columns')
otrs_staff_raw_xlsx.to_excel(temppath + '\FillPhones\ADT_otrs_staff_raw.xlsx', index=False)
otrs_staff_split_xlsx = otrs_staff_raw_xlsx['Name - Название'].str.split(' ',expand=True)
otrs_staff_split_xlsx.to_html(temppath + '\FillPhones\ADT_otrs_staff_split.html', index=False)
otrs_staff_split_html = pd.read_html(temppath + '\FillPhones\ADT_otrs_staff_split.html', encoding = 'utf-8')
otrs_staff_names_part1 = otrs_staff_split_html[0].drop(otrs_staff_split_html[0].columns[[0, 2]], axis='columns')
otrs_staff_names_part2 = otrs_staff_split_html[0].drop(otrs_staff_split_html[0].columns[[1, 2]], axis='columns')
otrs_staff_names_part1.to_excel(temppath + '\FillPhones\ADT_otrs_staff_names_part1.xlsx', index=False)
otrs_staff_names_part2.to_excel(temppath + '\FillPhones\ADT_otrs_staff_names_part2.xlsx', index=False)
otrs_staff_names = otrs_staff_names_part1.merge(otrs_staff_names_part2, left_index=True, right_index=True)
otrs_staff_names.to_excel(temppath + '\FillPhones\ADT_otrs_staff_names.xlsx', header=False, index=False)
otrs_staff_login = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[1, 2, 3, 4, 5]], axis='columns')
otrs_staff_email = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[0, 1, 3, 4, 5]], axis='columns')
otrs_staff = otrs_staff_login.merge(otrs_staff_names, left_index=True, right_index=True)
otrs_staff['Внутренний номер'] = ""
otrs_staff = otrs_staff.merge(otrs_staff_email, left_index=True, right_index=True)
otrs_staff['Моб. номер'] = ""
otrs_staff.to_excel(temppath + '\FillPhones\ADT_otrs_staff.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер"])

# Tables merging
conf_staff = pd.read_excel(temppath + '\FillPhones\ADT_conf_staff.xlsx', encoding = 'utf-8')
otrs_staff = pd.read_excel(temppath + '\FillPhones\ADT_otrs_staff.xlsx', encoding = 'utf-8')
staff = conf_staff.merge(otrs_staff, on=['Фамилия', 'Имя'], how = 'outer')
staff.to_excel(temppath + '\FillPhones\ADT_staff_merge1.xlsx', index=False)
staff['Логин_x'] = staff[['Логин_x','Логин_y']].fillna('').sum(axis=1)
staff['Внутренний номер_x'] = staff['Внутренний номер_x'].fillna(0).astype(int)
staff['Внутренний номер_x'] = staff[['Внутренний номер_x','Внутренний номер_y']].sum(axis=1)
staff['Внутренний номер_x'] = staff['Внутренний номер_x'].replace(0, np.nan, regex=True)
staff['E-mail_x'] = staff['E-mail_x'].fillna(staff['E-mail_y'])
staff['Моб. номер_x'] = staff[['Моб. номер_x','Моб. номер_y']].fillna('').sum(axis=1)
staff.to_excel(temppath + '\FillPhones\ADT_staff_merge2.xlsx', index=False)
staff = staff.drop(staff.columns[[6, 7, 8, 9]], axis='columns')
staff.to_excel(temppath + '\FillPhones\ADT_staff.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер"])

browser.close()
browser.quit()

# Deleting temp files
os.remove(temppath + '\FillPhones\ADT_conf_raw.html')
os.remove(temppath + '\FillPhones\ADT_conf_staff_raw.html')
os.remove(temppath + '\FillPhones\ADT_conf_staff_raw.xlsx')
os.remove(temppath + '\FillPhones\ADT_conf_staff_split.html')
os.remove(temppath + '\FillPhones\ADT_conf_staff_names.xlsx')
os.remove(temppath + '\FillPhones\ADT_conf_staff_info.xlsx')
os.remove(temppath + '\FillPhones\ADT_otrs_staff_raw.xlsx')
os.remove(temppath + '\FillPhones\ADT_otrs_staff_split.html')
os.remove(temppath + '\FillPhones\ADT_otrs_staff_names_part1.xlsx')
os.remove(temppath + '\FillPhones\ADT_otrs_staff_names_part2.xlsx')
os.remove(temppath + '\FillPhones\ADT_otrs_staff_names.xlsx')
os.remove(temppath + '\FillPhones\ADT_staff_merge1.xlsx')
os.remove(temppath + '\FillPhones\ADT_staff_merge2.xlsx')