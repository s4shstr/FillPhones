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

# Confluence company names parsing
company_name = 'ADT'

# Confluence staff info parsing and tables preparing
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
file = codecs.open(temppath + '\FillPhones\\' + company_name + '_conf_raw.html', 'w', encoding = 'utf-8')
file.write(page)
file.close()
with open(temppath + '\FillPhones\\' + company_name + '_conf_raw.html', 'r', encoding = 'utf-8') as file:
    for line in file:
        if "Телефоны сотрудников" in line:
            start_table = line.find("Телефоны сотрудников")
            end_table = line.find("2. Доступ в Интернет/Телефония", start_table - 1)
            out = line[start_table:end_table]
            with open(temppath + '\FillPhones\\' + company_name + '_conf_staff_raw.html', 'w', encoding = 'utf-8') as file1:
                file1.write(out + "\n")           
conf_staff_raw_html = pd.read_html(temppath + '\FillPhones\\' + company_name + '_conf_staff_raw.html', encoding = 'utf-8')
conf_staff_raw_xlsx = conf_staff_raw_html[0].drop(conf_staff_raw_html[0].columns[[0, 4, 6, 7]], axis='columns')
tset1 = temppath + '\FillPhones\\' + company_name + '_conf_staff_raw.xlsx'
conf_staff_raw_xlsx.to_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff_raw.xlsx', index=False)
conf_staff_split_xlsx = conf_staff_raw_xlsx['Ф.И.О.'].str.split(' ',expand=True)
conf_staff_split_xlsx.to_html(temppath + '\FillPhones\\' + company_name + '_conf_staff_split.html', index=False)
conf_staff_split_html = pd.read_html(temppath + '\FillPhones\\' + company_name + '_conf_staff_split.html', encoding = 'utf-8')
conf_staff_names = conf_staff_split_html[0].drop(conf_staff_split_html[0].columns[[2]], axis='columns')
conf_staff_names.to_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff_names.xlsx', index=False)
conf_staff_info = conf_staff_raw_html[0].drop(conf_staff_raw_html[0].columns[[0, 1, 4, 6, 7]], axis='columns')
conf_staff_info.to_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff_info.xlsx', index=False)
conf_staff_login = pd.read_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff_names.xlsx', encoding = 'utf-8')
conf_staff_login['Логин'] = ""
conf_staff = conf_staff_login.drop(conf_staff_login.columns[[0, 1]], axis='columns')
conf_staff = conf_staff.merge(conf_staff_names, left_index=True, right_index=True)
conf_staff = conf_staff.merge(conf_staff_info, left_index=True, right_index=True)
conf_staff.to_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер"])

# Confluence staff tables preparing, collisions searching
digit_count = 0
conf_staff = pd.read_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff.xlsx', encoding = 'utf-8')
conf_staff["Коллизии внутреннего номера"] = ''
conf_staff["Коллизии мобильного номера"] = ''
phone_extension_list = conf_staff['Внутренний номер'].tolist()
phone_mobile_list = conf_staff['Моб. номер'].tolist()
for index, row in conf_staff.iterrows():
    phone_extension_start = ''
    phone_extension = ''
    phone_extension_mobile = ''
    phone_extension_raw = ''
    phone_mobile_multiple_extension = ''
    phone_mobile_start = ''
    phone_mobile = ''
    phone_mobile_extension = ''
    phone_mobile_multiple = ''
    phone_mobile_probably = ''
    phone_mobile_raw = ''
    if str(phone_extension_list[index]) != 'nan':
        phone_extension_raw = int(phone_extension_list[index])
        phone_extension_raw = str(phone_extension_raw)
    if str(phone_mobile_list[index]) != 'nan':
        phone_mobile_raw = str(phone_mobile_list[index])
    if phone_extension_raw != '':
        flag_start = False
        start_number_digit = 0
        start_number_digit_flag = False
        phone_number_creating = ''
        phone_number_creating_cache = ''
        start_digit_count = 0
        for i in range(len(phone_extension_raw)):
            if phone_extension_raw[i] == '0' or phone_extension_raw[i].isdigit() != True:
                continue
                start_digit_count = start_digit_count + 1
            else:
                phone_extension_start = phone_extension_raw[i:]
                break
        if phone_extension_start != '':
            for i in range(len(phone_extension_start)):
                if phone_extension_start[i].isdigit() == True:
                    phone_extension = phone_extension + phone_extension_start[i]
            if len(phone_extension) >= 2 and len(phone_extension) < 6:
                conf_staff.iloc[index, 3] = phone_extension
            elif len(phone_extension) >= 11:
                for i in range(len(phone_extension)):
                    if phone_extension[i] == '9' and i + 10 <= len(phone_extension):
                        start_number_digit = i
                        for k in range(i,i+10):
                            phone_number_creating = phone_number_creating + phone_extension[k]
                            if (phone_extension[k] == '9' and k != start_number_digit) or (k == 10 and start_number_digit_flag == False):
                                start_number_digit = k
                                start_number_digit_flag = True
                        i = start_number_digit
                        if phone_number_creating != phone_number_creating_cache and phone_number_creating != '':
                            phone_number_creating_cache = phone_number_creating
                            phone_mobile_multiple_extension = phone_mobile_multiple_extension + '\n' + '7' + phone_number_creating
                        phone_number_creating = ''
                if phone_extension[:2] != '79' and phone_extension[:2] != '89':
                    conf_staff.iloc[index, 6] = 'Неверный внутренний номер. \nВероятно присутствует несколько мобильных номеров телефонов: ' + '\nИзначальный номер: ' + phone_extension_raw + ' \nОбработанный номер: ' + phone_mobile_multiple_extension
            elif len(phone_extension) == 11 and phone_extension[0] == '7':
                conf_staff.iloc[index, 6] = 'Неверный внутренний номер. \nВероятно +' + phone_extension + ' номер мобильного телефона.'
                phone_extension_mobile = phone_extension_mobile + phone_extension + '\n'
            elif len(phone_extension) == 10 and int(phone_extension[0]) == 9:
                conf_staff.iloc[index, 6] = 'Неверный внутренний номер. \nВероятно это номер мобильного телефона. \nК номеру были добавлены +7, проверьте корректность номера \n+7' + phone_extension
                phone_extension_mobile = phone_extension_mobile + '7' + phone_extension
            elif len(phone_extension) > 0 and phone_extension_raw.isdigit() == True:
                conf_staff.iloc[index, 6] = 'Не удалось определить номер: ' + phone_extension_raw
            elif len(phone_extension_raw) > 0:
                conf_staff.iloc[index, 6] = 'Вероятно \"' + phone_extension_raw + '\" не является номером телефона.'
        elif len(phone_extension_raw) > 0:
            conf_staff.iloc[index, 6] = 'Вероятно \"' + phone_extension_raw + '\" не является номером телефона.'
    if phone_mobile_raw != '':
        flag_start = False
        start_number_digit = 0
        start_number_digit_flag = False
        phone_number_creating = ''
        phone_number_creating_cache = ''
        for i in range(len(phone_mobile_raw)):
            if phone_mobile_raw[i] == '0' or phone_mobile_raw[i].isdigit() != True:
                continue
            else:
                phone_mobile_start = phone_mobile_raw[i:]
                break
        if phone_mobile_start != '':
            for i in range(len(phone_mobile_start)):
                if phone_mobile_start[i].isdigit() == True:
                    phone_mobile = phone_mobile + phone_mobile_start[i]
            if len(phone_mobile) >= 11:
                for i in range(len(phone_mobile)):
                    if phone_mobile[i] == '9' and i + 10 <= len(phone_mobile):
                        start_number_digit = i
                        for k in range(i,i+10):
                            phone_number_creating = phone_number_creating + phone_mobile[k]
                            if (phone_mobile[k] == '9' and k != start_number_digit) or (k == 10 and start_number_digit_flag == False):
                                start_number_digit = k
                                start_number_digit_flag = True
                        i = start_number_digit
                        if phone_number_creating != phone_number_creating_cache and phone_number_creating != '':
                            phone_number_creating_cache = phone_number_creating
                            phone_mobile_multiple = phone_mobile_multiple + '\n' + '7' + phone_number_creating
                            phone_number_creating = ''
                if phone_mobile[:2] != '79' and phone_mobile[:2] != '89':
                    conf_staff.iloc[index, 7] = 'Вероятно присутствует несколько мобильных номеров телефонов: ' + '\nИзначальный номер: ' + phone_mobile_raw + ' \nОбработанный номер: ' + phone_mobile_multiple
                conf_staff.iloc[index, 5] = phone_mobile_multiple[1:]
            elif len(phone_mobile) == 11 and (phone_mobile[:2] == '79' or phone_mobile[:2] == '89'):
                conf_staff.iloc[index, 5] = phone_mobile
            elif len(phone_mobile) >= 2 and len(phone_mobile) < 6:
                conf_staff.iloc[index, 7] = 'Неверный мобильный номер телефона.\nВероятно номер ' + phone_mobile + ' является добавочным'
                phone_mobile_extension = phone_mobile_extension + phone_mobile + '\n'
            elif len(phone_mobile) == 10 and int(phone_mobile[0]) == 9:
                conf_staff.iloc[index, 7] = 'К номеру были добавлены +7, проверьте корректность номера. \n+7' + phone_mobile
                conf_staff.iloc[index, 5] = '7' + str(phone_mobile)
            elif len(phone_mobile) > 0 and phone_mobile_raw.isdigit() == True:
                conf_staff.iloc[index, 7] = 'Не удалось определить номер: ' + phone_mobile_raw
            elif (len(phone_mobile) > 9 and (phone_mobile[:2] == '79' or phone_mobile[:2] == '89')) or (len(phone_mobile) > 7 and len(phone_mobile) < 10 and phone_mobile[0] == '9'):
                conf_staff.iloc[index, 7] = 'Неполный номер \n' + phone_mobile_raw
            elif len(phone_mobile_raw) > 0:
                conf_staff.iloc[index, 7] = 'Вероятно \"' + phone_mobile_raw + '\" не является номером телефона.'
        elif len(phone_mobile_raw) > 0:
            conf_staff.iloc[index, 7] = 'Вероятно \"' + phone_mobile_raw + '\" не является номером телефона.'
    if phone_extension_mobile != '':
        if  str(conf_staff.iloc[index, 5]) == 'nan':
            conf_staff.iloc[index, 5] = phone_extension_mobile
        else:
            conf_staff.iloc[index, 5] = str(conf_staff.iloc[index, 5]) + '\n' + phone_extension_mobile
    if phone_mobile_extension != '':
        if str(conf_staff.iloc[index, 3]) == 'nan':
            conf_staff.iloc[index, 3] = phone_mobile_extension
        else:
            conf_staff.iloc[index, 3] = str(conf_staff.iloc[index, 3]) + '\n' + phone_mobile_extension
    if phone_mobile_multiple_extension != '':
        if str(conf_staff.iloc[index, 5]) == 'nan':
            conf_staff.iloc[index, 5] = phone_mobile_multiple_extension
        else:
            conf_staff.iloc[index, 5] = str(conf_staff.iloc[index, 5]) + '\n' + phone_mobile_multiple_extension
conf_staff.to_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff_collisions.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер", "Коллизии внутреннего номера", "Коллизии мобильного номера"])

# OTRS staff info parsing and tables preparing (preview info)
browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Nav=' + company_name)
s_username = browser.find_element_by_id('User')
s_password = browser.find_element_by_id('Password')
s_continue = browser.find_element_by_id('LoginButton')
s_username.send_keys(str(input('Введите логин: ')))
s_password.send_keys(str(getpass('Введите пароль: ')))
s_continue.click()
company_name_field = browser.find_element_by_xpath('//*[@id="Search"]')
company_search_button = browser.find_element_by_xpath('//form[contains(@class, "SearchBox")]//button[contains(@title,"Поиск")]')
browser.execute_script("arguments[0].click();", company_name_field)
browser.execute_script("arguments[0].value = '" + company_name + "';", company_name_field);
browser.execute_script("arguments[0].click();", company_search_button)
page = browser.page_source
otrs_staff_raw_html = pd.read_html(browser.find_element_by_id("CustomerTable").get_attribute('outerHTML'))
otrs_staff_raw_xlsx = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[3, 4, 5]], axis='columns')
otrs_staff_raw_xlsx.to_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff_raw.xlsx', index=False)
otrs_staff_split_xlsx = otrs_staff_raw_xlsx['Name - Название'].str.split(' ',expand=True)
otrs_staff_split_xlsx.to_html(temppath + '\FillPhones\\' + company_name + '_otrs_staff_split.html', index=False)
otrs_staff_split_html = pd.read_html(temppath + '\FillPhones\\' + company_name + '_otrs_staff_split.html', encoding = 'utf-8')
otrs_staff_names_part1 = otrs_staff_split_html[0].drop(otrs_staff_split_html[0].columns[[0, 2]], axis='columns')
otrs_staff_names_part2 = otrs_staff_split_html[0].drop(otrs_staff_split_html[0].columns[[1, 2]], axis='columns')
otrs_staff_names_part1.to_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff_names_part1.xlsx', index=False)
otrs_staff_names_part2.to_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff_names_part2.xlsx', index=False)
otrs_staff_names = otrs_staff_names_part1.merge(otrs_staff_names_part2, left_index=True, right_index=True)
otrs_staff_names.to_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff_names.xlsx', header=False, index=False)
otrs_staff_login = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[1, 2, 3, 4, 5]], axis='columns')
otrs_staff_email = otrs_staff_raw_html[0].drop(otrs_staff_raw_html[0].columns[[0, 1, 3, 4, 5]], axis='columns')
otrs_staff = otrs_staff_login.merge(otrs_staff_names, left_index=True, right_index=True)
otrs_staff['Внутренний номер'] = ""
otrs_staff = otrs_staff.merge(otrs_staff_email, left_index=True, right_index=True)
otrs_staff['Моб. номер'] = ""
otrs_staff.to_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер"])

# OTRS staff info parsing, tables preparing, collisions searching (extended info)
digit_count = 0
otrs_staff = pd.read_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff.xlsx', encoding = 'utf-8')
otrs_staff["Коллизии внутреннего номера"] = ''
otrs_staff["Коллизии мобильного номера"] = ''
for index, row in otrs_staff.iterrows():
    if index == 12: #bugcheck
        break       #bugcheck
    otrs_login = row['Логин'].split('@')[0]
    print(otrs_login + ' (' + str(index + 1) + ' of ' + str(len(otrs_staff)) + ')')
    browser.get('https://hd.itfb.ru/index.pl?Action=AdminCustomerUser;Subaction=Change;ID='+ otrs_login + '%40' + company_name + ';Search=' + company_name + ';Nav=Agent')
    phone_extension_raw = browser.find_element_by_xpath('//*[@id="UserPhone"]').get_attribute("value")
    phone_mobile_raw = browser.find_element_by_xpath('//*[@id="UserMobile"]').get_attribute("value")
    phone_extension_start = ''
    phone_extension = ''
    phone_extension_mobile = ''
    phone_mobile_multiple_extension = ''
    phone_mobile_start = ''
    phone_mobile = ''
    phone_mobile_extension = ''
    phone_mobile_multiple = ''
    phone_mobile_probably = ''
    if phone_extension_raw != '':
        flag_start = False
        start_number_digit = 0
        start_number_digit_flag = False
        phone_number_creating = ''
        phone_number_creating_cache = ''
        start_digit_count = 0
        for i in range(len(phone_extension_raw)):
            if phone_extension_raw[i] == '0' or phone_extension_raw[i].isdigit() != True:
                continue
                start_digit_count = start_digit_count + 1
            else:
                phone_extension_start = phone_extension_raw[i:]
                break
        if phone_extension_start != '':
            for i in range(len(phone_extension_start)):
                if phone_extension_start[i].isdigit() == True:
                    phone_extension = phone_extension + phone_extension_start[i]
            if len(phone_extension) >= 2 and len(phone_extension) < 6:
                otrs_staff.iloc[index, 3] = phone_extension
            elif len(phone_extension) >= 11:
                for i in range(len(phone_extension)):
                    if phone_extension[i] == '9' and i + 10 <= len(phone_extension):
                        start_number_digit = i
                        for k in range(i,i+10):
                            phone_number_creating = phone_number_creating + phone_extension[k]
                            if (phone_extension[k] == '9' and k != start_number_digit) or (k == 10 and start_number_digit_flag == False):
                                start_number_digit = k
                                start_number_digit_flag = True
                        i = start_number_digit
                        if phone_number_creating != phone_number_creating_cache and phone_number_creating != '':
                            phone_number_creating_cache = phone_number_creating
                            phone_mobile_multiple_extension = phone_mobile_multiple_extension + '\n' + '7' + phone_number_creating
                        phone_number_creating = ''
                if phone_extension[:2] != '79' and phone_extension[:2] != '89':
                    otrs_staff.iloc[index, 6] = 'Неверный внутренний номер. \nВероятно присутствует несколько мобильных номеров телефонов: ' + '\nИзначальный номер: ' + phone_extension_raw + ' \nОбработанный номер: ' + phone_mobile_multiple_extension
            elif len(phone_extension) == 11 and phone_extension[0] == '7':
                otrs_staff.iloc[index, 6] = 'Неверный внутренний номер. \nВероятно +' + phone_extension + ' номер мобильного телефона.'
                phone_extension_mobile = phone_extension_mobile + phone_extension + '\n'
            elif len(phone_extension) == 10 and int(phone_extension[0]) == 9:
                otrs_staff.iloc[index, 6] = 'Неверный внутренний номер. \nВероятно это номер мобильного телефона. \nК номеру были добавлены +7, проверьте корректность номера \n+7' + phone_extension
                phone_extension_mobile = phone_extension_mobile + '7' + phone_extension
            elif len(phone_extension) > 0 and phone_extension_raw.isdigit() == True:
                otrs_staff.iloc[index, 6] = 'Не удалось определить номер: ' + phone_extension_raw
            elif len(phone_extension_raw) > 0:
                otrs_staff.iloc[index, 6] = 'Вероятно \"' + phone_extension_raw + '\" не является номером телефона.'
        elif len(phone_extension_raw) > 0:
            otrs_staff.iloc[index, 6] = 'Вероятно \"' + phone_extension_raw + '\" не является номером телефона.'
    if phone_mobile_raw != '':
        flag_start = False
        start_number_digit = 0
        start_number_digit_flag = False
        phone_number_creating = ''
        phone_number_creating_cache = ''
        for i in range(len(phone_mobile_raw)):
            if phone_mobile_raw[i] == '0' or phone_mobile_raw[i].isdigit() != True:
                continue
            else:
                phone_mobile_start = phone_mobile_raw[i:]
                break
        if phone_mobile_start != '':
            for i in range(len(phone_mobile_start)):
                if phone_mobile_start[i].isdigit() == True:
                    phone_mobile = phone_mobile + phone_mobile_start[i]
            if len(phone_mobile) >= 11:
                for i in range(len(phone_mobile)):
                    if phone_mobile[i] == '9' and i + 10 <= len(phone_mobile):
                        start_number_digit = i
                        for k in range(i,i+10):
                            phone_number_creating = phone_number_creating + phone_mobile[k]
                            if (phone_mobile[k] == '9' and k != start_number_digit) or (k == 10 and start_number_digit_flag == False):
                                start_number_digit = k
                                start_number_digit_flag = True
                        i = start_number_digit
                        if phone_number_creating != phone_number_creating_cache and phone_number_creating != '':
                            phone_number_creating_cache = phone_number_creating
                            phone_mobile_multiple = phone_mobile_multiple + '\n' + '7' + phone_number_creating
                            phone_number_creating = ''
                if phone_mobile[:2] != '79' and phone_mobile[:2] != '89':
                    otrs_staff.iloc[index, 7] = 'Вероятно присутствует несколько мобильных номеров телефонов: ' + '\nИзначальный номер: ' + phone_mobile_raw + ' \nОбработанный номер: ' + phone_mobile_multiple
                otrs_staff.iloc[index, 5] = phone_mobile_multiple[1:]
            elif len(phone_mobile) == 11 and (phone_mobile[:2] == '79' or phone_mobile[:2] == '89'):
                otrs_staff.iloc[index, 5] = phone_mobile
            elif len(phone_mobile) >= 2 and len(phone_mobile) < 6:
                otrs_staff.iloc[index, 7] = 'Неверный мобильный номер телефона.\nВероятно номер ' + phone_mobile + ' является добавочным'
                phone_mobile_extension = phone_mobile_extension + phone_mobile + '\n'
            elif len(phone_mobile) == 10 and int(phone_mobile[0]) == 9:
                otrs_staff.iloc[index, 7] = 'К номеру были добавлены +7, проверьте корректность номера. \n+7' + phone_mobile
                otrs_staff.iloc[index, 5] = '7' + str(phone_mobile)
            elif len(phone_mobile) > 0 and phone_mobile_raw.isdigit() == True:
                otrs_staff.iloc[index, 7] = 'Не удалось определить номер: ' + phone_mobile_raw
            elif (len(phone_mobile) > 9 and (phone_mobile[:2] == '79' or phone_mobile[:2] == '89')) or (len(phone_mobile) > 7 and len(phone_mobile) < 10 and phone_mobile[0] == '9'):
                otrs_staff.iloc[index, 7] = 'Неполный номер \n' + phone_mobile_raw
            elif len(phone_mobile_raw) > 0:
                otrs_staff.iloc[index, 7] = 'Вероятно \"' + phone_mobile_raw + '\" не является номером телефона.'
        elif len(phone_mobile_raw) > 0:
            otrs_staff.iloc[index, 7] = 'Вероятно \"' + phone_mobile_raw + '\" не является номером телефона.'
    if phone_extension_mobile != '':
        if  str(otrs_staff.iloc[index, 5]) == 'nan':
            otrs_staff.iloc[index, 5] = phone_extension_mobile
        else:
            otrs_staff.iloc[index, 5] = str(otrs_staff.iloc[index, 5]) + '\n' + phone_extension_mobile
    if phone_mobile_extension != '':
        if str(otrs_staff.iloc[index, 3]) == 'nan':
            otrs_staff.iloc[index, 3] = phone_mobile_extension
        else:
            otrs_staff.iloc[index, 3] = str(otrs_staff.iloc[index, 3]) + '\n' + phone_mobile_extension
    if phone_mobile_multiple_extension != '':
        if str(otrs_staff.iloc[index, 5]) == 'nan':
            otrs_staff.iloc[index, 5] = phone_mobile_multiple_extension
        else:
            otrs_staff.iloc[index, 5] = str(otrs_staff.iloc[index, 5]) + '\n' + phone_mobile_multiple_extension
otrs_staff.to_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff_collisions.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер", "Коллизии внутреннего номера", "Коллизии мобильного номера"])


# Confluence and OTRS tables merging
conf_staff = pd.read_excel(temppath + '\FillPhones\\' + company_name + '_conf_staff.xlsx', encoding = 'utf-8')
otrs_staff = pd.read_excel(temppath + '\FillPhones\\' + company_name + '_otrs_staff.xlsx', encoding = 'utf-8')
staff = conf_staff.merge(otrs_staff, on=['Фамилия', 'Имя'], how = 'outer')
staff.to_excel(temppath + '\FillPhones\\' + company_name + '_staff_merge1.xlsx', index=False)
staff['Логин_x'] = staff[['Логин_x','Логин_y']].fillna('').sum(axis=1)
staff['Внутренний номер_x'] = staff['Внутренний номер_x'].fillna(0).astype(int)
staff['Внутренний номер_x'] = staff[['Внутренний номер_x','Внутренний номер_y']].sum(axis=1)
staff['Внутренний номер_x'] = staff['Внутренний номер_x'].replace(0, np.nan, regex=True)
staff['E-mail_x'] = staff['E-mail_x'].fillna(staff['E-mail_y'])
staff['Моб. номер_x'] = staff[['Моб. номер_x','Моб. номер_y']].fillna('').sum(axis=1)
staff.to_excel(temppath + '\FillPhones\\' + company_name + '_staff_merge2.xlsx', index=False)
staff = staff.drop(staff.columns[[6, 7, 8, 9]], axis='columns')
staff.to_excel(temppath + '\FillPhones\\' + company_name + '_staff.xlsx', index=False, header=["Логин", "Фамилия", "Имя", "Внутренний номер", "E-mail", "Моб. номер"])

browser.close()
browser.quit()

# Deleting temp files
os.remove(temppath + '\FillPhones\\' + company_name + '_conf_raw.html')
os.remove(temppath + '\FillPhones\\' + company_name + '_conf_staff_raw.html')
os.remove(temppath + '\FillPhones\\' + company_name + '_conf_staff_raw.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_conf_staff_split.html')
os.remove(temppath + '\FillPhones\\' + company_name + '_conf_staff_names.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_conf_staff_info.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_otrs_staff_raw.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_otrs_staff_split.html')
os.remove(temppath + '\FillPhones\\' + company_name + '_otrs_staff_names_part1.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_otrs_staff_names_part2.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_otrs_staff_names.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_staff_merge1.xlsx')
os.remove(temppath + '\FillPhones\\' + company_name + '_staff_merge2.xlsx')