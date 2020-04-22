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


number_a = ''
number_b = ''
digit_count = 0
number_a_raw = '+79265263784'
number_b_raw = '+79265263784+79265263785'
if number_a_raw[0] == '0':
    number_a_raw = number_a_raw[1:]
if len(number_a_raw) > 10:
    for i in range(len(number_a_raw)):
        if number_a_raw[i].isdigit() == True and digit_count < 11:
            number_a = number_a + number_a_raw[i]
            digit_count = digit_count + 1
        elif (number_a_raw[i].isdigit() == True) and (digit_count >= 11):
            print(number_a)
            print('Вероятно присутствует несколько номеров телефонов')
            digit_count = 1
            number_a = number_a_raw[i]
    if digit_count == 11:
        print(number_a)
    else:
        print('Неполный номер\n' + number_a)
elif len(number_a_raw) >= 2 and len(number_a_raw) < 6 and number_a_raw.isdigit() == True:
    print(number_a_raw)
    print('Вероятно это добавочный номер')
elif len(number_a_raw) == 10 and number_a_raw.isdigit() == True and int(number_a_raw[0]) == 9:
    print('+7' + number_a_raw)
    print('К номеру были добавлены +7, проверьте корректность номера')
elif len(number_a_raw) > 0 and number_a_raw.isdigit() == True:
    print('Не удалось определить номер: ' + number_a_raw)
elif len(number_a_raw) > 0 and number_a_raw.isdigit() == False:
    print('Вероятно \"' + number_a_raw + '\" не является номером')