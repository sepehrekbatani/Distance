import requests
import time
import csv
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import ui
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoSuchWindowException
import json
from slimit import ast  
from slimit.parser import Parser as JavascriptParser
from slimit.visitors import nodevisitor
from unidecode import unidecode
import xlrd
import xlwt
from xlwt import Workbook
from xlutils.copy import copy
import time

#chrome_driver_path = r"C:\chromedriver.exe"

loc = "C:/Users/sepeh/Desktop/Major Choice/Distance/Python/cities.xlsx"

# To open Workbook
wb = xlrd.open_workbook(loc)
s = wb.sheet_by_index(0)
n = 1114

class City(object):
    def __init__(self, name=None, state=None):
        self.name = name
        self.state = state

cities = []
for row in range(0, n):
    city = City()
    a = s.cell_value(row, 1)
    b = s.cell_value(row, 0)
    a = a.replace('ي', 'ی')
    b = b.replace('ي', 'ی')
    a = a.replace('ك', 'ک')
    b = b.replace('ك', 'ک')
    city.name = a
    city.state = b
    cities.append(city)
    #city(row - 1).name = s.cell_value(row, 1)
    #city(row - 1).state = s.cell_value(row, 0)

def get_distance(orig, dest, flag):
    dist_time = []

    if flag == 1:
        orig_text = orig.name + ', ' + 'استان ' + orig.state
        browser.find_element_by_id("search-1").clear()
        search1 = browser.find_element_by_id("search-1")
        search1.send_keys(orig_text)

    dest_text = dest.name + ', ' + 'استان ' + dest.state
    browser.find_element_by_id("search-3").clear()
    search3 = browser.find_element_by_id("search-3")
    search3.send_keys(dest_text)

    route = browser.find_element_by_id("route-button")
    click3 = None
    while click3 is None:
        try:
            route.click()
            click3 = 1
        except:
            pass

    dis = None
    while dis is None:
        try:
            dis = browser.find_element_by_class_name("dis")
        except:
            pass
    dis = browser.find_element_by_class_name("dis")
    distance = dis.text
    splitted = distance.split(' ')
    dist_time.append(unidecode(splitted[0]))

    zaman = browser.find_element_by_class_name("time")
    zaman = zaman.text
    splitted = zaman.split(' ')
    if "ثانیه" in zaman:
        zaman = 0
    elif "ساعت" in zaman:
        if "دقیقه" in zaman:
            zaman = 60*int(unidecode(splitted[0]))+int(unidecode(splitted[3]))
        else:
            zaman = 60*int(unidecode(splitted[0]))
    else:
        zaman = int(unidecode(splitted[0]))
    dist_time.append(zaman)
    #dist_time[۱] = 60*unidecode(splitted[0])+unidecode(splitted[0])
    return dist_time

loc = "C:/Users/sepeh/Desktop/Major Choice/Distance/Python/dist_python210.xls"
loc2 = "C:/Users/sepeh/Desktop/Major Choice/Distance/Python/dist_python239.xls"
loc3 = "C:/Users/sepeh/Desktop/Major Choice/Distance/Python/dist_python268.xls"
rb = xlrd.open_workbook(loc, formatting_info=True)
r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
sheet1 = wb.get_sheet(0)
#sheet1 = wb.add_sheet('Sheet 1', cell_overwrite_ok=True)

browser = webdriver.Chrome()
main_url = 'https://www.bahesab.ir/map/distance/'
browser.get(main_url)
start_time = time.time()
row = 0
for i in range(210, n):
    if i == 239:
        loc = loc2
        rb = xlrd.open_workbook(loc, formatting_info=True)
        r_sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        sheet1 = wb.get_sheet(0)
        row = 0
    if i == 268:
        loc = loc3
        rb = xlrd.open_workbook(loc, formatting_info=True)
        r_sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        sheet1 = wb.get_sheet(0)
        row = 0
    print(i)
    elpased_time = time.time() - start_time
    start_time = time.time()
    print(elpased_time)
    sheet1.write(row, 0, cities[i].state)
    sheet1.write(row, 1, cities[i].name)
    sheet1.write(row, 2, cities[i].state)
    sheet1.write(row, 3, cities[i].name)
    sheet1.write(row, 4, 0)
    sheet1.write(row, 5, 0)
    row = row + 1
    flag = 1
    for j in range(i+1, n):
        sheet1.write(row, 0, cities[i].state)
        sheet1.write(row, 1, cities[i].name)
        sheet1.write(row, 2, cities[j].state)
        sheet1.write(row, 3, cities[j].name)
        output = get_distance(cities[i], cities[j], flag)
        flag = 0
        sheet1.write(row, 4, output[0])
        sheet1.write(row, 5, output[1])
        sheet1.write(row+1, 0, cities[j].state)
        sheet1.write(row+1, 1, cities[j].name)
        sheet1.write(row+1, 2, cities[i].state)
        sheet1.write(row+1, 3, cities[i].name)
        sheet1.write(row+1, 4, output[0])
        sheet1.write(row+1, 5, output[1])
        row = row+2
    wb.save(loc)
