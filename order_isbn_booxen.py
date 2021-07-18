from os import pwrite
from tkinter import *
import tkinter.font as ft
from selenium import webdriver
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time , datetime

def main():
    main = Tk()
    main.title("북쎈 ISBN 봇")
    main.geometry("400x200")

    myFont = ft.Font(family = "나눔고딕")
    button_target = Button(main, text="시작", command=lambda : macro(), fg="green", width=5).place(x = 300, y = 18)
    main.mainloop()

def macro():
    url = "https://orderbook.booxen.com/"
    global ID, PW
    ID = "naerok2"
    PW = "ipass3309!"
    driver = webdriver.Safari()
    driver.maximize_window()
    driver.get(url)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="userId"]').send_keys(ID)
    driver.find_element_by_xpath('//*[@id="pass"]').send_keys(PW)
    driver.find_element_by_xpath('//*[@id="adm-intro"]/div[1]/div[1]/div/div[2]/div[2]/div/form/button/img').click()
    time.sleep(6)
    searchInfo(driver)

def searchInfo(driver):
    count = 1
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    for i in range(sheet.max_row):
        try:
            book = sheet.cell(count,2).value
            driver.find_element_by_xpath('//*[@id="kwdGnb"]').send_keys(book)
            time.sleep(0.5)
            driver.find_element_by_xpath('//*[@id="selForm"]/fieldset/div[2]/a/img').click()
            time.sleep(1)
            per = driver.find_elements_by_xpath('//*[@id="itemPrice0"]')[1].text
            per = per[-4:]
            per = per.replace(')', '')
            writeExcel(per, count)
            price = driver.find_element_by_xpath('//*[@id="itemPrice0"]').text
            print(price)
            okToSell(count, per, price)
        except:
            pass
        finally:
            count+=1
            pass

def writeExcel(per, count):
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    sheet.cell(count,18,per)
    bookList.save("test.xlsx")

def writeStar(count):
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    sheet.cell(count,1,"X")
    bookList.save("test.xlsx")

def okToSell(count, per, price):
    price = price[-7:]
    price = price.replace('원', '')
    price = price.replace(',', '')
    try:
        price = price.replace(' ', '')
    except:
        pass
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    ownPrice = sheet.cell(count,9).value
    okPrice = int(price) + int(ownPrice/10)

    if(ownPrice < okPrice):
        writeStar(count)
    else:
        bookList=load_workbook(filename='test.xlsx')
        sheet = bookList.active
        sheet.cell(count,1,"")
        bookList.save("test.xlsx")
main()