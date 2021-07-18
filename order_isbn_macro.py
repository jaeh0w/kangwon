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
    main.title("ISBN 봇")
    main.geometry("400x200")

    myFont = ft.Font(family = "나눔고딕")
    button_target = Button(main, text="시작", command=lambda : macro(), fg="green", width=5).place(x = 300, y = 18)
    main.mainloop()

def macro():
    url = "https://bscm.kyobobook.co.kr"
    global ID
    global PW
    ID = "1544200122"
    PW = "ipass2255-"
    driver = webdriver.Safari()
    driver.maximize_window()
    driver.get(url)
    time.sleep(1)

    driver.find_element_by_class_name("w2input.inp_id").send_keys(ID)
    driver.find_element_by_class_name("w2input.inp_pw").send_keys(PW)
    driver.find_element_by_class_name("w2anchor2.anc_login_pc").click()
    
    time.sleep(7)
    searchInfo(driver)

def searchInfo(driver):
    count = 1
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    for i in range(sheet.max_row):
        try:
            book = sheet.cell(count,10).value
            time.sleep(3)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mf_wfm_header_ibx_findName"]'))).text
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="mf_wfm_header_ibx_findName"]').send_keys(book)
            time.sleep(0.5)
            driver.find_element_by_class_name('w2textbox.search_tit').click()
            time.sleep(1)
            isbn = driver.find_element_by_xpath('//*[@id="mf_wfm_content_tac_main_contents_content1_body_gen_bksSrch_0_tbx_cmdtCode"]').text
            per = driver.find_element_by_xpath('//*[@id="mf_wfm_content_tac_main_contents_content1_body_gen_bksSrch_0_tbx_byngRate"]').text
            per = per.replace('(', '')
            per = per.replace(')', '')
            writeExcel(isbn, per, count)
            state = driver.find_element_by_xpath('//*[@id="mf_wfm_content_tac_main_contents_content1_body_gen_bksSrch_0_tbx_cmdtCdtnName"]').text
            price = driver.find_element_by_xpath('//*[@id="mf_wfm_content_tac_main_contents_content1_body_gen_bksSrch_0_tbx_byngPrce"]').text
            if(state != "정상"):
                writeStar(count)
            okToSell(count, per, price)
        except:
            pass
        finally:
            count+=1
            pass

def writeExcel(isbn, per, count):
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    sheet.cell(count,2,isbn)
    sheet.cell(count,17,per)
    bookList.save("test.xlsx")

def writeStar(count):
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    sheet.cell(count,1,"X")
    bookList.save("test.xlsx")

def okToSell(count, per, price):
    price = price.replace(',', '')
    price = price.replace('원', '')
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    ownPrice = sheet.cell(count,9).value
    okPrice = int(price) + int(ownPrice/10)
    if(ownPrice < okPrice):
        writeStar(count)


main()