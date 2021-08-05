from tkinter import *
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from selenium.webdriver.common.keys import Keys
import time , datetime

def main():
    main = Tk()
    main.title("YES24 ISBN 봇")
    main.geometry("400x200")

    button_target = Button(main, text="시작", command=lambda : init(1), fg="green", bg = "black", width=30, height=4).place(x = 40, y = 100)
    main.mainloop()

def init(t):
    url = "http://www.yes24.com/main/default.aspx"
    driver = webdriver.Safari()
    driver.maximize_window()
    driver.get(url)
    excel_name = "test.xlsx"
    wb = load_workbook(filename='test.xlsx')
    ws = wb.active
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="divYes24SCMEvent"]/div[2]/div[1]/label'))).click()
    for i in range(1, ws.max_row+1):
        try:
            bookName = ws.cell(i,10).value
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="yesSForm"]/fieldset/span[1]'))).send_keys(Keys.ENTER)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="query"]'))).send_keys(bookName)
            driver.find_element_by_xpath('//*[@id="yesSForm"]/fieldset/span[2]/button').send_keys(Keys.ENTER)
            time.sleep(1)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, 'img_bdr'))).send_keys(Keys.ENTER)
            isbn = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="infoset_specific"]/div[2]/div/table/tbody/tr[3]/td'))).text
            writeExcel(isbn, i)
        except:
            pass

def writeExcel(isbn, count):
    bookList=load_workbook(filename='test.xlsx')
    sheet = bookList.active
    sheet.cell(count,2,isbn)
    bookList.save("test.xlsx")

main()