from tkinter import *
import tkinter.font as ft
from selenium import webdriver
from openpyxl import load_workbook
import time , datetime

def main():
    main = Tk()
    main.title("타겟팅 봇")
    main.geometry("400x200")

    myFont = ft.Font(family = "나눔고딕")
    label_target = Label(main, text="타겟 대상 : ", font=myFont).place(x = 40, y = 20)
    entry_target = Entry(main)
    entry_target.place(x = 105, y= 20)
    button_target = Button(main, text="시작", command=lambda : targeting(entry_target.get()), fg="green", width=5).place(x = 300, y = 18)
    main.mainloop()


def targeting(t):
    baseUrl = "http://www.yes24.com//24/usedShop/mall/cjudeer/main"
    # url = baseUrl+str((plusUrl))
    url = baseUrl
    driver = webdriver.Safari()
    driver.maximize_window()
    driver.get(url)
    time.sleep(1)

    driver.find_element_by_xpath('//*[@id="divYes24SCMEvent"]/div[2]/div[1]/label').click()

    getInfo(driver)


def getInfo(driver):
    count = 0
    for k in range(10):
        for j in range(3,13):
            scrollDown(driver)
            lst = driver.find_elements_by_css_selector("strong.used-class-lowest") #최저가 탐색
            time.sleep(1)
            for i in range(len(lst)):
                try:
                    lst[i].click()
                    time.sleep(1)
                    title = driver.find_elements_by_tag_name('h2.gd_name')[0].text
                    price = driver.find_elements_by_tag_name('em.yes_m')[0].text
                    auth = driver.find_elements_by_tag_name('span.gd_auth')[0].text
                    publisher = driver.find_elements_by_tag_name('span.gd_pub')[0].text
                    date = driver.find_elements_by_tag_name('span.gd_date')[0].text
                    # isbn = getISBN(driver)
                    driver.back()
                    count+=1
                    price = price.replace(",", "")
                    Yes24WriteExcel(123, title, price, auth, publisher, date, count)
                except:
                    driver.back()
                    pass
            a = driver.find_elements_by_css_selector("p.page > *") #페이지 넘기
            a[j].click()
            scrollUp(driver)

def scrollUp(driver):
    start = datetime.datetime.now() 
    end = start + datetime.timedelta(seconds=1)
    while True:
        driver.execute_script('window.scrollTo(document.body.scrollHeight, 0);')
        if datetime.datetime.now() > end:
            break


def scrollDown(driver):
    start = datetime.datetime.now() 
    end = start + datetime.timedelta(seconds=1)
    while True:
        driver.execute_script('window.scrollTo(0, document.body.scrollHeight);')
        if datetime.datetime.now() > end:
            break

def getISBN(driver):
    try:
        driver.find_element_by_xpath('//*[@id="divFormatInfo"]/ul/li[1]').click()
        time.sleep(1)
        driver.switch_to_window(driver.window_handles[1]) 
        driver.get_window_position(driver.window_handles[1])
        isbn = driver.find_element_by_xpath('//*[@id="infoset_specific"]/div[2]/div/table/tbody/tr[3]/td').text
        driver.close()
        time.sleep(3)
        driver.switch_to.window(driver.window_handles[0])
        return isbn
    except:
        return 0



def Yes24WriteExcel(isbn, title, price, auth, publisher, date, count):

    wb = load_workbook("test.xlsx")
    ws = wb.active
    ws.cell(count,2,isbn)
    ws.cell(count,3,1111)
    ws.cell(count,4,"A")
    ws.cell(count,5,"A")
    ws.cell(count,6,"N")
    ws.cell(count,7,"1")
    ws.cell(count, 9, int(price)-50)
    ws.cell(count,10,title)
    ws.cell(count,13,"품질보장~!! (미사용 ★ 출판사에서 직접구매한 새★책 ^^)  ■  >□<")
    ws.cell(count,15,1)
    ws.cell(count,16,1)
    wb.save("test.xlsx")

main()


