from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter.font as ft
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import time , datetime

def main():
    main = Tk()
    main.title("타겟팅 봇")
    main.geometry("400x200")

    myFont = ft.Font(family = "나눔고딕")
    label_target = Label(main, text="타겟 대상 : ", font=myFont).place(x = 40, y = 20)
    label_excel = Label(main, text="엑셀 이름 : ", font=myFont).place(x = 40, y = 50)
    label_max = Label(main, text="최대 페이지 : \n", font=myFont).place(x = 28, y = 80)
    label_max = Label(main, text="시작 페이지 : \n ", font=myFont).place(x = 28, y = 110)
    entry_target = Entry(main)
    entry_target.place(x = 105, y= 20)
    global entry_excel 
    entry_excel= Entry(main)
    entry_excel.place(x = 105, y= 50)
    global entry_max 
    entry_max = Entry(main)
    entry_max.place(x = 105, y= 80)
    entry_min = Entry(main)
    entry_min.place(x = 105, y= 110)
    button_target = Button(main, text="시작", command=lambda : targeting(entry_target.get()), fg="green", bg = "black", width=30, height=2).place(x = 40, y = 140)
    main.mainloop()


def targeting(t):
    if(str(t) == "북코치"):
        baseUrl = "http://www.yes24.com//24/usedShop/mall/cjudeer/main"
    elif(str(t) == "마음북"):
        baseUrl = "http://www.yes24.com//24/usedShop/mall/saguraba/main"
    elif(str(t) == "할인에할인"):
        baseUrl = "http://www.yes24.com//24/usedShop/mall/freeman001/main"
    else:
        messagebox.showinfo("알림", "올바른 대상을 입력하세요!")
    #http://www.yes24.com//24/usedShop/mall/saguraba/main 마음북
    #http://www.yes24.com//24/usedShop/mall/cjudeer/main 북코치
    #http://www.yes24.com//24/usedShop/mall/freeman001/main 할인에할인
    # url = baseUrl+str((plusUrl))
    url = baseUrl
    driver = webdriver.Safari()
    driver.maximize_window()
    driver.get(url)
    time.sleep(2)

    # WebDriverWait(driver,1).until(EC.element_to_be_clickable((By.XPATH('//*[@id="divYes24SCMEvent"]/div[2]/div[1]/label'))))
    driver.find_element_by_xpath('//*[@id="divYes24SCMEvent"]/div[2]/div[1]/label').click()

    getInfo(driver)


def getInfo(driver):
    count = 0
    xNo = 3
    max = int(entry_max.get())
    for j in range(1,max):
        href = "#entrno=242444&mallid=cjudeer&searchname=&filtername=NAME&PageNumber=" + str(xNo)
        driver.find_element_by_css_selector("a[href='{}']".format(href)).click()
        try:
            scrollDown(driver)
            time.sleep(1)
            lst = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'strong.used-class-lowest')))
            for i in range(len(lst)):
                try:
                    lst[i].click()
                    title = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'h2.gd_name'))).text
                    price = driver.find_element_by_tag_name('em.yes_m').text
                    auth = driver.find_element_by_tag_name('span.gd_auth').text
                    publisher = driver.find_element_by_tag_name('span.gd_pub').text
                    date = driver.find_element_by_tag_name('span.gd_date').text
                    price = price.replace(",", "")
                    Yes24WriteExcel(123, title, price, auth, publisher, date, count)
                finally:
                    count+=1
                    driver.back()
                    pass
        except:
            print(123123)
            pass
        time.sleep(1)
        print(xNo)
        try:
            driver.find_element_by_css_selector("a[href='{}']".format(href)).click()
        except:
            pass
        finally:
            xNo+=1
            #entrno=242444&mallid=cjudeer&searchname=&filtername=NAME&PageNumber=13

            # a = WebDriverWait(driver,5).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "p.page > *")))
            # a[j].click()
        
        # driver.find_element_by_xpath('//*[@id="div_ListContainer"]/div[4]/center/p/a['+ str(j) +']').click()

        # driver.find_element_by_xpath('//*[@id="div_ListContainer"]/div[4]/center/p/a['+str(xNo)+']').click()
        # xNo+=1
        # print(xNo)
        # if(xNo == 13):
        #     xN0 = 3
                        # driver.find_element_by_xpath('//a[contains(@href,"#entrno=242444&mallid=cjudeer&searchname=&filtername=NAME&PageNumber='+str(j)+'")]').click()
        # driver.find_element_by_xpath('//a[@href="'+"#entrno=242444&mallid=cjudeer&searchname=&filtername=NAME&PageNumber="+str(j)+'"]').click()
        # //*[@id="div_ListContainer"]/div[3]/center/p/a[3]
        # //*[@id="div_ListContainer"]/div[3]/center/p/a[3]

        scrollUp(driver)
#
# //*[@id="div_ListContainer"]/div[4]/center/p/a[4] # 2
# //*[@id="div_ListContainer"]/div[4]/center/p/a[5] # 4
# //*[@id="div_ListContainer"]/div[4]/center/p/a[8] # 7
# //*[@id="div_ListContainer"]/div[4]/center/p/a[11] # 10
# //*[@id="div_ListContainer"]/div[4]/center/p/a[11] # 10
# //*[@id="div_ListContainer"]/div[4]/center/p/a[12] # 11
# //*[@id="div_ListContainer"]/div[4]/center/p/a[4]
# //*[@id="div_ListContainer"]/div[3]/center/p/a[12] # 11
# //*[@id="div_ListContainer"]/div[3]/center/p/a[3] # 12
# //*[@id="div_ListContainer"]/div[3]/center/p/a[4] # 13
# //*[@id="div_ListContainer"]/div[4]/center/p/a[5] # 14
# //*[@id="div_ListContainer"]/div[4]/center/p/a[6] # 15
# //*[@id="div_ListContainer"]/div[4]/center/p/a[7] # 16
# //*[@id="div_ListContainer"]/div[4]/center/p/a[8] # 17
# //*[@id="div_ListContainer"]/div[4]/center/p/a[9] # 18
# //*[@id="div_ListContainer"]/div[4]/center/p/a[10] # 19
# //*[@id="div_ListContainer"]/div[4]/center/p/a[11] # 20
# //*[@id="div_ListContainer"]/div[4]/center/p/a[12] # 21
# //*[@id="div_ListContainer"]/div[4]/center/p/a[3] # 22
# //*[@id="div_ListContainer"]/div[4]/center/p/a[4] # 23
# //*[@id="div_ListContainer"]/div[4]/center/p/a[5] # 24
# //*[@id="div_ListContainer"]/div[4]/center/p/a[4] # 33
# //*[@id="div_ListContainer"]/div[4]/center/p/a[13] # 


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


def Yes24WriteExcel(isbn, title, price, auth, publisher, date, count):
    excel_name = "test.xlsx"
    wb = load_workbook(excel_name)
    ws = wb.active
    # ws.cell(count,2,isbn)
    # ws.cell(count,3,1111)
    # ws.cell(count,4,"A")
    # ws.cell(count,5,"A")
    # ws.cell(count,6,"N")
    # ws.cell(count,7,"1")
    # ws.cell(count,9, int(price)-50)
    # ws.cell(count,10,title)
    # ws.cell(count,13,"품질보장~!! (미사용 ★ 출판사에서 직접구매한 새★책 ^^)  ■  >□<")
    # ws.cell(count,15,1)
    # ws.cell(count,16,1)
    lst = ["",isbn, 1111, 'A', 'A', 'N', 1, "", int(price)-50, title]
    ws.append(lst)
    wb.save(excel_name)

main()


