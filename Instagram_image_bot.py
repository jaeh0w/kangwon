from re import T
from tkinter import *
from urllib.request import urlopen
from urllib.parse import quote_plus
from bs4 import BeautifulSoup
import tkinter.font as ft
from selenium import webdriver
import time
import shutil
import requests

def main():
    main = Tk()
    main.title("인스타그램 이미지 저장 봇")
    main.geometry("600x500")

    myFont = ft.Font(family = "나눔고딕", size = 20)
    label_title = Label(main, text="인스타그램 이미지 저장 봇", bg="green", font = myFont)
    label_title.place(x =180, y =10)

    label_tag = Label(main, text="저장 할 태그 : ")
    entry_tag = Entry(main)
    bt_tag = Button(main, text = "가동", command=lambda : init__(entry_tag.get(), text_log))

    label_tag.place(x=130, y=50)
    entry_tag.place(x=210, y =50)
    bt_tag.place(x = 410, y = 50)

    text_log = Text(main, width=82, height= 30)
    text_log.place(x = 10, y = 90)

    main.mainloop()

def init__(tag, log):
    baseUrl = "https://www.instagram.com/explore/tags/"
    plusUrl = tag
    url = baseUrl+quote_plus(plusUrl)
    driver = webdriver.Safari()
    driver.get(url)
    time.sleep(3)
    
    html = driver.page_source
    soup = BeautifulSoup(html)

    imgList = []
    for i in range(10):
        insta = soup.select('.v1Nh3.kIKUG._bz0w')
        for i in insta:
            log.insert(END, "{}\n".format("https://www.instagram.com" + i.a['href']))
            imageUrl = i.select_one('.KL4Bh').img['src']
            imgList.append(imageUrl)
            imgList = list(set(imgList))
            html = driver.page_source
            soup = BeautifulSoup(html)
            insta = soup.select('.v1Nh3.kIKUG._bz0w')
        driver.execute_script("window.scrollTo(0, document.body.scorllHeight);")
        time.sleep(2)

    n = 0

    for i in range(len(imgList)):
        print(n)
        image_url = imgList[n]
        resp = requests.get(image_url, stream = True)
        local_file = open('./img/' + plusUrl + str(n) + '.jpg', 'wb')
        resp.raw.decode_content = True
        shutil.copyfileobj(resp.raw, local_file)
        n+=1
        del resp

    driver.close()


main()