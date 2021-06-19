from tkinter import *
from tkinter import messagebox
import tkinter
import pyupbit as coin
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

def main():

    window = Tk()
    window.geometry("550x700")
    window.title("업비트 자동매매 프로그램")

    #전역변수
    global txt_AC 
    global txt_SC
    global showMoney_txt

    #Label - 라벨
    label_AC = Label(window, text="Access Key")
    label_SC = Label(window, text="Secreat Key")
    label_showMoney = Label(window, text="현재 잔액: ")

    #Text - 텍스트
    Text_trade = Text(window,width=68, height=35)


    #Button - 버튼
    bt_start = Button(window,command=lambda : init(),width=15, height= 3, text="프로그램 시작")
    bt_showMoney = Button(window,command=lambda : fn_showMoney(),width=15, height= 3, text="현재 잔액 표시")
    bt_func = Button(window,command=lambda : window_func(),width=15, height= 3, text="부가 기능")

    #Entry - 엔트리
    txt_AC = Entry(window)
    txt_SC = Entry(window)
    showMoney_txt = Entry(window)


    #배치
    label_AC.grid(row=0, column=0)
    txt_AC.grid(row=0, column=1)
    label_showMoney.grid(column=0, row=4)

    label_SC.grid(row=1, column=0)
    txt_SC.grid(row=1,column=1)

    bt_start.grid(row=3, column=0)
    bt_showMoney.grid(row=3, column=1)
    bt_func.grid(row=3, column=2)
    
    showMoney_txt.grid(row=4, column=1)

    Text_trade.place(x=30, y= 170)


    window.mainloop()


def init():
    access = txt_AC.get()
    secreat = txt_SC.get()
    global account

    account = coin.Upbit(access, secreat)
    messagebox.showinfo("알림", "시작완료")


def fn_showMoney():
    myMoney = account.get_balance()
    current_KRW = myMoney
    showMoney_txt.delete(0, END)
    showMoney_txt.insert(0, str(int(current_KRW)))
    messagebox.showinfo("알림", "현재 잔액: {}".format(int(myMoney)))

def window_func():
    window_func = Tk()
    window_func.geometry("600x600")
    window_func.title("부가 기능")

    bt_getCoins = Button(window_func,command=lambda : window_getCoin(),width=15, height= 3, text="KRW 코인 종류")
    bt_getgraph = Button(window_func,command=lambda : fn_getGraph(entry_coinName.get()),width=15, height= 3, text="그래프 표시")
    bt_empty = Button(window_func,width=15, height= 3, text="개발중")

    bt_getCoins.grid(row=0, column=0)
    bt_getgraph.grid(row=0, column=1)
    bt_empty.grid(row=0, column=2)

    entry_coinName = Entry(window_func)
    entry_coinName.grid(row=1, column=1)

    window_func.mainloop()


def window_getCoin():
    window_getCoin = Tk()
    window_getCoin.geometry("400x400")
    window_getCoin.title("KRW 코인 종류")
    label_coinNo = Label(window_getCoin, text = "KRW 코인 종류").grid(row=0, column=0)
    coin_txt = Text(window_getCoin, width=50)
    coinList = coin.get_tickers("KRW")
    coin_txt.delete('1.0', END)
    coin_txt.insert(END, coinList)
    coin_txt.place(x=10, y=20)
    window_getCoin.mainloop()

def fn_getGraph(ticker):
    plt.rcParams["123.figsize"] = (6,3)
    plt.rcParams["axes.formatter.limits"] = -10000, 10000
    price = coin.get_ohlcv(ticker, interval= 'minute1', count = 1000)
    price[["close", "volume"]].plot(secondary_y=["volume"])
    plt.show()

main()



