from tkinter import *
from tkinter import messagebox
import pyupbit as coin

window = Tk()
window.geometry("380x200")
window.title("업비트 자동매매 프로그램")
window.resizable(width=0, height=0)

label_AC = Label(window, text="Access Key")
label_AC.grid(column=0, row=0)
AC_txt = Entry(window)
AC_txt.grid(row=0, column=1)

label_SC = Label(window, text="Secreat Key")
label_SC.grid(column=0, row=1)
SC_txt = Entry(window)
SC_txt.grid(row=1,column=1)

label_showMoney = Label(window, text="현재 잔액: ")
label_showMoney.grid(column=0, row=4)
showMoney_txt = Entry(window)
showMoney_txt.grid(row=4, column=1)

bt_start = Button(window,command=lambda : init(),width=15, height= 3, text="프로그램 시작").grid(row=3, column=0)
bt_showMoney = Button(window,command=lambda : showMoney(),width=15, height= 3, text="현재 잔액 표시").grid(row=3, column=1)

def init():
    access = AC_txt.get()
    secreat = SC_txt.get()
    global account

    account = coin.Upbit(access, secreat)
    messagebox.showinfo("알림", "시작완료")


def showMoney():
    myMoney = account.get_balance()
    current_KRW = myMoney
    showMoney_txt.delete(0, END)
    showMoney_txt.insert(0, str(int(current_KRW)))
    messagebox.showinfo("알림", "현재 잔액: {}".format(int(myMoney)))

window.mainloop()

