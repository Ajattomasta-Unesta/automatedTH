

'''print("MAC Address :", get_mac_address())
f = open("moonlight.dat", "w")
f.write(get_mac_address())
f.close()

#from getmac import get_mac_address
try:
    ff = open("moonlight.dat", "r")
    if ff.read() == get_mac_address() : print("BENE")
    else: print("MALE")
    ff.close()
except : pass'''


#파일이름지정가능 미입력시 기존처럼 자동
from tkinter import *
from tkinter import filedialog

from sys import exit
from getmac import get_mac_address
def funcc() :
    print(stv_fn.get())
if __name__ == "__main__":

    ###################CERT####################
    try:
        ff = open("moonlight.dat", "r")
        if ff.read() == get_mac_address():
            print("VABENE")
        else: raise Exception
    except:
        exit()

    ff.close()
    ###################CERT####################

    root = Tk()
    root.title("경비원 배치폐지신고서 자동생성")

    lbl = Label(root, text="경비원 배치폐지신고서")
    lbl.grid(row=0, column=0, padx=15, pady=15)

    btn_xl = Button(root, text="엑셀 파일 선택")
    btn_xl.grid(row=1, column=0, padx=15, pady=15)

    stv_xl = StringVar()
    lb_xl = Label(root, text="엑셀 파일 경로", textvariable=stv_xl)
    lb_xl.grid(row=1, column=1, padx=15, pady=15)

    btn_ph = Button(root, text="내보낼 폴더 선택")
    btn_ph.grid(row=2, column=0, padx=15, pady=15)

    stv_ph = StringVar()
    lb_ph = Label(root, text="내보낼 폴더 경로", textvariable=stv_ph)
    lb_ph.grid(row=2, column=1, padx=15, pady=15)

    lb_xx = Label(root, text="파일명 입력란\n미입력시 이름 자동생성, 확장자 입력 필요 없음")
    lb_xx.grid(row=3, column=0, padx=25, pady=25)

    stv_fn = StringVar()
    sheetname = Entry(root, text="Sheet1", textvariable=stv_fn)
    sheetname.grid(row=3, column=1, padx=25, pady=25)

    btn_st = Button(root, text="생성 시작")
    btn_st.grid(row=3, column=2, padx=25, pady=25)

    root.mainloop()
    funcc()
