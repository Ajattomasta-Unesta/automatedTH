from tkinter import *
from tkinter import filedialog

def load_xl():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
    print(filename)
    stv_xl.set(filename)

def load_ph():
    filename = filedialog.askdirectory(initialdir="/", title="Select file")
    print(filename)
    stv_ph.set(filename)

if __name__ == "__main__":
    root = Tk()
    root.title("경비원 배치폐지신고서 자동생성")

    lbl = Label(root, text="경비원 배치폐지신고서")
    lbl.grid(row=0, column=0, padx=15, pady=15)

    btn_xl = Button(root, text="엑셀 파일 선택", command = load_xl)
    btn_xl.grid(row=1, column=0, padx=15, pady=15)

    stv_xl = StringVar()
    lb_xl = Label(root, text="엑셀 파일 경로", textvariable=stv_xl)
    lb_xl.grid(row=1, column=1, padx=15, pady=15)

    btn_ph = Button(root, text="내보낼 폴더 선택", command=load_ph)
    btn_ph.grid(row=2, column=0, padx=15, pady=15)

    stv_ph = StringVar()
    lb_ph = Label(root, text="내보낼 폴더 경로", textvariable=stv_ph)
    lb_ph.grid(row=2, column=1, padx=15, pady=15)

    sheetname = Entry(root, text="Sheet1")
    sheetname.grid(row=3, column=0, padx=25, pady=25)

    btn_st = Button(root, text="생성 시작")
    btn_st.grid(row=3, column=1, padx=25, pady=25)

    root.mainloop()