import tkinter as tk
from tkinter import scrolledtext
from tkinter import END
from main import *


class MyGui:
    def __init__(self):
        self.isRun = False

        w = tk.Tk()
        w.geometry('600x600')
        self.w = w

        lb1 = tk.Label(w, text='엑셀 파일명을 입력하세요.')
        lb1.grid(row=0, column=0, rowspan=1, columnspan=2, sticky='news', pady=5, padx=15)

        lb2 = tk.Label(w, text='x값을 입력하세요.')
        lb2.grid(row=0, column=2, rowspan=1, columnspan=2, sticky='news', pady=5, padx=15)

        lb3 = tk.Label(w, text='y값을 입력하세요.')
        lb3.grid(row=0, column=4, rowspan=1, columnspan=2, sticky='news', pady=5, padx=15)

        self.en1 = tk.Entry()
        self.en1.insert(0, 'input.xlsx')
        self.en1.grid(row=1, column=0, rowspan=1, columnspan=2, sticky='news', pady=5)

        self.en2 = tk.Entry()
        self.en2.insert(0, 50)
        self.en2.grid(row=1, column=2, rowspan=1, columnspan=2, sticky='news', pady=5)

        self.en3 = tk.Entry()
        self.en3.insert(0, 1000)
        self.en3.grid(row=1, column=4, rowspan=1, columnspan=2, sticky='news', pady=5)

        b1 = tk.Button(w, text='실행')
        b1.grid(row=2, column=2, rowspan=2, columnspan=2, sticky='news', pady=5)
        b1.config(command=self.run)

        self.lb4 = tk.Label(w, text='준비중')
        self.lb4.grid(row=2, column=4, rowspan=1, columnspan=1, sticky='news', pady=5, padx=15)

        self.scText = scrolledtext.ScrolledText(w)
        self.scText.grid(row=5, column=1, rowspan=3, columnspan=5, sticky='news', pady=5, padx=15)

        w.columnconfigure(tuple(range(6)), weight=1)
        w.rowconfigure(tuple(range(8)), weight=1)

        w.mainloop()

    def run(self):
        filename = self.en1.get()
        x = int(self.en2.get())
        y = int(self.en3.get())

        if not self.isRun:
            self.lb4.config(text = '작동중')
            self.w.update()
            self.isRun = True
            cleanCopyExcel(fileName=filename, MAX_X=x, MAX_Y=y, logF=self.logF)
            self.isRun = False
            self.lb4.config(text='끝')

    def logF(self, s):
        self.scText.insert(END, s + '\n')
        self.w.update()



if __name__ == '__main__':
    mygui = MyGui()
