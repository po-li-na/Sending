from tkinter import *
from tkinter import filedialog
import pandas as pd
import pywhatkit
from pathlib import Path
from tkinter.messagebox import showerror, showwarning, showinfo
import re
from time import sleep
import threading
import keyboard
import random
class Main():
   
    def main(self):
        
        self.root=Tk()
        self.root['bg']='#7EE7ED'
        self.root.title('спамчик')
        self.root.geometry('300x150')
        x = (self.root.winfo_screenwidth() - self.root.winfo_reqwidth()) / 2.1
        y = (self.root.winfo_screenheight() - self.root.winfo_reqheight()) / 2
        self.root.wm_geometry("+%d+%d" % (x, y))
     
        self.df=[]


        self.btn_open=Button(self.root, text='open file', bg='yellow', command=self.openf)
        self.btn_open.place(relx=0.3, rely=0.4, relwidth=0.4, relheight=0.2)
        self.frame = Frame(self.root, background='#7EE7ED')
        self.sec_str = StringVar(value='15')
        
        # создаем текстовые поля для текста рассылки (несколько, чтоб номер не банился)
        self.text_msg1 = Text(self.frame,  font=("Bold",10), bg='yellow')
        self.text_msg2 = Text(self.frame,  font=("Bold",10), bg='yellow')
        self.text_msg3 = Text(self.frame,  font=("Bold",10), bg='yellow')
        self.text_msg4 = Text(self.frame,  font=("Bold",10), bg='yellow')
        self.root.mainloop()



    def openf(self):
        file=filedialog.askopenfilename()
            
        if file != '':
            if Path(file).suffix=='.xlsx':
                self.df=pd.read_excel(file)
                self.btn_open.destroy()
                self.root.geometry('1000x900')
                check = (self.root.register(self.is_valid), "%P")

                self.frame.place(relx=0.0001, rely=0.0001, relwidth=1, relheight=1)
                
                lab_sec=Label(self.frame, text="время задержки (в секундах)",background='#7EE7ED')
                lab_sec.place(relx=.3, rely=.1, anchor="c")  

                sec_sleep = Entry(self.frame,textvariable=self.sec_str, validatecommand=check,bg='yellow') 
                sec_sleep.place(relx=.54, rely=.08, relwidth=0.4, relheight=0.05)  

                lab_msg=Label(self.frame, text="                    текст сообщения",background='#7EE7ED')
                 

               
                self.text_msg1.place(relx=.1, rely=.15, relwidth=0.4, relheight=0.3)
                self.text_msg2.place(relx=.1, rely=.5, relwidth=0.4, relheight=0.3)
                self.text_msg3.place(relx=.55, rely=.15, relwidth=0.4, relheight=0.3)
                self.text_msg4.place(relx=.55, rely=.5, relwidth=0.4, relheight=0.3)
                btn_start=Button(self.frame, text='start', bg='green', command=self.start)
                btn_start.place(relx=0.4, rely=0.85, relwidth=0.2, relheight=0.1)

            else:
                showerror(title="Ошибка", message="файл должен быть с расширением .xlsx")

    def start(self):
        sec=int(self.sec_str.get())
        msgr1=self.text_msg1.get('1.0', END)
        msgr2=self.text_msg2.get('1.0', END)
        msgr3=self.text_msg3.get('1.0', END)
        msgr4=self.text_msg4.get('1.0', END)
        if sec<15:
             showerror(title="Ошибка", message="не менее 15 секунд")
        else:

            self.frame.destroy()
           

            self.df.columns=['nom']
            
            self.tf=False
            self.tf2=False
            t = threading.Thread(target=self.pot) # поток для ctrl+с
            t.start()
            sleep(2)
            n=[]
            for i in range(len(self.df.index)):
                sleep(2)
                if self.tf: # рассылка остановлена
                    self.tf2=True
                    break
                # рассылка
                n=self.df['nom'][i]
                msgr=eval(f'msgr{random.randrange(1,5)}') # текст рандомно подирается
                pywhatkit.sendwhatmsg_instantly(f'+{n}', msgr, sec, True, 2 )
                self.df.drop(labels=[i], axis=0, inplace=True)
               
            if self.tf2==False:  # рассылка завершена
                lab=Label(self.root, text="программа завершена",background='#7EE7ED')
                lab.place(relx=.5, rely=.4, anchor="c")
            else: # рассылка остановлена, кнопка для сохранения файла
                 btn_save=Button(self.root, text='save file', bg='#CD04FB', command=self.savef)
                 btn_save.place(relx=0.3, rely=0.4, relwidth=0.4, relheight=0.2)

    
    def savef(self): # сохранение файла с номерами 
        filepath = filedialog.asksaveasfilename(defaultextension='.xlxs', filetypes=(('Excel files', '*.xlsx'), ('Any file', '*')))
        if filepath != "":
            with pd.ExcelWriter(filepath) as file:
                self.df.to_excel(file, index=False)
 
            self.root.destroy()    
       

    def is_valid(self ,newval):
        result=  re.match("^[0-9]+$", newval) is not None
        return result
    
    def pot(self):
        keyboard.add_hotkey('Ctrl + C', self.tff) # горячие клавиши для завершения рассылки
               
    def tff(self):
        self.tf=True

if __name__ == "__main__":
    Main().main()

    