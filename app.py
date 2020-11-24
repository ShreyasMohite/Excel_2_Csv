



from tkinter import *
from tkinter.ttk import Notebook,Progressbar,Combobox
import threading
import time
from tkinter import filedialog
import tkinter.messagebox
import pandas as pd






class Excel_2_Converting:
    def __init__(self,root):
        self.root=root
        self.root.title("Excel to Csv Convertion")
        self.root.geometry("500x300")
        self.root.iconbitmap("logo1102.ico")
        self.root.resizable(0,0)


        save=StringVar()


        def on_enter1(e):
            but_browse['background']="black"
            but_browse['foreground']="cyan"  
        def on_leave1(e):
            but_browse['background']="SystemButtonFace"
            but_browse['foreground']="SystemButtonText"

            

        def on_enter2(e):
            but_convert_png['background']="black"
            but_convert_png['foreground']="cyan"  
        def on_leave2(e):
            but_convert_png['background']="SystemButtonFace"
            but_convert_png['foreground']="SystemButtonText"

        def on_enter3(e):
            but_convert_jpg['background']="black"
            but_convert_jpg['foreground']="cyan"  
        def on_leave3(e):
            but_convert_jpg['background']="SystemButtonFace"
            but_convert_jpg['foreground']="SystemButtonText"



    #==============================command=======================================#

        def browse():
           global filename
           file_path = filedialog.askopenfilename(title = "Select file",filetypes = (("Excel","*.xlsx"),("all files","*.*")))
           if len(file_path)!=0:
               lab_file.config(text="File is selected")
           filename=file_path


        def convert_excel():
            try:

                if save.get()!="":
                    read_file=pd.read_excel(filename)
                    read_file.to_csv('{}.csv'.format(save.get(),index=None,heaader=True))
                    
                else:
                    tkinter.messagebox.showerror("Error","please enter  name to save")
            except Exception as e:
                print(e)

        def con_excel():
            t1=threading.Thread(target=convert_excel)
            t1.start()

        def clear():
            save.set("")



    #========================frame========================================#
        mainframe=Frame(self.root,width=500,height=300,relief="ridge",bd=3)
        mainframe.place(x=0,y=0)

        top_frame=Frame(mainframe,width=495,height=265,relief="ridge",bd=3)
        top_frame.place(x=0,y=0)

        task_bar_frame=Frame(mainframe,width=495,height=30,relief="ridge",bd=3)
        task_bar_frame.place(x=0,y=265)

    #=====================================================================#

        lab_frame=LabelFrame(top_frame,text="Excel to Csv Convertion",width=490,height=258,font=('times new roman',12,'bold'),bg="#1c45ea",fg="white")
        lab_frame.place(x=0,y=0)

    #========================lab_frame=============================================#

        but_browse=Button(lab_frame,text="Browse",width=15,font=('times new roman',12,'bold'),cursor="hand2",command=browse)
        but_browse.place(x=170,y=50)
        but_browse.bind("<Enter>",on_enter1)
        but_browse.bind("<Leave>",on_leave1)

        lab_file=Label(lab_frame,text="Select the file",font=('times new roman',12,'bold'),bg="#1c45ea",fg="white")
        lab_file.place(x=190,y=10)

        lab_enter_name=Label(lab_frame,text="Enter Name to save",font=('times new roman',11,'bold'),bg="#1c45ea",fg="white")
        lab_enter_name.place(x=180,y=120)

        ent_name=Entry(lab_frame,width=18,font=('times new roman',14,'bold'),relief="ridge",bd=3,textvariable=save)
        ent_name.place(x=150,y=150)

        but_convert_jpg=Button(lab_frame,text="Convert-Csv",width=15,font=('times new roman',12,'bold'),cursor="hand2",command=con_excel)
        but_convert_jpg.place(x=20,y=190)
        but_convert_jpg.bind("<Enter>",on_enter3)
        but_convert_jpg.bind("<Leave>",on_leave3)


        but_convert_png=Button(lab_frame,text="Clear",width=15,font=('times new roman',12,'bold'),cursor="hand2",command=clear)
        but_convert_png.place(x=330,y=190)
        but_convert_png.bind("<Enter>",on_enter2)
        but_convert_png.bind("<Leave>",on_leave2)

    #=============================task_bar=========================================================#

        prg=Progressbar(task_bar_frame,length=489,orient=HORIZONTAL,mode='indeterminate')
        prg.place(x=0,y=0)


if __name__ == "__main__":
    root=Tk()
    app=Excel_2_Converting(root)
    root.mainloop()
