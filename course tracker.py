import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkcalendar import *
import datetime
from tkinter import *
from tkinter import ttk




dt=datetime.datetime.now()#creating date and time instances
#creating time list
Class_time= ["8:00 AM","9:00 AM","10:00 AM", "11:00 AM" , "12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM"] 


#adding asistacne
win=tk.Tk()

#adding geometry
win.geometry('700x500')

#adding title
win.title(" Attendence Data")
#shutting down the resizable
#win.resizable(False,False)

#adding a level
ttk.Label(win,text="বিষয়  নির্বাচন  :").grid(column=0,row=0)

#creating combbox
number = tk.StringVar()
numberChosen = ttk.Combobox(win, width = 25, textvariable = number,font = 'fontExample')# Adding Values
numberChosen['values'] = ("Power Plant Engineering", "Industrial Management ", "Applied Engineering Mathematics","Noise and Vibration","Basic Mechatronics","Internal Combustion Engines	",	
" Fluid Machinery",		
" Machine Tools",		
"Production Planning and Control",					
"Refrigeration and Building Mechanical System")				



numberChosen.grid(column = 0, row = 1)
numberChosen.current(1)# Calling Main()



#creating option Menu of Class Time
variable = tk.StringVar(win)
variable.set(Class_time[3])

opt = tk.OptionMenu(win, variable, *Class_time)
opt.config(height=1,width=10, font=('Helvetica', 12))
opt.grid(row=3, column=0)


#creating Topic entry
topic_var=tk.StringVar()
ttk.Label(win,text="আলোচিত বিষয়").grid(row=3, column=1)
topic=tk.Entry(win,bd=2,width=30,font=('Helvetica', 14))
topic.grid(row=3, column=2,columnspan=20,padx=5, pady=5)









#creating calener
cal=Calendar(win,selectmode="day",year=dt.year,month=dt.month,day=dt.day)
cal.grid(row=10,column=0)



#Creating CheackButton
CheckVar1 =tk.IntVar()
CheckVar1.set(1)

C1=tk.Checkbutton(win,text="হাল",onvalue = 1,offvalue = 0,variable = CheckVar1)
C1.grid(row=4,column=0)
C1.configure(height=5,width = 10,bd=10,font = ("Helvetica", "16"))

#data confarming lavel




#Adding a button
def click_me():


    
    path='Attendence.xlsx'
    df1=pd.read_excel(path)
    SeriesA=df1['subject']
    SeriesB=df1['date']
    SeriesC=df1['time']
    SeriesD=df1['status']
    SeriesE=df1['Topic']
    A=pd.Series(numberChosen.get())
    B=pd.Series(str(cal.get_date()))
    C=pd.Series(variable.get())
    D=pd.Series(CheckVar1.get())
    E=pd.Series(topic.get())
    SeriesA=SeriesA.append(A)
    SeriesB=SeriesB.append(B)
    SeriesC=SeriesC.append(C)
    SeriesD=SeriesD.append(D)
    SeriesE=SeriesE.append(E)
    df2=pd.DataFrame({"subject":SeriesA,"date":SeriesB,"time":SeriesC,"status":SeriesD,"Topic":SeriesE})
    df2.to_excel(path,index=False)

    topic.delete(0,tk.END)
    #e2.delete(0,tk.END)



action=ttk.Button(win,text="জমা ",command=click_me)
action.grid(row=5,column=2)


#adding quit button:
class quitButton(ttk.Button):
    def __init__(self, parent):
        ttk.Button.__init__(self, parent)
        self['text'] = 'খোদাহাফেজ'
        # Command to close the window (the destory method)
        self['command'] = parent.destroy
        self.grid(row=11,column=2)


quitButton(win)














win.mainloop()
