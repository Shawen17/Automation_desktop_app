#Declare sites up

from declare_functions import pastetext
from tkinter import *
from tkinter import filedialog as fd
import pandas as pd
import os
from datetime import datetime,date,time


window=Tk()
window.title('3G Report Declaration')
window.geometry('600x500')

link='O:\SEUN\2' 
link1='C://Users//NOC//'

    
    
label1=Label(window,text='File name:')
label1.place(relx=0,rely=0,relheight=0.05,relwidth=0.1)
label2=Label(window,text='File name:')
label2.place(relx=0,rely=0.05,relheight=0.05,relwidth=0.1)
label3=Label(window,text='File name:')
label3.place(relx=0,rely=0.1,relheight=0.05,relwidth=0.1)
label4=Label(window,text='File name:')
label4.place(relx=0,rely=0.15,relheight=0.05,relwidth=0.1)

e1=Entry(window)
e1.place(relx=0.1,rely=0,relheight=0.05,relwidth=0.5)
e2=Entry(window)
e2.place(relx=0.1,rely=0.05,relheight=0.05,relwidth=0.5)
e3=Entry(window)
e3.place(relx=0.1,rely=0.1,relheight=0.05,relwidth=0.5)
e4=Entry(window)
e4.place(relx=0.1,rely=0.15,relheight=0.05,relwidth=0.5)


button1=Button(window,text='Open',bg='grey',command=lambda: pastetext(e1))
button1.place(relx=0.6,rely=0,relheight=0.05,relwidth=0.1)
button2=Button(window,text='Open',bg='grey',command=lambda: pastetext(e2))
button2.place(relx=0.6,rely=0.05,relheight=0.05,relwidth=0.1)    
button3=Button(window,text='Open',bg='grey',command=lambda: pastetext(e3))
button3.place(relx=0.6,rely=0.1,relheight=0.05,relwidth=0.1)
button4=Button(window,text='Open',bg='grey',command=lambda: pastetext(e4))
button4.place(relx=0.6,rely=0.15,relheight=0.05,relwidth=0.1)    
entry_list=[e1,e2,e3,e4]
def merge_logs(entry_list):
    active_files=[pd.read_excel(i.get(),header=4)[:-2] for i in entry_list if len(i.get())!=0]
    comb_files=pd.concat(active_files,ignore_index=True,sort=False)
    comb_files=comb_files[['Location Information','Occurred On (NT)','Cleared On (NT)']]
    comb_files.rename(columns={'Location Information':'SITE_ID','Occurred On (NT)':'FROM (Date & Time)','Cleared On (NT)':'CLEARED'},inplace=True)
    comb_files['SITE_ID']=comb_files['SITE_ID'].str.extract('(U\w*)')
    comb_files['SITE_ID']=comb_files['SITE_ID'].str[1:]
    comb_files.to_excel('output_test.xlsx')
    

log_files='output_test.xlsx'
def run_report(log_files):
    comb_files=pd.read_excel(log_files)
    comb_files['FROM (Date & Time)']=pd.to_datetime(comb_files['FROM (Date & Time)'])
    comb_files['CLEARED']=pd.to_datetime(comb_files['CLEARED'])
    comb_files.dropna(inplace=True)
    comb_files['FROM (Date & Time)']=comb_files['FROM (Date & Time)'].apply(lambda x: x.strftime('%d/%b/%Y %H:%M'))
    comb_files['FROM (Date & Time)']=comb_files['FROM (Date & Time)'].apply(lambda x: datetime.strptime(x,'%d/%b/%Y %H:%M'))
    comb_files['CLEARED']=comb_files['CLEARED'].apply(lambda x: x.strftime('%d/%b/%Y %H:%M'))
    comb_files['CLEARED']=comb_files['CLEARED'].apply(lambda x: datetime.strptime(x,'%d/%b/%Y %H:%M'))

    excel_files=[file for file in os.listdir() if file.endswith('.xlsm')]
    lastest_file=max(excel_files,key=os.path.getmtime)
    report=pd.read_excel(lastest_file,header=2)
    report=report.fillna('')
    for index,row in report.iterrows():
        if row['FROM (Date & Time)']=='':
            report.at[index,'FROM (Date & Time)']=row['FROM (Date & Time)']
        else:
            report.at[index,'FROM (Date & Time)']=row['FROM (Date & Time)'].strftime('%d/%b/%Y %H:%M')


    report['FROM (Date & Time)']=pd.to_datetime(report['FROM (Date & Time)'])
    report=report.merge(comb_files,how='left',on=['SITE_ID','FROM (Date & Time)'])
    for index,row in report.iterrows():
        if row['TO (Date & Time)']=='':
            report.at[index,'TO (Date & Time)']=row['CLEARED']
        else:
            row['TO (Date & Time)']
    report.drop(columns='CLEARED',axis=1,inplace=True)
    report.to_excel('test_report.xlsx')
    
button5=Button(window,text='Run',bg='#00FFFF',command=merge_logs)
button5.place(relx=0.3,rely=0.25,relheight=0.05,relwidth=0.1) 
button6=Button(window,text='Declare',bg='#00FFFF',command=lambda: run_report(log_files))
button6.place(relx=0.45,rely=0.25,relheight=0.05,relwidth=0.1)
window.mainloop()