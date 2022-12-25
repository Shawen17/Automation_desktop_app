#functions associated with Escalation,Update Request and Reporting
from pathlib import Path
import re
from tkinter import *
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter import filedialog as fd
import pandas as pd
import os
from datetime import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from time import sleep
from datetime import datetime,date,time
from PIL import ImageTk, Image
import sys



def resource_path(relative_path):
    try:
        base_path=sys._MEIPASS
    except Exception:
        base_path=os.path.abspath('.')
    return os.path.join(base_path,relative_path)
    
    


ref_date=datetime.now()
ref_data=ref_date.strftime('%d/%B/%Y %H:%M')
ref_data=ref_data.split('/')
month=ref_data[1].upper()
year=ref_data[2].split(' ')[0]
year
path1=Path(f'R:\\noc\\NOC REPORT\\NOC REPORT 3G\\{year}\\{month}')
path2=Path(f'R:\\noc\\NOC REPORT\\NOC REPORT 2G\\{year}\\{month}')
rd_base=pd.read_excel('R:\\noc\\RDBASE\\REPORTDBASE2.xlsm')[['SITE_ID','TECH STATE']]

def pastetext(entry):
    link='O:\SEUN\2' 
    name=fd.askopenfilenames(initialdir=link,title='select file',filetypes=(('all files','*.*'),('jpeg files','*.jpg')))
    entry.insert(END,name)

def clean_3g(string):
    if 'NodeB' in string:
        t=re.findall('U\w*',string)
        return(t[0][1:])
    elif '_' in string and len(string)<=11:
        t=string[1:-4]
        return (t)
    elif len(string)==6:
        return(string)
    else:
        t=string[1:]
        return(t)

def clean_2g(string):
    if 'D' not in string:
        return(string)
    elif (len(string))>=6 and string.index('D')==0:
        t=string[1:]
        return t   

def clean_alarm(string):
    if 'OML Fault' in string:
        return string.replace('OML Fault','DOWN')
    elif 'CSL Fault' in string:
        return string.replace('CSL Fault','DOWN')
    elif 'Mains' in string:
        matches=re.findall('M\w{4}',string)
        return matches[0].replace('Mains','Rectifier Main')
    elif 'Urgent' in string:
        matches=re.findall('U\w{5}',string)
        return matches[0].replace('Urgent','Rectifier Urgent')
    elif 'System' in string:
        matches=re.findall('S\w{5}',string)
        return matches[0].replace('System','Gen System')    
    elif 'Gen' in string or 'Gen2' in string:
        return (string.replace(string,'Gen System'))
    
    
def merge_log(entry_list,v,activity):
    if v.get()==2:
        active_files=[a.get()[1:-1] for a in entry_list if len(a.get())!=0]
        active_files=[pd.read_excel(i,header=5)[:-2] for i in active_files]
        
        comb_files=pd.concat(active_files,ignore_index=True,sort=False)
        comb_files=comb_files[['Location Information','Occurred On (NT)','Cleared On (NT)']]
        comb_files.rename(columns={'Location Information':'SITE_ID','Occurred On (NT)':'FROM (Date & Time)','Cleared On (NT)':'CLEARED'},inplace=True)
        comb_files['SITE_ID']=comb_files['SITE_ID'].apply(lambda x: clean_3g(x))
        comb_files.to_excel('output_test.xlsx')
    elif v.get()==1:
        active_files=[a.get()[1:-1] for a in entry_list if len(a.get())!=0]
        active_files=[pd.read_excel(i,header=5)[:-2] for i in active_files]
        comb_files=pd.concat(active_files,ignore_index=True,sort=False)
        if activity.get()=='Escalation':
            comb_files=comb_files[['MO Name','Name']]
            comb_files.rename(columns={'MO Name':'SITE_ID','Name':'Alarm'},inplace=True)
        else:
            comb_files=comb_files[['MO Name','Occurred On (NT)','Cleared On (NT)']]
            comb_files.rename(columns={'MO Name':'SITE_ID','Occurred On (NT)':'FROM (Date & Time)','Cleared On (NT)':'CLEARED'},inplace=True)
            comb_files['SITE_ID']=comb_files['SITE_ID'].apply(lambda x: clean_2g(x))
        comb_files.to_excel('output_test.xlsx')

def run_report(log_files,v):
    global latest_file
    if v.get()==2:
        comb_files=pd.read_excel(log_files)
        comb_files['FROM (Date & Time)']=pd.to_datetime(comb_files['FROM (Date & Time)'])
        comb_files['CLEARED']=pd.to_datetime(comb_files['CLEARED'])
        comb_files.dropna(inplace=True)
        comb_files['FROM (Date & Time)']=comb_files['FROM (Date & Time)'].apply(lambda x: x.strftime('%d/%b/%Y %H:%M'))
        comb_files['FROM (Date & Time)']=comb_files['FROM (Date & Time)'].apply(lambda x: datetime.strptime(x,'%d/%b/%Y %H:%M'))
        comb_files['CLEARED']=comb_files['CLEARED'].apply(lambda x: x.strftime('%d/%b/%Y %H:%M'))
        comb_files['CLEARED']=comb_files['CLEARED'].apply(lambda x: datetime.strptime(x,'%d/%b/%Y %H:%M'))
 
        report=pd.read_excel(latest_file,header=2)
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
        
        report.to_excel('test_report.xlsx')
    elif v.get()==1:
        comb_files=pd.read_excel(log_files)
        comb_files['FROM (Date & Time)']=pd.to_datetime(comb_files['FROM (Date & Time)'])
        comb_files['CLEARED']=pd.to_datetime(comb_files['CLEARED'])
        comb_files.dropna(inplace=True)
        comb_files['FROM (Date & Time)']=comb_files['FROM (Date & Time)'].apply(lambda x: x.strftime('%d/%b/%Y %H:%M'))
        comb_files['FROM (Date & Time)']=comb_files['FROM (Date & Time)'].apply(lambda x: datetime.strptime(x,'%d/%b/%Y %H:%M'))
        comb_files['CLEARED']=comb_files['CLEARED'].apply(lambda x: x.strftime('%d/%b/%Y %H:%M'))
        comb_files['CLEARED']=comb_files['CLEARED'].apply(lambda x: datetime.strptime(x,'%d/%b/%Y %H:%M'))
        excel_files=[i for i in path2.glob('*.xlsm')]
        file=max(excel_files,key=os.path.getmtime)
        report=pd.read_excel(file,header=2)
        report=report.fillna('')
        for index,row in report.iterrows():
            if row['FROM (Date & Time)']=='':
                report.at[index,'FROM (Date & Time)']=row['FROM (Date & Time)']
            else:
                report.at[index,'FROM (Date & Time)']=row['FROM (Date & Time)'].strftime('%d/%b/%Y %H:%M')


        report['FROM (Date & Time)']=pd.to_datetime(report['FROM (Date & Time)'])
        report=report.merge(comb_files,how='left',on=['SITE_ID','FROM (Date & Time)'])
        isna=report['TO (Date & Time)'].isnull()
        report.loc[isna,'TO (Date & Time)']=report['CLEARED'].values
        '''for index,row in report.iterrows():
            if row['TO (Date & Time)']=='':
                report.at[index,'TO (Date & Time)']=row['CLEARED']
            else:
                row['TO (Date & Time)']'''
        report.drop(columns='CLEARED',axis=1,inplace=True)
        report.to_excel('test_report.xlsx')
    messagebox.showinfo('Declare Up','Task Done')

def refresh(e,P,C):
    e.delete(1.0,END)
    P['value']=0
    P.update()
    C.current(0)



def perform(downloaded,v,V,browser,e,operation,newest_file,RDB,progress):
    should_restart=True
    while should_restart:
        if  (v.get()==2 or  v.get()==1):
            should_restart=False
        else:
            msgbox=messagebox.showinfo('CHOOSE','choose 2G or 3G')
            return
    should_restart=True    
    while should_restart:
        if (V.get()==3 or V.get()==4 or V.get()==5):
            should_restart=False
        else:
            msgbox=messagebox.showinfo('CHOOSE','choose preffered sms platform')
            return
    
    if V.get()==3:
        url='http://10.100.111.22/smsfeeder'
    elif V.get()==4:
        url='http://10.100.111.222/smsfeeder'
    else:
        url='http://10.100.111.221/smsfeeder'
    
    browser.implicitly_wait(15)
    browser.get(url)
    user_input=browser.find_elements_by_class_name("TextField")
    input_username=user_input[0].send_keys('nmc')
    input_password=user_input[1].send_keys('nmc123')
    login_button = browser.find_element_by_class_name("Button" )
    login_button.click()
    message_type=browser.find_element_by_xpath("//*[@id='mtype']/option[text()='Text Message']").click()
    
   
    present=datetime.now()
    if present.time() >= time(1,0,0,0) and present.time() < time(7,0,0,0):
        greeting='Morning,'
        t='5AM'
    elif present.time() >= time(7,1,0,0) and present.time() < time(11,59,0,0):
        greeting='Morning,'
        t='7AM'
    elif present.time() >= time(12,0,0,0) and present.time() <= time(19,0,0,0):
        greeting='Afternoon,'
        t='1PM'
    elif present.time() >= time(19,0,0,0) and present.time() <= time(21,59,0,0):
        greeting='Good Evening,'
        t='7PM'
    elif present.time() >= time(22,0,0,0) and present.time() <= time(23,59,0,0):
        greeting='Good Evening,'
        t='10PM'
        
    
    
    
    
    if operation.get()=='Escalation':
        if v.get()==2:
            alarm_columns=['Normal','Severity','Acknowledgment','Time','SITE_ID','RNC','alarm','ne type']
            file=pd.read_excel(downloaded,names=alarm_columns)
            file['SITE_ID']=file['SITE_ID'].apply(clean_3g)
            file=file[['SITE_ID']]
    
            esc_file=file.merge(RDB,on='SITE_ID',how='left')
            esc_file.rename(columns={'TECH STATE':'STATE'},inplace=True)
            esc_file.fillna('others')
            grouping=esc_file.groupby('STATE')['SITE_ID'].apply(list).reset_index(name='SITES')
            messagebox.askokcancel('ESCALATE?','Do u wish to continue')
            sender_id=browser.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/input").send_keys('3G DOWN')
            progress['maximum']=len(grouping.STATE)-1
            progress['value']=0
            for row in grouping.itertuples():
                sleep(0.001)
                progress['value']=progress['value']+1
                with open('test1.txt','w+') as filehandle:
                    filehandle.writelines('%s\n' % a for a in (['DOWN:'] + row.SITES))
                    filehandle.seek(0)
                    msg=filehandle.read()
                    e.insert(END,msg)
                progress.update()
                if len(msg)>=391:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                else:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
        elif v.get()==1:
            messagebox.askokcancel('ESCALATE?','Do u wish to continue')
            sender_id=browser.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/input").send_keys('NMC')
            comb_file=pd.read_excel('output_test.xlsx')
            
            for index,row in comb_file.iterrows():
                if 'E' in row.SITE_ID  and len(row.SITE_ID)>6 and row.SITE_ID.index('E')==0:
                    comb_file.at[index,'SITE_ID']=row.SITE_ID.lstrip('E')
                elif row.Alarm == 'CSL Fault':
                    comb_file.at[index,'SITE_ID']=row.SITE_ID.replace('D',' ',1).strip()
                elif row.Alarm in ['Rectifier Mains Power Failure',
                                   'Rectifier Urgent', 'Rectifier Mains Failure Alarm',
                                   'Rectifier Urgent Alarm'] and len(row.SITE_ID)> 6:
                    comb_file.at[index,'SITE_ID']=row.SITE_ID.replace('C',' ',1).strip()
        
            
            esc_page=pd.merge(comb_file,RDB,on='SITE_ID',how='left')
            esc_page['cond']=''

            for index,row in esc_page.iterrows():
                if row.Alarm=='CSL Fault':
                    esc_page.at[index,'SITE_ID']='D' + row.SITE_ID
    
            esc_page['Alarm']=esc_page.Alarm.apply(clean_alarm)
            esc_page.sort_values(by=['TECH STATE','Alarm'],inplace=True)
            esc_page=esc_page.drop_duplicates(subset=['SITE_ID','Alarm','TECH STATE'],keep='first')
            esc_page.fillna('others',inplace=True)
            grouping=esc_page.groupby('TECH STATE')
            state_list=esc_page['TECH STATE'].unique()
            progress['maximum']=len(state_list)
            progress['value']=0
            for state in state_list:
                ib=grouping.get_group(state)
                ib=ib.copy()
                ib.loc[:,'cond']=ib.Alarm.duplicated()
                ib['merger']='okay'
                msg_list=[]
                sleep(0.001)
                progress['value']=progress['value']+1
                for index,row in ib.iterrows():
                    if row['cond']==False:
                        row.merger=row.Alarm +':' + row.SITE_ID
                    elif row['cond']==True:
                        row.merger=row['SITE_ID']
                    msg_list.append(row.merger)
                with open('text1.txt','w+') as f:
                    f.writelines('%s\n' % a for a in msg_list)
                    f.seek(0)
                    msg=f.read()
                    e.insert(END,msg)
                progress.update()
                if len(msg)>=391:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                    if state=='others':
                        pass
                    else:
                        region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{state}.txt')
                        browser.find_element_by_name('btn1').click()
                        browser.back()
                        browser.find_element_by_xpath("//*[@id='msg']").clear()
                        message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                        region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{state}.txt')
                        browser.find_element_by_name('btn1').click()
                        browser.back()
                        browser.find_element_by_xpath("//*[@id='msg']").clear()
                else:
                    if state=='others':
                        pass
                    else:
                        message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                        region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{state}.txt')
                        browser.find_element_by_name('btn1').click()
                        browser.back()
                        browser.find_element_by_xpath("//*[@id='msg']").clear()
        messagebox.showinfo('TASK STATUS','Done')
        
    elif operation.get()=='Update Request':
        messagebox.askokcancel('REQUEST?','Do u wish to continue')
        sender_id=browser.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/input").send_keys('UPDATE')
        if v.get()==2:
            try:
                report=pd.read_excel(newest_file,header=2)
                report.dropna(subset=['SUBSYSTEM'],inplace=True)
            except PermissionError:
                messagebox.showinfo('','this report is opened')
                pass
            report=report[report['ROOT CAUSE ANALYSIS 3'].str.contains('POWER SUSPECTED')]
            report[['SITE_ID','STATE']]
            grouping=report.groupby('STATE')['SITE_ID'].apply(list).reset_index(name='SITES')
            progress['maximum']=len(grouping.STATE)-1
            for index,row in grouping.iterrows():
                sleep(0.001)
                progress['value']=index
                m=[greeting] + ['kindly send the updates of these 3g site(s) to 4940 '] + row.SITES + ['thanks']
                with open('test1.txt','w+') as filehandle:
                    filehandle.writelines('%s\n' % a for a in m)
                    filehandle.seek(0)
                    msg=filehandle.read()
                    e.insert(END,msg)
                progress.update()
                if len(msg)>=391:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                else:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
        elif v.get()==1:
            try:
                excel_files=[i for i in path2.glob('*.xlsm')]
                file=max(excel_files,key=os.path.getmtime) 
                report=pd.read_excel(file,header=2)
                report.dropna(subset=['ROOT CAUSE ANALYSIS 3'],inplace=True)
            except PermissionError:
                messagebox.showinfo('','this report is opened')
            report=report[report['ROOT CAUSE ANALYSIS 3'].str.contains('POWER SUSPECTED')]
            report[['SITE_ID','STATE']]
            grouping=report.groupby('STATE')['SITE_ID'].apply(list).reset_index(name='SITES')
            progress['maximum']=len(grouping.STATE)-1
            for index,row in grouping.iterrows():
                sleep(0.001)
                progress['value']=index
                m=[greeting] + ['kindly send the updates of these 2g site(s) to 09060986935 '] + row.SITES + ['thanks']
                with open('test1.txt','w+') as filehandle:
                    filehandle.writelines('%s\n' % a for a in m)
                    filehandle.seek(0)
                    msg=filehandle.read()
                    e.insert(END,msg)
                progress.update()
                if len(msg)>=391:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                else:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
        messagebox.showinfo('TASK STATUS','Done')
                
    elif operation.get()=='Send Report':
        messagebox.askokcancel('SEND REPORT?','Do u wish to continue')
        if v.get()==2:
            try:
                report=pd.read_excel(newest_file,header=2)
                report.dropna(subset=['ROOT CAUSE ANALYSIS 2'],inplace=True)
            except PermissionError:
                messagebox.showinfo('','this report is opened')
                pass
            df1=pd.read_excel('3G_Site_per_State.xlsx')
            df1_list=df1.STATE.tolist()
            df1_dict=dict(zip(df1.STATE,df1.SITE_COUNT))
            sender_id=browser.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/input").send_keys(f'3G DWN {t}')
            report=report[~report['ROOT CAUSE ANALYSIS 2'].str.contains('ROLLOUT') & ~report.NOP_SITES.str.contains('NOP') & pd.isnull(report['TO (Date & Time)'])]
            report[['SITE_ID','STATE']]
            grouping=report.groupby('STATE')['SITE_ID'].apply(list).reset_index(name='SITES')
            progress['maximum']=len(grouping.STATE)
            for index,row in grouping.iterrows():
                a= row.STATE + ' ' + str(len(row.SITES)) + '/' + str(df1_dict[row.STATE])
                m=[a] + row.SITES
                sleep(0.001)
                progress['value']=index
                with open('test1.txt','w+') as filehandle:
                    filehandle.writelines('%s\n' % a for a in m)
                    filehandle.seek(0)
                    msg=filehandle.read()
                    e.insert(END,msg)
                progress.update()
                if t=='5AM' or t=='10PM':
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                else:
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'Z:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG,EMC&SALES\\{row.STATE}.txt')
                if len(msg)>=391:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                    region_selection
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                    region_selection
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                else:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                    region_selection
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
            
    
                
        elif v.get()==1:
            try:
                excel_files=[i for i in path2.glob('*.xlsm')]
                file=max(excel_files,key=os.path.getmtime) 
                report=pd.read_excel(file,header=2)
                report.dropna(subset=['ROOT CAUSE ANALYSIS 2'],inplace=True)
            except PermissionError:
                messagebox.showinfo('','this report is opened')
                pass
            df1=pd.read_excel('2G_Site_per_State.xlsx')
            df1_list=df1.STATE.tolist()
            df1_dict=dict(zip(df1.STATE,df1.SITE_COUNT))
            sender_id=browser.find_element_by_xpath("/html/body/div[1]/div[2]/form/div[2]/input").send_keys(f'2G DWN {t}')
            report=report[~report['ROOT CAUSE ANALYSIS 2'].str.contains('ROLLOUT/NEW') & ~report.NOP_SITES.str.contains('NOP') & pd.isnull(report['TO (Date & Time)'])]
            report[['SITE_ID','STATE']]
            grouping=report.groupby('STATE')['SITE_ID'].apply(list).reset_index(name='SITES')
            progress['maximum']=len(grouping.STATE)
            for index,row in grouping.iterrows():
                a= row.STATE + ' ' + str(len(row.SITES)) + '/' + str(df1_dict[row.STATE])
                m=[a] + row.SITES
                sleep(0.001)
                progress['value']=index
                with open('test1.txt','w') as filehandle:
                    filehandle.writelines('%s\n' % a for a in m)
                with open('test1.txt','r') as v:
                    msg=v.read()
                    e.insert(END,msg)
                progress.update()
                if t=='5AM' or t=='10PM':
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG&EMC\\{row.STATE}.txt')
                else:
                    region_selection=browser.find_element_by_name('userfile').send_keys(f'R:\\noc\\BULK SMS NEWKINGZ\\NEW STRUCTURE REG,EMC&SALES\\{row.STATE}.txt')
                if len(msg)>=391:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[:392])
                    region_selection
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg[390:])
                    region_selection
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
                else:
                    message=browser.find_element_by_xpath("//*[@id='msg']").send_keys(msg)
                    region_selection
                    browser.find_element_by_name('btn1').click()
                    browser.back()
                    browser.find_element_by_xpath("//*[@id='msg']").clear()
        messagebox.showinfo('TASK STATUS','Done')
        
       
    