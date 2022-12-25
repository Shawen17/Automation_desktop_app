#ESCALATE,SEND REPORT AND SEND REQUEST
from declare_functions import pastetext,merge_log,run_report,refresh,perform

window=Tk()
window.title('Escalation')
window.geometry('600x600')
window.maxsize(700,700)
canvas=Canvas(window,background='#4B8BBE')
canvas.place(relx=0.05,rely=0.05,relheight=0.9,relwidth=0.9)
    
options=('Escalation','Send Report','Update Request','Declare Up')
comb=Combobox(canvas,values=options)
comb.current(0)
comb.place(relx=0.4,rely=0.1,relheight=0.06,relwidth=0.4)    
    

v1=IntVar()
def open_toplevel1():
    top=Toplevel(bg='#3EB049')
    top.title('SITES_STATUS')
    top.geometry('400x400')
    top.maxsize(500,500)
    label1=Label(top,text='log1:')
    label1.place(relx=0.05,rely=0.1,relheight=0.05,relwidth=0.1)
    label2=Label(top,text='log2:')
    label2.place(relx=0.05,rely=0.2,relheight=0.05,relwidth=0.1)
    label3=Label(top,text='log3:')
    label3.place(relx=0.05,rely=0.3,relheight=0.05,relwidth=0.1)
    label4=Label(top,text='log4:')
    label4.place(relx=0.05,rely=0.4,relheight=0.05,relwidth=0.1)

    e1=Entry(top)
    e1.place(relx=0.15,rely=0.1,relheight=0.05,relwidth=0.55)
    e2=Entry(top)
    e2.place(relx=0.15,rely=0.2,relheight=0.05,relwidth=0.55)
    e3=Entry(top)
    e3.place(relx=0.15,rely=0.3,relheight=0.05,relwidth=0.55)
    e4=Entry(top)
    e4.place(relx=0.15,rely=0.4,relheight=0.05,relwidth=0.55)


    button1=Button(top,text='Open',command=lambda: pastetext(e1))
    button1.place(relx=0.7,rely=0.1,relheight=0.05,relwidth=0.2)
    button2=Button(top,text='Open',command=lambda: pastetext(e2))
    button2.place(relx=0.7,rely=0.2,relheight=0.05,relwidth=0.2)    
    button3=Button(top,text='Open',command=lambda: pastetext(e3))
    button3.place(relx=0.7,rely=0.3,relheight=0.05,relwidth=0.2)
    button4=Button(top,text='Open',command=lambda: pastetext(e4))
    button4.place(relx=0.7,rely=0.4,relheight=0.05,relwidth=0.2)
    logs=[e1,e2,e3,e4]    
    merged_files='output_test.xlsx'
    button5=Button(top,text='Run',command=lambda : merge_log(logs,v1,comb))
    button5.place(relx=0.3,rely=0.6,relheight=0.05,relwidth=0.1) 
    button6=Button(top,text='Declare Up',command=lambda: run_report(merged_files,v1))
    button6.place(relx=0.45,rely=0.6,relheight=0.05,relwidth=0.2)
    top.mainloop()
    
e1=ScrolledText(canvas)
e1.place(relx=0.4,rely=0.3,relheight=0.5,relwidth=0.4)

def open_toplevel2():
    top1=Toplevel(bg='#3EB049')
    top1.title('2G_ESCALATION')
    top1.geometry('400x400')
    top1.maxsize(500,500)
    label1=Label(top1,text='alarm1:')
    label1.place(relx=0.05,rely=0.1,relheight=0.05,relwidth=0.1)
    label2=Label(top1,text='alarm2:')
    label2.place(relx=0.05,rely=0.2,relheight=0.05,relwidth=0.1)
    label3=Label(top1,text='alarm3:')
    label3.place(relx=0.05,rely=0.3,relheight=0.05,relwidth=0.1)
    label4=Label(top1,text='alarm4:')
    label4.place(relx=0.05,rely=0.4,relheight=0.05,relwidth=0.1)
    label5=Label(top1,text='alarm5:')
    label5.place(relx=0.05,rely=0.5,relheight=0.05,relwidth=0.1)
    label6=Label(top1,text='alarm6:')
    label6.place(relx=0.05,rely=0.6,relheight=0.05,relwidth=0.1)
    
    a1=Entry(top1)
    a1.place(relx=0.15,rely=0.1,relheight=0.05,relwidth=0.55)
    e2=Entry(top1)
    e2.place(relx=0.15,rely=0.2,relheight=0.05,relwidth=0.55)
    e3=Entry(top1)
    e3.place(relx=0.15,rely=0.3,relheight=0.05,relwidth=0.55)
    e4=Entry(top1)
    e4.place(relx=0.15,rely=0.4,relheight=0.05,relwidth=0.55)
    e5=Entry(top1)
    e5.place(relx=0.15,rely=0.5,relheight=0.05,relwidth=0.55)
    e6=Entry(top1)
    e6.place(relx=0.15,rely=0.6,relheight=0.05,relwidth=0.55)


    button1=Button(top1,text='Open',command=lambda: pastetext(a1))
    button1.place(relx=0.7,rely=0.1,relheight=0.05,relwidth=0.2)
    button2=Button(top1,text='Open',command=lambda: pastetext(e2))
    button2.place(relx=0.7,rely=0.2,relheight=0.05,relwidth=0.2)    
    button3=Button(top1,text='Open',command=lambda: pastetext(e3))
    button3.place(relx=0.7,rely=0.3,relheight=0.05,relwidth=0.2)
    button4=Button(top1,text='Open',command=lambda: pastetext(e4))
    button4.place(relx=0.7,rely=0.4,relheight=0.05,relwidth=0.2)
    button5=Button(top1,text='Open',command=lambda: pastetext(e5))
    button5.place(relx=0.7,rely=0.5,relheight=0.05,relwidth=0.2)
    button6=Button(top1,text='Open',command=lambda: pastetext(e6))
    button6.place(relx=0.7,rely=0.6,relheight=0.05,relwidth=0.2)
    
    logs=[a1,e2,e3,e4,e5,e6]    
    merged_files='output_test.xlsx'
    button7=Button(top1,text='Run',command=lambda : merge_log(logs,v1,comb))
    button7.place(relx=0.3,rely=0.7,relheight=0.05,relwidth=0.1) 
    button8=Button(top1,text='SEND',command=lambda : perform(export,v1,v2,web,e1,comb,latest_file,rd_base,bar))
    button8.place(relx=0.45,rely=0.7,relheight=0.05,relwidth=0.2)
    top1.mainloop()




l1=Label(canvas,text='PREFFERED SMS',background='#FFD43B')
v2=IntVar()
rad3=Radiobutton(canvas,text='22',value=3,variable=v2)
rad4=Radiobutton(canvas,text='222',value=4,variable=v2)
rad5=Radiobutton(canvas,text='221',value=5,variable=v2)
l1.place(relx=0.07,rely=0.2,relheight=0.05,relwidth=0.2)
rad3.place(relx=0.07,rely=0.25)
rad4.place(relx=0.07,rely=0.3)
rad5.place(relx=0.07,rely=0.35)





rad1=Radiobutton(canvas,text='2G',value=1,variable=v1)
rad2=Radiobutton(canvas,text='3G',value=2,variable=v1)
rad1.place(relx=0.07,rely=0.1)
rad2.place(relx=0.07,rely=0.15)



option= webdriver.ChromeOptions()
option.add_argument('headless')
web = webdriver.Chrome(executable_path='C:\\Program Files\\SeleniumBasic\\chromedriver.exe', options=option)

export='O:\\SEUN\\1\\ABUJA DOWN.xlsx'
excel_files=[i for i in path1.glob('*.xlsm')]
latest_file=max(excel_files,key=os.path.getmtime) 

style=Style()
style.configure('black.Horizontal.TProgressbar',background='#BFEFFF')
bar=Progressbar(window,length=100,style='black.Horizontal.TProgressbar',mode='determinate')
bar.place(relx=0.2,rely=0.85,relheight=0.05,relwidth=0.3)

def open_declare():
    if (v1.get()==2 or v1.get()==1) and comb.get()=='Declare Up':
        open_toplevel1()
    elif v1.get()==1 and comb.get()=='Escalation':
        open_toplevel2()
    else:
        perform(export,v1,v2,web,e1,comb,latest_file,rd_base,bar)
    

button=Button(canvas,text='EXECUTE',command=open_declare)
button.place(relx=0.7,rely=0.85,relheight=0.05,relwidth=0.12)
button1=Button(canvas,text='REFRESH',command=lambda : refresh(e1,bar,comb))
button1.place(relx=0.82,rely=0.85,relheight=0.05,relwidth=0.12)

def on_closing():
    if messagebox.askokcancel('Quit','Do you want to quit'):
        web.quit()
        window.destroy()


window.protocol('WM_DELETE_WINDOW',on_closing)
window.mainloop()