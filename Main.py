import tkinter
from tkinter import ttk
import FirstTestCase
import SecondTestCase
import ThirdTestCase
import FourthTestCase
import FifthTestCase
from PIL import Image,ImageTk
import dominate
from dominate.tags import *
import pandas as pd
import matplotlib.pyplot as plt

def text():
    chrome_path=txt1.get()
    folder_path=txt2.get()
    ter=FirstTestCase.test_one(chrome_path,folder_path)
    Re=ter.result()
    status=''
    if(Re[0]==True):
        status='passed'
    else:
        status='failed'
   
    ter2=SecondTestCase.test_two(chrome_path,folder_path)
    Re2=ter2.result()
    if(Re2[0]==True):
        status='passed'
    else:
        status='failed'
    
    ter3=ThirdTestCase.test_three(chrome_path,folder_path)
    Re3=ter3.result()
    if(Re3[0]==True):
        status='passed'
    else:
        status='failed'
    
    ter4=FourthTestCase.test_four(chrome_path,folder_path)
    Re4=ter4.result()
    if(Re4[0]==True):
        status='passed'
    else:
        status='failed'
    Suite1={'1':{'name':'Testcase_1','status':status,'data':'data','path':Re[1]},'2':{'name':'Testcase_2','status':status,'data':'data','path':Re2[1]},'3':{'name':'Testcase_3','status':status,'data':'data','path':Re3[1]}}
    
    ter5=FifthTestCase.test_five(chrome_path,folder_path)
    Re5=ter5.result()
    if(Re5[0]==True):
        status='passed'
    else:
        status='failed'
    #t5={'5':{'name':'Testcase_1','status':status,'data':'data','path':Re5[1]}}
    Suite2={'4':{'name':'Testcase_4','status':status,'data':'data','path':Re4[1]},'5':{'name':'Testcase_5','status':status,'data':'data','path':Re5[1]}}
    alltest={'suite1':Suite1,'Suite2':Suite2}
    table_headers = [ 'Testcase Name', 'Status','Test Data','Documentation']
    x1=[]
    w=0.4
    Passed=[]
    Failed=[]
    countn=0
    countp=0
    for s in alltest:
        for t in alltest[s]:
            if(alltest[s][t]['status']=='passed'):
                countp=countp+1
            else:
                countn=countn+1
        Passed.append(countp)
        Failed.append(countn)
        countn=0
        countp=0
    for i in range(len(alltest)):
        x1.append("Suite " +str(i+1))
    
    plt.bar(x1,Passed,w,label="Passed",color='green')
    plt.bar(x1,Failed,w,bottom=Passed,label="Failed",color='red')
    plt.xlabel("Test Suits")
    plt.ylabel("Number of test cases")
    plt.title("Test Result")
    plt.legend()
    plt.savefig(folder_path+'//'+'chart.jpg',bbox_inches='tight')
    plt.show()
    doc = dominate.document(title='Report')
    with doc.head:
        link(rel='stylesheet', href='style.css')

    with doc:
        with div(cls='container'):
            h1('Test Report')
            dominate.tags.img(src=folder_path+'//'+'chart.jpg')
            with table(id='main', cls='table table-striped'):
                caption(h3('Test Result'))
                with thead(style="color: black;background-color: coral;"):
                    with tr():
                        for table_head in table_headers:
                            th(table_head)
                with tbody():
                    for n in alltest:
                        for m in alltest[n]:
                            with tr():
                                td(alltest[n][m]['name'])
                                td(alltest[n][m]['status'])
                                td(alltest[n][m]['data'])
                                td(a("Testflow",href=alltest[n][m]['path']))


    with open(folder_path+'//'+'Final_Report.html', 'w') as file:
        file.write(doc.render())
    
               
frm=tkinter.Tk()
frm.title("Test Framework")
w=700
h=500
sw=frm.winfo_screenwidth()
sh=frm.winfo_screenheight()
x=(sw-w)/2
y=(sh-h)/2
frm.geometry('%dx%d+%d+%d' % (w,h,x,y))
frm.config(background='#424242')
load=Image.open('pic1.jpg')
load=load.resize((700,500),Image.ANTIALIAS)
render=ImageTk.PhotoImage(load)
img=ttk.Label(frm,image=render)
img.place(x=0,y=0)
lbl1=ttk.Label(frm,text='Automation Framework')
lbl1.grid(row=2,column=2)
lbl1.config(background='#B0BEC5',font=('impact',20))
lbl2=ttk.Label(frm,text='ChromePath')
lbl2.grid(row=4,column=0)
txt1=ttk.Entry(frm)
txt1.grid(row=4,column=1)
lbl3=ttk.Label(frm,text='Folderpath')
lbl3.grid(row=5,column=0)
txt2=ttk.Entry(frm)
txt2.grid(row=5,column=1)
btn=ttk.Button(frm,text='Start',command=text)
btn.place(x=200,y=200)
frm.mainloop()