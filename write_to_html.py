class html_report:
    def  __init__(self,folder_path,tests):
        self.path=folder_path
        self.alltest=tests

        import dominate
        from dominate.tags import *
        import pandas as pd
        import matplotlib.pyplot as plt

#Suite1={'1':{'name':'t1','status':'passed','data':'data','path':'my_report.docx'},
        #'2':{'name':'t2','status':'passed','data':'data','path':'my_report.docx'}}
#Suite2={'3':{'name':'t3','status':'failed','data':'data','path':'my_report.docx'},
        #'4':{'name':'t4','status':'passed','data':'data','path':'my_report.docx'}}
#alltest={}
#alltest['Suite1']=Suite1
#alltest['Suite2']=Suite2
        table_headers = [ 'Testcase Name', 'Status','Test Data','Documentation']
        x1=[]
        w=0.4
        Passed=[]
        Failed=[]
        countn=0
        countp=0
        for s in self.alltest:
            for t in self.alltest[s]:
                if(self.alltest[s][t]['status']=='passed'):
                    countp=countp+1
                else:
                    countn=countn+1
            Passed.append(countp)
            Failed.append(countn)
            countn=0
            countp=0
        for i in range(len(self.alltest)):
            x1.append("Suite " +str(i+1))
        plt.bar(x1,Passed,w,label="Passed",color='green')
        plt.bar(x1,Failed,w,bottom=Passed,label="Failed",color='red')
        plt.xlabel("Test Suits")
        plt.ylabel("Number of test cases")
        plt.title("Test Result")
        plt.legend()
        plt.savefig(self.path +'\\'+'to.jpg',bbox_inches='tight')
        plt.show()
        doc = dominate.document(title='Report')
    
        with doc.head:
            link(rel='stylesheet', href='style.css')
    
        with doc:
            with div(cls='container'):
                h1('Test Report')
                img(src="to.jpg")
                with table(id='main', cls='table table-striped'):
                    caption(h3('Test Result'))
                    with thead(style="color: black;background-color: coral;"):
                        with tr():
                            for table_head in table_headers:
                                th(table_head)
                    with tbody():
                        for n in self.alltest:
                            for m in self.alltest[n]:
                                with tr():
                                    td(self.alltest[n][m]['name'])
                                    td(self.alltest[n][m]['status'])
                                    td(self.alltest[n][m]['data'])
                                    td(a("Testflow",href=self.alltest[n][m]['path']))
    
    
        with open(self.path +'\\'+'Final_Report.html', 'w') as file:
            file.write(doc.render())


