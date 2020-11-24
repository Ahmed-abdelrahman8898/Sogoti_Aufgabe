class test_four:
    def  __init__(self,chrome_path,folder_path):
        self.chrome=chrome_path
        self.path=folder_path

        import json
        import requests
        import pandas as pd
        import sys
        from datetime import date
        try:
            self.out_passed=True
            Testdata=pd.read_excel(self.path +'\\'+'4_testcase.xlsx')
            result=[]
            today = date.today()
            d4 = today.strftime("%b-%d-%Y")
            d3=today.strftime("%b_%d_%Y")
            respond =requests.get('http://api.zippopotam.us/de/bw/stuttgart')
            if(respond.status_code!=200):
                print("wrong")
                result.append('Failed: '+str(respond.status_code))
            else:
                result.append('Passed: '+str(respond.status_code))
            if(respond.elapsed.total_seconds()<1):
                result.append('Passed: '+str(respond.elapsed.total_seconds()))
            else:
                result.append('Failed: '++str(respond.elapsed.total_seconds()))
            respond_json=respond.json()
            if(respond_json['country']=='Germany'):
                result.append('Passed: '+respond_json['country'])
            else:
                result.append('Failed: '+respond_json['country'])
            if(respond_json['state']=='Baden-Württemberg'):
                result.append('Passed: '+respond_json['state'])
                
            else:
                result.append('Failed: '+respond_json['state'])
                
            
            
            i=0
            while (i <len(respond_json['places'])):
                placename=respond_json['places'][i]['post code']
                if (placename=='70597' and respond_json['places'][i]['place name']=='Stuttgart Degerloch'):
                    
                    result.append('Passed: '+respond_json['places'][i]['place name'])
                
                i=i+1
        except:
            result.append(sys.exc_info())
            if (len(result)!=len(Testdata)):
                while (len(result)<len(Testdata)):
                    result.append('N/A')
            self.out_passed=False
        finally:
            Testdata['Result']=result
            Testdata=pd.DataFrame(Testdata)
            self.report_path=self.path +'\\'+'4_testcase'+d3+'.xlsx'
            self.report=[self.out_passed,self.report_path]
            writer = pd.ExcelWriter(self.report_path, engine='xlsxwriter')
            Testdata.to_excel(writer, sheet_name='Sheet1')
            writer.save()
    def result (self):
        
        return self.report
