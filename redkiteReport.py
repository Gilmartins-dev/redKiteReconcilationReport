from os import times
import pandas as pd
import requests
import json
import hashlib
import time

#FUNCTION TO REQUEST AND RETURN TOKEN
def getToken():
    url = "https://login.microsoftonline.com/5b370688-b179-45c6-8271-628b64c03723/oauth2/token"
    payload='client_id=d5d35f45-e502-4ba6-90a9-8a0e72a5f683&client_secret=7~6z9tW-1nW8f.N~Okes12l4vK4a61U3C5&grant_type=client_credentials&resource=https%3A%2F%2Fgilmartins.crm11.dynamics.com'
    headers = {
  'Content-Type': 'application/x-www-form-urlencoded',
  'Cookie': 'buid=0.ASAAiAY3W3mxxkWCcWKLZMA3IwcAAAAAAAAAwAAAAAAAAAAgAAA.AQABAAEAAAD--DLA3VO7QrddgJg7Wevr4jffD5W1yOlpEWY5x9ycJYIRfeYKw6YQZn-GfDQihIvD-GMAMNyFlZ6VJDydLdwhwUBUFK6UHwvlFTq5J8JfjpTNVnCJOrxn7Dqd6ry4D0QgAA; fpc=Aq6MgxF1xdpCmpYjXELNUo1mM4adAQAAAGyIUtkOAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd'
    }
    response = requests.request("POST", url, headers=headers, data=payload)
    responseJson = response.json()

    resToken = responseJson['access_token']

    return resToken

#CALL FUNCTION TO RETURN TOKEN AND STORE IN VARIABLE TOKEN
token = getToken()

class Workorder:
    def __init__(self,jobnumber,rk_value,gilm_value,statusOfWork):
        self.jobnumber = jobnumber
        self.rk_value = rk_value
        self.gilm_value = gilm_value
        self.statusOfWork = statusOfWork
#USE AN XLSX EXCEL FILE NAMED DAILYREPORT.XLSX AND CREATE A SHEET, SETNAME TO VARIABLE BELOW           
sheetName = "23122021"          
# READ FILE            .
pd.set_option('precision', 0)   
df = pd.read_excel('dailyreport.xlsx', sheet_name=sheetName)

# READ THE JOBNUMBERS FROM THE FILE
listOfJobnumbers = df['Job Number']

#CREATE EMPTY LISTS
listOfJobs = []
failedList = []

#FUNCTION TO GET GILMARTINS TOTAL JOB VALUE USING THE CLIENT REF SUPPLIED BY REDKITE 
def getGilmartinsValue(clientRef):
    headers = {
                'Authorization': "Bearer " + token,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json',
                'Content-Type': 'application/json; charset=utf-8',
                'Prefer': 'return=representation'
            }
   
    
    url ="https://gilmartins.crm11.dynamics.com/api/data/v9.2/msdyn_workorders?$select=msdyn_estimatesubtotalamount,gilm_statusofwork&$filter=(gilm_clientref eq '" + str(clientRef) + "')"
    #print(url)
    a = requests.get(url, headers = headers)
    resJson = a.json()
       
    return resJson["value"]
    
#ITERATE THROUGH THE READ FILE WHICH WAS CREATED ON LINE 36 CHECK EACH ROW IS PRINTING TO CONSOLE
for index, row in df.iterrows():
    print("CLIENTREF ", row['Job Number'])
    print("VALUE ", row['Total Value'])
    
    response = getGilmartinsValue(str(row['Job Number']))
    #print(response)
    if response != []:
        gilm_estimated = response[0]["msdyn_estimatesubtotalamount"]
        gilm_status = response[0]["gilm_statusofwork"]
        #A WORKORDER CLASS IS CREATED AND ALL ',' AND '£' ARE REMOVED SO ONLY DIGITS REMAIN OTHERWISE CHANGING STRING TO INTEGER WILL FAIL LATER
        workorder = Workorder(str(row['Job Number']), row['Total Value'].replace('£', '').replace(',',""),gilm_estimated,gilm_status)
        listOfJobs.append(workorder)

#CREATE TIME STAMP AND STORE IN VARIABLE
timestr = time.strftime("%d%m%y-%H%M")
#CREATE JOBS TITLES VARIABLE WITH WILL HOLD THE SHEET NAME, '_jobs_lOG_, TIME STRING AND TXT SUFFIX
#THREE LOGS ARE CREATED A FULL LOG OF ALL JOB NUMBERS WITH A 'FAIL' APPENDED TO EACH LINE THAT FAILS, A LIST OF FAILURES AND A LIST OF FAILURES WIITH DUPLICATES REMOVED
jobs_log = sheetName + '_jobs_log_' + timestr +  '.txt'
failed_list_withduplicates = sheetName + '_failed_list_dup_' + timestr + '.txt'
failed_log = sheetName + '_failed_log_' + timestr + '.txt'

#A COUNTER VARIABLE WHICH IS ADDED TO EACH TIME A FAILURE IS IDENTIFIED
counter = 0

#ITERATES THROUGH EACH LINE IN THE listOfJobs APPENDING FAIL TO EACH LINE WHICH FAILS THE CRITERIA AND CREATES A TEXT FILE
with open(jobs_log, 'a') as j:
    for job in listOfJobs:
        print("Gilm number" + str(job.gilm_value))
        if float(job.rk_value) < float(job.gilm_value) or job.statusOfWork != 870110000:
            failedList.append(job)
            counter += 1
            j.write("Job no" + " : " + str(job.jobnumber) + ' : ' + "\tRK : " + str(job.rk_value) + ' : ' +  "\tGil" + ' : ' + str(round(job.gilm_value, 2))  + ' : ' + "\tStatus : " +  str(job.statusOfWork) +  " : FAILED")
            j.write('\n')
        else:
            j.write("Job no" + " : " + str(job.jobnumber) + ' : ' + "\tRK : " + str(job.rk_value) + ' : ' +  "\tGil" + ' : ' + str(round(job.gilm_value, 2))  + ' : ' + "\tStatus : " +  str(job.statusOfWork))
            j.write('\n')
    #CREATE A LOG OF ALL FAILED JOBS WHICH JUST CONTAINS THE JOB NUMBER
    with open(failed_list_withduplicates, 'a') as f:
            f.write('Hi Tom,')
            f.write('\n')
            f.write('These jobs failed on ' + timestr[0:2] + '/' + timestr[2:4] + '/' + timestr[4:6])
            f.write('\n')
            f.write('-------------------------')
            f.write('\n')
            for failed in failedList:                
                f.write(str(failed.jobnumber))
                f.write('\n')
            

#THIS CODE REMOVES DUPLICATES FROM failed_list_withduplicates AND CREATES A NEW LIST NAMED failed_log
output_file_path = failed_log
input_file_path = failed_list_withduplicates

completed_lines_hash = set()

output_file = open(output_file_path, "w")

for line in open(input_file_path, "r"):
    hashValue = hashlib.md5(line.rstrip().encode("utf-8")).hexdigest()
    if hashValue not in completed_lines_hash:
        output_file.write(line)
        completed_lines_hash.add(hashValue)
output_file.close()

num_lines = sum(1 for line in open(failed_log))

with open(failed_log, 'r+') as j:
    j.write("There are  " + str(num_lines - 3) + " failures in 1000")
    j.write('\n')