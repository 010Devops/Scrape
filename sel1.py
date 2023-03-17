from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import xlsxwriter
import os
projectTitle = []
projectValue = []

TitleText = ''
TitleList = []
TitleContent = []

jobDetails=[]
jobTitle = []
jobContent = []
jobList = []
jobListData = []

apdexval = []
row=1
column=0
Trash = 0
count = 0

username = "nick-cit@outlook.com"
password = "WT@dmin24"

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors-spki-list')
options.add_argument('--ignore-ssl-errors')
options.add_argument('log-level=3')
driver = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))

driver.maximize_window()
driver.get("https://login.newrelic.com/login")
driver.find_element(By.ID,"login_email").send_keys(username)
driver.find_element(By.NAME,"button").click()
driver.find_element(By.ID,"login_password").send_keys(password)
driver.find_element(By.NAME,"button").click()

try:
    if driver.find_element(By.ID,"end_sessions").is_displayed():
        driver.find_element(By.ID,"end_sessions").click()
        driver.find_element(By.ID,"login_submit").click()
  
except:
    print("except")

time.sleep(10)

projectTitle = driver.find_elements(By.CLASS_NAME,"nr1-EntityTitleTableRowCell")
test = False
for pro in range(len(projectTitle)):

    time.sleep(5)

    projectValue = driver.find_elements(By.CLASS_NAME,"nr1-EntityTitleTableRowCell")
    
    cnt = len(projectValue) / 2
    cnt = round(cnt)
    if(pro == cnt):
        break
    else:
        projectValue[pro].click()

        time.sleep(10)
        scrollValue = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/section/div[4]/div/section/div/div/div/div[1]/div/section')
        driver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollValue)
   
        Title = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/section/div[1]/div[1]/div[1]/div[2]/div/div/div[1]/span')
        print(Title.text, "print")
        TitleList.append(Title.text)

        time.sleep(20)
        driver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollValue)
        if(not(bool(pro))):
            listval = driver.find_elements(By.XPATH, '//*[@id="tabpanel-hosts-table-id-'+str(pro)+'"]/div/div/div[4]/div/div')
     
        for i in range(1,len(listval),1):
            ind = str(i)
            if(not(bool(pro))):
                jobLis = driver.find_element(By.XPATH, '//*[@id="tabpanel-hosts-table-id-0"]/div/div/div[4]/div/div['+ind+']/div/span[1]')
                jobList.append(jobLis.text)

            check = driver.find_elements(By.XPATH, '//*/div/div/div[5]/div/div/div[1]/div')
            
            if len(check) > 0:
                jobDat = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/section/div[4]/div/section/div/div/div[1]/div[1]/div/section/div[4]/div/div/div[2]/div[2]/div/div/div[5]/div/div/div['+ind+']/div')
                jobListData.append(jobDat.text)
               
        # D A T A  S C R A P I N G

        count = len(jobList)
        # time.sleep(15)

        workbook = xlsxwriter.Workbook('Data4.xlsx')
        worksheet = workbook.add_worksheet()
        print(jobList)
        for index,val in enumerate(jobList):
            worksheet.write(0, index+1, val)
               
        for index, job in enumerate(TitleList, start = 1):
            worksheet.write(index,column, job)

        col = 0
        for index, job in enumerate(jobListData, start=1):
            col += 1
            if(col > count):
                row +=1
                col = 1
                worksheet.write(row, col, job)
            else:
                worksheet.write(row, col, job)

        check1 = driver.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/section/div[4]/div/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div/div[1]/div[1]/h2')
        if len(check1) > 0:
            apdexName = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/section/div[4]/div/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div/div[1]/div[1]/h2').text
            worksheet.write(0,count+1,apdexName)

            apdexValue = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/section/div[4]/div/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div[1]/div/div[1]/div[2]/div/span[1]').text
            apdexval.append(apdexValue)

            for index, job in enumerate(apdexval, start=1):
                worksheet.write(index, count+1, job)

        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/section/div[1]/div[1]/div[1]/div[1]/a').click()
        time.sleep(5)
        workbook.close()

time.sleep(10)

driver.quit()

os.system('start "excel.exe" "Data4.xlsx"')