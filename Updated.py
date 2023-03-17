from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import xlsxwriter
import os
import PySimpleGUI as dialogBox
from os import environ as env
from azure.identity import ClientSecretCredential
from azure.keyvault.secrets import SecretClient
from wakepy import keepawake

def notFoundException():
    chromeDriver.quit()
    webDriver()
# Getting Azure Secret
def keyVaultData():
    global KV_USERNAME
    global KV_PASSWORD
    KV_USERNAME = env.get("NEWRELIC_USERNAME","")
    KV_PASSWORD = env.get("NEWRELIC_PASSWORD","")

def webDriver():
    with keepawake(keep_screen_awake=True):
        keyVaultData()
        webDriverErrorResolver = webdriver.ChromeOptions()
        webDriverErrorResolver.add_argument('--ignore-certificate-errors-spki-list')
        webDriverErrorResolver.add_argument('--ignore-ssl-errors')
        webDriverErrorResolver.add_argument('log-level=3')
        global chromeDriver
        chromeDriver = webdriver.Chrome(options=webDriverErrorResolver,service=Service  (ChromeDriverManager().install()))
        chromeDriver.maximize_window()
        chromeDriver.get("https://login.newrelic.com/login")
        chromeDriver.find_element(By.ID,"login_email").send_keys(KV_USERNAME)
        chromeDriver.find_element(By.NAME,"button").click()
        chromeDriver.find_element(By.ID,"login_password").send_keys(KV_PASSWORD)
        chromeDriver.find_element(By.NAME,"button").click()
        try:
            if chromeDriver.find_element(By.ID,"end_sessions").is_displayed():
                chromeDriver.find_element(By.ID,"end_sessions").click()
                chromeDriver.find_element(By.ID,"login_submit").click()
                scrapingProcess()
        except:
            scrapingProcess()

def scrapingProcess():
    time.sleep(10)
    global excelWorkbook
    currentDateAndTime = time.strftime("D-%d-%m-%Y-T-%H-%M-%S")
    excelWorkbook = xlsxwriter.Workbook(''+currentDateAndTime+'.xlsx')
    excelWorksheet = excelWorkbook.add_worksheet()
    currentTableRow = 0
    apdexColumnCount = 0
    indexOfCurrentSite = 2
    scrollCountInEntities = 0  # main page site count
    isNotLive = 0
    siteVisibleCheck = chromeDriver.find_elements(By.XPATH,"//div[contains(@class,'-wnd-CardBaseHeader-heading')]/a")
    if(len(siteVisibleCheck) == 0):
        time.sleep(20)
        scrapingProcess()
    totalSiteCount =  chromeDriver.find_element(By.XPATH,"//div[contains(@class,'-wnd-CardBaseHeader-heading')]/a").text
    totalSiteCount = totalSiteCount[totalSiteCount.find("(")+1:totalSiteCount.find(")")] 
    chromeDriver.find_element(By.XPATH,"//div[contains(@class,'-wnd-CardBaseHeader-heading')]/a").click()#view all click
    # below loop is used to select a specific site 
    for index in range(int(totalSiteCount)):
        time.sleep(5)
        totalColumnHeaderForApdexCount = []
        entityPageUrl = chromeDriver.current_url
        print(index,'index')
        isLive = False
        if(index != 0 and index % 9 == 0):
            scrollCountInEntities += 1
            indexOfCurrentSite = 4
            time.sleep(5)
        isScrollBarInEntities = chromeDriver.find_elements(By.XPATH,"//div[contains(@class,'-wnd-DataTable')]/div")
        if(len(isScrollBarInEntities)==0):
            time.sleep(15)
        scrollBarInEntities = chromeDriver.find_element(By.XPATH,"//div[contains(@class,'-wnd-DataTable')]/div")       
        # below loop is used to scroll the page after scraping the data from set of sites
        for i in range(scrollCountInEntities):
            chromeDriver.execute_script('arguments[0].scrollBy(0,360);', scrollBarInEntities)
            time.sleep(2)
        pathOfCurrentSite = chromeDriver.find_element(By.XPATH,'//div['+str(indexOfCurrentSite)+']/span[contains(@class,"-wnd-DataTableBaseRowCell--ellipsisType-end")]')
        indexOfCurrentSite += 1
        isLiveSite = pathOfCurrentSite.text[pathOfCurrentSite.text.find("(")+1:pathOfCurrentSite.text.find(")")]
        if(isLiveSite.lower() == 'live'):
            isLive = True
            isNotLive = 1
        elif(isNotLive == 3):
            break
        else:
            chromeDriver.refresh()
            isNotLive += 1
            continue
        if(isLive):
            pathOfCurrentSite.click()   # clicking the current site
            time.sleep(12)
            isScrollBarInSelectedSite = chromeDriver.find_elements(By.XPATH,"//section[contains(@class,'u-chart-columns-primary')]")
            if(len(isScrollBarInSelectedSite)==0):
                time.sleep(30)
            scrollBarInSelectedSite = chromeDriver.find_element(By.XPATH, "//section[contains(@class,'u-chart-columns-primary')]")
            chromeDriver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollBarInSelectedSite)  
            pathOfCurrentSiteName = chromeDriver.find_element(By.XPATH,"//div[contains(@class,'-wnd-EntitySwitcher-entityName')]/span")
            currentTableRow += 1
            siteNameForExcel = pathOfCurrentSiteName.text.find("(")
            siteNameForExcel = pathOfCurrentSiteName.text[0:siteNameForExcel-1]
            print(siteNameForExcel,'name',pathOfCurrentSiteName.text)
            excelWorksheet.write(currentTableRow, 0 , siteNameForExcel)
            time.sleep(5)
            chromeDriver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollBarInSelectedSite)
            numberOfColumnsInSelectedSite = chromeDriver.find_elements(By.XPATH, "//div/span[contains(@class,'-wnd-TableHeaderCell-title')]")   
            if(len(numberOfColumnsInSelectedSite) == 0):
                time.sleep(20)
                chromeDriver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollBarInSelectedSite)

            for i in range(1,len(numberOfColumnsInSelectedSite),1):
                columnHeaderInSelectedSite = chromeDriver.find_element(By.XPATH, '//div['+str(i)+']/div/span[contains(@class,"-wnd-TableHeaderCell-title")]')
                totalColumnHeaderForApdexCount.append(columnHeaderInSelectedSite.text)
                excelWorksheet.write(0,i,columnHeaderInSelectedSite.text)
                apdexColumnCount = len(totalColumnHeaderForApdexCount) + 1
            time.sleep(5)
            isMonitoringData = chromeDriver.find_elements(By.XPATH, "//div[contains(@class,'-wnd-TableCell-content')]")   
            # apdexCheck = chromeDriver.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[1]/div[1]/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/h2')
            apdexCheck = chromeDriver.find_elements(By.XPATH,"//h2[text()='Apdex score']")
            if len(apdexCheck) > 0:
                # apdexName = chromeDriver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[1]/div[1]/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/h2').text
                apdexName = chromeDriver.find_element(By.XPATH,"//h2[text()='Apdex score']").text
                excelWorksheet.write(0,apdexColumnCount,apdexName)
                # apdexData = chromeDriver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[1]/div[1]/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[2]/div/span[1]').text 
                apdexData = chromeDriver.find_element(By.XPATH,"//div[2]/div/div/div[1]/div/div[1]/div[2]/div/span[contains(@class,'--vz--widget-chart-summary__value')]").text 
            if len(isMonitoringData) > 0:
                currentTableColumn = 0
                print(len(isMonitoringData),'mm')
                for i in range (1,len(isMonitoringData),1):
                    if(i % 7 == 0):
                        currentTableRow += 1
                        currentTableColumn = 0
                        continue
                    else:     
                        currentTableColumn += 1   
                        monitoringDataOfSelectedSite = chromeDriver.find_element(By.XPATH, '//div['+str(i)+']/div[contains(@class,"-wnd-TableCell-content")]')
                        print(monitoringDataOfSelectedSite.text,'val')
                        excelWorksheet.write(currentTableRow,currentTableColumn,monitoringDataOfSelectedSite.text)
                        excelWorksheet.write(currentTableRow, apdexColumnCount, apdexData)
            chromeDriver.get(entityPageUrl)
            time.sleep(5)  
    with keepawake(keep_screen_awake = False):   
        excelWorkbook.close()
        chromeDriver.find_element(By.XPATH,"//div[contains(@class,'-wnd-UserMenu')]/button").click() 
        chromeDriver.find_element(By.XPATH,"//div[contains(@class,'-wnd-PopoverListItem-content')]").click()                         
        chromeDriver.quit()
        os.system('start "excel.exe" "'+currentDateAndTime+'.xlsx"')
    

def beginScraping():
    try:
        webDriver()
    except:
        excelWorkbook.close()
        notFoundException()
    finally:
        return True

beginScraping()