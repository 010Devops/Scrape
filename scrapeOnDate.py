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
    print('2')
    global KV_USERNAME
    global KV_PASSWORD
    # KV_USERNAME = env.get("NEWRELIC_USERNAME","")
    # KV_PASSWORD = env.get("NEWRELIC_PASSWORD","")
    KV_USERNAME = "girish.p@congruentindia.com"
    KV_PASSWORD = "Congruent@123"


def webDriver():
    print('1')
    with keepawake(keep_screen_awake=True):
        keyVaultData()
        print('3')
        webDriverErrorResolver = webdriver.ChromeOptions()
        webDriverErrorResolver.add_argument('--ignore-certificate-errors-spki-list')
        webDriverErrorResolver.add_argument('--ignore-ssl-errors')
        webDriverErrorResolver.add_argument('log-level=3')
        global chromeDriver
        print('chrome')
        # chromeDriver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=webDriverErrorResolver)
        # chromeDriver = webdriver.Chrome(ChromeDriverManager().install())
        chromeDriver = webdriver.Chrome()
        # chromeDriver = webdriver.Firefox()
        #chromeDriver = webdriver.Edge()
        print('chromedriver')
        chromeDriver.maximize_window()
        chromeDriver.get("https://login.newrelic.com/login")
        print('4')
        chromeDriver.find_element(By.ID, "login_email").send_keys(KV_USERNAME)
        chromeDriver.find_element(By.NAME, "button").click()
        chromeDriver.find_element(By.ID, "login_password").send_keys(KV_PASSWORD)
        chromeDriver.find_element(By.NAME, "button").click()
        try:
            print('try')
            if chromeDriver.find_element(By.ID, "end_sessions").is_displayed():
                chromeDriver.find_element(By.ID, "end_sessions").click()
                chromeDriver.find_element(By.ID, "login_submit").click()
                scrapingProcess()
        except:
            print('except')
            scrapingProcess()


def scrapingProcess():
    print('scrapingProcess()')
    # time.sleep(10)
    global excelWorkbook
    currentDateAndTime = time.strftime("D-%d-%m-%Y-T-%H-%M-%S")
    excelWorkbook = xlsxwriter.Workbook(''+currentDateAndTime+'.xlsx')
    excelWorksheet = excelWorkbook.add_worksheet()
    currentTableRow = 0
    apdexColumnCount = 0
    indexOfCurrentSite = 2
    scrollCountInEntities = 0  # main page site count
    isNotLive = 0
    siteVisibleCheck = chromeDriver.find_elements(By.XPATH, "//div[contains(@class,'-wnd-CardBaseHeader-heading')]/a")
    if (len(siteVisibleCheck) == 0):
        time.sleep(20)
        scrapingProcess()
    totalSiteCount = chromeDriver.find_element(By.XPATH, "//div[contains(@class,'-wnd-CardBaseHeader-heading')]/a").text
    totalSiteCount = totalSiteCount[totalSiteCount.find("(")+1:totalSiteCount.find(")")]
    chromeDriver.find_element(By.XPATH, "//div[contains(@class,'-wnd-CardBaseHeader-heading')]/a").click()  # view all click
    excelWorksheet.write(0, 0, 'Site name')
    # below loop is used to select a specific site
    for index in range(int(totalSiteCount)):
        time.sleep(5)
        totalColumnHeaderForApdexCount = []
        entityPageUrl = chromeDriver.current_url
        print(index, 'index')
        # index +=9
        isLive = False
        if (index != 0 and index % 9 == 0):
            scrollCountInEntities += 1
            indexOfCurrentSite = 8  # 4
            time.sleep(5)
        isScrollBarInEntities = chromeDriver.find_elements(By.XPATH, "//div[contains(@class,'-wnd-DataTable')]/div")
        if (len(isScrollBarInEntities) == 0):
            time.sleep(15)
        scrollBarInEntities = chromeDriver.find_element(By.XPATH, "//div[contains(@class,'-wnd-DataTable')]/div")
        # below loop is used to scroll the page after scraping the data from set of sites
        for i in range(scrollCountInEntities):
            chromeDriver.execute_script('arguments[0].scrollBy(0,362);', scrollBarInEntities)
            time.sleep(2)
        pathOfCurrentSite = chromeDriver.find_element(By.XPATH, '//div['+str(indexOfCurrentSite)+']/span[contains(@class,"-wnd-DataTableBaseRowCell--ellipsisType-end")]')
        indexOfCurrentSite += 1
        isLiveSite = pathOfCurrentSite.text[pathOfCurrentSite.text.find("(")+1:pathOfCurrentSite.text.find(")")]
        if (isLiveSite.lower() == 'live'):
            isLive = True
            isNotLive = 1
        elif (isNotLive == 3):
            break
        else:
            chromeDriver.refresh()
            isNotLive += 1
            continue
        if (isLive):
            pathOfCurrentSite.click()   # clicking the current site
            time.sleep(15)
            print('d1')
            chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[1]/div[2]/div/div/div/button[2]").click() #since
            time.sleep(2)
            chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div/div/div/div/div/div/div[11]/span[1]").click() #custom
            if (index < 3):
                print('set-date')
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div[1]/div/div[2]/div/div[1]/div/div/div/span/input").click() #date1
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div[2]/div/div/div[1]/div/div[2]/div[3]/div[2]/div[6]").click() #sept8
                time.sleep(2)
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div/div/div[2]/div/div[2]/div/div/div/span/input").click()    #time1
                scrollBarInSelectedSiteTime = chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div[2]/div/div/div/div")
                chromeDriver.execute_script("arguments[0].scroll(100, 600);",   scrollBarInSelectedSiteTime)
                time.sleep(2)
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div[2]/div/div/div/div/div/button[4]").click() #2:30pm
                time.sleep(2)
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div/div/div[2]/div/div[3]/div/div/div/span/input").click() #date2
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div[2]/div/div/div[1]/div/div[2]/div[3]/div[2]/div[6]").click() #sept8
                time.sleep(2)
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div/div/div[2]/div/div[4]/div/div/div/span/input").click() #time2
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div[2]/div/div/div/div/div/div/button[2]/span/span[1]").click() #3:00pm
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div/div/div[2]/button[1]").click() #apply
            else:
                print('recent')
                chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[1]/div[2]/div/div/div[2]/footer/ol/li/a/span").click() #recent  
            
            # chromeDriver.refresh()
            time.sleep(10)
            print('d2')
            isScrollBarInSelectedSite = chromeDriver.find_elements(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]")
            print(len(isScrollBarInSelectedSite),'isScrollBarInSelectedSite')
            if (len(isScrollBarInSelectedSite) == 0):
                chromeDriver.refresh()
                time.sleep(15)                
            scrollBarInSelectedSite = chromeDriver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]")
            chromeDriver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollBarInSelectedSite)
            pathOfCurrentSiteName = chromeDriver.find_element(By.XPATH, "//div[contains(@class,'-wnd-EntitySwitcher-entityName')]/span")
            currentTableRow += 1
            siteNameForExcel = pathOfCurrentSiteName.text.find("(")
            siteNameForExcel = pathOfCurrentSiteName.text[0:siteNameForExcel-1]
            print(siteNameForExcel, 'name', pathOfCurrentSiteName.text)
            # excelWorksheet.write(currentTableRow, 0, siteNameForExcel)
            print(currentTableRow, 0, siteNameForExcel,'current')
            time.sleep(10)
            chromeDriver.execute_script("arguments[0].scroll(0, 2000);", scrollBarInSelectedSite)
            # print(scrollBarInSelectedSite)
            time.sleep(10)
            host = chromeDriver.find_element(By.CLASS_NAME, "HostTable")
            # print(host, 'host')
            numberOfColumnsInSelectedSite = host.find_elements(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[3]/div/div/div[2]/div[1]/div/div/div[1]/div")
            #/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div[4]/div/div/div[2]/div[1]/div/div/div[1]/div
            if (len(numberOfColumnsInSelectedSite) == 0):
                time.sleep(20)
                chromeDriver.execute_script("arguments[0].scroll(0, arguments[0].scrollHeight);", scrollBarInSelectedSite)
                numberOfColumnsInSelectedSite = host.find_elements(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[3]/div/div/div[2]/div[1]/div/div/div[1]/div")
                #/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[5]/div/div/div[2]/div[1]/div/div/div[1]/div[1]
            # if(currentTableRow == 1):
            #     excelWorksheet.write(0,0,'Site name')
            #     for i in range(1,len(numberOfColumnsInSelectedSite),1):
            #         columnHeaderInSelectedSite = chromeDriver.find_element(By.XPATH, '//div['+str(i)+']/div/span[contains(@class,"-wnd-TableHeaderCell-title")]')
            #         totalColumnHeaderForApdexCount.append(columnHeaderInSelectedSite.text)
            #         excelWorksheet.write(0,i,columnHeaderInSelectedSite.text)
            #         print(columnHeaderInSelectedSite.text,'sitee')
            #         apdexColumnCount = len(totalColumnHeaderForApdexCount) + 1
            print(len(numberOfColumnsInSelectedSite), 'numberOfColumnsInSelectedSite')
            for i in range(0, len(numberOfColumnsInSelectedSite), 1):
                # print(len(numberOfColumnsInSelectedSite),'nociss')
                columnHeaderInSelectedSite = host.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[3]/div/div/div[2]/div[1]/div/div/div[1]/div['+str(i+1)+']/button/span[1]')
                # /html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div[4]/div/div/div[2]/div[1]/div/div/div[1]/div[1]/button/span[1]
                # /html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div[4]/div/div/div[2]/div[1]/div/div/div[1]/div[2]/button/span[1]
                print(i,columnHeaderInSelectedSite.text)
                totalColumnHeaderForApdexCount.append(columnHeaderInSelectedSite.text)
                #excelWorksheet.write(row,column,value)
                excelWorksheet.write(0, i+1, columnHeaderInSelectedSite.text)
                # print(columnHeaderInSelectedSite.text, 'sitee')
                apdexColumnCount = len(totalColumnHeaderForApdexCount) + 1
            print(currentTableRow, apdexColumnCount, 'apd')
            time.sleep(5)
            # print(chromeDriver.find_element(By.XPATH, "//span[contains(@class,'-wnd-DataTableEntityRowCell-name')]").text,'moni')
            # print(len(chromeDriver.find_elements(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[4]/div/div/div[2]/div[1]/div/div/div[2]/span")),'isMonitoringData')
            isMonitoringData = chromeDriver.find_elements(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[3]/div/div/div[2]/div[1]/div/div/div[2]/span")
            print(len(isMonitoringData), 'monitor')
            totalMonitoringData = chromeDriver.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[3]/div/div/div[2]/div[1]/div/div/div[contains(@class,"-wnd-DataTableRow")]')
            print(len(totalMonitoringData), 'totalMonitoringData')
            # apdexCheck = chromeDriver.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[1]/div[1]/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/h2')
            # print(chromeDriver.find_element(By.XPATH, "//h2[text()='Apdex score']").text)
            apdexCheck = chromeDriver.find_element(By.XPATH, "//h2[text()='Apdex score']").text

            if (apdexCheck):
                print(apdexCheck,'apdexCheck')
                # apdexName = chromeDriver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[1]/div[1]/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/h2').text
                apdexName = chromeDriver.find_element(By.XPATH, "//h2[text()='Apdex score']").text
                excelWorksheet.write(0, apdexColumnCount, apdexName)
                # apdexData = chromeDriver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[1]/div[1]/section/div/div/div/div[1]/div/section/div[2]/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[2]/div/span[1]').text
                apdexData = chromeDriver.find_element(By.XPATH, "//div[2]/div/div/div[1]/div/div[1]/div[2]/div/span[contains(@class,'--vz--widget-chart-summary__value')]").text
                print(apdexData,'apdexData')
            if len(isMonitoringData) > 0:
                currentTableColumn = 0
                print(len(isMonitoringData),'mm')
                for totalindex in range(0,len(totalMonitoringData),1):
                    print(totalindex,'>>> totalindex')
                    for index in range(1, len(isMonitoringData)+2, 1):
                        print(index,'>>>',index % 8)
                        if (index % 8 == 0):
                            # print('one')
                            if(totalindex == len(totalMonitoringData)-1):     
                                print('continue')                           
                                continue
                            else:
                                print('continue-else')
                            currentTableRow += 1
                            currentTableColumn = 0
                            continue
                        else:
                            currentTableColumn += 1                            
                            # print(currentTableColumn,'currentTableColumn')
                            # monitoringDataOfSelectedSite = chromeDriver.find_element(By.XPATH, '//span['+str(i)+']/   div[contains(@class,"-wnd-DataTableEntityRowCell-name")]')
                            monitoringDataOfSelectedSite = chromeDriver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div[1]/div[2]/section/div[4]/div[2]/div[1]/section/div/div/div/div[1]/div/section/div/div[2]/div[3]/div/div/div[2]/div[1]/div/div/div['+str(totalindex+2)+']/span['+str(index)+']')
                            print( currentTableRow, currentTableColumn,monitoringDataOfSelectedSite.text,'val')
                            excelWorksheet.write(currentTableRow, 0, siteNameForExcel)
                            excelWorksheet.write(
                                currentTableRow, currentTableColumn, monitoringDataOfSelectedSite.text)
                            excelWorksheet.write(currentTableRow, apdexColumnCount, apdexData)
            chromeDriver.get(entityPageUrl)
            time.sleep(5)
    with keepawake(keep_screen_awake=False):
        # excelWorksheet.write(0, 0, 'Site name')
        time.sleep(5)
        excelWorkbook.close()
        chromeDriver.find_element(By.XPATH, "//div[contains(@class,'-wnd-UserMenu')]/button").click()
        chromeDriver.find_element(By.XPATH, "//div[contains(@class,'-wnd-PopoverListItem-content')]").click()
        chromeDriver.quit()
        os.system('start "excel.exe" "'+currentDateAndTime+'.xlsx"')
        time.sleep(5)


def beginScraping():
    try:
        webDriver()
    except:
        excelWorkbook.close()
        notFoundException()
    finally:
        return True


beginScraping()
