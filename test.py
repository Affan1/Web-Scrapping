from logging import exception
from turtle import delay
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import csv
import pyodbc
import pandas as pd
import time
import traceback


browser = webdriver.Chrome('C:\\Users\\Affan\\Desktop\\Web Scrapping\\chromedriver.exe')

conn = pyodbc.connect('')

c_p = "-company-profile"
i_s = "-income-statement"
b_s = "-balance-sheet"
c_f = "-cash-flow"
r_t = "-ratios"
h_d = "-historical-data"



AccTypes = ['A','Q']




for yr in range(3,10):
    page_num = str(yr)
    url = 'https://www.investing.com/stock-screener/?sp=country::44|sector::a|industry::a|equityType::a%3Ceq_market_cap;' + page_num
    browser.get(url)
    delay = 8

    try:
        myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'symbol left bold elp')))
        print("Page is ready!")
    except Exception:
        time.sleep(0) 
    tables = browser.find_elements(By.ID ,'resultsTable')
    hreflist=[]
    namelist=[]

    for table in tables:
        rows = table.find_elements(By.TAG_NAME,'tr') # get all of the rows in the table
        for row in rows[1:]:
            # print(row.text)
            cols = row.find_elements(By.TAG_NAME,'td')
            
            a = (cols[1].find_elements(By.TAG_NAME,'a')[0])
            href = a.get_attribute('href').split('?')[0]
            name = a.text
            hreflist.append(href)
            namelist.append(name)
    for href in range(len(hreflist)):


        # for AccType in AccTypes:


            vals1= []
            vals2= []
            vals3= []
            vals4= []
            year = []
            i=16
            
            
            #balance sheet code

            browser.get(hreflist[href]+c_p)
            try:
                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'companyProfileHeader')))
                company_code = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[2]/div[3]/span[2]")[0]
                vals1.append(company_code.text)
                company_name = browser.find_elements(By.XPATH,"//h1[@class='float_lang_base_1 relativeAttr']")[0]
                vals1.append(company_name.text)
                company_symbol = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[2]/div[3]/span[2]")[0]
                vals1.append(company_symbol.text)
                # company_profile = browser.find_elements(By.XPATH,"//p[@id='profile-fullStory-showhide']")[0]
                # vals1.append(company_profile.text)
                # vals2.append(company_profile.text)
                # vals3.append(company_profile.text)
                # vals4.append(company_profile.text)
                exchange = browser.find_elements(By.XPATH,"//i[@class='btnTextDropDwn arial_12 bold']")[0]
                vals1.append(exchange.text)
                currency = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[1]/div[2]/div[2]/span[4]")[0]
                vals1.append(currency.text)
                header = browser.find_elements(By.XPATH,"//div[@class='companyProfileHeader']")[0]
                for divs in header.find_elements(By.CSS_SELECTOR,'div'):
                    data = [a.text for a in divs.find_elements(By.CSS_SELECTOR,'a')]
                    if(len(data)>0):
                        vals1.append(data[0])
                num =0
                for divs in header.find_elements(By.CSS_SELECTOR,'div'):
                    data = [a.text for a in divs.find_elements(By.CSS_SELECTOR,'p')]
                    if(len(data)>0 and num==0):
                        vals1.append(data[0])
                        num = num +1
                    
                locality = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[1]/span[3]/a/span[1]")[0]
                country = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[1]/span[3]/a/span[3]")[0]
                website = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[4]/span[3]/a")[0]
                vals1.append(locality.text)
                vals1.append(country.text)
                vals1.append(website.text)
                vals1.append(hreflist[href]+ c_p)
            except Exception as error:
                traceback.print_exc()
            try:
                # Start of Balance Sheet
                browser.get(hreflist[href]+b_s)
                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.XPATH, "//a[@data-ptype='Annual']")))        
                python_button = browser.find_elements(By.XPATH,"//a[@data-ptype='Annual']")[0]
                time.sleep(2)
                python_button.click()
                vals1.append(python_button.text)

                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.XPATH, "//table[@class='genTbl reportTbl']")))        
            
                # years of sheet
                years = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/th")

            
                for h in years:
                    year.append(h.text.splitlines())
                # print(vals1)

                # acc heads name in table
                accheads = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/td/span")

            
                for r in accheads:
                     vals2.append(r.text)
                # print(vals2)

                # value in table
                val = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/td")
            
                for d in val:
                    vals3.append(d.text.splitlines()[0])
                # print(vals3)
            
                # 1st column of year in balance sheet
                ran0= len(vals1);
                ran1= len(vals2);
                ran2 = len(vals3);
                date = int(year[1][0]);
                vals1.append(year[1][0])
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+1])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],vals1[13],vals1[11],b_s,vals3[j],vals3[j+1])
                            cursor.commit()
                yeardate = date -1
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+2])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],b_s,vals3[j],vals3[j+2])
                            cursor.commit()
                yeardate = date -2
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+3])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],b_s,vals3[j],vals3[j+3])
                            cursor.commit()
                yeardate = date -3
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+4])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],b_s,vals3[j],vals3[j+4])
                            cursor.commit()
            except Exception as error:
                traceback.print_exc()
            
            # Income Statement Code
            vals1.clear()
            vals2.clear()
            vals3.clear()
            vals4.clear()
            year.clear()
            browser.get(hreflist[href]+c_p)
            try:
                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'companyProfileHeader')))
                company_code = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[2]/div[3]/span[2]")[0]
                vals1.append(company_code.text)
                company_name = browser.find_elements(By.XPATH,"//h1[@class='float_lang_base_1 relativeAttr']")[0]
                vals1.append(company_name.text)
                company_symbol = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[2]/div[3]/span[2]")[0]
                vals1.append(company_symbol.text)
                # company_profile = browser.find_elements(By.XPATH,"//p[@id='profile-fullStory-showhide']")[0]
                # vals1.append(company_profile.text)
                # vals2.append(company_profile.text)
                # vals3.append(company_profile.text)
                # vals4.append(company_profile.text)
                exchange = browser.find_elements(By.XPATH,"//i[@class='btnTextDropDwn arial_12 bold']")[0]
                vals1.append(exchange.text)
                currency = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[1]/div[2]/div[2]/span[4]")[0]
                vals1.append(currency.text)
                header = browser.find_elements(By.XPATH,"//div[@class='companyProfileHeader']")[0]
                for divs in header.find_elements(By.CSS_SELECTOR,'div'):
                    data = [a.text for a in divs.find_elements(By.CSS_SELECTOR,'a')]
                    if(len(data)>0):
                        vals1.append(data[0])
                num =0
                for divs in header.find_elements(By.CSS_SELECTOR,'div'):
                    data = [a.text for a in divs.find_elements(By.CSS_SELECTOR,'p')]
                    if(len(data)>0 and num==0):
                        vals1.append(data[0])
                        num = num +1
                    
                locality = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[1]/span[3]/a/span[1]")[0]
                country = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[1]/span[3]/a/span[3]")[0]
                website = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[4]/span[3]/a")[0]
                vals1.append(locality.text)
                vals1.append(country.text)
                vals1.append(website.text)
                vals1.append(hreflist[href]+ c_p)
            except Exception as error:
                traceback.print_exc()


             # Start of Income Statement
            
            
            browser.get(hreflist[href]+i_s)
            try:
                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.XPATH, "//a[@data-ptype='Annual']")))        
                python_button = browser.find_elements(By.XPATH,"//a[@data-ptype='Annual']")[0]
                time.sleep(2)
                python_button.click()
                vals1.append(python_button.text)

                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.XPATH, "//table[@class='genTbl reportTbl']")))        
            
                # years of sheet
                years = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/th")

            
                for h in years:
                    year.append(h.text.splitlines())
                # print(vals1)

                # acc heads name in table
                accheads = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/td/span")

            
                for r in accheads:
                     vals2.append(r.text)
                # print(vals2)

                # value in table
                val = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/td")
            
                for d in val:
                    vals3.append(d.text.splitlines()[0])
                # print(vals3)
            
                # 1st column of year in income_Statement
                ran0= len(vals1);
                ran1= len(vals2);
                ran2 = len(vals3);
                date = int(year[1][0]);
                vals1.append(year[1][0])
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+1])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],vals1[13],vals1[11],i_s,vals3[j],vals3[j+1])
                            cursor.commit()
                yeardate = date -1
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+2])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],i_s,vals3[j],vals3[j+2])
                            cursor.commit()
                yeardate = date -2
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+3])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],i_s,vals3[j],vals3[j+3])
                            cursor.commit()
                yeardate = date -3
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+4])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],i_s,vals3[j],vals3[j+4])
                            cursor.commit()
            except Exception as error:
                traceback.print_exc()

                # Cash Flow code
                vals1.clear()
                vals2.clear()
                vals3.clear()
                vals4.clear()
                year.clear()
                browser.get(hreflist[href]+c_p)
            try:
                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'companyProfileHeader')))
                company_code = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[2]/div[3]/span[2]")[0]
                vals1.append(company_code.text)
                company_name = browser.find_elements(By.XPATH,"//h1[@class='float_lang_base_1 relativeAttr']")[0]
                vals1.append(company_name.text)
                company_symbol = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[2]/div[3]/span[2]")[0]
                vals1.append(company_symbol.text)
                # company_profile = browser.find_elements(By.XPATH,"//p[@id='profile-fullStory-showhide']")[0]
                # vals1.append(company_profile.text)
                # vals2.append(company_profile.text)
                # vals3.append(company_profile.text)
                # vals4.append(company_profile.text)
                exchange = browser.find_elements(By.XPATH,"//i[@class='btnTextDropDwn arial_12 bold']")[0]
                vals1.append(exchange.text)
                currency = browser.find_elements(By.XPATH,"//*[@id='quotes_summary_current_data']/div[1]/div[2]/div[2]/span[4]")[0]
                vals1.append(currency.text)
                header = browser.find_elements(By.XPATH,"//div[@class='companyProfileHeader']")[0]
                for divs in header.find_elements(By.CSS_SELECTOR,'div'):
                    data = [a.text for a in divs.find_elements(By.CSS_SELECTOR,'a')]
                    if(len(data)>0):
                        vals1.append(data[0])
                num =0
                for divs in header.find_elements(By.CSS_SELECTOR,'div'):
                    data = [a.text for a in divs.find_elements(By.CSS_SELECTOR,'p')]
                    if(len(data)>0 and num==0):
                        vals1.append(data[0])
                        num = num +1
                    
                locality = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[1]/span[3]/a/span[1]")[0]
                country = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[1]/span[3]/a/span[3]")[0]
                website = browser.find_elements(By.XPATH,"//*[@id='leftColumn']/div[10]/div[1]/div[4]/span[3]/a")[0]
                vals1.append(locality.text)
                vals1.append(country.text)
                vals1.append(website.text)
                vals1.append(hreflist[href]+ c_p)
            except Exception as error:
                traceback.print_exc()


             # Start of Cash flow 
            
            
            browser.get(hreflist[href]+c_f)
            try:
                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.XPATH, "//a[@data-ptype='Annual']")))        
                python_button = browser.find_elements(By.XPATH,"//a[@data-ptype='Annual']")[0]
                time.sleep(2)
                python_button.click()
                vals1.append(python_button.text)

                myElem = WebDriverWait(browser, delay).until(EC.presence_of_element_located((By.XPATH, "//table[@class='genTbl reportTbl']")))        
            
                # years of sheet
                years = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/th")

            
                for h in years:
                    year.append(h.text.splitlines())
                # print(vals1)

                # acc heads name in table
                accheads = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/td/span")

            
                for r in accheads:
                     vals2.append(r.text)
                # print(vals2)

                # value in table
                val = browser.find_elements(By.XPATH,"//table[@class='genTbl reportTbl']/tbody/tr/td")
            
                for d in val:
                    vals3.append(d.text.splitlines()[0])
                # print(vals3)
            
                # 1st column of year in cash flow
                ran0= len(vals1);
                ran1= len(vals2);
                ran2 = len(vals3);
                date = int(year[1][0]);
                vals1.append(year[1][0])
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+1])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],vals1[13],vals1[11],c_f,vals3[j],vals3[j+1])
                            cursor.commit()
                yeardate = date -1
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+2])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],c_f,vals3[j],vals3[j+2])
                            cursor.commit()
                yeardate = date -2
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+3])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],c_f,vals3[j],vals3[j+3])
                            cursor.commit()
                yeardate = date -3
                for i in range(0,ran1):
                    for j in range(0,ran2):
                        if vals2[i] == vals3[j]:
                            print(vals3[j])
                            print(vals3[j+4])
                        
                            cursor = conn.cursor()
                            cursor.execute("insert into Web_Sheets(Company_Name,Comp_ID,Sector,Industry,Acc_Type,Year,Link,Sheet,Acc_Head,Acc_Value) values (?,?,?,?,?,?,?,?,?,?)", vals1[1],vals1[0],vals1[6],vals1[5],vals1[12],yeardate,vals1[11],c_f,vals3[j],vals3[j+4])
                            cursor.commit()
            except Exception as error:
                traceback.print_exc() 
browser.close()
