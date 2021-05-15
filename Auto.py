from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from time import sleep, strftime
from random import randint
import pandas as pd
from openpyxl import load_workbook
import os
import sys

dirname = 'D:\\Report\\1400\\TabdilVaz'
progdirname = os.path.dirname(__file__)
       
chromedriver_path = r'D://chromedriver/chromedriver.exe' # Change this to your own chromedriver path!
webdriver = webdriver.Chrome(executable_path=chromedriver_path)
# webdriver.set_window_size(1920, 1080)
webdriver.maximize_window()
webdriver.get('https://amar.imo.org.ir/')
html = webdriver.page_source
userData = open("userData.udata", "r")

username = webdriver.find_element_by_name('_58_INSTANCE_dehyari_login')
username.send_keys(userData.readline().rstrip())
password = webdriver.find_element_by_name('_58_INSTANCE_dehyari_password')
password.send_keys(userData.readline().rstrip())

webdriver.find_element_by_xpath('//button[text()=" ورود "]').click()
companies = userData.readline().rstrip().split(',')
print(companies)
# companies = ['SazMotori','SazMotori']
forms =  userData.readline().rstrip().split(',')
excelfile =  userData.readline().rstrip()
if len(sys.argv) >= 2:
    excelfile = sys.argv[1]
if len(sys.argv) >= 3:
    forms = sys.argv[2].split(',')
print(forms)

companies = ['SH']
forms = ['Frm3']
i = 0
delaytime = 3
attachmentError = ''
# داشبورد مدیریت
webdriver.get('https://amar.imo.org.ir/member/-_-irn-l-14000277232')
webdriver.find_element_by_xpath("//*[contains(text(),'نمایش داده های آماری')]/..").click()
sleep(delaytime)
mainFrame = webdriver.find_element_by_xpath("//iframe")
# webdriver.switch_to_frame('_cartable_WAR_Cartableportlet_dashboard-application_iframe_')
webdriver.switch_to.frame(mainFrame)
webdriver.find_element_by_xpath("//*[contains(text(),'مجموعه اطلاعات جامع اداری و استخدامی')]").click()
webdriver.find_element_by_xpath("//*[contains(text(),'فرم شماره 3- اطلاعات جامع نیروهای شرکتی')]").click()

for company in companies:
    wb = load_workbook(os.path.join(dirname, company+excelfile))

    if('Frm2' in forms):
        ws = wb['Frm2']

    if('Frm3' in forms):
        ws = wb['Frm3']
        added = 0
        withError = 0
        col = []
        values = []

        ws['AM1'] = 'state'
        ws['AN1'] = 'error'
        ws['AO1'] = 'place'
        ws['AP1'] = 'serial'
        # http://amarnameh.imo.org.ir/Input/Update.aspx?Id=8023&cid=281
        index = 0
        total = ws.max_row
        for rownum in ws.iter_rows():
            index = index + 1
            try:
                if index == 1:
                    col = [(u"" if cell.value is None else str(cell.value).strip()) for cell in rownum]
                else:
                    values = [(u"" if cell.value is None else str(cell.value).strip()) for cell in rownum]

                    if(values[38] == 'added'):
                        continue
                    elif(values[41] == ''):
                        webdriver.find_element_by_xpath("//label[contains(text(),'کد ملی')]/../input").clear()
                        webdriver.find_element_by_xpath("//label[contains(text(),'کد ملی')]/../input").send_keys(values[0])
                        webdriver.find_element_by_xpath("//label[contains(text(),'کد ملی')]/../input").send_keys(u'\ue007')

                        sleep(delaytime)

                        webdriver.switch_to_default_content()
                        webdriver.switch_to.frame(mainFrame)
                        msg = "پیدا نشد"
                        el = webdriver.find_element_by_xpath("//a/img[contains(@title,'ویرایش')]/..")
                        msg = "ویرایش نشد"
                        webdriver.execute_script("arguments[0].click();", el)

                        
                    else:                   
                        webdriver.find_element_by_id("_srp_amar_amarmanagement_dynamic_dataviewer_WAR_AmarManagementportlet_add").click()

                    i = 2
                    xpString = "/html/body/div/div/div/div/div/div/div/section[2]/form/fieldset/div/div/fieldset/div/div[1]/div["+str(i)+"]/input"
                    webdriver.find_element_by_xpath(xpString).clear()
                    webdriver.find_element_by_xpath(xpString).send_keys(values[i-1])#نام
                    
                    msg = "موفق"
                    # webdriver.find_element_by_xpath().click()                                                             
                    print('{0} - {1} {2:3.0f}% '.format(index,total,index/total*100))

                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_dialog_954246153"]/div/div[2]/button[1]').click() 
                    # i = 6
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44839').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44839').send_keys(values[i])#نام
                    # i = 7
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44840').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44840').send_keys(values[i])#نام خانوادگی
                    # i = 8
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44841').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44841').send_keys(values[i])#نام پدر
                    # i = 9
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44842').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44842').send_keys(values[i])#کدملی
                    # i = 10
                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FACT_FIELD_44843"]/option['+values[i]+']').click()# آخرين مدرک تحصیلی
                    # i = 11
                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FACT_FIELD_44844"]/option['+values[i]+']').click()# رشته تحصيلي
                    # i = 12
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44845').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44845').send_keys(values[i])# ساير رشته تحصيلي
                    # i = 13
                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FACT_FIELD_44846"]/option['+values[i]+']').click()# جنسیت
                    # i = 14
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44847').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44847').send_keys(values[i])# تاريخ تولد
                    # i = 15
                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FACT_FIELD_44848"]/option['+values[i]+']').click()# وضعيت تاهل
                    # i = 16
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44849').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44849').send_keys(values[i])# تعداد فرزند
                    # i = 17
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44850').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44850').send_keys(values[i])# شماره تلفن همراه
                    # webdriver.find_element_by_xpath('//*[@id="mnuNext"]').click()# next page 2

                    # i = 18
                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FACT_FIELD_44851"]/option['+values[i]+']').click()# سابقه ايثارگري ؟
                    # if(values[i]=='2'):
                    #     i = 19
                    #     webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_FACT_FIELD_44852"]/option['+values[i]+']').click()# وضعيت ايثارگري
                    #     i = 20
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44853').clear()
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44853').send_keys(values[i])# تاريخ شروع رزمندگي
                    #     i = 21
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44854').clear()
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44854').send_keys(values[i])# تاريخ پايان رزمندگيi = 2
                    #     i = 22
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44855').clear()
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44855').send_keys(values[i])# تاريخ جانبازي
                    #     i = 23
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44856').clear()
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44856').send_keys(values[i])# درصد جانبازي
                    #     i = 24
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44857').clear()
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44857').send_keys(values[i])# تاریخ شروع اسارت
                    #     i = 25
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44858').clear()
                    #     webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44858').send_keys(values[i])# تاریخ پایان اسارت
                    # webdriver.find_element_by_xpath('//*[@id="mnuNext"]').click()# next page 3

                    # i = 26
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44859').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44859').send_keys(values[i])# سوابق شركتي در شهرداري از تاريخ
                    # i = 27
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44860').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44860').send_keys(values[i])# سوابق شركتي در شهرداري تا تاريخ(29-4-98
                    # i = 28
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44861').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44861').send_keys(values[i])# مجموع سوابق شركتي در شهرداري -سال
                    # i = 29
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44862').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44862').send_keys(values[i])# مجموع سوابق شركتي در شهرداري -ماه
                    # i = 30
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44863').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44863').send_keys(values[i])# مجموع سوابق شركتي در شهرداري -روز
                    # i = 31
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44864').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44864').send_keys(values[i])# محل دقيق خدمت
                    # i = 32
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44865').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44865').send_keys(values[i])# شغل فعلي در شهرداري
                    # i = 33
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44866').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44866').send_keys(values[i])# نام شركت پيمانكاري مربوطه
                    # i = 34
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44867').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44867').send_keys(values[i])# كد بيمه تامين اجتماعي
                    # webdriver.find_element_by_xpath('//*[@id="mnuNext"]').click()# next page 4

                    # temp = os.path.join(progdirname,'1234567890.pdf')
                    # attachmentError = ''
                    # i=-1
                    # attachmentFile = os.path.join(dirname,company+'\\'+values[9]+'t.pdf')
                    # if(not(os.path.exists(attachmentFile))):
                    #     attachmentFile = os.path.join(dirname,company+'\\'+values[9]+' t.pdf')
                    #     if(not(os.path.exists(attachmentFile))):
                    #         attachmentFile = temp
                    #         attachmentError = attachmentError + ' t'
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44868').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44868').send_keys(attachmentFile)# مدرک تحصیلی

                    # attachmentFile = os.path.join(dirname,company+'\\'+values[9]+'s.pdf')
                    # if(not(os.path.exists(attachmentFile))):
                    #     attachmentFile = os.path.join(dirname,company+'\\'+values[9]+'.pdf')
                    #     if(not(os.path.exists(attachmentFile))):
                    #         attachmentFile = temp
                    #         attachmentError = attachmentError + ' s'
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44869').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44869').send_keys(attachmentFile)# سابقه بیمه

                    # attachmentFile = os.path.join(dirname,company+'\\'+values[9]+'g.pdf')
                    # if(not(os.path.exists(attachmentFile))):
                    #     attachmentFile = os.path.join(dirname,company+'\\'+values[9]+'j.pdf')
                    #     if(not(os.path.exists(attachmentFile))):
                    #         attachmentFile = temp
                    #         attachmentError = attachmentError + ' g'
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44870').clear()
                    # webdriver.find_element_by_name('ctl00$ContentPlaceHolder1$FACT_FIELD_44870').send_keys(attachmentFile)# قزازداد        

                    # webdriver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_mnuSave"]').click()# finish                    
                    if(attachmentError == ''):
                        added += 1
                        ws['AM'+str(index)] = "added"
                        ws['AN'+str(index)] = ''
                        ws['AO'+str(index)] = ''
                        print('{0} Frm3 : {1} - {2} {3:3.0f}% Adeed {4} '.format(company,index,total,index/total*100,values[9]))

                    else:
                        withError += 1
                        ws['AM'+str(index)] = "error"
                        ws['AN'+str(index)] = 'attachment'
                        ws['AO'+str(index)] = attachmentError
                        print('{0} Frm3 : {1} - {2} {3:3.0f}% Adeed AttachmentError {4} - {5}'.format(company,index,total,index/total*100,values[9], attachmentError))
                    try:
                        webdriver.switch_to.alert.accept()
                    except:
                        print('alert error')

                   
            except:
                ws['AM'+str(index)] = "error"
                withError += 1
                ws['AN'+str(index)] = col[i]
                ws['AO'+str(index)] = values[i]
                print('{0} Frm3 : {1} - {2} {3:3.0f}%  Error {4} - {5} {6} {7}'.format(company,index,total,index/total*100,values[9], i,col[i], values[i]))                                 
                continue

        try:
            wb.save(os.path.join(dirname, company+excelfile))
        except:
            print('file open') 

        print('Frm3 {} Added {} Person.'.format(company,added))
        print('Frm3 {} With {} Error.'.format(company,withError))
webdriver.stop_client()
webdriver.close()