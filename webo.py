from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import datetime,time
driver = webdriver.Chrome()
driver.get("https://web.whatsapp.com")
wait5 = WebDriverWait(driver, 20)
wb = load_workbook('1.Family_Master.xlsx')
sheet = wb.active
sheet = wb.get_sheet_by_name('Dadaka')
sheet1 = wb.active
sheet1 = wb.get_sheet_by_name('Nanka')
sheet_list = [sheet,sheet1]
flag1, flag2, flag3, flag4, flag5, flag6 = (0, )*6
while(True):
    for g in sheet_list:
        clk_ob = datetime.datetime.now()
        date  = clk_ob.day
        month = clk_ob.month
        colfval, colgval, colival, coljval = ([], )*4
        fc = g['F']
        gc = g['G']
        ic = g['I']
        jc = g['J']
        colfval = [c1.value for c1 in fc ]
        flag1 = 1 if(date in colfval) else 0
        colgval = [g1.value for g1 in gc ]
        flag2 = 1 if(month in colgval) else 0
        colival = [i1.value for i1 in ic ]
        flag3 = 1 if(date in colival) else 0
        coljval = [j1.value for j1 in jc ]
        flag4 = 1 if(month in coljval) else 0
        if(((clk_ob.hour==int(0) and clk_ob.minute==int(0))) and ((flag1==1 and flag2==1) or (flag3==1 and flag4==1))):
                index1 = []
                index2 = []
                if(flag1==1 and flag2==1):
                    for i in range(len(colfval)):
                        if(colfval[i]==date and colgval[i]==month):
                            index1.append(i)
                            flag5 = 1
                if(flag3==1 and flag4==1):
                    for i in range(len(colival)):
                        if(colival[i]==date and coljval[i]==month):
                            index2.append(i)
                            flag6 = 1
                bc = g['C']
                colcval = []
                colcval = [b1.value for b1 in bc]
                def webfire(msg1,msg2):
                    try:
                        wait5.until(EC.presence_of_element_located((By.XPATH,"//span[(@title ='Incredible Family')]")))
                        wait5.until(EC.presence_of_element_located((By.XPATH,"//span[(@title ='If test')]")))
                        wait5.until(EC.presence_of_element_located((By.XPATH,"//span[(@title ='Raghav Kundra')]")))
                        wait5.until(EC.presence_of_element_located((By.XPATH,"//span[(@title ='Bareily ki Burfi')]")))
                        wait5.until(EC.presence_of_element_located((By.XPATH,"//span[(@title ='Bkb test')]")))
                        if(g==sheet):
                            user = driver.find_element_by_xpath("//span[(@title ='Incredible Family')]")
                            #user = driver.find_element_by_xpath("//span[(@title ='If test')]")
                        elif(g==sheet1):
                            user = driver.find_element_by_xpath("//span[(@title ='Bareily ki Burfi')]")
                            #user = driver.find_element_by_xpath("//span[(@title ='Bkb test')]")
                        user.click()
                        msg_box = driver.find_element_by_class_name("_2S1VP")
                        msg_box.send_keys(msg1)
                        button = driver.find_element_by_class_name("_35EW6")
                        button.click()
                        time.sleep(2)
                        user = driver.find_element_by_xpath("//span[(@title ='Raghav Kundra')]")
                        user.click()
                        msg_box = driver.find_element_by_class_name("_2S1VP")
                        msg_box.send_keys(msg2)
                        button = driver.find_element_by_class_name("_35EW6")
                        button.click()
                        time.sleep(2)
                    except:
                        print("Chat Box not found. \n")
                        print("not completed")
                if(flag5==1):
                    for i in index1:
                        msg1 = 'Happy Birthday Dear' + ' ' + str(colcval[i])
                        #msg1 = '-'
                        msg2 = "Wished " + ' ' + str(colcval[i])
                        #msg2 = '-'
                        print(msg1,msg2)
                        webfire(msg1,msg2)
                time.sleep(2)
                if(flag6==1):
                    for i in index2:
                        msg1 = 'Happy Anniversary Dear' + ' ' + str(colcval[i])
                        #msg1 = '-'
                        msg2 = "Wished " + ' ' + str(colcval[i])
                        #msg2 = '-'
                        print(msg1,msg2)
                        webfire(msg1,msg2)
