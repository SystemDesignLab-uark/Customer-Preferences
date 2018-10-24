import pandas as pd
import time
import xlwt
import sys
import xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from importlib import reload


#reload(sys)
#sys.setdefaultencoding('utf8')
workbook = xlwt.Workbook()
#print('Begin')
sheet1 = workbook.add_sheet("Hub")
sheet1.write(0,0,"Hub ID")
sheet1.write(0,1,"Hub Name")
sheet1.write(0,2,"Area")
sheet1.write(0,3,"Description")
sheet1.write(0,4,"Invoice")
sheet1.write(0,5,"Active Since")
sheet1.write(0,6,"Response time")
sheet1.write(0,7,"Average")
sheet1.write(0,8,"Print quality")
sheet1.write(0,9,"Service")
sheet1.write(0,10,"Speed")
sheet1.write(0,11,"Communication")
sheet1.write(0,12,"Customer_1")
sheet1.write(0,13,"Review Days_1")
sheet1.write(0,14,"Review_1")
sheet1.write(0,15,"Customer_2")
sheet1.write(0,16,"Review Days_2")
sheet1.write(0,17,"Review_2")
sheet1.write(0,18,"Customer_3")
sheet1.write(0,19,"Review Days_3")
sheet1.write(0,20,"Review_2")
sheet1.write(0,21,"Total_Reviews")

sheet2 = workbook.add_sheet("Delivery method")
sheet2.write(0,0,"Hub ID")
sheet2.write(0,1,"Hub Name")
sheet2.write(0,2,"Delivery_method")
sheet2.write(0,3,"Date")
sheet2.write(0,4,"Price")

sheet3 = workbook.add_sheet("Printer")
sheet3.write(0,0,"Supplier ID")
sheet3.write(0,1,"Supplier Name")
sheet3.write(0,2,"Marvin Size")
sheet3.write(0,3,"Printer name")
sheet3.write(0,4,"Material")
sheet3.write(0,5,"Strength")
sheet3.write(0,6,"Cost")
sheet3.write(0,7,"Color")

links = open('links.txt')

row1 = 0
row2 = 1
row=1

for link in links:
    Hub_ID = link
#   try:
#   driver = webdriver.Firefox(executable_path='C:\Users\charan_vuda\geckodriver.exe')
    driver = webdriver.Firefox(executable_path='/Users/collincorcoran/Documents/CustomerPreferenceProject/geckodriver')
    #driver.get("https://www.3dhubs.com/service/cubiforminc")
    driver.get(link)
    row1=row1+1
    #addition to website
    try:
       Got_it = driver.find_element_by_xpath("//div[contains(@class,'cc-window')]/div[contains(@class,'cc-compliance')]/a[contains(@class,'cc-btn cc-dismiss')]")
       Got_it.click()
    except:
       print("got it link missing")

    try:
        h3d_buttons = driver.find_element_by_xpath("//div[contains(@class,'h3d-user__content')]/div[contains(@class,'h3d-user__name')]")
        Hub_Name = h3d_buttons.find_element_by_xpath("./h1").text
        sheet1.write(row1,0,Hub_ID)
        sheet1.write(row1,1,Hub_Name)
    except:
        print("Hub_ID and Hub_Name not found in :"+link)

    h3d_area = ""
    try:
        h3d_buttons = driver.find_elements_by_xpath("//ol[contains(@class,'h3d-breadcrumbs')]/li[contains(@class,'h3d-breadcrumbs')]")
        for h3d_button in h3d_buttons:
            h3d_area = h3d_area + ',' + h3d_button.find_element_by_xpath(".//a").text
        sheet1.write(row1,2,h3d_area)
    except:
        print("Hub_area not found in :"+link)

    try:
        Description= driver.find_element_by_xpath("//div [contains(@class,'h3d-panel__body')]/p").text
        sheet1.write(row1,3,Description)
    except:
        print("Description not found in :"+link)

    try:
        Invoice=driver.find_element_by_xpath("//div [contains(@class,'h3d-grid__col-md-4')]/div[@class='h3d-panel']/div[contains(@class,'h3d-panel__body')]/div[contains(@class,'h3d-grid__col-md-8')]").text
        sheet1.write(row1,4,Invoice)
    except:
        print("Invoice not found in :"+link)

    try:
        Active_Since=driver.find_element_by_xpath("//div [contains(@class,'h3d-grid__col-md-4')]/div[@class='h3d-panel']/div[contains(@class,'h3d-panel__body')]/div[contains(@class,'ng-binding')]").text
        sheet1.write(row1,5,Active_Since)
    except:
        print("Active_Since not found in :"+link)

    try:
        Response_time=driver.find_element_by_xpath("//div [contains(@class,'h3d-grid__col-md-8')]/listing-response-time[@class='ng-isolate-scope']/span[@class='u-text-primary']/span[contains(@class,'u-text-strong')]").text
        sheet1.write(row1,6,Response_time)
    except:
        #print("Respone_time 1st path not found in :"+link)
        try:
            Response_time=driver.find_element_by_xpath("//div [contains(@class,'h3d-grid__col-md-8')]/listing-response-time[@class='ng-isolate-scope']/span[contains(@data-ng-class,'u-text-primary')]/span[contains(@class,'u-text-strong')]").text
            sheet1.write(row1,6,Response_time)
        except:
            print("Respone_time not found in :"+link)

#ratings collected but amount of reviews is not
    col1 = 7
    try:
        h3d_ratings = driver.find_elements_by_xpath("//div[contains(@class,'h3d-rating ng-isolate-scope h3d-rating--dark')]")
        for h3d_rating in h3d_ratings:
            sheet1.write(row1,col1,h3d_rating.find_element_by_xpath(".//span[contains(@data-ng-bind,'vm.rating.value | rating5')]").text)

            col1=col1+1
    except:
        print("Ratings not found in :"+link)
#test to collect total reviews
    try:
        Total_Reviews = driver.find_element_by_xpath("//div [contains(@class, 'h3d-grid__col-md-8')]/div[@class='h3d-panel']/div[contains(@class, 'h3d-collection')]/div[contains(@class, 'h3d-collection__item')]/div[contains(@class, 'h3d-grid')]/div[contains(@class, 'h3d-grid__col-md-6')]").text
        sheet1.write(row1,21,Total_Reviews)
    except:
        print("Total Reviews not found in :" +link)

    try:
        Reviews = driver.find_elements_by_xpath("//div [contains(@class,'h3d-grid__col-md-10')]")
        for Review in Reviews:
            Customer_1=Review.find_element_by_xpath("./strong/a").text
            sheet1.write(row1,col1,Customer_1)
            col1= col1+1

            Days_ago_1=Review.find_element_by_xpath("./strong/span[contains(@data-ng-bind,'reviewed')]").text
            sheet1.write(row1,col1,Days_ago_1)
            col1= col1+1

            Review_1=Review.find_element_by_xpath("./p").text
            sheet1.write(row1,col1,Review_1)
            col1= col1+1
    except:
        print("Reviews not found in :"+link)


    try:
        rowd=row2
        h3d_buttons_3 = driver.find_elements_by_xpath("//tr[@class='ng-scope']/td[contains(@class,'h3d-table__cell h3d-hub-deliveries__cell--narrow h3d-hub-deliveries__cell--wide')]")
        for h3d_button_3 in h3d_buttons_3:
            Delivery_method=h3d_button_3.find_element_by_xpath("./span").text
            #sheet2.write(rowd,2,(Delivery_method.encode("utf-8")))
            sheet2.write(rowd,2,Delivery_method)
            h3d_buttons=driver.find_element_by_xpath("//div[contains(@class,'h3d-user__content')]/div[contains(@class,'h3d-user__name')]")
            Hub_name = h3d_buttons.find_element_by_xpath("./h1").text
            sheet2.write(rowd,0,Hub_ID)
            sheet2.write(rowd,1,Hub_name)
            rowd=rowd+1
    except:
        print("Hub_ID or Hub_Name not found in :"+link)

    rowd1=row2

    try:
        h3d_buttons_1 = driver.find_elements_by_xpath("//tr[@class='ng-scope']")
        for h3d_button_1 in h3d_buttons_1:
            Date=h3d_button_1.find_element_by_xpath("td[contains(@class,'ng-binding')]").text
            sheet2.write(rowd1,3,Date)
            rowd1=rowd1+1
    except:
        print("Date not found in :"+link)
    rowp=row2

    try:
        h3d_buttons_2 = driver.find_elements_by_xpath("//tr[@class='ng-scope']/td[contains(@data-ng-class,'u-text-primary')]/delivery-price[contains(@class,'ng-isolate-scope')]")
        for h3d_button_2 in h3d_buttons_2:
            Price=h3d_button_2.find_element_by_xpath(".//span").text
            sheet2.write(rowp,4,Price)
            rowp=rowp+1
        row2 = rowp
    except:
        print("Price not found in :"+link)

	
    #Click on each Mervin and get the details
    try:
        h3d_buttons = driver.find_elements_by_xpath("//div[contains(@class,'h3d-buttons-group')]/div[contains(@class,'h3d-button')]")
        for h3d_button in h3d_buttons:
            h3d_button.click()
            time.sleep(5)

            h3d_panels = driver.find_elements_by_xpath("//div[@class='h3d-panel']")
            #print(len(h3d_panels))

            for h3d_panel in h3d_panels:
                h3d_h4s = h3d_panel.find_elements_by_xpath(".//h4[@class='h3d-h4']")
                if len(h3d_h4s) > 0:
                    for h3d_h4 in h3d_h4s:
                        heading = h3d_h4.find_element_by_xpath("./span")
                        #print(heading.text)
                        h3d_collections = h3d_panel.find_elements_by_xpath(".//div[contains(@class,'h3d-collection__item')]")
                        #print(len(h3d_collections))
                        for h3d_collection in h3d_collections:
                            collection = h3d_collection.find_element_by_xpath(".//h5[contains(@class,'h3d-h5')]")
                            #print(collection.text)
                            h3d_lists = h3d_collection.find_elements_by_xpath(".//li[contains(@class,'h3d-list')]")
                            for h3d_list in h3d_lists:
                                strong = h3d_list.find_element_by_xpath(".//strong")
                                span = h3d_list.find_element_by_xpath(".//span[2]")
                                #f.write(h3d_button.text+'\t'+heading.text+'\t'+collection.text+'\t'+strong.text+'\t'+span.text+'\n')
                                colors = driver.find_elements_by_xpath("//div[contains(@class,'h3d-collection')]/div[contains(@class,'h3d-grid')]/div[contains(@class,'h3d-grid__col-sm-8')]/div[contains(@class,'h3d-grid__cell')]")
                                sheet3.write(row,0,Hub_ID)
                                sheet3.write(row,1,Hub_Name)
                                sheet3.write(row,2,h3d_button.text)
                                sheet3.write(row,3,heading.text)
                                sheet3.write(row,4,collection.text)
                                sheet3.write(row,5,strong.text)
                                sheet3.write(row,6,span.text)
                                #colors = driver.find_elements_by_xpath("//div[contains(@class,'h3d-collection')]/div[contains(@class,'h3d-grid')]/div[contains(@class,'h3d-grid__col-sm-8')]/div[contains(@class,'h3d-grid__cell')]")
                                col =7
                                for color in colors:
                                    items=color.find_elements_by_xpath("./span")
                                    for item in items:
                                        sheet3.write(row,col,item.text)
                                        col= col+1
                                row = row+1
    except:
        print("Issue in marvin details :"+link)
#   except:
#   print('Error reading :'+str(link))
    workbook.save("HubDetails.xls")
    driver.quit()

#workbook.save("HubDetails.xls")
