import time
import xlwt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

f = open('HubLinks_Atlanta.txt','w')
#driver = webdriver.Chrome('C:\Users\charan_vuda\chromedriver.exe')
#driver = webdriver.Firefox(executable_path='C:\Users\charan_vuda\geckodriver.exe')
driver = webdriver.Firefox(executable_path='/Users/collincorcoran/Documents/CustomerPreferenceProject/geckodriver')
driver.get("https://www.3dhubs.com/3dprint#?place=Atlanta,%20GA,%20USA&latitude=33.7489954&longitude=-84.3879824&shipsToCountry=US&shipsToState=GA&materialSubsets=fdm_standard-pla")
#driver.get("https://www.3dhubs.com/3dprint#?place=New%20York,%20NY,%20USA&latitude=40.7127753&longitude=-74.0059728&shipsToCountry=US&shipsToState=NY")
time.sleep(30)
print('Back from sleep')

#Click on each material to populate the service providers
h3d_materials = driver.find_elements_by_xpath("//div[contains(@class,'h3d-grid__col-md-6')]")

for h3d_material in h3d_materials:
	h3d_material.click()
	time.sleep(4)
	driver.find_element_by_xpath("//div[contains(@class,'h3d-grid__col-md-5')]/div[contains(@class,'ng-scope')]/small[contains(@class,'h3d-link')]").click()
	time.sleep(5)
	move=ActionChains(driver)

	#Set the distance to 250km
	distance=driver.find_element_by_xpath("//input[contains(@class,'h3d-simple-slider')]")
	move.click_and_hold(distance).move_by_offset(60,0).release().perform()
	time.sleep(3)

	while True:
		#print('Processing page: '+str(page_nbr))
		#Scan through all the hubs
		h3d_hubs = driver.find_elements_by_xpath("//div[contains(@class,'h3d-hub-row__name')]/button[contains(@class,'h3d-hub-row__name')]")
		for h3d_hub in h3d_hubs :
			try :		
				h3d_hub.click()		#Click on each hub
				time.sleep(3)
				driver.switch_to.window(window_name=driver.window_handles[-1])		#Switch control to new tab to get the url
				#print('Link of new tab: '+driver.current_url)
				f.write(driver.current_url+'\n')  	#Write the url to a file
				driver.close()	#Close the new tab
				driver.switch_to.window(window_name=driver.window_handles[0]) #Switch control back to old tab
			except :
				print ('Problem reading a link in page: '+str(page_nbr))
		
		try:
			driver.find_elements_by_xpath("//nav[contains(@class,'h3d-pagination')]/button[contains(@data-ng-click,'vm.next')]")[0].click()
			time.sleep(10)
			page_nbr = page_nbr + 1
		except:
			break
f.close()
