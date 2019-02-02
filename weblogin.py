from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time

#checks the webpage source for a string, and if there is a string, will tab so that the link isnt clicked
def needtab(driver,string1):

    driver.switch_to_frame(driver.find_element_by_id("contentAreaFrame"))
    html_source = driver.page_source
    if string1 in html_source:
        actions = ActionChains(driver)
        actions.send_keys(Keys.TAB)
        actions.perform()
        driver.switch_to.default_content()
        return 1
    else:
        driver.switch_to.default_content()

#get mylam downloads and access the page
def accessmylam(selected, accountvalue, partno, name_password1, name_password2):
    #convert passed values into intelligible strings
    selected = selected.get()
    accountvalue = accountvalue.get()
    partno.strip()
    #add browser settings
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--start-maximized")
    #change the default download directory
    prefs = {
                 "download.default_directory": r"C:\Users\wchu\Documents\\", # IMPORTANT - ENDING SLASH V IMPORTANT
                 "directory_upgrade": True}
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("detach", True)
    browser = webdriver.Chrome(chrome_options=options)
    #launch browser
    browser.get("https://username:passwor@portal.mylam.com/irj/portal?&EPPAP19_0")
    driver = browser
    username = driver.find_element_by_id('logonuidfield')
    password = driver.find_element_by_id('logonpassfield')

   #check which account was chosen - use the correct username/password for that account
    if accountvalue == 2:
        username.send_keys(name_password2[0])
        password.send_keys(name_password2[1])
    else:
        username.send_keys(name_password1[0])
        password.send_keys(name_password1[1])

    
    #print (password.text )#There's no text under div main, what would you expect?
    password.send_keys(Keys.RETURN)
    
    #skip through the useless pages
    mylam = driver.find_element_by_id('subTabIndex2')
    mylam.send_keys(Keys.RETURN)
    
    
    contentarea = driver.find_element_by_id('contentAreaFrame')

    contentarea.send_keys(Keys.TAB * 3)


    contentarea.send_keys(Keys.RETURN)
    
    
    
    ok = driver.find_element_by_id("contentAreaFrame")
    ok.send_keys(Keys.TAB* 2)

    ok.send_keys(Keys.RETURN)
    
    
    ok = driver.find_element_by_id("contentAreaFrame")
    ok.send_keys(Keys.TAB*4)

    ok.send_keys(partno)
    ok.send_keys(Keys.RETURN *2)
    time.sleep(1)


    ok = driver.find_element_by_id("contentAreaFrame")
    

    ok.send_keys(Keys.TAB*14)
    #use the needtab definition to skip useless links
    needtab(driver,'">C1</a')
    needtab(driver,'">C4</a')
    #check which option needs to be downloaded and hit tab accordingly
    if selected ==2 or selected ==4:
        ok.send_keys(Keys.ENTER)
    
    if selected ==3 or selected ==4:
        if needtab(driver, "displaybom") ==1:
            needtab(driver, "images/stpviewer.bmp" )

            needtab(driver,  "images/jt2go.bmp")
            needtab(driver,'futurerev=y')
            needtab(driver,'getecn.asp')
            needtab(driver,'ExcelIcon')
            ok.send_keys(Keys.ENTER)
            ok = driver.find_element_by_id("contentAreaFrame")
            ok.send_keys(Keys.TAB *9)
            ok.send_keys(Keys.ENTER)
    if selected != 1:
        import os
        os.startfile("C://Users//wchu//Documents") 


if __name__ == '__main__':

    accessmylam()
