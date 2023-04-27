from selenium import webdriver 
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

class freevee():
    data = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/IDMSDSN.xlsx', 'HomeBase')
    dsn = list(data.HomeBase)
    print('No. of HomeBase Devices:', len(dsn))

    data1 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/IDMSDSN.xlsx', 'HomeNonBase')
    dsn1 = list(data1.HomeNonBase)
    print('No. of HomeNonBase Devices:', len(dsn1))

    data2 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/IDMSDSN.xlsx', 'Office')
    dsn2 = list(data2.Office)
    print('No. of Office Devices:', len(dsn2))

    print('~' * 20)

    driver = webdriver.Firefox (service=Service (GeckoDriverManager().install()))
    action = ActionChains(driver)
    print('~' * 20)
    url = 'https://prod.idms.gdq.amazon.dev/#/device-checkin-dashboard'   # New
    userid_locator = '//*[@id="user_name_field"]'
    username = "hdavisvi"
    signin_locator = '//*[@id="user_name_btn"]'
    dsni_locator = '//*[@id="app"]/div[2]/div/div/div/main/div/div/div/div/div/div[5]/div/div[1]/div/div[2]/div[1]/div/div/span/div/input'   # New
    Check_Box = '/html/body/div[1]/div[2]/div/div/div/main/div/div/div/div/div/div[5]/div/div[2]/div[1]/table/thead/tr/th[1]/label/div/div/div[1]/input'   # New
    CheckIn1 = '//*[@id="app"]/div[2]/div/div/div/main/div/div/div/div/div/div[5]/div/div[1]/div/div[1]/div[2]/div/div[2]/button'   # New
    validate_locator = '//*[@id="return-device-dsn-input"]/div/div/div[2]/div/div/div/div[2]/button'
    dropd1_locator = '/html/body/div[1]/div[2]/div/div/div/main/div/div/div/div/div/div[6]/div/div/div/div/div[1]/div/div/div/div[1]/div/div/div/div[3]/div/div/div/div/div/div/div[1]/button/span[2]/span'
    location_locator = '/html/body/div[1]/div[2]/div/div/div/main/div/div/div/div/div/div[4]/div/div[2]/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div/div[3]/div/div/div/div/div/div/div[1]/button'
    location2_locator = '/html/body/div[1]/div[2]/div/div/div/main/div/div/div/div/div/div[4]/div/div[2]/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div/div[1]/div/div/div/div[2]/div/div/div/div[3]/div/div/div/div/div/div/div[1]/button'
    checkin_locator = '/html/body/div[1]/div[2]/div/div/div/main/div/div/div/div/div/div[4]/div/div[2]/div/div[2]/div/div/div/div[2]/div/div/div[2]/div/div/div[2]/div/div/div/div[3]/button'
    reset_locator = '//*[@id="app"]/div[2]/div/div/div/main/div/div/div/div/div/div[5]/div/div[1]/div/div[2]/div[1]/div/div/span/div/span[2]'

    def login(self):
        self.driver.get(self.url)
        self.driver.find_element(by=By.XPATH, value=self.userid_locator).send_keys(self.username)
        self.driver.find_element(by=By.XPATH, value=self.signin_locator).click()
        self.driver.implicitly_wait(45)

    def checkin(self):
        self.login()
        for i in self.dsn:
            self.driver.find_element(by=By.XPATH, value=self.dsni_locator).send_keys(i)
            time.sleep(4)

            try:
                if self.driver.find_element(by=By.XPATH, value=self.Check_Box).is_enabled():
                    self.driver.find_element(by=By.XPATH, value=self.Check_Box).click()
                    self.driver.find_element(by=By.XPATH, value=self.CheckIn1).click()
                    time.sleep(1)
                    self.driver.find_element(by=By.XPATH, value= self.location_locator).send_keys("Home")
                    self.driver.find_element(by=By.XPATH, value= self.location2_locator).send_keys("Base")
                    self.driver.find_element(by=By.XPATH, value=self.checkin_locator).click()
                    time.sleep(2)
            except NoSuchElementException:
                self.driver.implicitly_wait(3)
                self.driver.find_element(by=By.XPATH, value=self.reset_locator).click()
                print('HomeBase DSN Not Listed:', i)
                time.sleep(2)
        print('Check-in Successful for HomeBase Valid DSNs')
        print('~' * 20)
        for i in self.dsn1:
            self.driver.find_element(by=By.XPATH, value=self.dsni_locator).send_keys(i)
            time.sleep(4)

            try:
                if self.driver.find_element(by=By.XPATH, value=self.Check_Box).is_enabled():
                    self.driver.find_element(by=By.XPATH, value=self.Check_Box).click()
                    self.driver.find_element(by=By.XPATH, value=self.CheckIn1).click()
                    time.sleep(1)
                    self.driver.find_element(by=By.XPATH, value= self.location_locator).send_keys("Home")
                    self.driver.find_element(by=By.XPATH, value= self.location2_locator).send_keys("Non-Base")
                    self.driver.find_element(by=By.XPATH, value=self.checkin_locator).click()
                    time.sleep(2) 
            except NoSuchElementException:
                self.driver.implicitly_wait(3)
                self.driver.find_element(by=By.XPATH, value=self.reset_locator).click()
                print('HomeNonBase DSN Not Listed:', i)
                time.sleep(2)
        print('Check-in Successful for HomeNonBase Valid DSNs')
        print('~' * 20)
        for i in self.dsn2:
            self.driver.find_element(by=By.XPATH, value=self.dsni_locator).send_keys(i)
            time.sleep(4)

            try:
                if self.driver.find_element(by=By.XPATH, value=self.Check_Box).is_enabled():
                    self.driver.find_element(by=By.XPATH, value=self.Check_Box).click()
                    self.driver.find_element(by=By.XPATH, value=self.CheckIn1).click()
                    time.sleep(1)
                    self.driver.find_element(by=By.XPATH, value= self.location_locator).send_keys("Office")
                    self.driver.find_element(by=By.XPATH, value= self.location2_locator).send_keys("MAA15")
                    self.driver.find_element(by=By.XPATH, value=self.checkin_locator).click()
                    time.sleep(2)
            except NoSuchElementException:
                self.driver.implicitly_wait(3)
                self.driver.find_element(by=By.XPATH, value=self.reset_locator).click()
                print('Office DSN Not Listed:', i)
                time.sleep(2)
        print('Check-in Successful for Office Valid DSNs')
        print('~' * 20)
        time.sleep(5)
        self.driver.close()

fv = freevee()
fv.checkin()