#coding:utf-8
import time
import selenium
from selenium import webdriver
def openBrowser():
    driver = webdriver.Firefox()
    return driver
def openUrl(driver):
    driver.get('https://www.imooc.com/')
def wait_web_ele(func):
    flage = True
    times = 0
    while flage and times <100:
        f = False
        print('while')
        try:
            element = func
        except:
            f = True
        if f:
            time.sleep(0.1)
            times += 1
        else:
            flage = False
    # print('element')
    # element = func
    return element
def find_ele(driver):
    wait_web_ele(driver.find_element_by_link_text('运维&测试')).click()
    # driver.find_element_by_link_text('运维&测试').click()
if __name__ =='__main__':
    driver = openBrowser()
    openUrl(driver)
    find_ele(driver)

