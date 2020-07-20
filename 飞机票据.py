import selenium
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import time
import xlrd
import re
import random
import csv
import requests


class www_singaporeair_com():
    # def excel(self):
    def excel(self, ):
        wb = xlrd.open_workbook('1.xls')  # 打开Excel文件
        sheet = wb.sheet_by_name('Sheet1')  # 通过excel表格名称(rank)获取工作表
        dat = []  # 创建空list
        for s in range(sheet.nrows):  # 循环读取表格内容（每次读取一行数据）
            cells = sheet.row_values(s)  # 每行数据赋值给cells
            data = (cells[0])  # 因为表内可能存在多列数据，0代表第一列数据，1代表第二列，以此类推
            # data = int(data)
            dat.append(data)  # 把每次循环读取的数据插入到list
        return dat

    def excel_two(self):
        wb = xlrd.open_workbook('1.xls')  # 打开Excel文件
        sheet = wb.sheet_by_name('Sheet1')  # 通过excel表格名称(rank)获取工作表
        dat = []  # 创建空list
        for b in range(sheet.nrows):  # 循环读取表格内容（每次读取一行数据）
            cells = sheet.row_values(b)  # 每行数据赋值给cells
            data = (cells[2])  # 因为表内可能存在多列数据，0代表第一列数据，1代表第二列，以此类推
            # data = int(data)
            dat.append(data)  # 把每次循环读取的数据插入到list
        return dat

    def click_webdriver(self):
        # for senddata, senddatatwo, senddatathree , senddatafour ,senddatafive,senddatasix in zip(a, b, c ,d ,e ,s):
        #     print(senddata)
        #     print(f)
        for name, password in zip(a, b):
            driver = webdriver.Chrome(r'chromedriver.exe')
            time.sleep(1)
            driver.get('https://www.singaporeair.com/zh_TW/hk/home#/book/bookflight')
            # 点击管理预约
            driver.find_element_by_xpath('//*[@id="hwidget"]/div[2]/div/div[1]/li[2]/div/div/span').click()
            time.sleep(2)
            data = re.findall('(.*)/.*', name)[0]
            # 输入
            driver.find_element_by_id('last_familyNameFieldBR').send_keys(data)
            driver.find_element_by_id('bookingReferenceFieldBR').send_keys(password)
            # 利用JS滑动
            js = "var q=document.documentElement.scrollTop=200"  # javascript语句
            driver.execute_script(js)  # 模拟滑动200
            time.sleep(5)
            # 点击确定
            driver.find_element_by_xpath('//*[@id="hwidget"]/div[2]/div[1]/div[2]/div[1]/div/section/div[2]/form/div/div[3]/div').click()
            time.sleep(30)
            try:
                name123 = driver.find_element_by_xpath('//*[@id="app-container"]/div/section[1]/div/aside/span/div/span').text
            except:
                name123 = ''
                pass
            try:
                name123 = driver.find_element_by_xpath('//div[@class="geetest_btn"]/div[@class="geetest_radar_btn"]/div[@class="geetest_radar_tip"]/span[@class="geetest_radar_tip_content"]').text
                print('======================='+str(name123)+'=======================')
            except:
                name123 = ''
                pass

            print(str(data)+'---------'+str(password))
            if str(name123) == str('点击按钮进行验证'):
                # 等待验证
                time.sleep(5)
                driver.find_element_by_xpath('//*[@id="captcha-box"]/div/div[2]/div[1]/div[3]').click()
                time.sleep(5)
                img=driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div/div[2]/div[1]/div/div[2]/img')
                url=img.get_attribute('src')
                print(url)
                r = requests.get(url)
                with open("验证码.jpeg", "wb") as fp:
                    fp.write(r.content)
                time.sleep(5)
                zidonghua.main('',
                        '',
                        '验证码.jpeg',
                        "http://v1-http-api.jsdama.com/api.php?mod=php&act=upload",
                        '1',
                        '8',
                        '1303',
                        '')
                for code in self.code:
                    x=int(220)+int(code.split(',')[0])      
                    y=int(170)+int(code.split(',')[1])    
                    print(x,y)
                    # js = "var q=document.elementFromPoint(476, 444).click()"  # javascript语句
                    # driver.execute_script(js)
                    ActionChains(driver).move_by_offset(x, y).click().perform()
                    ActionChains(driver).move_by_offset(-x, -y).perform()
                    time.sleep(2)
                time.sleep(2)
                driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div/div[3]/a').click()
                WebDriverWait(driver, 600000000, 0.5).until(EC.presence_of_element_located((By.XPATH, '//*[text()="验证成功"]')))
                print('验证成功')
                time.sleep(30)
                try:
                    key = driver.find_element_by_xpath('//*[@id="app-container"]/div/section[1]/div/aside/span/div/span').text
                except:
                    key = ''
                    pass
                # 判断是什么界面
                if str(key) == str('下载机票和收据'):
                    print('在下载机票和收据的界面')
                    # 打印数据
                    name1 = driver.find_element_by_xpath('//div[@class="flight-info--trigger-content_schedule-wrap"]/div/div/span[@class="flight-info_inline-content flight-info_airport-code"]').text
                    name2 = driver.find_element_by_xpath('//div[@class="flight-info--trigger-content_schedule-wrap"]/div/div/span[@class="flight-info_inline-content flight-info_flight-tracker"]').text
                    print(name1, name2)
                    time.sleep(2)
                    # 刷新界面
                    shuaxin = driver.current_url
                    driver.get(shuaxin)
                    time.sleep(10)
                    driver.find_element_by_xpath('//*[@id="captcha-box"]/div/div[2]/div[1]/div[3]').click()
                    time.sleep(5)
                    img=driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div/div[2]/div[1]/div/div[2]/img')
                    url=img.get_attribute('src')
                    print(url)
                    r = requests.get(url)
                    with open("验证码.jpeg", "wb") as fp:
                        fp.write(r.content)
                    time.sleep(5)
                    zidonghua.main('',
                            '',
                            '验证码.jpeg',
                            "http://v1-http-api.jsdama.com/api.php?mod=php&act=upload",
                            '1',
                            '8',
                            '1303',
                            '')
                    for code in self.code:
                        x=int(220)+int(code.split(',')[0])      
                        y=int(170)+int(code.split(',')[1])    
                        print(x,y)
                        # js = "var q=document.elementFromPoint(476, 444).click()"  # javascript语句
                        # driver.execute_script(js)
                        ActionChains(driver).move_by_offset(x, y).click().perform()
                        ActionChains(driver).move_by_offset(-x, -y).perform()
                        time.sleep(2)
                    time.sleep(2)
                    driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div/div[3]/a').click()
                    # 等待刷新后返回到下载几篇和收据的界面
                    WebDriverWait(driver, 600000000, 0.5).until(EC.presence_of_element_located((By.XPATH, '//*[text()="下载机票和收据"]')))
                    time.sleep(10)
                    # 点击下载机票和收据
                    driver.find_element_by_xpath('//div[@class="wrap-control"]/aside[@class="booking-reference booking-reference--w-tab-float"]/span/div[@class="booking-reference_print"]').click()
                    time.sleep(5)
                    try:
                        num2 = driver.find_element_by_xpath('/html/body/aside[13]/div/div/h2').text
                        print(num2)
                    except:
                        num2 = ''
                        pass
                    # 判断是否在打印电子机票和收据的界面
                    if str(num2) == str('打印电子机票和收据'):
                        num = driver.find_element_by_xpath('//div[@class="list-receipt__info-group"]/div[@class="list-receipt-info"]/form[@id="retrieveTicket"]/span').text
                        print(num)
                        print('结束')
                        driver.quit()
                    else:
                        print('点击下载收据后并没有打开')
                        driver.quit()
                else:
                    print('在管理预订的界面')
                    time.sleep(5)
                    driver.quit()
            elif str(name123) == str('下载机票和收据'):
                print('在下载机票和收据的界面')
                # 打印数据
                name1 = driver.find_element_by_xpath('//div[@class="flight-info--trigger-content_schedule-wrap"]/div/div/span[@class="flight-info_inline-content flight-info_airport-code"]').text
                name2 = driver.find_element_by_xpath('//div[@class="flight-info--trigger-content_schedule-wrap"]/div/div/span[@class="flight-info_inline-content flight-info_flight-tracker"]').text
                print(name1, name2)
                time.sleep(2)
                # 刷新界面
                shuaxin = driver.current_url
                driver.get(shuaxin)
                time.sleep(5)
                # 等待刷新后返回到下载几篇和收据的界面
                WebDriverWait(driver, 600000000, 0.5).until(EC.presence_of_element_located((By.XPATH, '//*[text()="下载机票和收据"]')))
                time.sleep(10)
                 # 点击下载机票和收据
                driver.find_element_by_xpath('//div[@class="wrap-control"]/aside[@class="booking-reference booking-reference--w-tab-float"]/span/div[@class="booking-reference_print"]').click()
                time.sleep(5)
                try:
                    num2 = driver.find_element_by_xpath('/html/body/aside[13]/div/div/h2').text
                    print(num2)
                except:
                    num2 = ''
                    pass
                    # 判断是否在打印电子机票和收据的界面
                if str(num2) == str('打印电子机票和收据'):
                    num = driver.find_element_by_xpath('//div[@class="list-receipt__info-group"]/div[@class="list-receipt-info"]/form[@id="retrieveTicket"]/span').text
                    print(num)
                    print('结束')
                    driver.quit()
                else:
                    print('点击下载收据后并没有打开')
                    driver.quit()
            else:
                print('2')
                driver.quit()



    def main(self,api_username, api_password, file_name, api_post_url, yzm_min, yzm_max, yzm_type, tools_token):
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0',
            # 'Content-Type': 'multipart/form-data; boundary=---------------------------227973204131376',
            'Connection': 'keep-alive',
            'Host': 'v1-http-api.jsdama.com',
            'Upgrade-Insecure-Requests': '1'
        }

        files = {
            'upload': (file_name, open(file_name, 'rb'), 'image/png')
        }

        data = {
            'user_name': api_username,
            'user_pw': api_password,
            'yzm_minlen': yzm_min,
            'yzm_maxlen': yzm_max,
            'yzmtype_mark': yzm_type,
            'zztool_token': tools_token
        }
        s = requests.session()
        # r = s.post(api_post_url, headers=headers, data=data, files=files, verify=False, proxies=proxies)
        r = s.post(api_post_url, headers=headers, data=data, files=files, verify=False)
        req=r.json()['data']['val']
        print(req)
        self.code=req.split('|')

if __name__ == '__main__':
    zidonghua = www_singaporeair_com()
    a = (zidonghua.excel())
    b = (zidonghua.excel_two())
    zidonghua.click_webdriver()