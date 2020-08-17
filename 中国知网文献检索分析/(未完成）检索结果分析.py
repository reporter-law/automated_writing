"""程序说明"""
# -*-  coding: utf-8 -*-
# Author: cao wang
# Datetime : 2020
# software: PyCharm
# 收获:电脑端知网无法检索，只能手机端

import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from lxml import etree
import  time


logging.disable(logging.DEBUG)
def start_logger():
    """日志初始化设置、文件名（时间）、DEBUG为调试级别(级别导致输出内容的不同）、日志的记录格式、日期格式"""

    logging.basicConfig(  #filename='daily_report_error_%s.log' %

        #datetime.strftime(datetime.now(), '%m%d%Y_%H%M%S'),

        level=logging.DEBUG,
        format='%(asctime)s %(message)s',
        datefmt='%m-%d %H:%M:%S')
start_logger()
def search_result(keyword):
    """返回检索初步结果"""
    browser = webdriver.Chrome()
    wait = WebDriverWait(browser, 20)

    """网页获取"""
    browser.get('http://wap.cnki.net/touch/web/guide')
    click = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="keyword"]')))
    click.click()
    input_wprd = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="keyword_ordinary"]')))
    input_wprd.clear()
    input_wprd.send_keys(keyword)
    button = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchbody_ordinary"]/div/div[1]/div/div[1]/a[2]')))
    button.click()
    #time.sleep(1111)
    time.sleep(11111)
    try:
        while True:
            Btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'c-company__body-item-more')))

            if Btn:
                browser.execute_script('window.scrollTo(0,document.body.scrollHeight)')
                Btn.click()

            else:
                with open('xxxxxxxxx.txt', "a+", encoding='utf-8')as f:
                    f.write(browser.page_source)
                return browser.page_source
    except:
        with open('xxxxxxxxx.txt', "a+", encoding='utf-8')as f:
            f.write(browser.page_source)
        return browser.page_source






def get_page(html):
    """解析函数：首先要进行返回html"""

    html_ = etree.HTML(html)
    contents = html_.xpath('//*[@id="searchlist_div"]/div')
    # 内容遍历每一页
    print(contents)

    # logging.info(contents.xpath('/div//text()'))
    data = []
    for texts in contents:
        # logging.info(texts)
        """对每一页进行内容提取"""
        titles = ''.join(
            texts.xpath('//div[@class="c-company__body-title  c-company__body-title-blue"]//text()')).strip()
        # print(titles)
        title = [i.strip() for i in titles.split('\n')]
        print(len(title))
        print(title[0] + '\n')
        with open('titles.txt', "a+", encoding='utf-8')as f:
            f.write(str(title) + "\n")




    """
        其余内容
        authors = ''.join(texts.xpath('//div[@class="c-company__body-author"]//text()')).replace(' ','')
        author = [i.strip() for i in authors.split('\n') if i !='']
        print(author)


        contents = ''.join(texts.xpath('//div[@class="c-company__body-content"]//text()'))
        content = [i.strip().replace(' ', '') for i in contents.split('\n') if i !='']
        #print(content)


        times = ''.join(texts.xpath('//span[@class="color-green"]//text()'))
        time_ = [i.strip().replace(' ', '').replace('\xa0',',') for i in times.split('\n') if i !='']
        #print(time_)

        downloads = ''.join(texts.xpath('//a[@class="c-company__body-info"]//text()'))
        download = [i.replace('\xa0','-').replace(' ','')  for i in downloads.split('\n') if i !='']
        downloads = []
        for i in download:
            if i == '':
                pass
            else:
                downloads.append(i)
        print(downloads)




        data_list = []
        for t,a,c,ti,d in zip(title,author,content,time_,download):
            dict_datas={'title':t,'author':a,'content':c,'time':ti,'download':d}
            data_list.append(dict_datas)
        print(data_list)
        """
        # time.sleep(11111)




def main(keyword):
    """主函数"""
    html = search_result(keyword)
    get_page(html)





if __name__ == "__main__":
    list_ = ['权利告知','米兰达规则','犯罪嫌疑人知情权','侦查讯问告知']
    for i in list_:
        main(i)


