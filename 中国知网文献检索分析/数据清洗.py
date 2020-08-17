"""程序说明"""
# -*-  coding: utf-8 -*-
# Author: cao wang
# Datetime : 2020
# software: PyCharm
# 收获:
import time
import logging
import json
import re


def start_logger():
    """日志初始化设置、文件名（时间）、DEBUG为调试级别(级别导致输出内容的不同）、日志的记录格式、日期格式"""

    logging.basicConfig(  #filename='daily_report_error_%s.log' %

        #datetime.strftime(datetime.now(), '%m%d%Y_%H%M%S'),

        level=logging.DEBUG,
        format='%(asctime)s %(message)s',
        datefmt='%m-%d %H:%M:%S')
start_logger()
set_ =[]
"""
with open('titles.txt', 'r+', encoding='utf-8') as f:
    words = f.readlines()
    for word in words:
        word.replace('[','').replace(']','').replace("'",'').replace('[','').replace(']','')
        #print(word)

        print(type(word))
        words = word.strip().replace('[','').replace(']','').replace("'",'')
        abs = words.split(',')
        for i in abs:
            set_.append(i)
    print(len(set_))
    set_q = set(set_)
    print(set_q)
    print(len(set_q))#257,215\208\95\62
"""
with open('/项目/论文写作自动化/中国知网文献检索分析/data_json.json', 'r+', encoding='utf-8') as f:
    words = f.read()
    pattern = re.compile('"name":(.*?),')
    word = re.findall(pattern,words)
    print(word)
    print(len(word))
    set_word = set(word)
    print(len(set_word))










