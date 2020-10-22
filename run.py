# coding=utf-8

import time
import shelve
import os
import signal

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait

DATA_PATH = "votedata"
#微信投票的地址
URL = 'https://mp.weixin.qq.com/s/THi2tZEqI8zPlAUxTpZlZA'
#显示多少次取值
COUNT = 10

def pushData(data):
    with shelve.open(DATA_PATH) as write:
        write[time.strftime('%Y-%m-%d %H:%M:%S',
                            time.localtime(time.time()))] = data


def getData():
    data = {}
    with shelve.open(DATA_PATH) as read:
        keys = list(read.keys())
        try:
            keys.sort()
        except:
            pass
        keys = keys[-COUNT:]
        for key in keys:
            data[key] = read[key]
    return data


def getVote():
    # 投票的微信页面地址
    BROWER.get(URL)
    # 投票的iframe的class
    voteiframe = BROWER.find_elements_by_class_name('js_editor_vote_card')[0]
    # 获取iframe的实际地址
    data_src = voteiframe.get_attribute('data-src')

    # 访问实际的投票地址
    BROWER.get("https://mp.weixin.qq.com" + data_src)
    # 返回投票数据变量名（具体可在页面源码中找到）
    voteInfo = BROWER.execute_script('return voteInfo')
    return voteInfo

def run():
    votes = getVote()
    votes = votes['vote_subject'][0]['options']
    pushData(votes)
    votes = getData()
    data = {}
    title = []
    for votetime in votes:
        title.append(votetime.split(' '))
        for vote in votes[votetime]:
            if vote['name'] in data:
                data[vote['name']] = str(
                    data[vote['name']]) + '|' + str(vote['cnt']).rjust(10, ' ')
            else:
                data[vote['name']] = str(vote['cnt']).rjust(10, ' ')

    keys = data.keys()
    keyL = len(max(data.keys(), key=len))
    re = {}
    for key in keys:
        re[key.ljust(keyL, '　')] = data[key]
    title_date = '　'*keyL+' '
    title_time = '　'*keyL+' '
    for key in title:
        title_date += str(key[0])+'|'
        title_time += str(key[1]).center(10, ' ')+'|'
    
    print(title_date)
    print(title_time)

    for key in re:
        print(key+':'+re[key]+'|')

chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_argument('--headless')
BROWER = webdriver.Chrome(options=chromeOptions)
wait = WebDriverWait(BROWER, 10)
BROWER.maximize_window()
BROWER.implicitly_wait(6)

def quit(signum, frame):
    BROWER.quit()
    print('正在退出...')
    exit()

if __name__ == "__main__":
    signal.signal(signal.SIGINT, quit)
    signal.signal(signal.SIGTERM, quit)
    while True:
        os.system('clear')
        run();
        time.sleep(60)
