import json
import time
import requests
from loguru import logger
from DrissionPage import WebPage, ChromiumOptions, SessionOptions
import random
from datetime import datetime
import concurrent.futures
from threading import Lock
import html

import AOSCCOCR


def randomSleep():
    # 设置随机休眠时间，范围从min_seconds到max_seconds，整数秒
    min_seconds = 1
    max_seconds = 3
    # 生成一个随机整数，表示休眠的时间（秒）
    sleep_time = random.randint(min_seconds, max_seconds)
    print(f"Randomly selected sleep time: {sleep_time} seconds")
    # 让程序休眠这个随机时间
    time.sleep(sleep_time)
def rpapageshadow():
    co = ChromiumOptions()
    co.existing_only(False)
    # co = ChromiumOptions().headless()
    so = SessionOptions()
    page = WebPage(chromium_options=co, session_or_options=so)
    page.get('https://www.nbsinglewindow.cn/user/login.do')
    logger.info('第一次cookie状态检测')
    # cookiea = page.cookies(as_dict=True)  xpath://*[@id="cvf-page-content"]/div/div/div/div[2]/div/img
    cookiea = page.cookies()
    dictionary = {cookie['name']: cookie['value'] for cookie in cookiea}
    cookiea = dictionary
    logger.info(cookiea)
    logger.info('登录标识状态A')
    logger.info(cookiea.get('eporttoken'))
    if cookiea.get('eporttoken') is None:
        page.ele('xpath://*[@id="userId"]').input('nbfar')
        randomSleep()
        page.ele('xpath://*[@id="userPassword"]').input('nbfar666')
        randomSleep()

        # 调用 AOSCC 的OCR识别模式 开始
        page.ele('xpath://*[@id="certificationCodeImg"]').get_screenshot('thescreen.png')
        randomSleep()
        # OCR识别模块
        with open("thescreen.png", 'rb') as f:
            img_bytes = f.read()
        ocr = AOSCCOCR.AosccOcr()
        poses = ocr.classification(img_bytes)
        logger.info(poses)
        # 调用 AOSCC 的OCR识别模式 结束



        page.ele('xpath://*[@id="certificationCode"]').input(poses)
        # 这个按钮标签为了稳定安全考虑 后续更换为点控式按钮
        randomSleep()
        page.ele('xpath://html/body/div[2]/div[2]/a[1]').click()
        randomSleep()

    logger.info('第二次cookie状态检测')
    cookieb = page.cookies()
    dictionary = {cookie['name']: cookie['value'] for cookie in cookieb}
    cookieb = dictionary
    logger.info(cookieb)
    logger.info('登录标识状态B  ')
    logger.info(cookieb.get('eporttoken'))



    page.get('https://www.nbsinglewindow.cn/user/login.do?redirectURL=http%3A%2F%2Fp.nbsinglewindow.cn%2Fmember%2Freleaseinfo%2Freleaseinfo%21allquery.do')
    return page




if __name__ == '__main__':
    page = rpapageshadow()