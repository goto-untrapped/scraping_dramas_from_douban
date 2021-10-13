import requests
import re
from requests.exceptions import RequestException
from selenium import webdriver
import time
from time import sleep
import urllib.parse
import win32com.client

#dont need to use when you use selenium?
############################################
#login to douban (it replys error418 if you don't login)
def LOGIN(ClassName):
    if ClassName == "GlobalPhoneItem-district":
        buttuon = driver.find_elements_by_class_name("GlobalPhoneItem-district")
        button[RegionNumber].click()
        time.sleep(1)
    elif ClassName == "account-form-input":
        username_input = driver.find_element_by_class_name("account-form-input")
        username_input.send_keys(PhoneNumber)
        time.sleep(1)

    button = driver.find_element_by_class_name(ClassName)
    button.click()
    time.sleep(1)

#serach specific url and open description to get information about high evaluation dramas
#update 1 time site and get info that evalation >= 8.2
def GETPAGE(url):
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.text
        return None
    except RequestException:
        return None
############################################


def GETINFO():
    driver = webdriver.Chrome("where the driver is in")
    driver.get("https://movie.douban.com/tag/#/?sort=U&range=0,10&tags=%E5%8F%A4%E8%A3%85,%E7%94%B5%E8%A7%86%E5%89%A7" )

    elem_name_lst = []
    elem_score_lst = []
    elem_douban_url_lst = []
    youtube_url_lst = []
    
    for elem_p in driver.find_elements_by_xpath('//div/a/p'):
        elem_a = elem_p.find_element_by_xpath('..')  
        elem_douban_url = elem_a.get_attribute('href')   

        elem_name = elem_p.text
        elem_name_len = len(elem_name)
        elem_score = elem_name[elem_name_len - 3:]
        elem_name = elem_name[: elem_name_len - 4]

        try:
            elem_score = float(elem_score)
            if elem_score >= 8.2:
                elem_name_lst.append(elem_name)
                elem_score_lst.append(elem_score)
                elem_douban_url_lst.append(elem_douban_url)
                youtube_url = MAKEURL(elem_name)  
                youtube_url_lst.append(youtube_url)

        except:
            pass
        
    return elem_name_lst, elem_score_lst, elem_douban_url_lst, youtube_url_lst

#generate youtube link url
def MAKEURL(elem_name):
    qs_l = [('search_query', elem_name)]
    d_qs = urllib.parse.urlencode(qs_l)
    youtube_url ="https://www.youtube.com/results?"
    youtube_url = youtube_url + d_qs
    return youtube_url

#make text file
def MAKEFILE(elem_name_lst, elem_score_lst, elem_douban_url_lst, youtube_url_lst):
    path = "where your file to write in"
    with open(path, mode = 'w', encoding = 'utf-8') as f:
        for i in range(len(elem_name_lst)):
            f.writelines("{} ".format(elem_name_lst[i]))
            f.writelines("豆瓣评价={}\n".format(elem_score_lst[i]))
            f.writelines("{}\n".format(elem_douban_url_lst[i]))
            f.writelines("{}\n".format(youtube_url_lst[i]))
            f.write("\n")

#send text to email
def SENDMAIL():    
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.to = 'destination adress'
    mail.cc = ''
    mail.bcc = ''
    mail.subject = '古装电视剧 评价8.2以上'
    mail.bodyFormat = 3


    path = "where your file to write in"
    with open(path, mode = 'r', encoding = 'utf-8') as f:
        file_content = f.read()
        print(file_content)

    mail.body = file_content

    mail.display(True)
    #mail.send()


#execute function
############################################

#dont need to use when you use selenium?
############################################
if __name__ == '__main__':    
    LOGIN("account-form-field-area-code-label")
   
    RegionNumber = "Your region number"
    LOGIN("GlobalPhoneItem-district")
    
    PhoneNumber = "Your phone number to login"
    LOGIN("account-form-input")

    LOGIN("account-form-input")

    #you have to push buttuon to get code and before you have to move bar correct


if __name__ == '__main__':
    headers = "your browser header"
    url = "#https://movie.douban.com/tag/#/?sort=U&range=0,10&tags=%E5%8F%A4%E8%A3%85,%E7%94%B5%E8%A7%86%E5%89%A7' #dramas, guzhuang"
    res = requests.get(url, headers = headers)
    html = GETPAGE(url)
############################################

if __name__ == '__main__':
    GETINFO()

if __name__ == '__main__':
    MAKEFILE()

if __name__ == '__main__':
    SENDMAIL()