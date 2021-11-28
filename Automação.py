import time
import pandas as pd
import webbrowser
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

sheet = pd.read_excel('Produtos.xlsx')
datas = sheet[['Marca', 'Produto', 'Cor']]

pyautogui.PAUSE = 1


for value in datas.values:
    search = value[0] + ' ' + value[1] + ' ' + value[2]
    driver = webdriver.Chrome(r'C:\Users\Geovani Debastiani\.wdm\drivers\chromedriver\win32\95.0.4638.69\chromedriver.exe')
    driver.get('https://www.google.com.br/imghp?hl=pt-BR&tab=ri&ogbl')
    time.sleep(1)
    box = driver.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
    box.send_keys(search)
    box.send_keys(Keys.ENTER)
    time.sleep(1)

    try:
        box = driver.find_element_by_xpath('//*[@id="islrg"]/div[1]/div[1]/a[1]/div[1]/img').screenshot(f'Images/{search}.png')
    except:
        pass



