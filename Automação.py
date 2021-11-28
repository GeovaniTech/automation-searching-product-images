import time
import pandas as pd
import webbrowser
import pyautogui

from selenium import webdriver
from selenium.webdriver.common.keys import Keys


sheet = pd.read_excel('Produtos.xlsx')
datas = sheet[['Marca', 'Produto', 'Cor']]

pyautogui.PAUSE = 1

for value in datas.values:
    search = value[0] + ' ' + value[1] + ' ' + value[2]
    webbrowser.open('https://www.google.com.br/imghp?hl=pt-BR&tab=ri&ogbl')
    time.sleep(1)
    pyautogui.write(search)
    pyautogui.hotkey('Enter')


