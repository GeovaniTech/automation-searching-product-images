import time

import numpy as np
import pandas as pd
import pyautogui
import matplotlib.pyplot as plt

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import *
from PIL import Image
from openpyxl.drawing.image import Image as ImgOpen
from skimage.transform import resize


sheet = pd.read_excel('Produtos.xlsx')
datas = sheet[['Marca', 'Produto', 'Cor']]

pyautogui.PAUSE = 1

row = 1
ant_row = 0
for value in datas.values:
    search = value[0] + ' ' + value[1] + ' ' + value[2]

    driver = webdriver.Chrome(r'C:\Users\Geovani Debastiani\.wdm\drivers\chromedriver\win32\95.0.4638.69\chromedriver.exe')
    driver.get('https://www.google.com.br/imghp?hl=pt-BR&tab=ri&ogbl')
    box = driver.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
    box.send_keys(search)
    box.send_keys(Keys.ENTER)

    try:
        driver.find_element_by_xpath('//*[@id="islrg"]/div[1]/div[1]/a[1]/div[1]/img').click()
        driver.find_element_by_xpath('//*[@id="islrg"]/div[1]/div[1]/a[1]/div[1]/img').screenshot(f'Images/{search}.png')

        ant_row = row

        row += 1
    except:
        if row == ant_row:
            row += 1
        pass

    wb = load_workbook('Produtos.xlsx')
    sheet = wb['Produtos']

    loc_img = r'C:\Users\Geovani Debastiani\Desktop\Development\Projects\Automação - Busca por imagens\Images\{}.png'.format(search)

    #im = plt.imread(loc_img)
    #res = resize(im, (300, 300))


    img = Image.open(loc_img)

    new_img = img.resize((187, 156))
    new_img.save(loc_img)

    img = ImgOpen(loc_img)

    sheet.add_image(img, f'D{row}')
    wb.save('Produtos.xlsx')

