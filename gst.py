import time
from selenium import webdriver
import warnings
import pyautogui as a
import pandas as pd
import glob
import os
import subprocess
import sys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
path = 'D:\\Akshay\\AHK\\files\\temp_data'
warnings.filterwarnings("ignore", category=DeprecationWarning)
options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : path}
options.add_experimental_option('prefs', prefs)
options.add_argument('log-level=3')
options.add_argument('disable-gpu')
options.add_argument('disable-infobars')
options.add_argument('disable-extensions')
options.add_argument('disable-dev-shm-usage')
options.headless = True
options.add_experimental_option('excludeSwitches', ['enable-logging'])
FailSafeException=True
oldxl = glob.glob(path+'\\*.xlsx')
for f in oldxl:
    os.remove(f)
w = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))
start_time = time.time()
f = open(path+"\\gst.txt", "r")
gstt=f.read()
gst=gstt.strip()
f.close()
#if os.path.exists(path+"\\gst.txt"):
  #os.remove(path+"\\gst.txt")
w.get('https://my.gstzen.in/p/gstin-validator/')
WebDriverWait(w, 1).until(EC.presence_of_element_located((By.NAME,'text'))).send_keys(gst)
WebDriverWait(w, 1).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div/div[4]/div/form/div[2]/div/button'))).click()
WebDriverWait(w, 1).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div/div[2]/div[1]/h5/a'))).click()
time.sleep(0.75)
print("Details downloaded!")

sys.exit()
w.close()
w.quit()











WebDriverWait(w, 60).until(EC.presence_of_element_located((By.XPATH,'//*[@id="lottable"]/div[2]/div[2]/div/div[2]/p[2]')))

status = w.find_elements(By.XPATH,'//*[@id="lottable"]/div[2]/div[2]/div/div[2]/p[2]')[0].text
type = w.find_elements(By.XPATH,'//*[@id="lottable"]/div[2]/div[2]/div/div[3]/p[2]')[0].text
addr = w.find_elements(By.XPATH,'//*[@id="lottable"]/div[2]/div[3]/div/div[3]/p[2]')[0].text
print(addr)

