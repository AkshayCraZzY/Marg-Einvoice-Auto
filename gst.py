import time
from selenium import webdriver
import warnings
import glob
import os
import sys
import psutil
import csv
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

warnings.filterwarnings("ignore", category=DeprecationWarning)
options = webdriver.ChromeOptions()
path='D:\\Akshay\\AHK\\files\\temp_data'
prefs = {'download.default_directory' : path,'profile.managed_default_content_settings.images' : 2}
options.add_experimental_option('prefs', prefs)
options.add_argument('log-level=3')
options.add_argument('disable-gpu')
options.add_argument('disable-infobars')
options.add_argument('disable-extensions')
options.add_argument('disable-dev-shm-usage')
options.add_argument('--headless')
options.add_argument('window-size=1920x1080')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
FailSafeException=True
header=['GST','PINCODE','LOCATION','ADDRESS','STATUS']
count=0
#oldcsv = glob.glob(path+'\\*.csv')
#for f in oldxl:
    #os.remove(f)
w = webdriver.Chrome(options=options,service=Service(ChromeDriverManager().install()))
w.get('https://www.mastersindia.co/gst-number-search-and-gstin-verification/')
w.maximize_window()

def checksum(gst_check):
    check = gst_check[-1]
    gst_check = gst_check[:-1]
    l = [int(c) if c.isdigit() else ord(c)-55 for c in gst_check]
    l = [val*(ind % 2 + 1) for (ind, val) in list(enumerate(l))]
    l = [(int(x/36) + x%36) for x in l]
    csum = (36 - sum(l)%36)
    csum = str(csum) if (csum < 10) else chr(csum + 55)
    return True if (check == csum) else False

def main():
    global count
    while not os.path.exists(path+"\\gst.txt"):
      time.sleep(1)
      if "AutoHotkey.exe" in (i.name() for i in psutil.process_iter()):
        count += 1
        print(count)
      else:
        w.quit()
        sys.exit()
      
    if os.path.exists(path+"\\gst.txt"):
      f = open(path+"\\gst.txt", "r")
      gstt=f.read()
      gst=gstt.strip()
      f.close()
      if not checksum(gst):
        with open(path+'\\data.csv', 'w', encoding='UTF8',newline='') as f:
          writer = csv.writer(f)
          data=['INVALID','CHECKSUM']
          writer.writerow(data)
          os.remove(path+"\\gst.txt")
        main()
    else:
      print("ERROR")
    WebDriverWait(w, 1).until(EC.presence_of_element_located((By.XPATH,"//input[@placeholder='Search by GST Number']"))).send_keys(gst)
    time.sleep(.88)
    WebDriverWait(w, 1).until(EC.presence_of_element_located((By.XPATH,"//button[text()='Search']"))).click()
    try:
      gst = WebDriverWait(w, 3).until(EC.presence_of_element_located((By.XPATH,"//*[@id='panel-0']/div/table/tbody/tr[1]/td")))
    except TimeoutException as ex:
      print("Exception has been thrown.")
      #os.remove(path+"\\gst.txt")
      WebDriverWait(w, 1).until(EC.presence_of_element_located((By.XPATH,"//input[@placeholder='Search by GST Number']"))).clear()
      with open(path+'\\data.csv', 'w', encoding='UTF8',newline='') as f:
        writer = csv.writer(f)
        data=['INVALID','WEBSITE']
        writer.writerow(data)
        os.remove(path+"\\gst.txt")
      
      main()

    addr = WebDriverWait(w, 3).until(EC.presence_of_element_located((By.XPATH,"//*[@id='panel-0']/div/table/tbody/tr[3]/td")))
    state = WebDriverWait(w, 3).until(EC.presence_of_element_located((By.XPATH,"//*[@id='panel-0']/div/table/tbody/tr[5]/td")))
    centre = WebDriverWait(w, 3).until(EC.presence_of_element_located((By.XPATH,"//*[@id='panel-0']/div/table/tbody/tr[6]/td")))
    status = WebDriverWait(w, 3).until(EC.presence_of_element_located((By.XPATH,"//*[@id='panel-0']/div/table/tbody/tr[10]/td")))
    jurisdiction= ''.join([i for i in state.text if not i.isdigit()]).replace('_', '')
    print(jurisdiction)
    data=[gst.text,addr.text.rsplit(',', 1)[1].strip(),centre.text.replace('RANGE', '').strip(),','.join(addr.text.split(',')[:-2]).replace('floor-', ''),status.text] #List[gstno,addr,pin,location,status]
    print(data)
    with open(path+'\\data.csv', 'w', encoding='UTF8',newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerow(data)
    os.remove(path+"\\gst.txt")
    WebDriverWait(w, 1).until(EC.presence_of_element_located((By.XPATH,"//input[@placeholder='Search by GST Number']"))).clear()
    main()

main()

