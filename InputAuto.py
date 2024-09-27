import time
import pyautogui
import pyscreeze
import os
import pandas as pd

os.chdir('E:\\InputAuto')
df = pd.read_excel('dataset.xlsx')
#Anchor for click
anchor = ('anchor.png')

print(r'''
 /$$$$$$ /$$   /$$ /$$$$$$$  /$$   /$$ /$$$$$$$$/$$$$$$              /$$              
|_  $$_/| $$$ | $$| $$__  $$| $$  | $$|__  $$__/$$__  $$            | $$              
  | $$  | $$$$| $$| $$  \ $$| $$  | $$   | $$ | $$  \ $$ /$$   /$$ /$$$$$$    /$$$$$$ 
  | $$  | $$ $$ $$| $$$$$$$/| $$  | $$   | $$ | $$$$$$$$| $$  | $$|_  $$_/   /$$__  $$
  | $$  | $$  $$$$| $$____/ | $$  | $$   | $$ | $$__  $$| $$  | $$  | $$    | $$  \ $$
  | $$  | $$\  $$$| $$      | $$  | $$   | $$ | $$  | $$| $$  | $$  | $$ /$$| $$  | $$
 /$$$$$$| $$ \  $$| $$      |  $$$$$$/   | $$ | $$  | $$|  $$$$$$/  |  $$$$/|  $$$$$$/
|______/|__/  \__/|__/       \______/    |__/ |__/  |__/ \______/    \___/   \______/                                                                                                                                                                                                              
''')
time.sleep(1)
start = time.time()
def day_first():
    #OG 1 07:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    keyboard.press_and_release('alt+tab')
    time.sleep(1)
    pyautogui.locateOnScreen('p1og7.png')
    pyautogui.moveTo(p1og7)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 2]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OD 1 07:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    pyautogui.locateOnScreen('p1od1.png')
    pyautogui.moveTo(p1od1)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 3]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OD 2 07:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    pyautogui.locateOnScreen('p1od2.png')
    pyautogui.moveTo(p1od2)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OG 2 07:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    time.sleep(1)
    pyautogui.locateOnScreen('p2og7.png')
    pyautogui.moveTo(p2og7)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 5]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)    
    #Prod P1 07:00
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    time.sleep(1)
    pyautogui.locateOnScreen('p1prd7.png')
    pyautogui.moveTo(p1prd7)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[29:48, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod 2 07:00
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    time.sleep(1)
    pyautogui.locateOnScreen('p2prd7.png')
    pyautogui.moveTo(p2prd7)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P1 07:00
    pyautogui.press('pagedown', 2)
    time.sleep(1)
    pyautogui.locateOnScreen('p1wrn7.png')
    pyautogui.moveTo(p1wrn7)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[22:25, 11]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P2 07:00
    time.sleep(1)
    pyautogui.locateOnScreen('p2wrn7.png')
    pyautogui.moveTo(p2wrn7)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[28:31, 11]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
def day_second():
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    keyboard.press_and_release('alt+tab')
    pyautogui.click()
    pyautogui.press('pagedown')
    #Prod P1 11:00
    time.sleep(1)
    pyautogui.locateOnScreen('p1prd11.png')
    pyautogui.moveTo(p1prd11)
    pyautogui.moveRel(0, 10)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod 2 11:00
    pyautogui.press('enter')
    pyautogui.moveRel(0, -40)
    pyautogui.click()
    pyautogui.press('down', presses=3)
    time.sleep(1)
    pyautogui.locateOnScreen('p2prd11.png')
    pyautogui.moveTo(p2prd11)
    pyautogui.moveRel(0, 10)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P1 11:00
    pyautogui.press('pagedown')
    time.sleep(1)
    pyautogui.locateOnScreen('p1wrn11.png')
    pyautogui.moveTo(p1wrn11)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[22:26, 12]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P2 11:00
    time.sleep(1)
    pyautogui.locateOnScreen('p2wrn11.png')
    pyautogui.moveTo(p2wrn11)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[28:32, 12]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
def swing_first():
    #OG 1 15:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    keyboard.press_and_release('alt+tab')
    time.sleep(1)
    pyautogui.locateOnScreen('p1og15.png')
    pyautogui.moveTo(p1og15)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 2]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OG 2 15:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    time.sleep(1)
    pyautogui.locateOnScreen('p2og15.png')
    pyautogui.moveTo(p2og15)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 5]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OD 1 15:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    pyautogui.locateOnScreen('p2od1.png')
    pyautogui.moveTo(p2od1)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 6]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OD 2 15:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    pyautogui.locateOnScreen('p2od2.png')
    pyautogui.moveTo(p2od2)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)          
    #Prod P1 15:00
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    time.sleep(1)
    pyautogui.locateOnScreen('p1prd15.png')
    pyautogui.moveTo(p1prd15)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[29:48, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod 2 15:00
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    time.sleep(1)
    pyautogui.locateOnScreen('p2prd15.png')
    pyautogui.moveTo(p2prd15)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P1 15:00
    pyautogui.press('pagedown', 2)
    time.sleep(1)
    pyautogui.locateOnScreen('p1wrn15.png')
    pyautogui.moveTo(p1wrn15)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[22:25, 13]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P2 15:00
    time.sleep(1)
    pyautogui.locateOnScreen('p2wrn15.png')
    pyautogui.moveTo(p2wrn15)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[28:31, 13]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
def swing_second():
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    keyboard.press_and_release('alt+tab')
    time.sleep(1)
    pyautogui.click()
    pyautogui.press('pagedown')
    #Prod P1 19:00
    time.sleep(1)
    pyautogui.locateOnScreen('p1prd19.png')
    pyautogui.moveTo(p1prd19)
    pyautogui.moveRel(0, 10)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod 2 19:00
    pyautogui.press('enter')
    pyautogui.moveRel(0, -40)
    pyautogui.click()
    pyautogui.press('down', presses=3)
    time.sleep(1)
    pyautogui.locateOnScreen('p2prd19.png')
    pyautogui.moveTo(p2prd19)
    pyautogui.moveRel(0, 10)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P1 19:00
    pyautogui.press('pagedown')
    time.sleep(1)
    pyautogui.locateOnScreen('p1wrn19.png')
    pyautogui.moveTo(p1wrn19)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[22:26, 14]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P2 19:00
    time.sleep(1)
    pyautogui.locateOnScreen('p2wrn19.png')
    pyautogui.moveTo(p2wrn19)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[28:32, 14]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
def night_first():
    #OG 1 23:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    keyboard.press_and_release('alt+tab')
    time.sleep(1)
    pyautogui.locateOnScreen('p1og23.png')
    pyautogui.moveTo(p1og23)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 2]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #OG 2 23:00
    df = pd.read_excel('dataset.xlsx', sheet_name='OGOD')
    time.sleep(1)
    pyautogui.locateOnScreen('p2og23.png')
    pyautogui.moveTo(p2og23)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[32:41, 5]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod P1 23:00
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    time.sleep(1)
    pyautogui.locateOnScreen('p1prd23.png')
    pyautogui.moveTo(p1prd23)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[29:48, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod 2 23:00
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    time.sleep(1)
    pyautogui.locateOnScreen('p2prd23.png')
    pyautogui.moveTo(p2prd23)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P1 23:00
    pyautogui.press('pagedown', 2)
    time.sleep(1)
    pyautogui.locateOnScreen('p1wrn23.png')
    pyautogui.moveTo(p1wrn23)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[22:25, 15]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P2 23:00
    time.sleep(1)
    pyautogui.locateOnScreen('p2wrn23.png')
    pyautogui.moveTo(p2wrn23)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[28:31, 15]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
def night_second():
    df = pd.read_excel('dataset.xlsx', sheet_name='Fusion')
    keyboard.press_and_release('alt+tab')
    pyautogui.click()
    pyautogui.press('pagedown')
    #Prod P1 03:00
    time.sleep(1)
    pyautogui.locateOnScreen('p1prd3.png')
    pyautogui.moveTo(p1prd3)
    pyautogui.moveRel(0, 10)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 4]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Prod 2 03:00
    pyautogui.press('enter')
    pyautogui.moveRel(0, -40)
    pyautogui.click()
    pyautogui.press('down', presses=3)
    time.sleep(1)
    pyautogui.locateOnScreen('p2prd3.png')
    pyautogui.moveTo(p2prd3)
    pyautogui.moveRel(0, 10)
    pyautogui.click()
    data_to_input = df.iloc[29:49, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P1 03:00
    pyautogui.press('pagedown')
    time.sleep(1)
    pyautogui.locateOnScreen('p1wrn3.png')
    pyautogui.moveTo(p1wrn3)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[22:26, 16]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Warna P2 03:00
    time.sleep(1)
    pyautogui.locateOnScreen('p2wrn3.png')
    pyautogui.moveTo(p2wrn3)
    pyautogui.moveRel(0, 18)
    pyautogui.click()
    data_to_input = df.iloc[28:32, 16]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
def blending():
    df = pd.read_excel('dataset.xlsx', sheet_name='Blending')
    keyboard.press_and_release('ctrl+tab')
    #Blending 2
    time.sleep(1)
    pyautogui.locateOnScreen('bld2.png')
    pyautogui.moveTo(bld2)
    pyautogui.moveRel(0, 45)
    pyautogui.click()
    data_to_input = df.iloc[29:51, 4]
    for value in data_to_input:
        pyautogui.write(str(value))
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
    #Blending 3
    time.sleep(1)
    pyautogui.locateOnScreen('bld3.png')
    pyautogui.moveTo(bld3)
    pyautogui.moveRel(0, 45)
    pyautogui.click()
    data_to_input = df.iloc[29:51, 7]
    for value in data_to_input:
        pyautogui.write(str(value)) 
        pyautogui.press('enter')
        pyautogui.press('down')
        pyautogui.press('enter')   
        time.sleep(1)
consent = input('masukkan jam pengisian analisa: \n''ketik 1 dan tekan enter untuk jam 07:00\n''ketik 2 dan tekan enter untuk jam 11:00\n''ketik 3 dan tekan enter untuk jam 15:00\n''ketik 4 dan tekan enter untuk jam 19:00\n''ketik 5 dan tekan enter untuk jam 23:00\n''ketik 6 dan tekan enter untuk jam 03:00\n')
if consent == '1':
    day_first()
elif consent == '2':
    blend = input('apakah ada blending?(y/n): ')
    if blend == 'y':
        day_second()
        blending()
    else:
        day_second()
elif consent == '3':
    swing_first()
elif consent == '4':
    swing_second()
elif consent == '5':
    night_first()
elif consent == '6':
    night_second()
keyboard.press_and_release('alt+tab')
end = time.time()
print('Program berjalan selama: ', int(end-start), 'detik')
time.sleep(10)


