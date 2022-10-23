import webbrowser
import openpyxl
import os,time
import pyautogui as py
from math import ceil

wb = openpyxl.load_workbook("li.xlsx")
sheet = wb['Table 2']
listEntreprise = []
for row in sheet.iter_rows():
    for cell in row:
        listEntreprise.append(cell.value.split('  ')[0])

#print(listEntreprise)

for entreprise in listEntreprise:
    webbrowser.open("https://google.com")
    time.sleep(3)
    py.write(entreprise, interval=0.1)
    time.sleep(3)
    py.press('return')

    time.sleep(3)

    im = py.screenshot('img/'+entreprise+'.png',region=(800,300, 600, 400))

    time.sleep(3)

    py.hotkey('ctrl', 'w')

    time.sleep(3)