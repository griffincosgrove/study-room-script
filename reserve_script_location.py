import pyautogui as pg
import datetime as dt
import openpyxl
from openpyxl import Workbook

pg.PAUSE = 1
print(pg.position())

def login():
    pg.moveTo(x=106, y=189)
    pg.click()
    pg.typewrite("cosgrogc")
    pg.moveTo(x=102, y=214)
    pg.click()
    pg.typewrite("Fuckpike624$")
    pg.moveTo(x=235, y=233)
    pg.click()

def click_reserve():
    pg.moveTo(x=132, y=262)
    pg.click()


def weekday_reserve():
    pg.moveTo(x=281, y=299)
    pg.click()
    pg.moveTo(x=268, y=314)
    pg.click()


#login()
#click_reserve()
weekday_reserve()
