import pyautogui as py
import time as t
import pandas as pd
from pynput.keyboard import Controller
keyboard = Controller()

def down_f(i):
  count = 0
  while count < i:
    py.press('down')
    count += 1

def open_menu_f():  
    py.keyDown('alt')
    py.press('c')
    py.keyUp('alt')

def new_item_f(): 
  py.keyDown('ctrl')
  py.press('insert')
  py.keyUp('ctrl')

def cancel_f():
  py.keyDown('ctrl')
  py.press('backspace')
  py.keyUp('ctrl')

def close_f():
  py.keyDown('alt')
  py.press('f4')
  py.keyUp('alt')

table = pd.read_excel(f".\db\group.xlsx")

t.sleep(3)

open_menu_f()
down_f(2)
py.press('enter')

for i, row in table.iterrows():
  cod = row['COD']
  group = row['GRUPO']

  new_item_f()
  py.write(str(cod))
  py.press('tab')
  t.sleep(1)
  keyboard.type(str(group))

py.press('enter')
cancel_f()
close_f()