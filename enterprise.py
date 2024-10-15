import pyautogui as py
import time as t
import pandas as pd
from pynput.keyboard import Controller
keyboard = Controller()

table = pd.read_excel(f".\db\enterprise.xlsx")

t.sleep(3)

py.keyDown('alt')
py.press('c')
py.keyUp('alt')
py.press('down')
py.press('enter')

for i, row in table.iterrows():
  cod = row['COD']
  name = row['NOME']
  cnpj = row['CNPJ']

  py.keyDown('ctrl')
  py.press('insert')
  py.keyUp('ctrl')
  py.write(str(cod))
  py.press('tab')
  t.sleep(1)
  keyboard.type(str(name))
  # py.write(str(name))
  py.press('tab')
  t.sleep(1)
  py.write(str(cnpj))

py.press('enter')
py.keyDown('ctrl')
py.press('backspace')
py.keyUp('ctrl')
py.keyDown('alt')
py.press('f4')
py.keyUp('alt')