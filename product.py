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

def tab_f(i):
    count = 0
    while count < i:
        py.press('tab')
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

def create_item_f():
    py.keyDown('ctrl')
    py.press('insert')
    py.keyUp('ctrl')
    t.sleep(1)

def save_f():    
    py.keyDown('ctrl')
    py.press('s')
    py.keyUp('ctrl')

# table = pd.read_excel(f".\db\product.xlsx")
groups_input = input("Insira os grupos separados por vírgulas: ")
groups = groups_input.split(',')


t.sleep(3)

open_menu_f()
down_f(3)
py.press('enter')

for group in groups:
  group = group.strip()
  table = pd.read_excel(f".\db\product.xlsx", sheet_name=f"{group}")
  for i, row in table.iterrows():
        cod = row["COD"]
        description = row["DESCRICAO"]
        unit = row["UNIDADE"]
        cost_price = row["PRECO_CUSTO"]
        sale_price = row["PRECO_VENDA"]
        manufacturer = row["FABRICANTE"]    
        stock = row["ESTOQUE"]
        min_stock = row["ESTOQUE_MIN"]       
        
        create_item_f()
        py.write(str(cod))
        tab_f(2)
        down_f(1)
        tab_f(3)        
        keyboard.type(str(description))
        tab_f(1)
        py.write(str(unit))
        tab_f(2)
        py.write(str(cost_price))
        tab_f(2)
        py.write(str(sale_price))
        tab_f(3)
        py.write(str(manufacturer))
        tab_f(1)
        py.write(str(group))
        tab_f(1)
        py.write(str(stock))
        tab_f(1)
        py.write(str(min_stock))
        t.sleep(1)
        save_f()
        t.sleep(1)


py.press('enter')
cancel_f()
close_f()