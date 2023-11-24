import pyautogui as py
import time as t
import pandas as pd

def down_f(i):
    count = 0
    while count < i:
      py.press('down')      
      t.sleep(1)
      count += 1

def tab_f(i):
    count = 0
    while count < i:
      py.press('tab')      
      t.sleep(1)
      count += 1

def save_f():    
   py.keyDown('ctrl')
   py.press('s')
   py.keyUp('ctrl')
   t.sleep(1)

def finished_f():
   py.keyDown('alt')
   py.press('f4')
   py.keyUp('alt')
   t.sleep(1)   

name_table = str(input('INSIRA O NOME DA TABELA: '))
table = pd.read_excel(f"{name_table}.xlsx", dtype={'COD': str})

t.sleep(5)

for i, row in table.iterrows():
    code = row["COD"]
    description = row["DESCRICAO"]
    unit = row["UNIDADE"]
    cost_price = row["PRECO_CUSTO"]
    sale_price = row["PRECO_VENDA"]
    manufacturer = row["FABRICANTE"]
    group = row["GRUPO"]
    stock = row["ESTOQUE"]
    min_stock = row["ESTOQUE_MIN"]
    
    py.press('alt')
    py.press('c')
    down_f(3)
    py.press('enter')
    py.write(str(code))
    py.press('tab')
    down_f(1)
    tab_f(3)
    py.write(str(description))
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
    save_f()
    py.press('alt')
    py.press('r')
    down_f(4)
    finished_f()
    py.press('s')