import pyautogui as py
import time as t
import pandas as pd

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

def save_f():    
   py.keyDown('ctrl')
   py.press('s')
   py.keyUp('ctrl')

def finished_f():
   py.keyDown('alt')
   py.press('f4')
   py.keyUp('alt')
   t.sleep(1)   

name_table = str(input('INSIRA O NOME DA TABELA: '))
table = pd.read_excel(f"{name_table}.xlsx", dtype={'COD': str})

t.sleep(5)

py.keyDown('alt')
py.press('c')
py.keyUp('alt')
down_f(3)
py.press('enter')

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
    
    py.write(str(code))
    tab_f(2)
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
    t.sleep(1)
    save_f()
    t.sleep(1)    
    py.keyDown('alt')
    py.press('r')
    py.keyUp('alt')
    down_f(4)
    py.press('enter')
    t.sleep(1)

finished_f()
py.press('n')