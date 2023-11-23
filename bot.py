import pyautogui as py
import time as t

def down_f(i):
    count = 0
    while count < i:      
      py.press('down')      
      t.sleep(1)
      count += 1
    

def main():
    t.sleep(5)
    py.press('alt')
    py.press('c')
    down_f(3)
    py.press('enter')

if __name__ == '__main__':
    main()  