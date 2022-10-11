#from tkinter.tix import Select
from dataclasses import dataclass
from pyexpat import native_encoding
from re import T
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from turtle import onclick
#from selenium.webdriver.support import expected_conditions as EC
from screens import tirar_print
import schedule
import time, os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date, timedelta, datetime
import pyperclip
import pyautogui
from scrape2 import email2, email3
pyautogui.FAILSAFE = False



def email():
  yesterday = (date.today() - timedelta(days=1)).strftime('%d/%m/%Y')
  options = Options()
  #options.add_argument("--headless")
  options.add_experimental_option("excludeSwitches", ["enable-logging"]) #Eliminar erro Bluetooth adapter
  options.add_experimental_option("detach", True)
  navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) #Setando navegador
  
  #navegador.minimize_window()
  navegador.get("http://192.168.0.2/")
  time.sleep(2)

  user = navegador.find_element(By.ID, 'username') 
  user.send_keys("rafaelvilela@eficazcob.com.br")

  passw = navegador.find_element(By.ID, 'password') 
  passw.send_keys("AAaa11**")

  navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/form/table/tbody/tr[3]/td[2]/input[2]').click()
  time.sleep(2)

  busca = navegador.find_element(By.XPATH, '//*[@id="zi_search_inputfield"]') 
  busca.send_keys("no-reply@openlabs.com.br")
  busca.send_keys(Keys.SPACE)
  busca.send_keys(yesterday)
  navegador.find_element(By.XPATH, '//*[@id="zb__Search__SEARCH_left_icon"]/div').click()
  time.sleep(2)
  email_click = navegador.find_element(By.XPATH, "//*[starts-with(@id, 'zlif__CLV-SR-1')]") #selec primeiro email
  email_click.click()
  time.sleep(2)
  navegador.switch_to.frame(navegador.find_element(By.TAG_NAME, 'iframe'))

  table_id = navegador.find_element(By.XPATH, '//*[@id="tb01"]')
  ultima_linha = table_id.text.splitlines()[-1]
  print(ultima_linha[6:])
  string_final = ultima_linha[6:]
  #print (table_id.text.splitlines()[-1])

  pyperclip.copy(string_final)
  pyperclip.paste()
  navegador.quit()
  os.startfile('Daily.xlsm')
  time.sleep(4)
  
  pyautogui.click(x=1149, y=547)
  pyautogui.hotkey('ctrl', 't')
  time.sleep(1)
  pyautogui.press('enter')
  time.sleep(1)
  pyautogui.click(x=1149, y=547)
  tirar_print()
  time.sleep(1)

  email2()
  
def email_sabado():
  yesterday = (date.today() - timedelta(days=2)).strftime('%d/%m/%Y')
  options = Options()
  #options.add_argument("--headless")
  options.add_experimental_option("excludeSwitches", ["enable-logging"]) #Eliminar erro Bluetooth adapter
  options.add_experimental_option("detach", True)
  navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) #Setando navegador
  
  #navegador.minimize_window()
  navegador.get("http://192.168.0.2/")
  time.sleep(2)

  user = navegador.find_element(By.ID, 'username') 
  user.send_keys("rafaelvilela@eficazcob.com.br")

  passw = navegador.find_element(By.ID, 'password') 
  passw.send_keys("AAaa11**")

  navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/form/table/tbody/tr[3]/td[2]/input[2]').click()
  time.sleep(2)

  busca = navegador.find_element(By.XPATH, '//*[@id="zi_search_inputfield"]') 
  busca.send_keys("no-reply@openlabs.com.br")
  busca.send_keys(Keys.SPACE)
  busca.send_keys(yesterday)
  navegador.find_element(By.XPATH, '//*[@id="zb__Search__SEARCH_left_icon"]/div').click()
  time.sleep(2)
  email_click = navegador.find_element(By.XPATH, "//*[starts-with(@id, 'zlif__CLV-SR-1')]") #selec primeiro email
  email_click.click()
  time.sleep(2)
  navegador.switch_to.frame(navegador.find_element(By.TAG_NAME, 'iframe'))

  table_id = navegador.find_element(By.XPATH, '//*[@id="tb01"]')
  ultima_linha = table_id.text.splitlines()[-1]
  print(ultima_linha[6:])
  string_final = ultima_linha[6:]
  #print (table_id.text.splitlines()[-1])

  pyperclip.copy(string_final)
  pyperclip.paste()
  navegador.quit()
  os.startfile('Daily.xlsm')
  time.sleep(4)
  
  pyautogui.click(x=1149, y=547)
  pyautogui.hotkey('ctrl', 't')
  time.sleep(1)
  pyautogui.press('enter')
  time.sleep(1)
  pyautogui.click(x=1149, y=547)
  tirar_print()
  time.sleep(1)

  email3()

def email_sab_sexta():
  sexta = (date.today() - timedelta(days=3)).strftime('%d/%m/%Y')
  options = Options()
  #options.add_argument("--headless")
  options.add_experimental_option("excludeSwitches", ["enable-logging"]) #Eliminar erro Bluetooth adapter
  options.add_experimental_option("detach", True)
  navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) #Setando navegador
  
  #navegador.minimize_window()
  navegador.get("http://192.168.0.2/")
  time.sleep(2)

  user = navegador.find_element(By.ID, 'username') 
  user.send_keys("rafaelvilela@eficazcob.com.br")

  passw = navegador.find_element(By.ID, 'password') 
  passw.send_keys("AAaa11**")

  navegador.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/form/table/tbody/tr[3]/td[2]/input[2]').click()
  time.sleep(2)

  busca = navegador.find_element(By.XPATH, '//*[@id="zi_search_inputfield"]') 
  busca.send_keys("no-reply@openlabs.com.br")
  busca.send_keys(Keys.SPACE)
  busca.send_keys(sexta)
  navegador.find_element(By.XPATH, '//*[@id="zb__Search__SEARCH_left_icon"]/div').click()
  time.sleep(2)
  email_click = navegador.find_element(By.XPATH, "//*[starts-with(@id, 'zlif__CLV-SR-1')]") #selec primeiro email######
  email_click.click()
  time.sleep(2)
  navegador.switch_to.frame(navegador.find_element(By.TAG_NAME, 'iframe'))

  table_id = navegador.find_element(By.XPATH, '//*[@id="tb01"]')
  ultima_linha = table_id.text.splitlines()[-1]
  print(ultima_linha[6:])
  string_final = ultima_linha[6:]
  #print (table_id.text.splitlines()[-1])

  pyperclip.copy(string_final)
  pyperclip.paste()
  navegador.quit()
  os.startfile('Daily.xlsm')
  time.sleep(4)
  
  pyautogui.click(x=1149, y=547)
  pyautogui.hotkey('ctrl', 't')
  time.sleep(1)
  pyautogui.press('enter')
  time.sleep(1)
  pyautogui.click(x=1149, y=547)
  pyautogui.hotkey('alt', 'f4')
  pyautogui.press('enter')
  time.sleep(1)
  
  
# position = pyautogui.position()
# print(position)



sim = "s"
nao = "n"
sabado = "x"
sab_sex = "y"

while True:
    word = input("Deseja começar?(S/N)(X=Sabado/Y=Sab_Sexta): ").lower()
    if word == sim:
        email()
        sys.exit()
    if word == nao:
        break
    if word == sabado:
        email_sabado()
        sys.exit()
    if word == sab_sex:
        email_sab_sexta()
        email_sabado()
        sys.exit()
    



















