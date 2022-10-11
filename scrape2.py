from dataclasses import dataclass
from pyexpat import native_encoding
from re import T
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from turtle import onclick
import time, pyautogui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import date, timedelta, datetime


def email2():
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

  navegador.find_element(By.XPATH, '//*[@id="zb__NEW_MENU_title"]').click()
  time.sleep(1)

  #navegador.find_element(By.XPATH, '//*[@id="zti__main_Mail__3867_textCell"]').click()

  para = navegador.find_element(By.XPATH, '//*[@id="zv__COMPOSE-1_to_control"]')
  para.send_keys("coordenadores@eficazcob.com.br")
  pyautogui.press('space')

  cc = navegador.find_element(By.XPATH, '//*[@id="zv__COMPOSE-1_cc_control"]')
  cc.send_keys("monitoria@eficazcob.com.br")
  pyautogui.press('space')
  cc.send_keys("monique@eficazcob.com.br")
  pyautogui.press('space')
  cc.send_keys("discadores@eficazcob.com.br")
  pyautogui.press('space')
  ass = navegador.find_element(By.XPATH, '//*[@id="zv__COMPOSE-1_subject_control"]')
  ass.send_keys("Relatório Receptivos Tim")

  navegador.switch_to.frame("ZmHtmlEditor1_body_ifr")

  corpo = navegador.find_element(By.XPATH, '//*[@id="tinymce"]')
  corpo.click()
  pyautogui.hotkey('ctrl', 'a')
  pyautogui.hotkey('backspace')
  corpo.send_keys("Boa tarde! Segue consolidado dos receptivos da Tim referente ao dia de ontem:")
  pyautogui.hotkey('enter')
  pyautogui.hotkey('enter')
  pyautogui.hotkey('ctrl', 'v')
  #corpo.send_keys("ola")

def email3():
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

  navegador.find_element(By.XPATH, '//*[@id="zb__NEW_MENU_title"]').click()
  time.sleep(1)

  #navegador.find_element(By.XPATH, '//*[@id="zti__main_Mail__3867_textCell"]').click()

  para = navegador.find_element(By.XPATH, '//*[@id="zv__COMPOSE-1_to_control"]')
  para.send_keys("coordenadores@eficazcob.com.br")
  pyautogui.press('space')

  cc = navegador.find_element(By.XPATH, '//*[@id="zv__COMPOSE-1_cc_control"]')
  cc.send_keys("monitoria@eficazcob.com.br")
  pyautogui.press('space')
  cc.send_keys("monique@eficazcob.com.br")
  pyautogui.press('space')
  cc.send_keys("discadores@eficazcob.com.br")
  pyautogui.press('space')
  ass = navegador.find_element(By.XPATH, '//*[@id="zv__COMPOSE-1_subject_control"]')
  ass.send_keys("Relatório Receptivos Tim")

  navegador.switch_to.frame("ZmHtmlEditor1_body_ifr")

  corpo = navegador.find_element(By.XPATH, '//*[@id="tinymce"]')
  corpo.click()
  pyautogui.hotkey('ctrl', 'a')
  pyautogui.hotkey('backspace')
  corpo.send_keys("Boa tarde! Segue consolidado dos receptivos da Tim referente a semana passada:")
  pyautogui.hotkey('enter')
  pyautogui.hotkey('enter')
  pyautogui.hotkey('ctrl', 'v')
  #corpo.send_keys("ola")

