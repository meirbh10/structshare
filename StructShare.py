

from argparse import _ActionsContainer, Action
from ast import List
import ctypes
import curses
from curses import KEY_DOWN
from curses.ascii import TAB
# from curses import key_down
import datetime
import hashlib
import os
import tkinter
from tkinter.tix import Select
from unittest import enterModuleContext
from urllib.parse import parse_qs, parse_qsl, urlparse
from xmlrpc.client import Boolean
import pyautogui
import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# from selenium.webdriver.chrome.service import Service as ChromeService
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

import pyautogui
import os
from datetime import datetime

from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import *

import pynput
from pynput.keyboard import Key, Controller

import openpyxl
import pyotp

from selenium.webdriver.common.keys import Keys

import keyboard

import requests

from selenium.webdriver.common.action_chains import ActionChains




ExcelRowNumber = 0

def setIndex(index):
  global ExcelRowNumber
  ExcelRowNumber = index
  return ExcelRowNumber

def getIndex():
  # print("ExcelRowNumber = ", ExcelRowNumber)
  return ExcelRowNumber


def SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook):
  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)
  workbook.save(FullPathForExcelReportFile)
  print('E N D   T E S T: StructShare\n')
  return True



def StructShare(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook):
  # Create a new workbook
  # workbook = openpyxl.Workbook()
  # Select the active worksheet
  # worksheet = openpyxl.Workbook.active
  # Headers for the Excel Report
  # worksheet['A1'] = 'Test Case Name'
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)
  worksheet['A' + ExcelRowNumber] = 'StructShare:'
  # worksheet['B1'] = 'Check Name'
  # worksheet['C1'] = 'Status'
  # worksheet['D1'] = 'Comment (Exception)'
  # worksheet['E1'] = 'The Screenshot File'



  windows = {}
  driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
  # driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

  wait = WebDriverWait(driver, 20)


  def wait_for_window(timeout=2):
      time.sleep(timeout)
      wh_now = driver.window_handles
      wh_then = windows['window_handles']
      if len(wh_now) > len(wh_then):
          return set(wh_now).difference(set(wh_then)).pop()


  url = 'http://the-internet.herokuapp.com'


  credentials = {
    'username': 'meirb',
    'password': '!Mbh202322Qa@'
  }

  NewPasswordcredentials = {
    'username': 'meirb',
    'password': '!Mbh202322Qa@!'
  }

  print('\nS T A R T   T E S T: StructShare')
  print('Open the website: ', url)
  time.sleep(3)
  driver.get(url)
  driver.maximize_window()
  driver.implicitly_wait(5)
  driver.set_script_timeout(5)

  
  # CHECK 1: A/B Testing
  try:
    worksheet['B' + ExcelRowNumber] = 'A/B Testing'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[1]/a'))).click()
    time.sleep(2)
    Message = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'example')))
    assert "Also known as split testing. This is a way in which businesses are able to simultaneously test and learn different versions of a page to see which text and/or functionality works best towards a desired outcome (e.g. a user action such as a click-through)." in Message.text
    print ("PASS: The Error Message ", Message.text, " APPEARS - As Expected!\n")
    worksheet['C' + ExcelRowNumber] = "PASS"
    worksheet['D' + ExcelRowNumber] = 'The A/B Test Control text appeared'
  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: The A/B Test Control textis NOT appears")
    worksheet['B' + ExcelRowNumber] = 'The A/B Test Control text appeared'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()




  # CHECK 2: Add/Remove Elements
  try:
    worksheet['B' + ExcelRowNumber] = 'Add/Remove Elements'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[2]/a'))).click()
    time.sleep(2)
    # Add
    button = driver.find_element(By.TAG_NAME, 'button')
    button.click()
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="elements"]/button')))
    print("1st button clicked")
    button.click()
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="elements"]/button[2]')))
    print("2nd button clicked")
    button.click()
    time.sleep(1)
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="elements"]/button[3]')))
    print("3rd button clicked")

    # Remove
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="elements"]/button[3]'))).click()
    time.sleep(1)
    print("Remove 3 button clicked")
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="elements"]/button[2]'))).click()
    time.sleep(1)
    print("Remove 2 button clicked")
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="elements"]/button[1]'))).click()
    time.sleep(1)
    print("Remove 1 button clicked")
  
    
    # list_of_elements = driver.find_elements(By.XPATH, "elements")
    # print("list_of_elements = ", list_of_elements)
    # i=0
    # for element in list_of_elements:
    #  i+=i+1
    #  element.click()
    #  print("Delete ", i, " clicked")
    
    
    worksheet['C' + ExcelRowNumber] = "PASS"
    worksheet['D' + ExcelRowNumber] = 'The Add/Remove Elements passed'
  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: The Add/Remove Elements failed")
    worksheet['B' + ExcelRowNumber] = 'The Add/Remove Elements failed'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()

  

  # CHECK 4: Broken Images
  try:
    worksheet['B' + ExcelRowNumber] = 'Broken Images'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[4]/a'))).click()
    time.sleep(2)
    # Find all images on the page
    images = driver.find_elements(By.TAG_NAME, 'img')
    # Validate each image's source URL
    IfThereIsBrokenImageFlag = "false"
    for image in images:
      src = image.get_attribute('src')
      response = requests.head(src)
      response_code = response.status_code
      print(f"Image source: {src}, Response code: {response_code}")
      response_code = str(response_code)
      print("A", {response_code}, "B")
      response_code = response_code.replace(" ", "")
      response_code = response_code.replace("{", "")
      response_code = response_code.replace("}", "")
      print("A", {response_code}, "B")
      if str(response_code) == "200":
        IfThereIsBrokenImageFlag = "true"

    if IfThereIsBrokenImageFlag == "true":
      print ("PASS: The CHECK 4: Broken Images - PASSED")
      worksheet['C' + ExcelRowNumber] = "PASS"
      worksheet['D' + ExcelRowNumber] = 'The CHECK 4: Broken Images PASSED'
    else:
      print ("FAILED: The CHECK 4: Broken Images - FAILED")
      worksheet['C' + ExcelRowNumber] = "FAILED"
      worksheet['D' + ExcelRowNumber] = 'The CHECK 4: Broken Images FAILED'
  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: The CHECK 4: Broken Images FAILED")
    worksheet['B' + ExcelRowNumber] = 'The CHECK 4: Broken Images FAILED'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()


  


  # CHECK 5: Challenging DOM - The buttons "delete" and "edit" didn't response so I skip it




  # CHECK 6: Checkboxes
  try:
    worksheet['B' + ExcelRowNumber] = 'Checkboxes'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[6]/a'))).click()
    time.sleep(2)
    # Check the 1st checkbox, Uncheck the 2nd checkbox - and validate it
    # Find all checkboxes on the page
    checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
    # Check the first two checkboxes
    i=1
    for checkbox in checkboxes:
      checkbox.click()
      if i==1:
        assert checkbox.is_selected(), f"{checkbox} is not checked"
        print("1st checkbox is checked")
      else:
         # Check if the 2nd checkbox is unchecked
        if not checkbox.is_selected():
          print("Checkbox 2 is unchecked")
          print ("PASS: The CHECK 6: Checkboxes - PASSED")
          worksheet['C' + ExcelRowNumber] = "PASS"
          worksheet['D' + ExcelRowNumber] = 'The CHECK 6: Checkboxes PASSED'
        else:
          print("Checkbox 2 is checked")
          print("FAIL: The CHECK 6: Checkboxes FAILED")
          worksheet['B' + ExcelRowNumber] = 'The CHECK 6: Checkboxes FAILED'
          worksheet['C' + ExcelRowNumber] = "FAIL"
      i=i+1

  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: The CHECK 6: Checkboxes FAILED")
    worksheet['B' + ExcelRowNumber] = 'The CHECK 6: Checkboxes FAILED'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()

  

# CHECK 7: Context Menu
  try:
    worksheet['B' + ExcelRowNumber] = 'Context Menu'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[7]/a'))).click()
    time.sleep(2)
    button_xpath = "//div[@id='hot-spot']"
    button = driver.find_element(By.XPATH, button_xpath)
    action_chains = ActionChains(driver)
    action_chains.context_click(button).perform()
    time.sleep(3)
    alert = driver.switch_to.alert
    time.sleep(3)
    alert.accept()
    pyautogui.press('esc')
    time.sleep(2)    
    print ("PASS: The CHECK 7: Context Menu - PASSED")
    worksheet['C' + ExcelRowNumber] = "PASS"
    worksheet['D' + ExcelRowNumber] = 'The CHECK 7: Context Menu PASSED'
  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: The CHECK 7: Context Menu FAILED")
    worksheet['B' + ExcelRowNumber] = 'The CHECK 7: Context Menu FAILED'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()


  


  # CHECK 10: Disappearing Elements
  try:
    worksheet['B' + ExcelRowNumber] = 'CHECK 10: Disappearing Elements'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[10]/a'))).click()
    time.sleep(2)
    # Drag squre 1 and drop it "on" squre 2
    Square1 = driver.find_element(By.ID, "column-a")
    Square2 = driver.find_element(By.ID, "column-b")
    action_chains = ActionChains(driver)
    action_chains.drag_and_drop(Square1, Square2).perform()
    print ("PASS: The CHECK 10: Disappearing Elements - PASSED")
    worksheet['C' + ExcelRowNumber] = "PASS"
    worksheet['D' + ExcelRowNumber] = 'The CHECK 10: Disappearing Elements PASSED'
  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: CHECK 10: Disappearing Elements FAILED")
    worksheet['B' + ExcelRowNumber] = 'CHECK 10: Disappearing Elements FAILED'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()


  
# CHECK 18: File Uploader
  try:
    worksheet['B' + ExcelRowNumber] = 'CHECK 18: File Uploader'
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/ul/li[18]/a'))).click()
    time.sleep(2)
    choose_file_button = driver.find_element(By.ID, "file-upload")
    path_to_type = r"C:\Users\meirb\AppData\Local\Programs\Python\Python311\StructShare\StructShare Python Scripts\Some Text File.txt"
    choose_file_button.send_keys(path_to_type)
    upload_button = driver.find_element(By.ID, "file-submit")
    upload_button.click()
    time.sleep(2)
    # Validate file uploaded
    element = driver.find_element(By.ID, "uploaded-files")
    text = element.text
    if "Some Text File.txt" in text:
      print ("PASS: The CHECK 18: File Uploader - PASSED")
      worksheet['C' + ExcelRowNumber] = "PASS"
      worksheet['D' + ExcelRowNumber] = 'The CHECK 18: File Uploader PASSED'
    else:
      print("FAIL: CHECK 18: File Uploader FAILED")
      worksheet['B' + ExcelRowNumber] = 'CHECK 18: File Uploader FAILED'
      worksheet['C' + ExcelRowNumber] = "FAIL"
  except Exception as ExceptionError:
    print("ExceptionError: \n", ExceptionError)
    print("FAIL: CHECK 18: File Uploader FAILED")
    worksheet['B' + ExcelRowNumber] = 'CHECK 18: File Uploader FAILED'
    worksheet['C' + ExcelRowNumber] = "FAIL"
    worksheet['D' + ExcelRowNumber] = str(ExceptionError)
    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    print("TimeStamp = ", TimeStamp)
    FullPathForScreenshot = os.path.join("C:/Users/meirb/AppData/Local/Programs/Python/Python311/StructShare/Screenshots/StructShare/StructShare_Screenshot_" + TimeStamp + ".png")
    print("Full Path For Screenshot = ", FullPathForScreenshot)
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save(FullPathForScreenshot)
    time.sleep(2)

    # Add the screenshot Link to the Excel file
    FullScreenshotPathForExcelReport = FullPathForScreenshot.replace("/", "\\")
    print("FullScreenshotPathForExcelReport (In the except) = ", FullScreenshotPathForExcelReport)
    worksheet['E' + ExcelRowNumber] = FullScreenshotPathForExcelReport

    Bool = SaveExcelAndCloseDriver(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    if Bool:
      driver.close()
      return ExcelRowNumber


  ExcelRowNumber = int(ExcelRowNumber)
  ExcelRowNumber = ExcelRowNumber + 1
  ExcelRowNumber = str(ExcelRowNumber)

  # Go back to the Main screen "Welcome to the-internet" (https://the-internet.herokuapp.com/)
  driver.back()

  workbook.save(FullPathForExcelReportFile)

  driver.close()

  print('E N D   T E S T: StructShare\n')

  return ExcelRowNumber