#importing libraries
import selenium
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver import ActionChains
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from time import time
from time import sleep
import numpy as np
import base64
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import getpass
from tkinter import simpledialog



root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 350, bg = 'honeydew2', relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text='Star Application', bg = 'honeydew2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)



def treasury ():

	''' This function allows me to import the excel file comtaining the unique work item ID'''
    global items
    
    import_file_path = filedialog.askopenfilename()
    read_file = pd.read_excel(import_file_path , sheet_name = 'Abiola', header = None)
    items = read_file.iloc[:, 0].tolist()
    
browseButton_CSV = tk.Button(text="      Import Excel File     ", command=treasury, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_CSV)   
    

def MoveCRMS ():

	'''This function moves each work items from the general CRMS queue to the next queue for review and processed'''
    global items
    
    driver = webdriver.Chrome() #Put chromedriver in the same folder where your codes are
    driver.get("https:// ")
    user = ' ' 
    pswd = ' '
    #pswd = getpass.getpass(prompt='Password: ', stream=None) 
    username = driver.find_element_by_xpath('//*[@id="UserName"]')
    username.click()
    username.send_keys(user)
    password = driver.find_element_by_xpath('//*[@id="Password"]')
    password.click()
    password.send_keys(pswd)
    submit = driver.find_element_by_xpath('//*[@id="login-form"]/form/div[4]/button')
    submit.click()
    sleep(4)
    tasks_bar = driver.find_element_by_xpath('//*[@id="navbar-collapse-2"]/ul/li[3]/a')
    tasks_bar.click()
    
    data = []

    for data in items:
        searchbox = driver.find_element_by_xpath('//*[@id="tasks_filter"]/label/input')
        searchbox.click()
        sleep(2)
        searchbox.clear()
        searchbox.send_keys(data)
        sleep(2) 
        try:
            take_ownership = driver.find_elements_by_xpath('//*[@id="tasks"]/tbody/tr/td[12]/a')[0]   
            take_ownership.click()
            sleep(2)
            sleep(2)
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            sleep(2)
            select = driver.find_element_by_xpath('//*[@id="content"]/div/div/div/form/div/div[23]/task-review/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select').click()
            select = Select(driver.find_element_by_xpath('/html/body/section/div/div/div/div/div[1]/div/div[2]/div/div/div/form/div/div[23]/task-review/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select'))        
            sleep(1)
            select.select_by_value('Approved')
            selected = driver.find_elements_by_xpath('/html/body/section/div/div/div/div/div[1]/div/div[2]/div/div/div/form/div/div[26]/submit-review/input')[0].click()
            sleep(2)
            alert = driver.switch_to.alert
            alert.accept()
            print('{} has been treated'.format(data))
            sleep(1)
            driver.refresh()
            driver.get("https: ")          
        except NoSuchElementException:     
            driver.execute_script("window.history.go(-1)")     
        except IndexError as error:      
            pass     
        except ElementClickInterceptedException:
            driver.execute_script("window.history.go(-1)")
        except UnexpectedAlertPresentException:
            driver.execute_script("window.history.go(-1)") 
            driver.refresh()
    driver.close()

MoveButton_CSV = tk.Button(text="      Move Items From CRMS     ", command=MoveCRMS, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=MoveButton_CSV)



def PushTreasury ():

	'''This function moves each work items from the previous queue to the final destination where it would be cross-checked'''
    global items
    driver = webdriver.Chrome() #Put chromedriver in the same folder where your codes are
    driver.get("https: ")
    user = ' '
    pswd = ' '
    username = driver.find_element_by_xpath('//*[@id="UserName"]')
    username.click()
    username.send_keys(user)
    password = driver.find_element_by_xpath('//*[@id="Password"]')
    password.click()
    password.send_keys(pswd)
    submit = driver.find_element_by_xpath('//*[@id="login-form"]/form/div[4]/button')
    submit.click()
    sleep(4)
    tasks_bar = driver.find_element_by_xpath('//*[@id="navbar-collapse-2"]/ul/li[3]/a')
    tasks_bar.click()
    
    data = []
    for data in items:
        #searching for the work items we want using our list  
        searchbox = driver.find_element_by_xpath('//*[@id="tasks_filter"]/label/input')  
        searchbox.click() 
        sleep(2)
        searchbox.clear()
        searchbox.send_keys(data) 
        sleep(2)   
        try: 
            take_ownership = driver.find_elements_by_xpath('//*[@id="tasks"]/tbody/tr/td[12]/a')[0]      
            take_ownership.click()
            sleep(2)
            sleep(2)
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            sleep(2)
            select = driver.find_element_by_xpath('//*[@id="content"]/div/div/div/form/div/div[25]/task-review/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select').click()
            select = Select(driver.find_element_by_xpath('//*[@id="content"]/div/div/div/form/div/div[25]/task-review/div/div[2]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select'))        
            sleep(1)
            select.select_by_value('Approved')
            selected = driver.find_elements_by_xpath('/html/body/section/div/div/div/div/div[1]/div/div[2]/div/div/div/form/div/div[27]/submit-review/input')[0].click()
            sleep(2)
            alert = driver.switch_to.alert  
            alert.accept()      
            print('{} has been treated'.format(data))      
            sleep(3)
            driver.refresh() 
            driver.get("https: ")          
        except NoSuchElementException as nse:    
            driver.execute_script("window.history.go(-1)")     
        except IndexError as error:     
            pass       
        except ElementClickInterceptedException as intercept :        
            driver.execute_script("window.history.go(-1)")   
        except UnexpectedAlertPresentException as unexpected :       
            driver.execute_script("window.history.go(-1)")       
        except NoAlertPresentException:       
            driver.refresh()
            driver.refresh()

    logout =  driver.find_element_by_xpath('//*[@id="logoutForm"]/a')
    logout.click()
    driver.close() 
PushButton_CSV = tk.Button(text="      Push Items From Treasury     ", command=PushTreasury, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=PushButton_CSV)


def exitApplication():

	'''This function is used to exit the appliction and terminate all process'''
    MsgBox = tk.messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()
     
exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 280, window=exitButton)

root.mainloop()



