from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import time
service = Service(executable_path='./chromedriver-mac19')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
import string
import json

def remove(string):
    nm=string.replace(" ", "")
    nm=nm.split("\n")
    nm=nm[0]
    nm.replace(" ", "")
    return nm

# URL of the page to open
url = 'http://psd.bits-pilani.ac.in/Login.aspx'  # Replace this URL with the desired page URL

# Open the URL
driver.get(url)
userName= driver.find_element(By.ID, 'TxtEmail')
passWd= driver.find_element(By.ID, 'txtPass')
userName.send_keys("email-id")
passWd.send_keys("password")
time.sleep(2)
driver.find_element(By.ID,"Button1").click()
time.sleep(2)
driver.find_element(By.ID,"PSI").click()
time.sleep(2)
driver.find_element(By.ID,"StudentUserControl_stuStaPref").click()

time.sleep(2)
#LIst Import
from selenium.webdriver.common.action_chains import ActionChains
import openpyxl

# Load the Excel file to get the desired order
workbook = openpyxl.load_workbook('PS-Data-Copy.xlsx')
sheet = workbook.active
order_list = [cell.value for cell in sheet['A']] 
order_list1 = [cell.value for cell in sheet['C']] 
order_list2 = [cell.value for cell in sheet['B']] 
pslist=[]





# Find the sortable-nav element
sortable_nav = driver.find_element(By.ID,'sortable_nav')

# Get all items inside the sortable-nav
items = sortable_nav.find_elements(By.TAG_NAME,'li')

# Create a dictionary to store the order of items
item_dict = {item.text: item for item in items}

# firstoption=order_list1[0]+"-"+order_list[0]+","+order_list2[0]
# firstoption=remove(firstoption)
# key1=remove(key1).strip()
# print(len(firstoption),firstoption)
# print(len(key1),key1)
# print("T")
# print(key1 == remove(firstoption))
# print("T")
diction={}
for i in range(len(list(item_dict.keys()))):
    key_name = list(item_dict.keys())[i]
    key_name1=remove(key_name)
    
    print(key_name+"  T  ",key_name1)
    diction[key_name1]=item_dict[key_name]


with open("myfile.txt", 'w') as f:  
    for key, value in diction.items():  
        f.write('%s\n' % (key))


for i in range(len(order_list)):
    if order_list1[i]=="-":
        pslist.append(remove(order_list[i]+","+order_list2[i]))
        continue
    pslist.append(remove(order_list1[i]+"-"+order_list[i]+","+order_list2[i]))
# Filter the Excel order list to exclude items not found on the web page
filtered_order = [item_text for item_text in pslist if item_text in diction]

# Sort the items based on the order from the Excel file

sorted_items = [diction[item_text] for item_text in filtered_order]


# Perform drag and drop to reorder the items
actions = ActionChains(driver)
for item in sorted_items[::-1]:
    actions.drag_and_drop(item, sortable_nav).perform()

# Close the browser window after some time (for demonstration purposes)

while True:
    time.sleep(10)  # Wait for 5 seconds
