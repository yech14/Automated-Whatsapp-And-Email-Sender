
# Project Automated Whatsapp Sender
# Sendings messages to a list of numbers automatically

## To Do List:
# -


import pandas as pd
import time
import urllib.parse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def make_driver():   # פונקציית הגדרת הדרייבר והוספת אפשרויות לדפדפן
    PATH = "C:\Program Files (x86)\chromedriver.exe"  # מיקום הפאטצ'
    service = Service(PATH)  # הפעלת הפאטצ'
    options = webdriver.ChromeOptions()  # כדי שהחלון של הדפדפן לא ייסגר ישר
    options.add_experimental_option("detach", True)  # כדי שהחלון של הדפדפן לא ייסגר ישר
    return webdriver.Chrome(service=service, options=options)  # הגדרת הדרייבר שאיתו נשתמש



# Load the data frame
df = pd.read_excel('C:/Users/.../names_phones.xlsx')
phone_numbers = list(df['phone number:'])
names = list(df['name:'])


# Make subgroup for test
names1 = names[10:100]
phone_numbers1 = phone_numbers[10:100]

driver = make_driver()
driver.get('https://web.whatsapp.com')
time.sleep(25)

for i, name in enumerate(names1):
    if name not in []:
        try:
            main_window = driver.current_window_handle  # שומר את העמוד הנוכחי
            driver.execute_script("window.open('');")  # פותח טאב ריק חדש
            driver.switch_to.window(driver.window_handles[-1])  # מעביר את הדריביר לעמוד החדש

            phone_number = '+972' + phone_numbers1[i][1:].replace('-', '')

            message = f"""Hello  {name}, 
Type your message here..
"""

            encoded_message = urllib.parse.quote(message)
            url = f"https://web.whatsapp.com/send?phone={phone_number}&text={encoded_message}&app_absent=0"

            driver.get(url)
            wait = WebDriverWait(driver, 50)
            send_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="compose-btn-send"]')))
            # Click the send button
            send_button.click()
            time.sleep(5)
            # Close the Tab
            driver.close()
            driver.switch_to.window(main_window)   # מחזיר את הדרייבר לחלון המקורי הראשון
            time.sleep(5)
        except Exception as e:
            print(f'Failed to send message to {name}')
    else:
        print(f"The whatsapp app didn't send message to {name}, because i don't want to send him message")


driver.quit()




