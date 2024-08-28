import os
import time
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

load_dotenv()

class App():
    def __init__(self):
        self.msg = None
        self.data = None
        self.email_col_index = None
        self.name_col_index = None
        self.requires_password = None
        self.otc = None
        self.requires_session_click = None
        self.driver = None

    def run(self):
        self.print_title()
        self.load_excel_file()
        self.load_email_msg()
        self.setup_fields()
        self.init_selenium()
        self.login_outlook()
        self.send_emails()
        
        self.driver.close()

    def print_title(self):
        os.system('cls')
        print('OUTLOOK AUTO-SENDER V1.1')
        print('========================\n')

    def load_excel_file(self):
        filepath = input('Drag and drop an Excel file and press [Enter]:\n')
        self.data = pd.read_excel(filepath)

    def load_email_msg(self):
        with open('email_content.txt', 'r', encoding='utf-8') as f:
            self.msg = f.read()

    def setup_fields(self):
        print('\nColumns detected:')
        [print(f'{col_index + 1}) {col_name}') for col_index, col_name in enumerate(self.data.columns)]

        try:
            self.name_col_index = int(input('\nWhich column contains the first names?: ')) - 1
            self.email_col_index = int(input('Which column contains the emails?: ')) - 1
        except:
            print('\nIt seems that you haven\'t wrote a number.')
            print('[-] Exiting...')
            time.sleep(3)
            exit()

        print(f'\n> "Hi, {self.data.iloc[1, self.name_col_index]}! Your email address is {self.data.iloc[0, self.email_col_index]}"')
        confirmation_data = input('\nDoes that look right? (y/n): ')

        if confirmation_data != 'y':
            print('\n[-] Exiting...')
            time.sleep(3)
            exit()

        confirmation_passwd = input('\nDo you need to enter you password to access your email? (y/n): ')
        
        if confirmation_passwd == 'y':
            self.requires_password = True

        confirmation_2fa = input('\nDo you need to enter a 2FA code from an authentication app? (y/n): ')
        
        if confirmation_2fa == 'y':
            self.otc = input('Enter the 2FA code: ')

        confirmation_save_session = input('\nDoes the website ask you if you want to save your session? (y/n): ')
        
        if confirmation_save_session == 'y':
            self.requires_session_click = True

    def init_selenium(self):
        print('\n[+] Starting automation...')

        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        self.driver = webdriver.Chrome(options=options)
        
    def login_outlook(self):
        self.driver.get(os.getenv('OUTLOOK_URL'))

        time.sleep(5)

        login_user = self.driver.find_element(By.NAME, 'loginfmt')
        login_user.send_keys(os.getenv('EMAIL_ADDRESS'))
        login_user.send_keys(Keys.RETURN)

        time.sleep(5)

        if self.requires_password == True:
            login_password = self.driver.find_element(By.NAME, 'passwd')
            login_password.send_keys(os.getenv('EMAIL_PASSWORD'))
            login_password.send_keys(Keys.RETURN)

            time.sleep(5)
        else:
            time.sleep(12)

        if self.otc != None:
            login_otc = self.driver.find_element(By.NAME, 'otc')
            login_otc.send_keys(self.otc)
            login_otc.send_keys(Keys.RETURN)

            time.sleep(5)

        if self.requires_session_click == True:
            login_save_session = self.driver.find_element(By.XPATH, '//input[@type="submit"]')
            login_save_session.click()

            time.sleep(10)

    def send_emails(self):
        for row in range(self.data.shape[0]): 
            button_new_mail = self.driver.find_element(By.XPATH, '//button[@aria-label="New mail"]')
            button_new_mail.click()

            time.sleep(2.5)

            field_to = self.driver.find_element(By.XPATH, '//div[@aria-label="To"]')
            field_to.send_keys(self.data.iloc[row, self.email_col_index])
            field_to.send_keys(Keys.RETURN)

            time.sleep(0.3)

            field_subject = self.driver.find_element(By.XPATH, '//input[@aria-label="Add a subject"]')
            field_subject.send_keys('This is a test!')

            time.sleep(0.3)

            field_message = self.driver.find_element(By.XPATH, '//div[contains(@aria-label, "Message body")]')
            field_message.send_keys(self.msg.replace('[[NAME]]', self.data.iloc[row, self.name_col_index]))

            time.sleep(0.3)

            button_send_mail = self.driver.find_element(By.XPATH, '//button[@aria-label="Send"]')

            time.sleep(1)

            button_send_mail.click()

            time.sleep(1)

if __name__ == '__main__':
    app = App()
    app.run()
    print(f'\nDone! {len(app.data.columns)} emails sent.')