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
        self.otc = None
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
        print('OUTLOOK AUTO-SENDER V1.0')
        print('========================\n')

    def load_excel_file(self):
        filepath = input('Enter the full path of the Excel file:\n')
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
            print('Exiting...')
            time.sleep(3)
            exit()

        print(f'\n"Hi, {self.data.iloc[1, self.name_col_index]}! Your email address is {self.data.iloc[0, self.email_col_index]}"')
        confirmation = input('\nDoes that look right? (y/n): ')

        if confirmation != 'y':
            print('\nExiting...')
            time.sleep(3)
            exit()

        self.otc = input('\nEnter the 2FA code: ')

    def init_selenium(self):
        print('\nStarting automation...')

        time.sleep(1.5)

        self.driver = webdriver.Chrome()
        
    def login_outlook(self):
        self.driver.get('https://login.live.com/')

        time.sleep(3)

        login_user = self.driver.find_element(By.NAME, 'loginfmt')
        login_user.send_keys(os.getenv('EMAIL_ACCOUNT'))
        login_user.send_keys(Keys.RETURN)

        time.sleep(3)

        login_password = self.driver.find_element(By.NAME, 'passwd')
        login_password.send_keys(os.getenv('EMAIL_PASSWORD'))
        login_password.send_keys(Keys.RETURN)

        time.sleep(3)

        login_otc = self.driver.find_element(By.NAME, 'otc')
        login_otc.send_keys(self.otc)
        login_otc.send_keys(Keys.RETURN)

        time.sleep(3)

        login_save_session = self.driver.find_element(By.XPATH, '//input[@type="submit"]')
        login_save_session.click()

        time.sleep(3)

    def send_emails(self):
        for row in range(self.data.shape[0]): 
            button_new_mail = self.driver.find_element(By.XPATH, '//button[@aria-label="New mail"]')
            button_new_mail.click()

            time.sleep(1.5)

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