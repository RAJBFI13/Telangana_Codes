import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import json
import os

class WebDriverInitializer:
    def __init__(self,):
        self.driver = None
        self.wait = None
    
    def initialize(self):
        try:
            #service = Service(ChromeDriverManager(cache_valid_range=100, version="114.0.5735.90", latest_release_url="https://chromedriver.storage.googleapis.com/114.0.5735.90").install())
            s = Service(ChromeDriverManager().install())
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-extensions")
            options.add_argument('--ignore-certificate-errors')
            options.add_argument("--disable-gpu")
            options.add_argument("disable-blink-features=AutomationControlled")
            options.add_argument(
                "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/87.0.4280.88 Safari/537.36")

            self.driver = webdriver.Chrome(service=s, options=options)
            self.wait = WebDriverWait(self.driver, 20)
        except WebDriverException as e:
            print("Error while initializing WebDriver:", e)
    
    def close(self):
        if self.driver:
            self.driver.quit()

class LoginPage:
    def __init__(self, driver, wait, email, password):
        self.driver = driver
        self.wait = wait
        self.email = email
        self.password = password

    def login(self):
        try:
            user_type = "Citizen"
            user_type_dropdown = self.wait.until(
                EC.visibility_of_element_located((By.XPATH, "//select[@id='user_type']"))
            )
            select = Select(user_type_dropdown)
            select.select_by_visible_text(user_type)

            email_input = self.getElement(By.XPATH, selector="//input[@id='username']")
            password_input = self.getElement(By.XPATH, selector="//input[@id='password']")

            user_type_dropdown.send_keys(user_type)
            time.sleep(5)
            email_input.send_keys(self.email)
            time.sleep(10)
            password_input.send_keys(self.password)
            time.sleep(30)

            login_button = self.getElement(By.XPATH,
                                           selector='//div[@style="text-align: center;"]//button[@type="submit"]')
            login_button.click()
            time.sleep(5)
        except Exception as e:
            print("Error during login:", e)

    def getElement(self, by, selector):
        try:
            return self.driver.find_element(by, selector)
        except Exception as e:
            print("Error while finding element:", e)

class RegistrationPage:
    def __init__(self, driver, wait):
        self.driver = driver
        self.wait = wait

    def navigate_to_details(self):
        try:
            registered_doc_details_link = self.wait.until(
                EC.visibility_of_element_located((By.XPATH, '//a[@href="/districtList.htm"]'))
            )
            registered_doc_details_link.click()

            new_window = self.driver.window_handles[1]
            self.driver.switch_to.window(new_window)

            district_dropdown = self.wait.until(
                EC.visibility_of_element_located((By.ID, 'districtCode'))
            )
            district_select = Select(district_dropdown)
            district_select.select_by_visible_text('HYDERABAD')  # Change to your actual district option text

            sro_dropdown = self.wait.until(
                EC.visibility_of_element_located((By.ID, 'sroCode'))
            )
            sro_select = Select(sro_dropdown)
            sro_select.select_by_visible_text('HYDERABAD (R.O)')  # Change to your actual SRO option text


        except Exception as e:
            print("Error navigating to registration details:", e)
    
    def getElement(self, by, selector):
        try:
            return self.driver.find_element(by, selector)
        except Exception as e:
            print("Error while finding element:", e)

class DocumentProcessing:
    def __init__(self, driver, wait, registration_page, config):
        self.driver = driver
        self.wait = wait
        self.registration_page = registration_page
        self.config = config

    def process_documents(self, district, sro, document_number, year):
        try:
            self.driver.refresh()
            time.sleep(5)
            district_field = self.wait.until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="xs-hidden"]//div[@class="container col-md-offset-3 col-md-6 col-sm-offset-2 col-sm-8 info-data3"]//div[@id="document"]//form[@id="bean"]//select[@id="districtCode"]'))
            )

            district_field.send_keys(district)
            time.sleep(10)

            Sub_Registrar_Office_field = self.wait.until(
                EC.visibility_of_element_located((By.XPATH, '//div[@class="xs-hidden"]//div[@class="container col-md-offset-3 col-md-6 col-sm-offset-2 col-sm-8 info-data3"]//div[@id="document"]//form[@id="bean"]//select[@id="sroCode"]'))
            )
            select_sro = Select(Sub_Registrar_Office_field)
            sro_option_to_select = sro
            select_sro.select_by_visible_text(sro_option_to_select)
            time.sleep(5)
            document_number_input = self.getElement(By.XPATH, selector='//div[@class="xs-hidden"]//div[@class="container col-md-offset-3 col-md-6 col-sm-offset-2 col-sm-8 info-data3"]//div[@id="document"]//form[@id="bean"]//input[@id="doctno"]')
            document_number_input.clear()
            document_number_input.send_keys(str(document_number))
            time.sleep(5)

            year_input = self.getElement(By.XPATH, selector='//div[@class="xs-hidden"]//div[@class="container col-md-offset-3 col-md-6 col-sm-offset-2 col-sm-8 info-data3"]//div[@id="document"]//form[@id="bean"]//input[@id="regyear"]')
            year_input.clear()
            year_input.send_keys(year)
            time.sleep(5)
    
            submit_button = self.getElement(By.XPATH, selector='//div[@class="xs-hidden"]//div[@class="container col-md-offset-3 col-md-6 col-sm-offset-2 col-sm-8 info-data3"]//div[@id="document"]//form[@id="bean"]//button[@class="btn btn-default"]')
            submit_button.click()
            time.sleep(10)

            # excel_report_button = self.getElement(By.XPATH,
            #                                       selector='//div//input[@type="button" and @value="Excel Report"]')
            # self.driver.execute_script("arguments[0].click();", excel_report_button)
            # time.sleep(20)

            html_content = self.driver.page_source

            soup = BeautifulSoup(html_content, 'html.parser')

            # Find the table
            table = soup.find('table')
            # print("table: \n", table)

            # Initialize an empty list to store the table data
            table_data = []

            # Find all rows in the table body
            rows = table.find('tbody').find_all('tr')
            # print("rows: \n", rows)
            # Loop through each row and extract the data
            for row in rows[1:]:
                # print("row: \n", row)
                columns = row.find_all('td')
                s_no = columns[0].text.strip()
                description = columns[1].text.strip()
                reg_exe_pres_dates = columns[2].text.strip().split('\n')
                nature_mkt_con_values = columns[3].text.strip()
                parties = columns[4].text.strip()
                vol_pg_cd_doct = columns[5].text.strip()
                row_data = {
                    "Document Number": document_number,
                    "S.No.": s_no,
                    "Description of property": description,
                    "Reg.Date": reg_exe_pres_dates[0],
                    "Exe.Date": reg_exe_pres_dates[1],
                    "Pres.Date": reg_exe_pres_dates[2],
                    "Nature, Mkt.Value, Con. Value": nature_mkt_con_values,
                    "Name of Parties Executant(EX) & Claimants(CL)": parties,
                    "Vol/Pg No CD No Doct No/Year": vol_pg_cd_doct,
                }
                # Append the row data to the table_data list
                table_data.append(row_data)

            # Save the table data as a JSON file
            json_file_name = f'table_data.json'
            if os.path.exists(json_file_name):
                with open(json_file_name, 'r', encoding="utf-8") as json_file:
                    existing_data = json.load(json_file)
                    existing_data.extend(table_data)
            else:
                existing_data = table_data

            with open(json_file_name, 'w', encoding="utf-8") as json_file:
                json.dump(existing_data, json_file, ensure_ascii=False, indent=4)

            print(f"Table data has been extracted and saved to 'table_data.json' for Document Number: {document_number}.")
            
            back_button = self.getElement(By.XPATH, '//input[@value="Back" and @class="btn btn-custom"]')
            self.driver.execute_script("arguments[0].click();", back_button)
            time.sleep(5)
        except Exception as ex:
            logging.info(f"Skipped document number: {document_number}")
            print(ex)

    def getElement(self, by, selector):
        try:
            return self.driver.find_element(by, selector)
        except Exception as e:
            print("Error while finding element:", e)

class ScraperApp:
    def __init__(self, config: dict):
        self.config = config
        self.driver_initializer = WebDriverInitializer()

    def run(self):
        try:
            self.driver_initializer.initialize()
            if self.driver_initializer.driver:
                self.driver_initializer.driver.get(self.config.get('url', "https://registration.telangana.gov.in/auth_login.htm#!"))
                login_page = LoginPage(self.driver_initializer.driver, self.driver_initializer.wait, "sayali.gujrathi@brantfordindia.com", "Alpha@123")
                login_page.login()

                registration_page = RegistrationPage(self.driver_initializer.driver, self.driver_initializer.wait)
                registration_page.navigate_to_details()

                document_processor = DocumentProcessing(self.driver_initializer.driver, self.driver_initializer.wait, registration_page, self.config)
                for document_number in range(1, 4000):
                    document_processor.process_documents(district="HYDERABAD",sro="BANJARAHILLS (R.O)",document_number=str(document_number),year="2023")

        except Exception as e:
            print("Error during script execution:", e)
        finally:
            self.driver_initializer.close()

if __name__ == "__main__":
    # Configure logging settings
    logging.basicConfig(filename='scraper.log', level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

    config = {"url": "https://registration.telangana.gov.in/auth_login.htm#!"}
    scraper_app = ScraperApp(config)
    scraper_app.run()