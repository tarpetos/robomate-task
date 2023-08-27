import logging
import os
import time
from abc import ABC, abstractmethod
from typing import List, Dict, Any

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

DOWNLOADS_DIR = 'downloads'
DATA_FILE_NAME = 'challenge.xlsx'
PARSE_URL = 'https://rpachallenge.com/'


class Worker(ABC):
    @abstractmethod
    def open_url(self) -> Any:
        pass

    @abstractmethod
    def do_action(self, *args) -> Any:
        pass


class XlsxHandler(Worker):
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver

    def open_url(self):
        download_button_xpath = '//a[@href="./assets/downloadFiles/challenge.xlsx"]'
        element = self.driver.find_element(By.XPATH, download_button_xpath)
        element.click()
        self.wait_for_data()

    def do_action(self) -> List[Dict[str, str]] | None:
        file_path = os.path.join(DOWNLOADS_DIR, DATA_FILE_NAME)
        result_dict = None

        try:
            data = pd.ExcelFile(file_path)
            df = data.parse(data.sheet_names[0])
            df.columns = df.columns.str.strip()
            result_dict = df.to_dict(orient='records')
        except FileNotFoundError as e:
            logger.info(e)

        return result_dict

    @staticmethod
    def wait_for_data() -> None:
        seconds = 0
        timeout = 10
        while not os.path.exists(os.path.join(DOWNLOADS_DIR, DATA_FILE_NAME)) and seconds < timeout:
            time.sleep(1)
            seconds += 1


class FormFiller(Worker):
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver

    def open_url(self):
        start_challenge_button_xpath = '//button[contains(@class, "btn")]'
        element = self.driver.find_element(By.XPATH, start_challenge_button_xpath)
        element.click()

    def do_action(self, insertion_data: List[Dict[str, str]], form_apply_count: int):
        if insertion_data is None:
            return

        submit_page_xpath = '//input[contains(@class, "btn")]'

        input_mapping = {
            'Address': 'labelAddress',
            'Company Name': 'labelCompanyName',
            'Phone Number': 'labelPhone',
            'Last Name': 'labelLastName',
            'Role in Company': 'labelRole',
            'First Name': 'labelFirstName',
            'Email': 'labelEmail'
        }

        for counter, user_data in enumerate(insertion_data, 1):
            counter_check = self.apply_counter_check(counter, form_apply_count, len(insertion_data))
            if counter_check != 0:
                return

            self.insert_data(user_data, input_mapping)
            element = self.driver.find_element(By.XPATH, submit_page_xpath)
            element.click()

    def insert_data(self, user_data: Dict[str, str], input_mapping: Dict[str, str]):
        for attribute, value in user_data.items():
            user_data_element_xpath = f'//input[@ng-reflect-name="{input_mapping[attribute]}"]'
            input_element = self.driver.find_element(By.XPATH, user_data_element_xpath)
            input_element.send_keys(value)

    @staticmethod
    def apply_counter_check(loop_counter: int, form_apply_count: int, max_counter_value: int) -> int:
        if form_apply_count < 0 or form_apply_count > max_counter_value:
            # time.sleep(5)
            print(f'Form cannot be filled. Invalid input!')
            return -1
        elif form_apply_count == loop_counter - 1:
            # time.sleep(5)
            print(f'Form was filled successfully {loop_counter - 1} time (s)!')
            return -2
        else:
            return 0


def downloads_path_checker() -> str:
    project_path = os.path.dirname(os.path.abspath(__file__))
    downloads_folder = os.path.join(project_path, DOWNLOADS_DIR)

    existing_file_path = os.path.join(downloads_folder, DATA_FILE_NAME)
    if os.path.exists(existing_file_path):
        os.remove(existing_file_path)

    return downloads_folder


def main():
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': downloads_path_checker()}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(PARSE_URL)

    xlsx_worker = XlsxHandler(driver)
    xlsx_worker.open_url()
    data = xlsx_worker.do_action()

    form_worker = FormFiller(driver)
    form_worker.open_url()
    form_worker.do_action(data, 5)


if __name__ == '__main__':
    main()
