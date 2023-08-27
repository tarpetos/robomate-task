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
    def open_url(self, driver: webdriver.Chrome) -> Any:
        pass

    @abstractmethod
    def do_action(self, *args) -> Any:
        pass


class XlsxHandler(Worker):

    def open_url(self, driver: webdriver.Chrome):
        download_button_xpath = '//a[@href="./assets/downloadFiles/challenge.xlsx"]'
        element = driver.find_element(By.XPATH, download_button_xpath)
        element.click()
        self.wait_for_data()

    def do_action(self, *args) -> List[Dict[str, str]] | None:
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
        while not os.path.exists(os.path.join(DOWNLOADS_DIR, DATA_FILE_NAME)):
            time.sleep(2)


class FormFiller(Worker):
    def __init__(self):
        self.input_mapping = {
            'Address': 'labelAddress',
            'Company Name': 'labelCompanyName',
            'Phone Number': 'labelPhone',
            'Last Name': 'labelLastName',
            'Role in Company': 'labelRole',
            'First Name': 'labelFirstName',
            'Email': 'labelEmail'
        }

    def open_url(self, driver: webdriver.Chrome) -> None:
        start_challenge_button_xpath = '//button[contains(@class, "btn")]'
        element = driver.find_element(By.XPATH, start_challenge_button_xpath)
        element.click()

    def do_action(self, insertion_data: List[Dict[str, str]], form_apply_count: int, driver: webdriver.Chrome) -> None:
        if insertion_data is None:
            return

        submit_page_xpath = '//input[contains(@class, "btn")]'
        for counter, user_data in enumerate(insertion_data, 1):
            counter_check = self.apply_counter_check(counter, form_apply_count, len(insertion_data))
            if counter_check is None:
                return

            self.insert_data(user_data, self.input_mapping, driver)
            element = driver.find_element(By.XPATH, submit_page_xpath)
            element.click()

    @staticmethod
    def insert_data(user_data: Dict[str, str], input_mapping: Dict[str, str], driver: webdriver.Chrome) -> None:
        for attribute, value in user_data.items():
            user_data_element_xpath = f'//input[@ng-reflect-name="{input_mapping[attribute]}"]'
            input_element = driver.find_element(By.XPATH, user_data_element_xpath)
            input_element.send_keys(value)

    @staticmethod
    def apply_counter_check(loop_counter: int, form_apply_count: int, max_counter_value: int) -> None | str:
        if form_apply_count < 0 or form_apply_count > max_counter_value:
            logger.info(msg='Form cannot be filled. Invalid input!')
        elif form_apply_count == loop_counter - 1:
            logger.info(msg=f'Form was filled successfully {loop_counter - 1} time (s)!')
        else:
            return 'Loop continues'


class Runner:
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver
        self.driver.get(PARSE_URL)

    def run(
            self, worker: Any, insertion_data: Dict[str, str] | None = None, form_apply_count: int = 0
    ) -> Dict[str, str] | None:
        worker.open_url(driver=self.driver)
        return worker.do_action(insertion_data, form_apply_count, self.driver)


def downloads_path_checker() -> str:
    project_path = os.path.dirname(os.path.abspath(__file__))
    downloads_folder = os.path.join(project_path, DOWNLOADS_DIR)

    existing_file_path = os.path.join(downloads_folder, DATA_FILE_NAME)
    if os.path.exists(existing_file_path):
        os.remove(existing_file_path)

    return downloads_folder


def main() -> None:
    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': downloads_path_checker()}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=chrome_options)

    runner = Runner(driver)
    xlsx_worker = XlsxHandler()
    form_worker = FormFiller()

    data = runner.run(xlsx_worker)
    runner.run(form_worker, insertion_data=data, form_apply_count=5)


if __name__ == '__main__':
    main()
