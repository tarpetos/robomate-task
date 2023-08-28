import logging
import os
import time
from abc import ABC, abstractmethod
from itertools import islice, cycle
from typing import List, Dict, Any, Optional

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

DOWNLOADS_DIR = "downloads"
DATA_FILE_NAME = "challenge.xlsx"
PARSE_URL = "https://rpachallenge.com/"


class Worker(ABC):
    def open_url(self, driver: webdriver.Chrome) -> Any:
        driver.get(PARSE_URL)

    @abstractmethod
    def do_action(self, *args, **kwargs) -> Any:
        pass


class XlsxHandler(Worker):
    def __init__(
        self, downloads_dir=DOWNLOADS_DIR, data_file_name=DATA_FILE_NAME, **kwargs
    ):
        super().__init__(**kwargs)
        self.output = None
        self.downloads_dir = downloads_dir
        self.data_file_name = data_file_name

    def open_url(self, driver: webdriver.Chrome):
        download_button_xpath = '//a[@href="./assets/downloadFiles/challenge.xlsx"]'
        element = driver.find_element(By.XPATH, download_button_xpath)
        element.click()
        self.wait_for_data()

    def do_action(self, *args) -> List[Dict[str, str]]:
        file_path = os.path.join(self.downloads_dir, self.data_file_name)

        data = pd.ExcelFile(file_path)
        df = data.parse(data.sheet_names[0])
        df.columns = df.columns.str.strip()
        result_dict = df.to_dict(orient="records")

        return result_dict

    @staticmethod
    def wait_for_data() -> None:
        while not os.path.exists(os.path.join(DOWNLOADS_DIR, DATA_FILE_NAME)):
            time.sleep(1)


class FormFiller(Worker):
    def __init__(self):
        self.input_mapping = {
            "Address": "labelAddress",
            "Company Name": "labelCompanyName",
            "Phone Number": "labelPhone",
            "Last Name": "labelLastName",
            "Role in Company": "labelRole",
            "First Name": "labelFirstName",
            "Email": "labelEmail",
        }

    def open_url(self, driver: webdriver.Chrome) -> None:
        super().open_url(driver)
        start_challenge_button_xpath = '//button[contains(@class, "btn")]'
        element = driver.find_element(By.XPATH, start_challenge_button_xpath)
        element.click()

    def do_action(
        self,
        driver: webdriver.Chrome,
        insertion_data: Optional[List[Dict[str, str]]],
        form_apply_count: int,
    ) -> None:
        if not insertion_data:
            raise ValueError("No insertion data provided")
        insertion_data = self.expand_to_length(insertion_data, form_apply_count)

        submit_page_xpath = '//input[contains(@class, "btn")]'
        for counter, user_data in enumerate(insertion_data, 1):
            logger.info("Inserting data for the %d worker...", counter)
            self.insert_data(user_data, self.input_mapping, driver)
            element = driver.find_element(By.XPATH, submit_page_xpath)
            element.click()
        logger.info("Form was filled successfully %d time (s)!", len(insertion_data))

    @staticmethod
    def expand_to_length(
        insertion_data: List[Dict[str, str]], x: int
    ) -> List[Dict[str, Any]]:
        return list(islice(cycle(insertion_data), x))

    @staticmethod
    def insert_data(
        user_data: Dict[str, str],
        input_mapping: Dict[str, str],
        driver: webdriver.Chrome,
    ) -> None:
        for attribute, value in user_data.items():
            user_data_element_xpath = (
                f'//input[@ng-reflect-name="{input_mapping[attribute]}"]'
            )
            input_element = driver.find_element(By.XPATH, user_data_element_xpath)
            input_element.send_keys(value)


class Runner:
    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver
        self.driver.get(PARSE_URL)

    def run(self, worker: Worker, **kwargs) -> None:
        worker.open_url(driver=self.driver)
        worker.output = worker.do_action(self.driver, **kwargs)


def downloads_path_checker(
    downloads_dir=DOWNLOADS_DIR, data_file_name=DATA_FILE_NAME
) -> str:
    project_path = os.path.dirname(os.path.abspath(__file__))
    downloads_folder = os.path.join(project_path, downloads_dir)

    existing_file_path = os.path.join(downloads_folder, data_file_name)
    if os.path.exists(existing_file_path):
        os.remove(existing_file_path)

    return downloads_folder


def main() -> None:
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": downloads_path_checker()}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chrome_options)

    runner = Runner(driver)
    xlsx_worker = XlsxHandler()
    form_worker = FormFiller()

    runner.run(xlsx_worker)
    data = xlsx_worker.output
    runner.run(form_worker, insertion_data=data, form_apply_count=5)


if __name__ == "__main__":
    main()
