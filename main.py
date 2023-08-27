import os
import time
from typing import List, Dict

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By


class RPAParser:
    DOWNLOADS_DIR = 'downloads'
    DATA_FILE_NAME = 'challenge.xlsx'
    PARSE_URL = 'https://rpachallenge.com/'

    def __init__(self):
        self.chrome_options = webdriver.ChromeOptions()
        self.prefs = {'download.default_directory': self.downloads_path_checker()}
        self.chrome_options.add_experimental_option('prefs', self.prefs)
        self.driver = webdriver.Chrome(options=self.chrome_options)
        self.driver.get(self.PARSE_URL)

    def make_click_on_specific_element(self, xpath: str, download_queue: bool = False) -> None:
        element = self.driver.find_element(By.XPATH, xpath)
        element.click()
        if download_queue:
            self.is_file_downloaded()

    def solve_challenge(self) -> None:
        download_button_xpath = '//a[@href="./assets/downloadFiles/challenge.xlsx"]'
        start_challenge_button_xpath = '//button[contains(@class, "btn")]'
        submit_page_xpath = '//input[contains(@class, "btn")]'

        self.make_click_on_specific_element(download_button_xpath, download_queue=True)
        self.make_click_on_specific_element(start_challenge_button_xpath)
        insertion_data = self.parse_excel(self.DOWNLOADS_DIR, self.DATA_FILE_NAME)

        if insertion_data is None:
            return

        input_mapping = {
            'Address': 'labelAddress',
            'Company Name': 'labelCompanyName',
            'Phone Number': 'labelPhone',
            'Last Name': 'labelLastName',
            'Role in Company': 'labelRole',
            'First Name': 'labelFirstName',
            'Email': 'labelEmail'
        }

        for user_data in insertion_data:
            self.insert_data(user_data, input_mapping)
            self.make_click_on_specific_element(submit_page_xpath)

        time.sleep(5)

    def insert_data(self, user_data: Dict[str, str], input_mapping: Dict[str, str]):
        for attribute, value in user_data.items():
            user_data_element_xpath = f'//input[@ng-reflect-name="{input_mapping[attribute]}"]'
            input_element = self.driver.find_element(By.XPATH, user_data_element_xpath)
            input_element.send_keys(value)

    def downloads_path_checker(self) -> str:
        project_path = os.path.dirname(os.path.abspath(__file__))
        downloads_folder = os.path.join(project_path, self.DOWNLOADS_DIR)

        existing_file_path = os.path.join(downloads_folder, self.DATA_FILE_NAME)
        if os.path.exists(existing_file_path):
            os.remove(existing_file_path)

        return downloads_folder

    def is_file_downloaded(self) -> None:
        while True:
            if os.path.exists(os.path.join(self.DOWNLOADS_DIR, self.DATA_FILE_NAME)):
                break
            time.sleep(1)

    @staticmethod
    def parse_excel(dirname: str, filename: str) -> List[Dict[str, str]] | None:
        file_path = os.path.join(dirname, filename)
        result_dict = None

        try:
            data = pd.ExcelFile(file_path)
            df = data.parse(data.sheet_names[0])
            df.columns = df.columns.str.strip()
            result_dict = df.to_dict(orient='records')
        except FileNotFoundError:
            print('Could not find the data file. Check the download path and try again.')

        return result_dict


def main():
    challenge_solver = RPAParser()
    challenge_solver.solve_challenge()


if __name__ == '__main__':
    main()
