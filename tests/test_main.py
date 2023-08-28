import os

import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By

from main import XlsxHandler, FormFiller, Runner, downloads_path_checker

from unittest.mock import Mock, patch, create_autospec


@pytest.fixture
def insertion_data():
    return [
        {
            "First Name": "John",
            "Last Name": "Smith",
            "Company Name": "IT Solutions",
            "Role in Company": "Analyst",
            "Address": "98 North Road",
            "Email": "jsmith@itsolutions.co.uk",
            "Phone Number": 40716543298,
        },
        {
            "First Name": "Jane",
            "Last Name": "Dorsey",
            "Company Name": "MediCare",
            "Role in Company": "Medical Engineer",
            "Address": "11 Crown Street",
            "Email": "jdorsey@mc.com",
            "Phone Number": 40791345621,
        },
        {
            "First Name": "Albert",
            "Last Name": "Kipling",
            "Company Name": "Waterfront",
            "Role in Company": "Accountant",
            "Address": "22 Guild Street",
            "Email": "kipling@waterfront.com",
            "Phone Number": 40735416854,
        },
        {
            "First Name": "Michael",
            "Last Name": "Robertson",
            "Company Name": "MediCare",
            "Role in Company": "IT Specialist",
            "Address": "17 Farburn Terrace",
            "Email": "mrobertson@mc.com",
            "Phone Number": 40733652145,
        },
        {
            "First Name": "Doug",
            "Last Name": "Derrick",
            "Company Name": "Timepath Inc.",
            "Role in Company": "Analyst",
            "Address": "99 Shire Oak Road",
            "Email": "dderrick@timepath.co.uk",
            "Phone Number": 40799885412,
        },
        {
            "First Name": "Jessie",
            "Last Name": "Marlowe",
            "Company Name": "Aperture Inc.",
            "Role in Company": "Scientist",
            "Address": "27 Cheshire Street",
            "Email": "jmarlowe@aperture.us",
            "Phone Number": 40733154268,
        },
        {
            "First Name": "Stan",
            "Last Name": "Hamm",
            "Company Name": "Sugarwell",
            "Role in Company": "Advisor",
            "Address": "10 Dam Road",
            "Email": "shamm@sugarwell.org",
            "Phone Number": 40712462257,
        },
        {
            "First Name": "Michelle",
            "Last Name": "Norton",
            "Company Name": "Aperture Inc.",
            "Role in Company": "Scientist",
            "Address": "13 White Rabbit Street",
            "Email": "mnorton@aperture.us",
            "Phone Number": 40731254562,
        },
        {
            "First Name": "Stacy",
            "Last Name": "Shelby",
            "Company Name": "TechDev",
            "Role in Company": "HR Manager",
            "Address": "19 Pineapple Boulevard",
            "Email": "sshelby@techdev.com",
            "Phone Number": 40741785214,
        },
        {
            "First Name": "Lara",
            "Last Name": "Palmer",
            "Company Name": "Timepath Inc.",
            "Role in Company": "Programmer",
            "Address": "87 Orange Street",
            "Email": "lpalmer@timepath.co.uk",
            "Phone Number": 40731653845,
        },
    ]


@pytest.fixture
def driver():
    mock_driver = create_autospec(webdriver.Chrome)
    return mock_driver


@pytest.fixture
def xlsx_handler():
    return XlsxHandler()


@pytest.fixture
def form_filler():
    return FormFiller()


@pytest.fixture
def runner(driver):
    return Runner(driver)


def test_xlsx_handler_open_url(driver, xlsx_handler):
    mock_element = Mock()
    mock_element.click = Mock()
    driver.find_element.return_value = mock_element

    with patch.object(xlsx_handler, "wait_for_data", return_value=None):
        xlsx_handler.open_url(driver)
        driver.find_element.assert_called_once_with(
            By.XPATH, '//a[@href="./assets/downloadFiles/challenge.xlsx"]'
        )
        mock_element.click.assert_called_once()
        xlsx_handler.wait_for_data.assert_called_once()


def test_xlsx_handler_do_action_success(insertion_data):
    xlsx_handler = XlsxHandler(
        downloads_dir="tests/test_data", data_file_name="test_challenge.xlsx"
    )
    actual_result = xlsx_handler.do_action()
    expected_result = insertion_data
    assert actual_result == expected_result


def test_xlsx_handler_do_action_fail():
    with pytest.raises(FileNotFoundError):
        xlsx_handler = XlsxHandler(
            downloads_dir="tests/test_data", data_file_name="fake_test_challenge.xlsx"
        )
        xlsx_handler.do_action()


def test_form_filler_open_url(driver, form_filler):
    mock_element = Mock()
    mock_element.click = Mock()
    driver.find_element.return_value = mock_element

    form_filler.open_url(driver)
    driver.find_element.assert_called_once_with(
        By.XPATH, '//button[contains(@class, "btn")]'
    )
    mock_element.click.assert_called_once()


def test_form_filler_do_action_success(driver, form_filler, insertion_data):
    mock_element = Mock()
    mock_element.click = Mock()
    driver.find_element.return_value = mock_element

    with patch.object(form_filler, "expand_to_length", return_value=insertion_data):
        form_filler.do_action(
            driver, insertion_data=insertion_data, form_apply_count=10
        )
        form_filler.expand_to_length.assert_called_once()
        assert mock_element.click.call_count == 10


def test_form_filler_do_action_fail(driver, form_filler):
    with pytest.raises(ValueError):
        form_filler = FormFiller()
        form_filler.do_action(driver, insertion_data=None, form_apply_count=5)


def test_runner_run(driver, runner):
    mock_worker = Mock()
    mock_worker.open_url = Mock()
    mock_worker.do_action = Mock()
    runner.run(mock_worker)

    mock_worker.open_url.assert_called_once()
    mock_worker.do_action.assert_called_once()


def test_downloads_path_checker():
    project_path = os.path.dirname(os.path.abspath("tests"))
    expected_result = os.path.join(project_path, "tests", "test_data")
    actual_result = downloads_path_checker(
        downloads_dir=os.path.join("tests", "test_data"),
        data_file_name="test_remove_challenge.xlsx",
    )
    assert actual_result == expected_result
