"""This module contains the main process of the robot."""

import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    browser = login(orchestrator_connection)
    search(browser, "case")
    numbers = get_phone_numbers(browser, "cpr")


def login(orchestrator_connection: OrchestratorConnection) -> webdriver.Chrome:
    browser = webdriver.Chrome()
    browser.maximize_window()

    eflyt_creds = orchestrator_connection.get_credential(config.EFLYT_LOGIN)

    browser.get("https://notuskommunal.scandihealth.net/")
    browser.find_element(By.ID, "Login1_UserName").send_keys(eflyt_creds.username)
    browser.find_element(By.ID, "Login1_Password").send_keys(eflyt_creds.password)
    browser.find_element(By.ID, "Login1_LoginImageButton").click()

    return browser


def search(browser: webdriver.Chrome, case_number: str):
    """Search for a case and open it.

    Args:
        browser: The browser object.
        case_number: The case number to search for.
    """
    browser.get("https://notuskommunal.scandihealth.net/web/Supersearch.aspx")
    browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_searchControl_imgLogo").click()
    browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_searchControl_btnClear").click()
    browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_searchControl_txtSagNr").send_keys(case_number)
    browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_searchControl_txtdatoFra").send_keys("01-01-2020")
    browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_searchControl_txtdatoTo").send_keys("01-01-2030")
    browser.find_element(By.ID, "ctl00_ContentPlaceHolder1_searchControl_btnSearch").click()
    browser.execute_script("__doPostBack('ctl00$ContentPlaceHolder1$searchControl$GridViewSearchResult','cmdRowSelected$0')")


def get_phone_numbers(browser: webdriver.Chrome, cpr_in: str) -> tuple[str, str]:
    """Search for a cpr number on an already open case and extract the phone numbers associated
    with the person.

    Args:
        browser: The browser object.
        cpr_in: The cpr number to look for.

    Returns:
        The persons phone and mobile numbers.
    """
    table = browser.find_element(By.ID, "ctl00_ContentPlaceHolder2_GridViewMovingPersons")
    rows = table.find_elements(By.TAG_NAME, "tr")

    # Remove header row
    rows.pop(0)

    for row in rows:
        cpr_link = row.find_element(By.XPATH, "td[2]/a[2]")
        cpr = cpr_link.text.replace("-", "")

        if cpr == cpr_in:
            cpr_link.click()
            break

    phone_number = browser.find_element(By.ID, "ctl00_ContentPlaceHolder2_ptFanePerson_stcPersonTab1_lblTlfnrTxt").text
    mobile_number = browser.find_element(By.ID, "ctl00_ContentPlaceHolder2_ptFanePerson_stcPersonTab1_lblMobilTxt").text

    return phone_number, mobile_number


if __name__ == '__main__':
    conn_string = os.getenv("OpenOrchestratorConnString")
    crypto_key = os.getenv("OpenOrchestratorKey")
    oc = OrchestratorConnection("Bøde test", conn_string, crypto_key, '{"approved users":["az68933"]}')
    process(oc)
