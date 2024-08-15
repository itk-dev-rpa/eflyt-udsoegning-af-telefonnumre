"""This module contains the main process of the robot."""

import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from itk_dev_shared_components.eflyt import eflyt_login, eflyt_search
from itk_dev_shared_components.graph import mail, authentication
from itk_dev_shared_components.graph.authentication import GraphAccess
from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    eflyt_credentials = orchestrator_connection.get_credential(config.EFLYT_LOGIN)
    graph_credentials = orchestrator_connection.get_credential(config.GRAPH_API)
    graph_access = authentication.authorize_by_username_password(graph_credentials.username, *graph_credentials.password)

    # Create a queue from email input
    emails = mail.get_emails_from_folder("itk-rpa@mkb.aarhus.dk", config.MAIL_SOURCE_FOLDER, graph_access)
    for email in emails:
        references, data = _read_input_from_email(email, graph_access)
        orchestrator_connection.bulk_create_queue_elements(
            config.QUEUE_NAME,
            references = references,
            data = data,
            created_by = "Robot")

    # Read queue and handle cases
    return_data = list[list[str]]
    browser = eflyt_login.login(eflyt_credentials.username, eflyt_credentials.password)
    queue_elements_processed = 0
    while (queue_element := orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)) and queue_elements_processed < config.MAX_TASK_COUNT:
        # Find a case to add note to
        case = queue_element.reference
        cpr = queue_element.data
        eflyt_search.open_case(browser, case)
        numbers = _get_phone_numbers(browser, cpr)
        return_data.append([case, cpr] + numbers)

    # Generate a CSV and send it off


def _read_input_from_email(email, graph_access: GraphAccess) -> list[list[str], list[str]]:
    """Read input and return pair of cases and cpr numbers"""
    attachments = mail.list_email_attachments(email, graph_access)
    cases = []
    cprs = []
    for attachment in attachments:
        email_attachment = mail.get_attachment_data(attachment, graph_access)
        lines = email_attachment.read().decode().split(",")
        for line in lines:
            cases.append(line[0].strip())
            cprs.append(line[1].strip().replace("-", ""))
    return cases, cprs


def _get_phone_numbers(browser: webdriver.Chrome, cpr_in: str) -> tuple[str, str]:
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
