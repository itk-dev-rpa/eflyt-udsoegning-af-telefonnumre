"""This module contains the main process of the robot."""

import os
import csv
import json
import re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from itk_dev_shared_components.eflyt import eflyt_login, eflyt_search
from itk_dev_shared_components.graph import mail, authentication
from itk_dev_shared_components.graph.authentication import GraphAccess
from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    graph_credentials = orchestrator_connection.get_credential(config.GRAPH_API)
    graph_access = authentication.authorize_by_username_password(graph_credentials.username, **json.loads(graph_credentials.password))

    # Create a queue from email input
    emails = mail.get_emails_from_folder("itk-rpa@mkb.aarhus.dk", config.MAIL_SOURCE_FOLDER, graph_access)
    for email in emails:
        cprs, cases, recipient = _read_input_from_email(email, graph_access)
        orchestrator_connection.bulk_create_queue_elements(
            config.QUEUE_NAME,
            references = cprs,
            data = cases,
            created_by = recipient)

    # Read queue and handle cases
    eflyt_credentials = orchestrator_connection.get_credential(config.EFLYT_LOGIN)
    browser = eflyt_login.login(eflyt_credentials.username, eflyt_credentials.password)
    queue_elements_processed = 0
    return_data = []
    while (queue_element := orchestrator_connection.get_next_queue_element(config.QUEUE_NAME)) and queue_elements_processed < config.MAX_TASK_COUNT:
        # Find a case to add note to
        cpr = queue_element.reference
        case = queue_element.data
        eflyt_search.open_case(browser, case)
        try:
            numbers = _get_phone_numbers(browser, cpr)
            return_data.append([cpr, case] + numbers)
        except NoSuchElementException:
            return_data.append([cpr, case] + ["no phone found"])

    # Generate a CSV and send it off
    with open('output.csv', mode='w', newline='', encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerows(return_data)


def _read_input_from_email(email: mail.Email, graph_access: GraphAccess) -> tuple[list[str], list[str], str]:
    """Read input and return pair of cases and cpr numbers"""
    recipient = _get_recipient_from_email(email.body)
    attachments = mail.list_email_attachments(email, graph_access)
    cprs = []
    cases = []
    for attachment in attachments:
        email_attachment = mail.get_attachment_data(attachment, graph_access)
        lines = email_attachment.read().decode().split()
        for line in lines:
            cpr, case = line.split(",")
            cases.append(case.strip())
            cprs.append(cpr.strip().replace("-", ""))
    return cases, cprs, recipient


def _get_phone_numbers(browser: webdriver.Chrome, cpr_in: str) -> list[str]:
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

    numbers = []
    if phone_number:
        numbers.append(phone_number)
    if mobile_number:
        numbers.append(mobile_number)

    return numbers


def _get_recipient_from_email(user_data: str) -> str:
    '''Find email in user_data using regex'''
    pattern = r"E-mail: (\S+)"
    return re.findall(pattern, user_data)[0]


if __name__ == '__main__':
    conn_string = os.getenv("OpenOrchestratorConnString")
    crypto_key = os.getenv("OpenOrchestratorKey")
    oc = OrchestratorConnection("Telefon test", conn_string, crypto_key, '{"approved users":["az68933"]}')
    process(oc)
