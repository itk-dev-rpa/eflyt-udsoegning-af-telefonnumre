"""This module contains the main process of the robot."""

import os
import json
from dataclasses import dataclass
from io import BytesIO
import io

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
from selenium import webdriver
from selenium.webdriver.common.by import By
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from itk_dev_shared_components.eflyt import eflyt_login, eflyt_search
from itk_dev_shared_components.graph import mail, authentication
from itk_dev_shared_components.graph.authentication import GraphAccess
from itk_dev_shared_components.smtp import smtp_util
from robot_framework import config


@dataclass
class CprCaseRow:
    """A dataclass representing a row from input and output"""
    case: str
    cpr: str
    name: str
    phone_number: list[str] | None


@dataclass
class EmailInput:
    '''A dataclass representing input from an email'''
    cpr_cases: list[CprCaseRow]
    requester: str
    email: mail.Email | None


def process(email_data: EmailInput | None, graph_access: GraphAccess, orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    # Login
    eflyt_credentials = orchestrator_connection.get_credential(config.EFLYT_LOGIN)
    browser = eflyt_login.login(eflyt_credentials.username, eflyt_credentials.password)

    recipient = json.loads(orchestrator_connection.process_arguments)["return_email"]

    if email_data:
        handle_email(email_data, browser)
        write_excel(email_data.cpr_cases)
        send_status_emails(email_data, recipient)
        if email_data.email:
            mail.delete_email(email_data.email, graph_access)
        os.remove(config.EMAIL_ATTACHMENT)


def handle_email(email_input: EmailInput, browser: webdriver.Chrome) -> None:
    """Handle an email by looking up each pair of CPR and cases in eflyt and adding a phone number to the instance.

    Args:
        email_input: An EmailInput object containing a list of CPR/Case pairs.
        browser: A WebDriver to use selenium.
    """
    for cpr_case_row in email_input.cpr_cases:
        if cpr_case_row.phone_number is not None:
            continue
        eflyt_search.open_case(browser, cpr_case_row.case)
        numbers = _get_phone_numbers(browser, cpr_case_row.cpr)
        if len(numbers) > 0:
            cpr_case_row.phone_number = numbers
        else:
            cpr_case_row.phone_number = ["No phone number found."]


def send_status_emails(email: EmailInput, recipient: str):
    """Send an email to the requesting party and to the controller.

    Args:
        email: The email that has been processed.
    """
    with open(config.EMAIL_ATTACHMENT, "rb") as file:
        smtp_util.send_email(
            recipient,
            config.EMAIL_STATUS_SENDER,
            "RPA: Udsøgning af telefonnumre",
            "Robotten til udsøgning af telefonnumre er nu færdig.\n\nVedhæftet denne mail finder du et excel-ark, som indeholder sags- og CPR-numre på navngivne borgere, for hvem robotten har slået op i Notus og udsøgt deres telefonnumre. Bemærk, at robotten kan have mødt fejl i systemet, hvilket vil være noteret i arket.\n\n Mvh. ITK RPA",
            config.SMTP_SERVER,
            config.SMTP_PORT,
            False,
            [smtp_util.EmailAttachment(file, config.EMAIL_ATTACHMENT)]
        )


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
    cpr_in = cpr_in.replace("-", "")

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


def write_excel(cases: list[CprCaseRow]) -> BytesIO:
    """Write a list of task objects to an excel sheet.

    Args:
        tasks: The list of task objects to write.

    Returns:
        A BytesIO object containing the Excel sheet.
    """
    wb = Workbook()
    sheet: Worksheet = wb.active
    header = ["Sagsnr.", "CPR", "Navn", "Telefonnumre"]
    sheet.append(header)

    # Populate the sheet
    for cpr_case in cases:
        phone_numbers = convert_phone_number((cpr_case.phone_number))
        row = [cpr_case.case, cpr_case.cpr, cpr_case.name, phone_numbers]
        sheet.append(row)

    # Styling
    set_column_width(sheet)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    # Save the file
    file = BytesIO()
    wb.save(file)
    with open(config.EMAIL_ATTACHMENT, 'wb') as f:
        f.write(file.getvalue())


def set_column_width(work_sheet: Worksheet):
    """Adjust column width to the widest content in the worksheet

    Args:
        work_sheet: The worksheet to mutate
    """
    for col in work_sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        adjusted_width = max_length + 2
        work_sheet.column_dimensions[column].width = adjusted_width


def convert_phone_number(phone_numbers: list[str] | None) -> str:
    """Convert a list of phone numbers to a single string

    Args:
        phone_numbers: A list of phone numbers

    Returns:
        A string with all the numbers from the list
    """
    if phone_numbers is None:
        phone_numbers = ""
    else:
        phone_numbers = ", ".join(phone_numbers)
    return phone_numbers


def _read_csv(email_attachment: BytesIO) -> list[CprCaseRow]:
    """Read data from a CSV. Only used for testing.

    Args:
        email_attachment: Attachment to read from.

    Returns:
        Return a CPR case with data from attachment.
    """
    lines = email_attachment.read().decode().split("\r\n")
    csv_cases = []
    for line in lines[1:]:
        case, cpr, name = line.split(";")
        csv_cases.append(CprCaseRow(case, cpr, name, None))
    return csv_cases


if __name__ == '__main__':
    test_csv = input("Please enter path of test data (CSV):\n")
    return_email = input("Please enter email to receive output:\n")

    conn_string = os.getenv("OpenOrchestratorConnString")
    crypto_key = os.getenv("OpenOrchestratorKey")
    oc = OrchestratorConnection("Telefon test", conn_string, crypto_key, f'{{"return_email": "{return_email}"}}')

    graph_credentials = oc.get_credential(config.GRAPH_API)
    ga = authentication.authorize_by_username_password(graph_credentials.username, **json.loads(graph_credentials.password))

    test_cases = []
    with open(test_csv, "rb") as test_file:
        file_bytes = io.BytesIO(test_file.read())
        test_cases = _read_csv(file_bytes)

    test_email = EmailInput(test_cases, return_email, None)
    process(test_email, ga, oc)
