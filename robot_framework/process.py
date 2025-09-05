"""This module contains the main process of the robot.
Emails come in through initialize, queue elements are created  and then queue elements are compiled to excel and sent out for each email found."""

import json
from dataclasses import dataclass, asdict
from io import BytesIO
import hashlib

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font
from selenium import webdriver
from selenium.webdriver.common.by import By
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection, QueueStatus, QueueElement
from OpenOrchestrator.common import crypto_util

from itk_dev_shared_components.eflyt import eflyt_login, eflyt_search
from itk_dev_shared_components.graph import mail
from itk_dev_shared_components.graph.authentication import GraphAccess
from itk_dev_shared_components.smtp import smtp_util
from robot_framework import config


@dataclass
class CprCaseRow:
    """A dataclass representing a row from input and output"""
    case: str
    cpr: str
    name: str
    phone_numbers: list[str]


@dataclass
class EmailInput:
    '''A dataclass representing input from an email'''
    cpr_cases: list[CprCaseRow]
    requester: str
    email: mail.Email


def process(email_data: EmailInput | None, graph_access: GraphAccess, orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")
    # Login
    eflyt_credentials = orchestrator_connection.get_credential(config.EFLYT_LOGIN)
    browser = eflyt_login.login(eflyt_credentials.username, eflyt_credentials.password)

    if email_data:
        # Read data
        add_phonenumbers_to_queue_elements(email_data, browser, orchestrator_connection)
        recipient = json.loads(orchestrator_connection.process_arguments)["return_email"]
        cases = list[CprCaseRow]
        while q := orchestrator_connection.get_next_queue_element(config.QUEUE_NAME):
            cases.append(convert_queue_element_to_cpr_case_row(q))
            orchestrator_connection.set_queue_element_status(q.id, QueueStatus.DONE)
        compile_results(cases, recipient, email_data, graph_access)


def add_phonenumbers_to_queue_elements(email_input: EmailInput, browser: webdriver.Chrome, orchestrator_connection: OrchestratorConnection) -> None:
    """Handle an email by looking up each pair of CPR and cases in eflyt and adding a phone number to the instance.

    Args:
        email_input: An EmailInput object containing a list of CPR/Case pairs.
        browser: A WebDriver to use selenium.
        orchestrator_connection: Connection used for creating queue elements
    """
    for cpr_case_row in email_input.cpr_cases:
        case_reference = _hash_cpr(cpr_case_row.cpr)
        if cpr_case_row.phone_numbers is not None or cpr_case_row.case == "Manuel" or any(orchestrator_connection.get_queue_elements(config.QUEUE_NAME, case_reference)):
            continue
        eflyt_search.open_case(browser, cpr_case_row.case)
        cpr_case_row.phone_numbers = _get_phone_numbers(browser, cpr_case_row.cpr)
        orchestrator_connection.create_queue_element(config.QUEUE_NAME, reference=case_reference, data=crypto_util.encrypt_string(json.dumps(asdict(cpr_case_row))))


def convert_queue_element_to_cpr_case_row(queue_element: QueueElement) -> CprCaseRow:
    """Convert a QueueElement to a CprCaseRow object.

    Args:
        queue_element: QueueElement to convert.

    Returns:
        CprCaseRow with the same data.
    """
    data = json.loads(crypto_util.decrypt_string(queue_element.data))
    cpr_case_row = CprCaseRow(**data)
    return cpr_case_row


def compile_results(cases: list[CprCaseRow], recipient: str, email: mail.Email, graph_access: GraphAccess):
    """Write excel with results, send a reply with results and remove the email.

    Args:
        cases: _description_
        recipient: _description_
        email: _description_
        graph_access: _description_
    """
    # Generate output
    file = write_excel(cases)
    send_status_emails(recipient, file)
    mail.delete_email(email, graph_access)


def send_status_emails(recipient: str, file: BytesIO):
    """Send an email to the requesting party and to the controller.

    Args:
        email: The email that has been processed.
    """
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


def _hash_cpr(cpr: str) -> str:
    return hashlib.sha256(cpr.encode()).hexdigest()


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

    # Find the phone numbers if they exists
    phone_number_fields = browser.find_elements(By.ID, "ctl00_ContentPlaceHolder2_ptFanePerson_stcPersonTab1_lblTlfnrTxt")
    phone_number = phone_number_fields[0].text if phone_number_fields else None
    mobile_number_fields = browser.find_elements(By.ID, "ctl00_ContentPlaceHolder2_ptFanePerson_stcPersonTab1_lblMobilTxt")
    mobile_number = mobile_number_fields[0].text if mobile_number_fields else None

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
        if cpr_case.phone_numbers == ["N/A"] or cpr_case.phone_numbers is None:  # Skip any entries without a phone number
            continue
        phone_numbers = convert_phone_number((cpr_case.phone_numbers))
        row = [cpr_case.case, cpr_case.cpr, cpr_case.name, phone_numbers]
        sheet.append(row)

    # Styling
    set_column_width(sheet)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    # Save the file
    file = BytesIO()
    wb.save(file)
    return file


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
