"""This module defines any initial processes to run when the robot starts."""

import re
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.graph import mail
from itk_dev_shared_components.graph.authentication import GraphAccess
from robot_framework import config
from robot_framework.process import EmailInput, CprCase


def initialize(graph_access: GraphAccess, orchestrator_connection: OrchestratorConnection) -> EmailInput | None:
    """Do all custom startup initializations of the robot."""
    orchestrator_connection.log_trace("Initializing.")

    # Create a work list of EmailInput from read emails
    emails = mail.get_emails_from_folder("itk-rpa@mkb.aarhus.dk", config.MAIL_SOURCE_FOLDER, graph_access)
    if len(emails) > 0:
        return _read_input_from_email(emails[0], graph_access)
    return None


def _read_input_from_email(email: mail.Email, graph_access: GraphAccess) -> EmailInput:
    """Read input and return pair of cases and cpr numbers"""
    requester = _get_recipient_from_email(email.body)
    attachments = mail.list_email_attachments(email, graph_access)
    cpr_cases = []
    for attachment in attachments:
        email_attachment = mail.get_attachment_data(attachment, graph_access)
        cpr_cases = _read_xlsx(email_attachment)
    return EmailInput(cpr_cases, requester, email)


def _read_xlsx(email_attachment: BytesIO) -> list[CprCase]:
    """Read data from XLSX

    Args:
        email_attachment: Attachment to read from.

    Returns:
        Return a CPR case with data from attachment.
    """
    input_sheet: Worksheet = load_workbook(email_attachment, read_only=True).active

    cases = []

    iter_ = iter(input_sheet)
    next(iter_)  # Skip header row
    for row in iter_:
        case = CprCase(
            case=row[0].value,
            cpr=row[1].value,
            name=row[2].value,
            phone_number=None
        )
        cases.append(case)
    return cases


def _get_recipient_from_email(user_data: str) -> str:
    '''Find email in user_data using regex'''
    pattern = r"E-mail: (\S+)"
    return re.findall(pattern, user_data)[0]
