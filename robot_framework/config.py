"""This module contains configuration constants used across the framework"""

# The number of times the robot retries on an error before terminating.
MAX_RETRY_COUNT = 3
MAX_TASK_COUNT = 100

# Whether the robot should be marked as failed if MAX_RETRY_COUNT is reached.
FAIL_ROBOT_ON_TOO_MANY_ERRORS = True

# Error screenshot config
SMTP_SERVER = "smtp.aarhuskommune.local"
SMTP_PORT = 25
SCREENSHOT_SENDER = "robot@friend.dk"

# Constant/Credential names
ERROR_EMAIL = "Error Email"
EFLYT_LOGIN = "Eflyt"
GRAPH_API = "Graph API"

# Email
MAIL_SOURCE_FOLDER = "Indbakke/Eflyt udsøgning af telefonnumre"
# EMAIL_RECEIVER = "debitorstyring@mkb.aarhus.dk"
EMAIL_RECEIVER = "kriba@aarhus.dk"
EMAIL_STATUS_SENDER = "itk-rpa@mkb.aarhus.dk"
EMAIL_COMPLETED_SUBJECT = "RPA: Udsøgning af telefonnumre"
EMAIL_COMPLETED_BODY = "Listen af telefonnumre er blevet færdigbehandlet og vedhæftet denne mail."
EMAIL_ATTACHMENT = "eflyt_telefonnumre.xlsx"

# Orchestrator
QUEUE_NAME = "Eflyt Udsøgning af Telefonnumre"
