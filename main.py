import email
import io
import logging
import os
import platform
import re
import socket
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import imaplib
import smtplib
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

debugging = False

log_stream = io.StringIO()
log_format = "%(message)s"

load_dotenv()

LOG_DIR = os.getenv("LOG_DIR")
LOGIN_PAGE = os.getenv("LOGIN_PAGE")
USER_PAGE = os.getenv("USER_PAGE")
ACCOUNT_EMAIL = os.getenv("ACCOUNT_EMAIL")
ACCOUNT_PASSWORD = os.getenv("ACCOUNT_PASSWORD")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")
EMAIL_SERVER = os.getenv("EMAIL_SERVER")
EMAIL_PORT = int(os.getenv("EMAIL_PORT"))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT")
PROCESSED_UIDS_FILE = os.getenv("PROCESSED_UIDS_FILE")

script_directory = Path(__file__).resolve().parent
script_name = Path(__file__).name
driver_path = script_directory.joinpath("edgedriver_macarm64", "msedgedriver")

edge_options = Options()
if debugging:
    edge_options.add_experimental_option("detach", True)
else:
    edge_options.add_argument("--headless")

log_format = "%(asctime)s - %(levelname)s - %(message)s"
log_time_format = "%H:%M:%S"

log_dir = Path(LOG_DIR)
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / f"{Path(script_name).stem}_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
absolute_log_path = log_file.resolve()

handlers = [
    logging.FileHandler(log_file),
    logging.StreamHandler()
]

if debugging:
    handlers.append(logging.StreamHandler(stream=log_stream))

logging.basicConfig(level=logging.INFO, format=log_format, datefmt=log_time_format, handlers=handlers)

def wait_for_element(driver, by, element_identifier, timeout=10):
    try:
        element_present = EC.presence_of_element_located((by, element_identifier))
        WebDriverWait(driver, timeout).until(element_present)
    except TimeoutException:
        logging.error(f"Timeout waiting for {element_identifier}")
        return None
    return driver.find_element(by, element_identifier)

def get_processed_uids():
    if not os.path.exists(PROCESSED_UIDS_FILE):
        return set()
    with open(PROCESSED_UIDS_FILE, "r") as file:
        return set(line.strip() for line in file)

def save_processed_uids(uid):
    with open(PROCESSED_UIDS_FILE, "a") as file:
        file.write(f"{uid}\n")

def fetch_latest_email():
    try:
        mail = imaplib.IMAP4_SSL(EMAIL_SERVER, EMAIL_PORT)
        mail.login(EMAIL_USER, EMAIL_PASSWORD)
        mail.select("inbox")
        logging.info(f"Successfully Authenticated: Checking for mail in {EMAIL_USER} inbox.")
        logging.info(f"Email Subject: {EMAIL_SUBJECT}")

        status, messages = mail.search(None, f'(SUBJECT "{EMAIL_SUBJECT}")')
        if status != "OK":
            logging.info("No emails found with the specified subject.")
            return ""

        email_ids = messages[0].split()
        if not email_ids:
            logging.info("No matching emails found.")
            return ""

        processed_uids = get_processed_uids()

        for email_id in reversed(email_ids):
            status, response = mail.fetch(email_id, "(UID)")
            uid = response[0].split()[-1].decode()

            if uid in processed_uids:
                logging.info(f"Email with UID {uid} already processed. Skipping.")
                continue


            status, msg_data = mail.fetch(email_id, "(RFC822)")
            if status != "OK" or not msg_data or not isinstance(msg_data[0], tuple):
                logging.error(f"Failed to fetch email with UID {uid}.")
                continue

            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            save_processed_uids(uid)
            mail.logout()

            logging.info(f"Processing email with UID: {uid}.")
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        return part.get_payload(decode=True).decode()
            else:
                return msg.get_payload(decode=True).decode()

    except Exception as e:
        logging.error(f"Error fetching email: {e}")
        return ""

def parse_users_from_email():
    users = []
    email_content = fetch_latest_email()
    if not email_content:
        logging.info("No email content to parse.")
        return users

    user_pattern = re.compile(
        r"(?P<first_name>\w+),\s*(?P<last_name>[\w'-]+),.*?,.*?,(?P<email>[\w.-]+@[\w.-]+\.\w+)"
    )
    matches = user_pattern.finditer(email_content)
    for match in matches:
        full_name = f"{match.group('first_name')} {match.group('last_name')}"
        user_email = match.group('email')
        users.append({
            "name": full_name,
            "email": user_email,
        })
        logging.info(f"Parsed users: {full_name} ({user_email})")
    return users

def login_to_account(driver):
    driver.get(LOGIN_PAGE)
    account_input = wait_for_element(driver, By.ID, "email")
    password_input = wait_for_element(driver, By.ID, "password")
    if account_input and password_input:
        account_input.send_keys(ACCOUNT_EMAIL)
        password_input.send_keys(ACCOUNT_PASSWORD)
        submit_button = wait_for_element(driver, By.XPATH, '//input[@type="submit" and @value="Log in"]')
        if submit_button:
            submit_button.click()
            logging.info(f"Successfully Authenticated: {LOGIN_PAGE}")

def navigate_to_user_page(driver):
    driver.get(USER_PAGE)

def scroll_to_element(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    WebDriverWait(driver, 10).until(EC.visibility_of(element))

def enter_email_address(driver, user):
    try:
        teacher_rows = driver.find_elements(By.XPATH, '//tr[starts-with(@id, "teacher-")]')

        for row in teacher_rows:
            teacher_name_element = row.find_element(By.CLASS_NAME, "teacher-name")
            teacher_name = teacher_name_element.text.strip()

            if user["name"].lower() in teacher_name.lower():
                unique_id = row.get_dom_attribute("id").split('-')[1]

                email_input = row.find_element(By.ID, f"email-{unique_id}")
                email_input.clear()
                email_input.send_keys(user["email"])
                logging.info(f"Email address '{user['email']}' entered for '{user['name']}' (ID:{unique_id})")

                analytics_checkbox = row.find_element(By.ID, f"analytics-{unique_id}")
                if analytics_checkbox.get_dom_attribute("data-value") == "no":
                    scroll_to_element(driver, analytics_checkbox)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(analytics_checkbox))
                    analytics_checkbox.click()
                    logging.info(f"Permissions for {unique_id}: Analytics Enabled.")

                sen_checkbox = row.find_element(By.ID, f"provisionmap-{unique_id}")
                if sen_checkbox.get_dom_attribute("data-value") == "no":
                    scroll_to_element(driver, sen_checkbox)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(sen_checkbox))
                    sen_checkbox.click()
                    logging.info(f"Permissions for {unique_id}: SEN Enabled.")

                detentions_checkbox = row.find_element(By.ID, f"detentions-{unique_id}")
                if detentions_checkbox.get_dom_attribute("data-value") == "no":
                    scroll_to_element(driver, detentions_checkbox)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(detentions_checkbox))
                    detentions_checkbox.click()
                    logging.info(f"Permissions for {unique_id}: Detentions Enabled.")

                return unique_id

        logging.info(f"User '{user['name']}' not found in the portal.")
        return None
    except TimeoutException:
        logging.error("Timeout while searching for user elements.")
        return None

def set_password(driver, unique_id, user):
    try:
        ellipsis_icon_xpath = f'//*[@id="{unique_id}"]//i[contains(@class, "fa-ellipsis-v")]'
        ellipsis_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, ellipsis_icon_xpath))
        )
        driver.execute_script("arguments[0].click();", ellipsis_icon)

        set_password_link_xpath = f'//a[@href="#{unique_id}" and contains(@class, "set-password")]'
        set_password_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, set_password_link_xpath))
        )
        driver.execute_script("arguments[0].click();", set_password_link)

        ok_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "change-teacher-password"))
        )
        driver.execute_script("arguments[0].click();", ok_button)
        logging.info(f"Password reset email sent to: {user['email']}.")
    except TimeoutException:
        logging.error(f"Timeout while setting password for {user['name']}.")
    except Exception as e:
        logging.error(f"Error occurred while setting password for {user['name']}: {e}")

def send_summary_email(successful_users, failed_users, general_errors, start_time, end_time):
    msg = EmailMessage()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECEIVER_EMAIL
    msg['Subject'] = "Staff Onboarding - Class Charts User Activation Summary"

    content = (
        "This is a summary email of an automated script for Class Charts staff user activation.\n\n"
        f"Environment:\n"
        f"Operating System: {platform.system()} {platform.release()}\n"
        f"Hostname: {socket.gethostname()}\n"
        f"Script Name: {script_name}\n"
        f"Start Time: {start_time.strftime('%H:%M:%S')}\n"
        f"End Time: {end_time.strftime('%H:%M:%S')}\n\n"
    )

    if successful_users:
        content += "Successfully Processed Users:\n"
        for user in successful_users:
            content += f"{user['name']} - {user['email']}\n"
            content += f"Password reset email sent to: {user['email']}.\n\n"
    if failed_users:
        content += "Failed Users:\n"
        for user in failed_users:
            content += f"{user['name']} - {user['email']}\n"
    else:
        content += "No failures occurred.\n"

    if general_errors:
        content += "General Errors:\n"
        for error in general_errors:
            content += f"{error}\n\n"

    content += f"\n\nLog files can be found at: {absolute_log_path}\n"

    try:
        msg.set_content(content)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
            smtp.send_message(msg)
        logging.info(f"Summary email sent successfully to {RECEIVER_EMAIL}")
    except Exception as e:
        logging.error(f"Failed to send summary email: {e}")

def main():
    start_time = datetime.now()
    logging.info(f"{script_name} started at: {start_time.strftime('%H:%M:%S')}")

    successful_users = []
    failed_users = []
    driver = None

    try:
        users = parse_users_from_email()
        if not users:
            logging.info("No users found in the email to process.")
            return

        service = Service(str(driver_path))
        driver = webdriver.Edge(service=service, options=edge_options)

        login_to_account(driver)
        navigate_to_user_page(driver)

        for user in users:
            logging.info(f"Processing user: {user['name']} with email ({user['email']})")
            try:
                user_unique_id = enter_email_address(driver, user)
                if user_unique_id:
                    set_password(driver, user_unique_id, user)
                    successful_users.append({"name": user["name"], "email": user["email"]})
                    logging.info(f"Successfully processed user: {user['name']}")
                else:
                    failed_users.append({"name": user["name"], "email": user["email"]})
                    logging.info(f"Failed to find user: {user['name']}")
            except Exception as e:
                failed_users.append({"name": user["name"], "email": user["email"]})
                logging.error(f"Error processing user {user['name']}: {e}")

    except WebDriverException as e:
        logging.error(f"General WebDriver error: {e}")
        end_time = datetime.now()
        send_summary_email(
            start_time=start_time,
            successful_users=[],
            failed_users=[],
            general_errors=[f"General WebDriver error: {e}"],
            end_time=end_time
        )
    finally:
        end_time = datetime.now()
        logging.info(f"{script_name} finished at: {end_time.strftime('%H:%M:%S')}")
        if successful_users or failed_users:
            send_summary_email(successful_users, failed_users, [], start_time, end_time)
        else:
            logging.info("No action taken.")

        if debugging:
            print(log_stream.getvalue())
        else:
            if driver is not None:
                driver.quit()

if __name__ == "__main__":
    main()
