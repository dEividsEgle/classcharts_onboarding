from idlelib.debugger_r import debugging
import os
import logging
import io
import smtplib
from pathlib import Path
from email.message import EmailMessage

from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

debugging = True

log_stream = io.StringIO()
log_format = "%(message)s"

logging.basicConfig(stream=log_stream, level=logging.INFO, format=log_format)

load_dotenv()

LOGIN_PAGE = os.getenv("LOGIN_PAGE")
USER_PAGE = os.getenv("USER_PAGE")
ACCOUNT_EMAIL = os.getenv("ACCOUNT_EMAIL")
ACCOUNT_PASSWORD = os.getenv("ACCOUNT_PASSWORD")
TARGET_NAME = os.getenv("TARGET_NAME")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
ANALYTICS = os.getenv("ANALYTICS")
SEN = os.getenv("SEN")
DETENTIONS = os.getenv("DETENTIONS")
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
RECEIVER_EMAIL = os.getenv("RECEIVER_EMAIL")

script_directory = Path(__file__).resolve().parent
driver_path = script_directory.joinpath("edgedriver_macarm64", "msedgedriver")

edge_options = Options()
if debugging:
    edge_options.add_experimental_option("detach", True)
else:
    edge_options.add_argument("--headless")

def wait_for_element(driver, by, element_identifier, timeout=10):
    try:
        element_present = EC.presence_of_element_located((by, element_identifier))
        WebDriverWait(driver, timeout).until(element_present)
    except TimeoutException:
        logging.info(f"Time out waiting for {element_identifier}")
        return None
    return driver.find_element(by, element_identifier)

def login_to_account(driver):
    driver.get(LOGIN_PAGE)

    account_input = wait_for_element(driver, By.ID, "email")
    password_input = wait_for_element(driver, By.ID, "password")

    if account_input and password_input:
        account_input.send_keys(ACCOUNT_EMAIL)
        password_input.send_keys(ACCOUNT_PASSWORD)

    submit_button = wait_for_element(
        driver, By.XPATH, '//input[@type="submit" and @value="Log in"]'
    )

    if submit_button:
        submit_button.click()

def navigate_to_user_page(driver):
    driver.get(USER_PAGE)

def enter_email_address(driver):
    try:
        teacher_rows = driver.find_elements(By.XPATH, '//tr[starts-with(@id, "teacher-")]')

        for row in teacher_rows:
            full_name = row.get_dom_attribute("data-full_name")
            if full_name == TARGET_NAME:
                unique_id = row.get_dom_attribute("id").split('-')[1]

                email_input = row.find_element(By.ID, f"email-{unique_id}")
                email_input.clear()
                email_input.send_keys(EMAIL_ADDRESS)

                analytics_checkbox = row.find_element(By.ID, f"analytics-{unique_id}")
                if analytics_checkbox.get_dom_attribute("data-value") != ANALYTICS:
                    analytics_checkbox.click()

                sen_checkbox = row.find_element(By.ID, f"provisionmap-{unique_id}")
                if sen_checkbox.get_dom_attribute("data-value") != SEN:
                    sen_checkbox.click()

                detentions_checkbox = row.find_element(By.ID, f"detentions-{unique_id}")
                if detentions_checkbox.get_dom_attribute("data-value") != DETENTIONS:
                    detentions_checkbox.click()

                return unique_id

        logging.info(f"User with name '{TARGET_NAME} not found.")
        return None
    except TimeoutException:
        logging.info("Timeout while searching for elements.")
        return None

def set_password(driver, unique_id):
    try:

        ellipsis_icon = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f'//*[@id="{unique_id}"]/span/i[contains(@class, "fa-ellipsis-v")]'))
        )
        driver.execute_script("arguments[0].click();", ellipsis_icon)

        set_password_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f'//a[@href="#{unique_id}" and contains(@class, "set-password")]'))
        )
        driver.execute_script("arguments[0].click();", set_password_link)

        ok_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "change-teacher-password"))
        )
        driver.execute_script("arguments[0].click();", ok_button)
        logging.info(f"Password reset email sent to: {EMAIL_ADDRESS}.")
    except TimeoutException:
        logging.info(f"Timeout while trying to set password for user ID {unique_id}.")
    except Exception as e:
        logging.info(f"Error occurred while setting password: {e}")

def send_message(subject, receiver):
    sender = SENDER_EMAIL

    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = receiver
    msg['Subject'] = subject
    msg.set_content(log_stream.getvalue())

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender, SENDER_PASSWORD)
        smtp.send_message(msg)

def main():
    service = Service(driver_path)

    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        login_to_account(driver)
        navigate_to_user_page(driver)
        user_unique_id = enter_email_address(driver)

        if user_unique_id:
            logging.info(f"Email address {EMAIL_ADDRESS} entered successfully for user {TARGET_NAME} with ID {user_unique_id}.")
            set_password(driver, user_unique_id)
            logging.info(f"Password set for user {TARGET_NAME} with ID {user_unique_id}.")
            send_message(
                subject=f"Success: Class Charts User {TARGET_NAME} Account Has Been Activated.",
                receiver=RECEIVER_EMAIL
            )
        else:
            logging.info(f"Failed to find or process user {TARGET_NAME}.")
            send_message(
                subject=f"Failure: Unable To Update Class Charts User {TARGET_NAME}",
                receiver=RECEIVER_EMAIL
            )
    except WebDriverException as e:
        logging.info(f"General WebDriver error: {e}")
        send_message(
            subject=f"General Error: Failed To Process Class Charts User {TARGET_NAME}",
            receiver=RECEIVER_EMAIL
        )
    finally:
        if debugging:
            print(log_stream.getvalue())
        else:
            driver.quit()

if __name__ == "__main__":
    main()
