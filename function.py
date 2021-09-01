import re
import smtplib
from email.message import EmailMessage
from config import my_email_account, my_email_password, HtmlFile


def regex_kod(text):
    regex = r"^\d\S+"
    matches = re.finditer(regex, text, re.MULTILINE)

    for matchNum, match in enumerate(matches, start=1):
        kod = match.group()
        return kod


def category(kod):
    if kod == "3":
        return "Контекст"
    if kod == "4":
        return "Люди"
    if kod == "5":
        return "Практика"


def html_file_create(name):
    with open(f"email/{name}.html", "w") as file:
        mail = {"Email address": name}
        html_file = HtmlFile(mail)
        html_file = html_file.get_html_file()

        file.write(html_file)


def mail_send_message(emails):
    print(emails)
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.starttls()
    smtpObj.login(my_email_account, my_email_password)
    msg = EmailMessage()
    msg["From"] = my_email_account
    msg["Subject"] = "DPM Test Result"

    for email in emails:
        try:
            msg["To"] = email
            msg.set_content("Downland and open this file")
            msg.add_attachment(open(f"email/{email}.html", "r").read(), filename="DPM_test_result.html")
            smtpObj.send_message(msg)
        except [ValueError, OSError]:
            pass
    smtpObj.quit()