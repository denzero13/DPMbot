import re


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