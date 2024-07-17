import win32com.client as win32
import datetime
import os
import pytz
import yaml

email_list = ["subham.aggarwal@etezazicorps.com", "siddharth.vyas619@gmail.com", "shubham.smvit@gmail.com"]
save_path = r"C:\PythonProjects\email-excel-scraper\output"

# TODO: add logging


def scrape_data() -> list:
    yest_date = datetime.datetime.now(pytz.utc) - datetime.timedelta(days=1)
    if get_date():
        yest_date = get_date()
    return_paths = []
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    for message in messages:
        try:
            addr = message.SenderEmailAddress
            time = message.ReceivedTime
            if addr in email_list and time > yest_date:
                print(addr)
                for attachment in message.Attachments:
                    attachment_path = os.path.join(save_path, attachment.FileName)
                    attachment.SaveAsFile(attachment_path)
                    return_paths.append(attachment_path)
        except AttributeError as e:
            continue
    return return_paths


def write_run_datetime() -> datetime.datetime:
    with open(r"C:\PythonProjects\email-excel-scraper\configs\last_run.yml", "w") as file:
        yaml.dump({"last_run":datetime.datetime.now(pytz.utc)}, file)
    return datetime.datetime.now(pytz.utc)


def get_date() -> datetime.datetime:
    with open(r"C:\PythonProjects\email-excel-scraper\configs\last_run.yml", "r") as file:
        data = yaml.safe_load(file)
        if data:
            return data['last_run']
        else: 
            return None

