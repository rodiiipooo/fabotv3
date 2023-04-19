### imports
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from datetime import datetime
import matplotlib.pyplot as plt
import pandas as pd
from datetime import datetime, timedelta
import win32com
import win32com.client
from tkinter import filedialog as fd
from workingFiles import *
import json
import requests

### REST API for website interaction
URL = "http://192.168.0.34/comGpsGate/api/v.1/applications/40/tokens"
headers = {
"accept": "application/json",
"Content-Type": "application/json"
}
params = {
"username": "your user name",
"password": "your password"
}
resp = requests.post(URL, headers = headers ,data=json.dumps(params))
tk = json.loads(resp.text)['token']
if resp.status_code != 200:
    print('error: ' + str(resp.status_code))
else:
    print('token: ' + str(tk))
    print('Success')

### function to select files // tied to "Prepare Reports" button
def select_docs():
    selected_files = []
    # set file names as global so other classes and their objects can reference and edit 
    global billing_register, csp_transactions, unposted_invoices, gbs_export, pie_extract

    filetypes = (('text files', '*.txt'),('All files', '*.*'))
    selected_files = list(fd.askopenfilenames(
        title='Open files',
        initialdir='/',
        filetypes=filetypes))
    ### read files selected by user and assign them to variables for program usage
    for i in selected_files:
        if ("xlsm" | "xlsx") in i:
            if "Register" in i:
                billing_register = pd.read_excel(i)
            elif "Transactions" in i:
                csp_transactions = pd.read_excel(i)
            elif "Review" in i:
                unposted_invoices = pd.read_excel(i)
        elif "csv" in i:
            if "export" in i.lower():
                gbs_export = pd.read_csv(i)
            elif "Extract" in i:
                pie_extract = pd.read_csv(i)
        else:
            break

### read distribution lists file and clear missing values
all_distributions = pd.read_excel("distributions/all_distributions.xlsx").dropna()
# set distribution lists
test_dir, posted_unposted_dir, focus_file_dir, overdue_invoices_dir = ['rcelisduran@ibm.com', 'jack.waldron@ibm.com'],\
    all_distributions.posted_unposted.values.tolist(),\
    all_distributions.focus_file.values.tolist(),\
    all_distributions.overdue_invoices.values.tolist()

### time and dates for emails and reports
def day_date():
    # define variables and se their values
    global last_friday, yesterday, subject_date
    last_friday, yesterday =\
        str(datetime.now() - timedelta(3)),\
        str(datetime.now() - timedelta(1))
    # if today is Monday the date for the subject will be Friday's date, otherwise it will be yesterday's
    if datetime.weekday(datetime.now()) != 0:
        subject_date = last_friday
    else:
        subject_date = yesterday

### email class
class Emails():
    ### billing process functions
    def test():
        message_subject = "Testing..." + subject_date
        message_body = "This is an automated email test"
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = "; ".join(test_dir)
        message.Subject = message_subject
        message.Body = message_body
        message.Send()

    def d01(attachment):
        message_subject = "Posted Invoices as of 7:00 PM EST " + subject_date
        message_body = "Attached are the posted/unposted invoices as of 7pm EST " + subject_date
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = "; ".join(posted_unposted_dir)
        message.Subject = message_subject
        message.Body = message_body
        message.Attachments.Add(attachment)
        message.Send()
        
    def d02(attachment):
        message_subject = "Billing Focus Forms Actuals..." + subject_date
        message_body = "Please see the attached file containing the billing focus template for "  + subject_date
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = "; ".join(focus_file_dir)
        message.Subject = message_subject
        message.Body = message_body
        message.Attachments.Add(attachment)
        message.Send()

    def d03(attachment):
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = "; ".join(overdue_invoices_dir)
        message.Subject = "Overdue GBS Submissions " + subject_date
        message.Body =\
            "The following invoices were not found in the GBS repository during the daily audit. Invoices are due within 48 hours of posting. Please enter the invoices ASAP. If the workdays past posted date exceeds 4, please respond explaining why it is late with your manager CC'd. "
        message.Attachments.Add(attachment)
        message.Send()

### tasks class with sub classes and their functions
class Tasks():
    class Daily():
        def posted_unposted():
            # edit files as needed for final attachment
            posted_unposted_invoices = pd.DataFrame(unposted_invoices.groupby(unposted_invoices, by="") \
                .groupby(pie_extract, by="") \
                .groupby(billing_register, by="") \
                .groupby(gbs_export, by="")\
                .groupby(csp_transactions, by=""))
            # prepare visuals needed
            plt.plot()

            attachment_d01 = None
            Emails.d01(attachment=attachment_d01)

        def focus_file():

            attachment_d02 = None
            Emails.d02(attachment=attachment_d02)

        def overdue_invoices():

            attachment_d03 = None
            Emails.d03(attachment=attachment_d03)

    class Weekly():
        def contract_mapping():
            pass
        
        def ar_report():
            pass

        def pie_import():
            pass

        def budget_spend():
            pass

        def sst_sbliw():
            pass

    class Monthly():
        def quick_pulse():
            pass

        def mi45_reminder():
            pass

        def fiwlr():
            pass
        
        def fiwlr_misc():
            pass
        
        def cp_actuals():
            pass
        
        def cadence_files():
            pass
        
        def eom_ar_report():
            pass
        
        def oem_files():
            pass
        
        def odd_day():
            pass
        
        def billing_audit():
            pass
        
        def cummulative_overdue():
            pass
        
        def planner_checks():
            pass

### function to perform tasks // tied to "Prepare Reports" button
def perform_tasks(task_list):
        # call day_date function to set dates for emails
        day_date()
        # perform tasks (call functions going down list of tasks)
        for task in task_list:
            if task == "d-Posted/Unposted":
                Tasks.Daily.posted_unposted()
            elif task == "d-Focus File":
                Tasks.Daily.focus_file()
            elif task ==  "d-Overdue Invoices":
                Tasks.Daily.overdue_invoices()
            elif task == "d-All Daily":
                Tasks.Daily()

