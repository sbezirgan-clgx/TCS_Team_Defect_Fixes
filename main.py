from jira import JIRA
import os
import pandas as pd
from pathlib import Path
from pyexcel.cookbook import merge_all_to_a_book
import openpyxl
import os
import csv
import glob
import win32com.client as win32
import my_date
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
window = Tk()
date_info = my_date.My_date()
print(type(date_info.jquery_start))
print(date_info.jquery_end)
JIRA_USERNAME = os.getenv("JIRA_USERNAME")
JIRA_PASSWORD = os.getenv("JIRA_PASSWORD")
FOLDER_PATH = rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}'
#file_input = "my_csv.csv"
#FULL_PATH = rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}\{file_input}'
fields = []
rows = []

def start_connection():
    '''Jira Server Connection'''
    jiraOptions = {'server': "https://jira-corelogic.valiantys.net"}
    jira = JIRA(options=jiraOptions, basic_auth=(JIRA_USERNAME, JIRA_PASSWORD))
    return jira

def check_if_folder_exists(folder_path:str) -> bool:
    path = folder_path
    isExist = os.path.exists(path)
    return isExist

def convert_csv_to_xlsx(csv_path:str, file_name:str):
    read_file = pd.read_csv(csv_path)
    read_file.to_excel(file_name, index=None, header=True)


def remove_the_csv_file(location:str, file_name:str):
    path = os.path.join(location, file_name)
    os.remove(path)


def line_count(filename):
    with open(filename, "r") as inf:
        # creating a csv reader object

        csvreader = csv.reader(inf)
        # get total number of rows
        return csvreader.line_num

    inf.close()


#def get_line_num_from_csv(filepath):
    #with open(filepath, 'r') as csvfile:
        #csvreader = csv.reader(csvfile)
        #print("Total no. of rows: %d" % (csvreader.line_num))

def send_an_email(recipient:str,mail_subject:str,mail_body:str,mail_cc:str):
    ol = win32.Dispatch('Outlook.Application')
    mail = ol.CreateItem(0)
    mail.to= recipient
    mail.Subject = mail_subject
    mail.Body=mail_body
    mail.cc = mail_cc
    mail.Send()

jira =start_connection()

all_fields = jira.fields()
nameMap = {field['id']:field['name'] for field in all_fields}
#print(nameMap)
jira_Search_result = jira.search_issues(jql_str=f'project = "EDG - QA Transaction TCS" AND  "Created"  >= {date_info.jquery_start} AND "Created" <= {date_info.jquery_end} AND Status = Open AND Sub-projects in ("Transaction QA","IRM monitoring")',json_result=True,maxResults=1000)
print(jira_Search_result.keys())
#print(type(jira_Search_result["issues"]))
def create_the_file():
    csv_list = []
    for single_issue in jira_Search_result["issues"]:
        print(single_issue["fields"]["assignee"])
        csv_row = {}
        csv_row["TCS Status"] = ''
        csv_row["TCS Comments"] = ''
        csv_row["Key"] = single_issue["key"]
        csv_row["OnShore QA Validation"] = ''
        csv_row["Onshore QA Review Date"] = ''
        csv_row["Status"] = single_issue["fields"]["status"]["name"]
        csv_row["Sub-Projects"] = single_issue["fields"]["customfield_24502"]["value"]
        csv_row["Vendor"] = single_issue["fields"]["customfield_24513"]["value"]
        csv_row["State"] = single_issue["fields"]["customfield_18909"]["value"]
        csv_row["County"] = single_issue["fields"]["customfield_25100"]["value"]
        csv_row["Document Year"] = single_issue["fields"]["customfield_26002"]
        csv_row["Doc Number"] = single_issue["fields"]["customfield_24519"]
        csv_row["Recording Date"] = single_issue["fields"]["customfield_24526"]
        csv_row["Issue Type"] = single_issue["fields"]["issuetype"]["name"]
        csv_row["Recording Book"] = single_issue["fields"]["customfield_26900"]
        csv_row["Recording Page"] = single_issue["fields"]["customfield_27309"]
        csv_row["Deed"] = single_issue["fields"]["customfield_24501"]["value"]
        csv_row["DAMAR Type"] = single_issue["fields"]["customfield_24511"]["value"]
        if str(type(single_issue["fields"]["customfield_24512"])) == "<class 'dict'>":
            csv_row["3072 Fields"] = single_issue["fields"]["customfield_24512"]["value"]
        elif str(type(single_issue["fields"]["customfield_24512"])) == "<class 'NoneType'>":
            csv_row["3072 Fields"] = ''
        if str(type(single_issue["fields"]["customfield_24508"])) == "<class 'dict'>":
            csv_row["ADC"] = single_issue["fields"]["customfield_24508"]["value"]
        elif str(type(single_issue["fields"]["customfield_24508"])) == "<class 'NoneType'>":
            csv_row["ADC"] = ''
        csv_row["Summary"] = single_issue["fields"]["summary"]
        csv_row["Description"] = single_issue["fields"]["description"]
        csv_row["Transmission Date"] = single_issue["fields"]["customfield_24529"]
        csv_row["Created"] = single_issue["fields"]['created'].split("T")[0] + " " + \
                             single_issue["fields"]['created'].split("T")[1].split(".")[0]
        csv_row["Batch Date"] = single_issue["fields"]["customfield_24515"]
        csv_row["Type of Error"] = single_issue["fields"]["customfield_24505"]["value"]
        csv_row["Sample Data"] = single_issue["fields"]["customfield_24528"]
        if str(type(single_issue["fields"]["assignee"])) == "<class 'dict'>":
            csv_row["Assignee"] = single_issue["fields"]["assignee"]["value"]
        elif str(type(single_issue["fields"]["assignee"])) == "<class 'NoneType'>":
            csv_row["Assignee"] = ''
        csv_row["Critical/Non Critical"] = single_issue["fields"]["customfield_24506"]["value"]

        csv_list.append(csv_row)
    print(csv_list)
    df = pd.DataFrame(csv_list)
    result = check_if_folder_exists(folder_path=FOLDER_PATH)
    if result is not True:
        os.makedirs(
            rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}',
            exist_ok=True)
        df.to_csv(
            rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}\out.csv',
            index=False)
        df.to_csv('out1.csv', index=False)
        messagebox.showinfo(title="File Status", message="File export successful")
        os.startfile(rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}')
        print(line_count('out1.csv'))

    else:
        messagebox.showinfo(title="File Status", message="This folder already exists.")

def convert_and_send_email():
    if check_if_folder_exists(
            folder_path=rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}\out.csv'):
        with open('out1.csv', "r") as inf:
            # creating a csv reader object

            csvreader = csv.reader(inf)

            # extracting field names through first row
            fields = next(csvreader)

            # extracting each data row one by one
            for row in csvreader:
                rows.append(row)

            # get total number of rows
            a = (csvreader.line_num)
            print(a)
        if a > 1:
            try:
                convert_csv_to_xlsx(
                    csv_path=rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}\out.csv',
                    file_name=rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS_defect_fixes\Week_{date_info.get_query_Start()}_to_{date_info.get_query_end()}\TCS_Defect_Fixes.xlsx')
                remove_the_csv_file(location=FOLDER_PATH, file_name='out.csv')
            except:
                send_an_email(recipient='sbezirgan@corelogic.com', mail_subject='Test',
                              mail_body=f'Hi Varsha, No data has been found in the document',
                              mail_cc='sbezirgan@corelogic.com')
            else:
                send_an_email(recipient=str(email_list_entry.get()), mail_subject='Test',
                              mail_body=f'Hi Varsha, \n\nData is copied in the below location from Week {date_info.jquery_start} to {date_info.jquery_end} for TCS Jira Defect Fixes \n\n{FOLDER_PATH} ',
                              mail_cc='pgandluri@corelogic.com; taevans@corelogic.com; kschwarz@corelogic.com')
                messagebox.showinfo(title="E-mail Status", message="E-mail has been sent.")
        else:
            send_an_email(recipient='sbezirgan@corelogic.com; jkao@corelogic.com; mgundluru@corelogic.com', mail_subject='Possible Error for TCS Defect Fixes',
                          mail_body=f'TCS Defect Fixes tool could not locate any Jira defect log for the given week.',
                          mail_cc='sbezirgan@corelogic.com')
            messagebox.showinfo(title="E-mail Status", message="E-mail has been sent. But no data has been found for the last week")
    else:
        print("File has been deleted already.")

"""
with open(FULL_PATH, "r") as inf:
            # creating a csv reader object
    csvreader = csv.reader(inf)

            # extracting field names through first row
    fields = next(csvreader)

            # extracting each data row one by one
    for row in csvreader:
        rows.append(row)

            # get total number of rows
    print((csvreader.line_num))
"""
window.title("TCS-Defect-Fixes")
window.config(padx=50, pady=50)
canvas = Canvas(width=200, height=200)
logo_img = PhotoImage(file="../JQL_Project/Safeimagekit-resized-img.png")
canvas.create_image(100, 100, image=logo_img)
canvas.grid(row=0, column=1)



#Labels
isc_label = Label(text="Today's Date")
isc_label.grid(row=1, column=0)
file_label = Label(text="File Folder:")
file_label.grid(row=2, column=0)
file_label = Label(text="E-mail List:")
file_label.grid(row=3, column=0)


#Entries
isc_entry = Entry(width=25)
isc_entry.insert(0,str(date_info.get_todays_date()))
isc_entry.config(state="disabled")
isc_entry.grid(row=1, column=1)
file_entry = Entry(width=30)
file_entry.insert(0,f'Week {date_info.jquery_start} to {date_info.jquery_end}')
file_entry.grid(row=2, column=1)
file_entry.config(state="disabled")
email_list_entry = Entry(width=50,fg="blue")
email_list_entry.insert(0,"varssingh@corelogic.com; eehsan@corelogic.com; mgopinath@corelogic.com; vlkumar@corelogic.com")
email_list_entry.config(state=NORMAL)
email_list_entry.grid(row=3, column=1,ipady=3)
email_list_entry.focus()


# Buttons
add_button = Button(text="Create the file",command=create_the_file, width=36,bg="blue", fg="white")
add_button.grid(row=4, column=1, columnspan=2)
add_button.config(padx=3, pady=3)

add_button2 = Button(text="Send mass e-mail",command=convert_and_send_email, width=36,bg="blue", fg="white")
add_button2.grid(row=5, column=1, columnspan=2)
add_button2.config(padx=7, pady=7)


window.mainloop()
   # print(single_issue["customfield_12517"])
#print(jira.fields())
#print(jira_Search_result["issues"][0])
#for singleIssue in jira.search_issues(jql_str='Project = ETQA AND "Transmission Date" >= 2022-10-01 AND "Transmission Date" <= 2022-10-08 AND Status = Open AND Sub-projects not in ("Post production validation")',json_result=True):
    #print(singleIssue)
    #issue_dict["TCS Status"] = ''
    #issue_dict["TCS Comments"] = ''
    #issue_dict["Key"] = str(singleIssue.key)
    #issue_dict["Summary"] = str(singleIssue.fields.summary)


    #issue_list.append(issue_dict)

#print(issue_list)
