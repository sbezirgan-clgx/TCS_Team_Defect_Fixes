from jira import JIRA
import os
import pandas as pd
from pathlib import Path
from pyexcel.cookbook import merge_all_to_a_book
import openpyxl
import os
import glob
import win32com.client as win32
import my_date
date_info = my_date.My_date()
print(type(date_info.jquery_start))
print(date_info.jquery_end)
JIRA_USERNAME = os.getenv("JIRA_USERNAME")
JIRA_PASSWORD = os.getenv("JIRA_PASSWORD")
FOLDER_PATH = rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}'

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
jira_Search_result = jira.search_issues(jql_str=f'project = ETQA AND Sub-projects = "Transaction QA" AND  "Transmission Date"  >= 2023-03-01 AND "Transmission Date" <=2023-03-08 AND Status = Open AND Sub-projects not in ("Post production validation")',json_result=True)
print(jira_Search_result.keys())
#print(type(jira_Search_result["issues"]))
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
    os.makedirs(rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}', exist_ok=True)
    df.to_csv(rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}\out.csv', index=False)
    print("Yeay")
else:
    print("This folder already exists.")




if check_if_folder_exists(folder_path=rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}\out.csv'):
    try:
        convert_csv_to_xlsx(
            csv_path=rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}\out.csv',
            file_name=rf'\\filer-diablo-prd\data_acquisition\QA\EagleQC\TCS Team defect fixes June 2022\Week {date_info.jquery_start} to {date_info.jquery_end}\TCS_Defect_Fixes.xlsx')
        remove_the_csv_file(location=FOLDER_PATH,file_name='out.csv')
    except:
        send_an_email(recipient='sbezirgan@corelogic.com', mail_subject='Test',
                      mail_body=f'Hi Varsha, No data has been found in the document',
                      mail_cc='sbezirgan@corelogic.com')
    else:
        send_an_email(recipient='sbezirgan@corelogic.com', mail_subject='Test',
                      mail_body=f'Hi Varsha, \nData is copied in the below location from Week {date_info.jquery_start} to {date_info.jquery_end} for TCS Jira Defect Fixes \n{FOLDER_PATH} ',
                      mail_cc='sbezirgan@corelogic.com')
else:
    print("File has been deleted already.")

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
