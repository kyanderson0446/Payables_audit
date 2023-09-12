import io
import requests
import pandas as pd
import csv
import os
import time
from datetime import datetime
import xlwings as xw
from requests.auth import HTTPBasicAuth


# Variables
reporting_today = datetime.today().date()
reporting_name = 'Payables_detail'
folder_month = datetime.today().month-1
folder_year = datetime.today().year
#################################################

# API Login
user = 'ISU_Accounting_Reporting'
password = '{=D~64n(<(K59mHa'

login = HTTPBasicAuth(user, password)
#################################################

# API fstring
reporting_month = datetime.today().month-1
if reporting_month < 10:
    reporting_month = fr"{reporting_month:02d}"
else:
    reporting_month
reporting_month = datetime.today().strftime(f'%Y-{reporting_month}-%d')
print(reporting_month)

# For specific dates or rerunning. Comment out the reporting month and hardcode '2023-03-12' March 12 2023


#################################################
# wd_ids = ['id']
# For use when rerunning a specific site. Comment out the code below.

# In order to make this optimized, API call for current Reference ID needs to be fresh
result_wdid = requests.get(
        fr"https://services1.myworkday.com/ccx/service/customreport2/pacs/1015967/PACS___Company_Facility_List_with_Reference_ID?format=json",
        auth=login)
result = io.StringIO(result_wdid.text)
df_wdid = pd.read_csv(result)
df_wdid['workdayID'] = df_wdid['workdayID']
df_wdid['referenceID'] = df_wdid['referenceID']
wd_ids = list(df_wdid['workdayID'])
#################################################

# Create Folder
try:
    os.makedirs(fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\2024 - Workday_AP_Aging\{folder_month} - {folder_year}")
except:
    pass
#################################################

# Excel Set Up
xw.Interactive = False
xw.Visible = False
app = xw.App()
#################################################


def get_wd_data(wd_id, report_month):
    xw.Interactive = False
    xw.Visible = False
    app = xw.App()
    # PACS | Company Reference
    result_comp = requests.get(
        fr"https://services1.myworkday.com/ccx/service/customreport2/pacs/ISU_Accounting_Reporting/Integration_Payables_Aging_Details_Report?Company%21WID=c8373e1c1805100e50d37463dbef0000&Reporting_Date=2023-03-29-07%3A00&format=json",
        auth=login)

    result_url = io.StringIO(result_comp.text)
    df_comp = pd.read_csv(result_url)
    df_comp['workdayID'] = df_comp['workdayID']
    df_comp['referenceID'] = df_comp['referenceID']
    if df_comp.loc[df_comp['workdayID'] == wd_id, 'referenceID'].values[0]:
        new_name = df_comp.loc[df_comp['workdayID'] == wd_id, 'referenceID'].values[0]

    # Payables Aging Details
    result_ap = requests.get(
        fr"url",
    auth=login)

    data_ap = io.StringIO(result_ap.text)
    df_ap = pd.read_csv(data_ap)
    df_ap = df_ap.rename(columns={'XMLNAME_1__Month_Overdue': '1_Month_Overdue', 'XMLNAME_2__Month_Overdue': '2_Month_Overdue','XMLNAME_3__Month_Overdue': '3_Month_Overdue',
                                  'XMLNAME_1__Months_Overdue': '1_Month_Overdue','XMLNAME_2__Months_Overdue': '2_Month_Overdue','XMLNAME_3__Months_Overdue': '3_Month_Overdue'})
    aging_file = df_ap.to_csv(fr"{wd_id}.csv", index=False)

    # Open the workbook and insert values on top of API data
    # data = df_ap.to_excel(fr'{new_name} AP aging.xlsx')
    wb = xw.Book(fr"{wd_id}.csv", update_links=False)
    time.sleep(1)
    wb.sheets[0].range("1:1").insert()
    wb.sheets[0].range("2:2").insert()
    wb.sheets[0].range("3:3").insert()
    wb.sheets[0].range(f"a1").value = reporting_name
    wb.sheets[0].range(f"b1").value = new_name
    wb.sheets[0].range(f"a2").value = "Reporting Month:"
    wb.sheets[0].range(f"b2").value = fr"{folder_month} - {folder_year}"
    wb.sheets[0].range(f"a3").value = "Date Pulled:"
    wb.sheets[0].range(f"b3").value = reporting_today
    print(new_name + " _saving...")
    wb.save(
        fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\2024 - Workday_AP_Aging\{folder_month} - {folder_year}\{folder_year} {folder_month} {new_name} Payables Aging.xlsx")
    wb.close()
    os.remove(fr"{wd_id}.csv")

    app.quit()

if __name__ == '__main__':
    for wd_id in wd_ids:
        get_wd_data(wd_id, reporting_month)