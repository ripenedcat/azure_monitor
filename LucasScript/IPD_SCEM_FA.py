
'''
从Case Assignment表中获取所有人的credit，不计transfer out.
根据班表中准确的工作时间来计算IPD

'''

# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import json
from collections import defaultdict
from dateutil.relativedelta import relativedelta
import pandas as pd
import requests,openpyxl
import calendar
from enum import Enum
from datetime import date, datetime, timedelta
import tabulate
import datetime
import logging
from io import BytesIO
global command
try:
    command
except:
    command = "ipd_last_month"


scem_fte_se = ['Lucky', 'Leo',   "Ame", "Nicholas",
                 "Tina", "Tom",  "Wendi","Cheryl"]



#all_se = monitoring_fte_se + monitoring_vendor_se
all_se = scem_fte_se

name_mapping = {"Lucky Zhang":"Lucky","Leo Yu":"Leo","Ame Meng":"Ame","Nicholas Li":"Nicholas","Tina He":"Tina","Tom Tian":"Tom","Wendi Cai":"Wendi","Cheryl Huang":"Cheryl"}


leave_dict = {}
total_days_of_month =0
local_debug = True
downloaded_excel_path = "./CaseAssignment.xlsx" if local_debug else '/tmp/a.xlsx'

# 显示所有行
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

def get_json_result():
    data = get_markdown4excel()
    success = True
    message = ' '
    if command == "ipd_last_month":
        message = f'Note: IPD last month is calculated by real working days, including Transferred Out.'
    if command == "ipd_this_month":
        message = f'Note: IPD this month is calculated from volume by last Sunday, including Transferred Out.'
    elif "|" in command:
        message = f'Note: IPD within {command} is calculated from volume by real working days, including Transferred Out.'
    dto = json.dumps({"data":data,"success":success,"message":message})
    logging.info(f"IPD_FA.py: final data = {dto}")
    return dto

def get_markdown4excel():
    print_hi('Script is running, please wait until finish')
    begin_date,end_date = get_days_per_command()
    df_all = get_excel_data(getEveryDay(begin_date,end_date))
    logging.info(f'df_all = {df_all}')


    return df_all.to_markdown(stralign="center",numalign="center")


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    logging.info(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

def get_leave_days(cmd):
    url = "https://botcheckresourceschedule.azurewebsites.net/api/LucasFunctionAppAnalyzeResourceScheduleHTTPTrigger"
    payload = json.dumps({"date_list":[],"command":cmd,'target_team':'scem'})
    response = requests.post(url, data=payload)
    week_off_dict = json.loads(response.text)
    new_week_off_dict = {}
    for k,v in week_off_dict.items():
        if k in name_mapping.keys():
            name = k
            new_week_off_dict[name]=v
    # for scem, add cheryl separately
    url = "https://botcheckresourceschedule.azurewebsites.net/api/LucasFunctionAppAnalyzeResourceScheduleHTTPTrigger"
    payload = json.dumps({"date_list": [], "command": cmd, 'target_team': 'monitor'})
    response = requests.post(url, data=payload)
    week_off_dict = json.loads(response.text)
    new_week_off_dict["Cheryl Huang"] = week_off_dict["Cheryl Huang"]
    return new_week_off_dict

def days_of_month(year, month):
    """
    给定年份和月份返回这个月有多少天
    :param year:
    :param month:
    :return:
    """
    return calendar.monthrange(year, month)[1]

def getEveryDay(begin_date,end_date):
    date_list = []

    while begin_date <= end_date:
        date_str = begin_date.strftime("%Y-%m-%d")
        date_list.append(date_str)
        begin_date += datetime.timedelta(days=1)
    return date_list


def get_days_per_command():
    '''
    根据输入的指令，获取日期list
    :return:
    '''
    global leave_dict,total_days_of_month
    today = date.today()
    year = today.year
    logging.info(f"IPD_FA.py command in script = {command}")

    if command == "ipd_this_month":
        last_day_of_this_month = today.replace(day=1) + relativedelta(months=1) - timedelta(days=1)
        start_day_of_this_month = today.replace(day=1)
        start, end = start_day_of_this_month,last_day_of_this_month
        leave_dict = get_leave_days("this_month")
        now = date.today()
        last_week_start = now - timedelta(days=now.weekday() + 7)
        last_week_end = now - timedelta(days=now.weekday() + 1)
        if last_week_end.month != now.month:
            return []
        start = today.replace(day=1)
        end = last_week_end
        total_days_of_month = today.day


    elif command == "ipd_last_month":

        last_day_of_prev_month = today.replace(day=1) - timedelta(days=1)
        start_day_of_prev_month = today.replace(day=1) - timedelta(days=last_day_of_prev_month.day)
        start, end = start_day_of_prev_month,last_day_of_prev_month
        leave_dict = get_leave_days("last_month")

        month = (today.replace(day=1) - timedelta(days=1)).month
        total_days_of_month = days_of_month(year,month)
    elif "|" in command:
        start_str, end_str = command.split("|")
        leave_dict = get_leave_days(command)
        start = datetime.datetime.strptime(start_str,'%Y-%m-%d')
        end = datetime.datetime.strptime(end_str,'%Y-%m-%d')
        total_days_of_month = (end - start).days
    else:
        return []
    logging.info(f'leave_dict = {leave_dict}')
    logging.info(f"start = {start},end = {end}")
    return start,end
# Press the green button in the gutter to run the script.
def get_excel_data(date_list):

    # 读取工作簿和工作簿中的工作表
    # df_backlog = pd.read_excel(backlog_excel,engine='openpyxl')
    # df_backlog = df_backlog.dropna(axis=1, how='all')
    # df_backlog = df_backlog[["Names","Cases","All Items"]]


    # 获取在线case assignemnt
    response = requests.get("https://prod-28.eastus.logic.azure.com:443/workflows/bc18fc094a4346999ea15bdc8be86741/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=oDAqQTO8EcvUf0EvtL4L97oikFXafXPhCEMk-8Dj0CM")
    excel_response = requests.get(
        "https://lucasstorageaccount.blob.core.windows.net/token/CaseAssignmentSCEM.xlsx")
    excel_response = excel_response.content
    # with open(downloaded_excel_path, 'wb') as f:
    #     f.write(excel_response)
    excel_binary = openpyxl.load_workbook(filename=BytesIO(excel_response),data_only=True)

    case_excel = excel_binary
    df_monitoring_case = pd.read_excel(case_excel,sheet_name='FY23 Assignment',engine='openpyxl')
    df_monitoring_case = df_monitoring_case.dropna(axis=1, how='all')
    df_monitoring_case = df_monitoring_case.loc[df_monitoring_case['Date'].isin(date_list)]
    logging.info("------------df_monitoring_case=-------------")

    #print(df_monitoring_case)

    # 新建一个工作簿
    ret = pd.DataFrame(columns=['Case Volume', 'Collaboration Task Volume',"Outage Volume", "Total" , "Working days","Daily Volume"],
                       index=name_mapping.keys())
    ret.loc[:, :] = 0

    for se in name_mapping.keys():
        logging.info(f" ------------- {se}-----------------")
        # if se in integration_se:
        #     df_case = df_integration_case
        # else:
        df_case = df_monitoring_case

        # case this week
        temp_df_case_this_week = df_case[df_case['Case/Task?'].str.lower().str.contains("case")]
        #temp_df_case_this_week = temp_df_case_this_week[~temp_df_case_this_week['Is Outage?'].astype(str).str.lower().str.contains("y")]
        temp_df_case_this_week = temp_df_case_this_week[temp_df_case_this_week["Engineer"].str.strip() == se]
        ret.loc[se, "Case Volume"] = temp_df_case_this_week.shape[0]
        # collab this week
        temp_df_task_this_week = df_case[df_case['Case/Task?'].str.lower().str.contains("task")]
        #temp_df_task_this_week = temp_df_task_this_week[~temp_df_task_this_week['Is Outage?'].astype(str).str.lower().str.contains("y")]
        temp_df_task_this_week = temp_df_task_this_week[temp_df_task_this_week["Engineer"].str.strip() == se]
        ret.loc[se, "Collaboration Task Volume"] = temp_df_task_this_week.shape[0]

        #Outage
        # temp_df_outage_this_week = df_case[df_case['Is Outage?'].astype(str).str.lower().str.contains("y")]
        # temp_df_outage_this_week = temp_df_outage_this_week[temp_df_outage_this_week["Case owner"].astype(str).str.strip() == se]
        # ret.loc[se, "Outage Volume"] = temp_df_outage_this_week.shape[0]
        ret.loc[se, "Outage Volume"] = 0

        # weekly total
        ret.loc[se, "Total"] = ret.loc[se, "Case Volume"]+ ret.loc[se, "Collaboration Task Volume"] + ret.loc[se, "Outage Volume"]



        #working days

        if se in leave_dict:
            ret.loc[se, "Working days"] = total_days_of_month - leave_dict[se]
        else:
            ret.loc[se, "Working days"] = 0
        if ret.loc[se, "Working days"] != 0:
            ret.loc[se, "Daily Volume"] =  (ret.loc[se, "Case Volume"]+ ret.loc[se, "Collaboration Task Volume"]+ ret.loc[se, "Outage Volume"]*0.33)/ ret.loc[se, "Working days"]
        else:
            ret.loc[se, "Daily Volume"] =  0




    ret.sort_values(["Daily Volume","Total"], ascending= (False,False),
                                   inplace=True)



    return ret




if __name__=="__main__":
    logging.getLogger().setLevel(logging.INFO)
    print(get_markdown4excel())

# See PyCharm help at https://www.jetbrains.com/help/pycharm/


