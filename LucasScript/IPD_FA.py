
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
import requests
import calendar
from enum import Enum
from datetime import date, datetime, timedelta
import tabulate
import datetime
import logging
global command
try:
    command
except:
    command = "ipd_last_month"


monitoring_fte_se = ['Arthur', 'Anna',   "Junsen", "Kelly",
                 "Niki", "Nina",  "Qianqian",  "Wuhao","Hugh","Sophia","Howard","Jimmy","Lucas","Jason","Wenru","Jingjing","Chunyan"]

monitoring_tw_se = ["Jeff","Tina","Cheryl"]

monitoring_au_se = ["Chris", "Nicky"]


#all_se = monitoring_fte_se + monitoring_vendor_se
all_se = monitoring_fte_se  + monitoring_tw_se+monitoring_au_se

name_mapping = {"Nina Li": "Nina", "Maggie Dong": "Maggie", "Anna Gao": "Anna", "Andy Wu": "Andy",
                "Kelly Zhou": "Kelly", "Qi Chen": "Qi",
                "Wuhao Chen": "Wuhao", "Qianqian Liu": "Qianqian", "Junsen Chen": "Junsen", "Mark He": "Mark",
                "Hugh Chao": "Hugh",
                "Sophia Zhang": "Sophia", "Howard Pei": "Howard", "Ji Bian": "Jimmy", "Jimmy Bian":"Jimmy","Niki Jin": "Niki",
                "Wan Huang": "Wan",
                "Jack Bian": "Jack", "Jiaqi Deng": "Jiaqi", "Arthur Huang": "Arthur",
                "Jerome Guan": "Jerome", "Lucas Huang": "Lucas", "Aristo Liao": "Aristo",
                "Victor Zhang": "Victor", "Guangyu Zhang": "Victor", "Tony Li": "Tony", "Allen Liang": "Allen",
                "Jason Zhou": "Jason", "Adelaide Wu": "Adelaide", "Xichen Xue": "Cici", "Jack Zhou": "Jack Zhou",
                "Zheyi Zheng": "Alen", "Wenru Huang": "Wenru", "Ivan Tong": "Ivan", "Jingjing Cai": "Jingjing",
                "Chunyan Liu": "Chunyan",

                "Cheryl Huang": "Cheryl", "Tina Su": "Tina", "Jeff Lee": "Jeff",
                "Nicky Lian": "Nicky", "Chris Butrymowicz": "Chris"}


leave_dict = {}
total_days_of_month =0
local_debug = False
downloaded_excel_path = "./CaseAssignment.xlsx" if local_debug else '/tmp/a.xlsx'

# 显示所有行
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

def get_json_result():
    data = get_markdown4excel()
    success = True
    if command == "ipd_last_month":
        message = f'Note:IPD last month is calculated by real working days, including Transferred Out.'
    if command == "ipd_this_month":
        message = f'Note:IPD this month is calculated from volume by last Sunday, including Transferred Out.'
    return json.dumps({"data":data,"success":success,"message":message})

def get_markdown4excel():
    print_hi('Script is running, please wait until finish')
    begin_date,end_date = get_days_per_command()
    df_all = get_excel_data(getEveryDay(begin_date,end_date))



    return df_all.to_markdown(stralign="center",numalign="center")


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    logging.info(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

def get_leave_days(cmd):
    url = "https://botmonitoringcheckresourceschedule.azurewebsites.net/api/LucasFunctionAppAnalyzeResourceScheduleHTTPTrigger"
    payload = json.dumps({"date_list":[],"command":cmd})
    response = requests.post(url, data=payload)
    week_off_dict = json.loads(response.text)
    new_week_off_dict = {}
    for k,v in week_off_dict.items():
        if k in name_mapping.keys():
            name = name_mapping[k]
            if name in all_se:
                new_week_off_dict[name]=v
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
    logging.info(f"command in script = {command}")

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
        end = today
        total_days_of_month = today.day


    elif command == "ipd_last_month":

        last_day_of_prev_month = today.replace(day=1) - timedelta(days=1)
        start_day_of_prev_month = today.replace(day=1) - timedelta(days=last_day_of_prev_month.day)
        start, end = start_day_of_prev_month,last_day_of_prev_month
        leave_dict = get_leave_days("last_month")

        month = (today.replace(day=1) - timedelta(days=1)).month
        total_days_of_month = days_of_month(year,month)
    else:
        return []
    print(leave_dict)
    logging.info(f"start = {start},end = {end}")
    return start,end
# Press the green button in the gutter to run the script.
def get_excel_data(date_list):

    # 读取工作簿和工作簿中的工作表
    # df_backlog = pd.read_excel(backlog_excel,engine='openpyxl')
    # df_backlog = df_backlog.dropna(axis=1, how='all')
    # df_backlog = df_backlog[["Names","Cases","All Items"]]


    # 获取在线case assignemnt
    response = requests.get("https://prod-25.eastus.logic.azure.com:443/workflows/377e11f2f61646f7832bbf9623dfa0df/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=OzigXT86xZbL_qS9q4hrY4zw0DU3Hld8NEGxkuyFlTo")
    excel_response = requests.get(
        "https://lucasstorageaccount.blob.core.windows.net/token/CaseAssignment.xlsx")
    excel_response = excel_response.content
    with open(downloaded_excel_path, 'wb') as f:
        f.write(excel_response)
    case_excel = downloaded_excel_path
    df_monitoring_case = pd.read_excel(case_excel,sheet_name='Azure Monitoring',engine='openpyxl')
    df_monitoring_case = df_monitoring_case.dropna(axis=1, how='all')
    df_monitoring_case = df_monitoring_case.loc[df_monitoring_case['Date'].isin(date_list)]
    print("------------df_monitoring_case=-------------")

    #print(df_monitoring_case)

    # 新建一个工作簿
    ret = pd.DataFrame(columns=['Case Volume', 'Collaboration Task Volume', "Total" , "Working days","Daily Volume"],
                       index=all_se)
    ret.loc[:, :] = 0

    for se in all_se:
        print(" ------------- ",se,'-----------------')
        # if se in integration_se:
        #     df_case = df_integration_case
        # else:
        df_case = df_monitoring_case

        # case this week
        temp_df_case_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("case")]
        temp_df_case_this_week = temp_df_case_this_week[temp_df_case_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Case Volume"] = temp_df_case_this_week.shape[0]
        # collab this week
        temp_df_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("collab")]
        temp_df_task_this_week = temp_df_task_this_week[temp_df_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Collaboration Task Volume"] = temp_df_task_this_week.shape[0]


        # weekly total
        ret.loc[se, "Total"] = ret.loc[se, "Case Volume"]+ ret.loc[se, "Collaboration Task Volume"]

        #working days

        if se in leave_dict:
            ret.loc[se, "Working days"] = total_days_of_month - leave_dict[se]
        else:
            ret.loc[se, "Working days"] = 0
        if ret.loc[se, "Working days"] != 0:
            ret.loc[se, "Daily Volume"] =  ret.loc[se, "Total"] / ret.loc[se, "Working days"]
        else:
            ret.loc[se, "Daily Volume"] =  0




    ret.sort_values(["Daily Volume","Total"], ascending= (False,False),
                                   inplace=True)



    return ret




if __name__=="__main__":
    print(get_markdown4excel())

# See PyCharm help at https://www.jetbrains.com/help/pycharm/


