# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd
import time
import datetime
import json,requests,openpyxl
from io import BytesIO


NAME=0
Role=1
name_role_mapping={"Anna":["Anna Gao","SME"],"Sophia":["Sophia Zhang","SE"],"Hugh":["Hugh Chao","SE"],"Junsen":["Junsen Chen","SME"],"Kelly":["Kelly Zhou","TA"],
                   "Jason":["Jason Zhou","SE"],"Nina":["Nina Li","EE"],"Qianqian":["Qianqian Liu","SME"],"Wuhao":["Wuhao Chen","SME"],"Howard":["Howard Pei","SE"],"Jimmy":["Ji Bian","SE"],
                   "Arthur":["Arthur Huang","SE"],"Lucas":["Lucas Huang","SE"],
                   "Chris":["Chris Butrymowicz","AU SE(cap 1)"],"Nicky":["Nicky Lian","AU SE(cap 1)"],
                   "Tina":["Tina Su","TW SE(cap 1)"],"Cheryl":["Cheryl Huang","TW SE(cap 1)"],"Jeff":["Jeff Lee","TW SE"],
                   "Wenru":["Wenru Huang","CSG(cap 1)"],"Jingjing":["Jingjing Cai","New Hire(cap 1)"],"Chunyan":["Chunyan Liu","New Hire(cap 1)"]}

pd.set_option('display.max_columns', None)
# 显示所有行
pd.set_option('display.max_rows', None)

d = time.localtime()

weekday_shift = time.strftime("%w", d)
weekday_shift = int(weekday_shift)
today = time.strftime("%Y-%m-%d", d)
print("today:",today)
print("weekday_shift = ",weekday_shift)

local_debug = False
downloaded_excel_path = "./Monitoring Today's Cases and Credit This Week.xlsx" if local_debug else '/tmp/a.xlsx'

def get_json_result():
    data = get_markdown4excel()
    success = True
    message = f'This is global daily credit for {time.strftime("%Y-%m-%d", time.localtime())}.'
    return json.dumps({"data":data,"success":success,"message":message})

def get_markdown4excel():
    print_hi('Script is running, please wait until finish')
    df_excel = get_excel_data()
    # print(df_excel.info)
    # print(df_excel)
    df_fte = get_se_data(df_excel, name_role_mapping.keys())
    print(df_fte)

    df_all = concat_and_sort(df_fte)
    return df_all.to_markdown(stralign="center",numalign="center")

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
def get_excel_data():

    # 读取工作簿和工作簿中的工作表
    response = requests.get('https://prod-00.eastus.logic.azure.com:443/workflows/764229174611433581f584080a1c15c1/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=CekobRDm-H9Dx-tBTpXOblRXrJqihxBoTeCyilDCi2w')
    excel_response = requests.get("https://lucasstorageaccount.blob.core.windows.net/token/Monitoring Today's Cases and Credit This Week.xlsx")
    excel_response = excel_response.content
    with open(downloaded_excel_path,'wb') as f:
        f.write(excel_response)
    df = pd.read_excel(downloaded_excel_path,
                       sheet_name='Case this week', engine='openpyxl')
    # df_au = pd.read_excel(downloaded_excel_path,
    #                       sheet_name='AU Case Assignment', engine='openpyxl')
    #df_tw = pd.read_excel(r"./Monitoring Today's Cases and Credit This Week.xlsx",
    #                      sheet_name='TW Case Assignment', engine='openpyxl')
    # 新建一个工作簿
    df = df.dropna(axis=0, how='all')
    # df_au = df_au.dropna(axis=0, how='all')
    #df_tw = df_tw.dropna(axis=0, how='all')

    # df = df.dropna(axis=1, how='all')
    return df


def get_se_data(df_excel,monitoring_se):


    ret = pd.DataFrame(columns=['workdays', 'case today', 'task today',
                                "active case this week", "active task this week",
                                "active credit per day", "transferred out case", "transferred out task","name","Role",today],
                       index=monitoring_se)
    ret.loc[:, :] = 0
    ret.loc[:, "workdays"] = weekday_shift
    for se in monitoring_se:
        # name
        ret.loc[se, "name"] = name_role_mapping[se][NAME]
        ret.loc[se, "Role"] = name_role_mapping[se][Role]
        # case this week
        temp_df_case_this_week = df_excel[df_excel['Case/Task'].str.contains("Case")]
        temp_df_case_this_week = temp_df_case_this_week[temp_df_case_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "active case this week"] = temp_df_case_this_week.shape[0]
        # task this week
        temp_df_task_this_week = df_excel[df_excel['Case/Task'].str.contains("Task")]
        temp_df_task_this_week = temp_df_task_this_week[temp_df_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "active task this week"] = temp_df_task_this_week.shape[0]
        # case today
        # temp_df_case_today = df_excel[df_excel['Case/Task'].str.contains("Case")]
        # temp_df_case_today = temp_df_case_today[temp_df_case_today["Weekday Shift"] == weekday_shift]
        # temp_df_case_today = temp_df_case_today[temp_df_case_today["Case owner"] == se]
        temp_df_case_today = df_excel[df_excel['Case/Task'].str.contains("Case")]
        temp_df_case_today = temp_df_case_today[temp_df_case_today["Date"]==today]
        temp_df_case_today = temp_df_case_today[temp_df_case_today["Case owner"] == se]
        ret.loc[se, "case today"] = temp_df_case_today.shape[0]
        # task today
        # temp_df_task_today = df_excel[df_excel['Case/Task'].str.contains("Task")]
        # temp_df_task_today = temp_df_task_today[temp_df_task_today["Weekday Shift"] == weekday_shift]
        # temp_df_task_today = temp_df_task_today[temp_df_task_today["Case owner"] == se]
        temp_df_task_today = df_excel[df_excel['Case/Task'].str.contains("Task")]
        temp_df_task_today = temp_df_task_today[temp_df_task_today["Date"] == today]
        temp_df_task_today = temp_df_task_today[temp_df_task_today["Case owner"] == se]
        ret.loc[se, "task today"] = temp_df_task_today.shape[0]
        # transfer out case
        temp_df_transfer_out_case = df_excel[df_excel['Case/Task'].str.contains("Case")]
        temp_df_transfer_out_case = temp_df_transfer_out_case[temp_df_transfer_out_case["Case owner"].str.strip() == se]
        temp_df_transfer_out_case = temp_df_transfer_out_case[temp_df_transfer_out_case["Transfer Out?"] == "Y"]
        ret.loc[se, "transferred out case"] = temp_df_transfer_out_case.shape[0]
        ##transfer out task
        temp_df_transfer_out_task = df_excel[df_excel['Case/Task'].str.contains("Task")]
        temp_df_transfer_out_task = temp_df_transfer_out_task[temp_df_transfer_out_task["Case owner"].str.strip() == se]
        temp_df_transfer_out_task = temp_df_transfer_out_task[temp_df_transfer_out_task["Transfer Out?"] == "Y"]
        ret.loc[se, "transferred out task"] = temp_df_transfer_out_task.shape[0]
        # minus transfer out
        #ret.loc[se, "active case this week"] -= ret.loc[se, "transferred out case"]
        #ret.loc[se, "active task this week"] -= ret.loc[se, "transferred out task"]
        # credit
        ret.loc[se, today] = ret.loc[se, "case today"] + ret.loc[se, "task today"]
        if ret.loc[se, "workdays"] == 0:
            ret.loc[se, "active credit per day"] = 0
        else:
            ret.loc[se, "active credit per day"] = (ret.loc[se, "active case this week"] +
                                                    ret.loc[se, "active task this week"] ) / ret.loc[
                                                       se, "workdays"]
    ret = ret.sort_index()

    ret = ret.drop(columns= ["transferred out case", "transferred out task"])
    return ret







def concat_and_sort(df_fte):
    sum_case_today = df_fte["case today"].sum()
    sum_task_today = df_fte["task today"].sum()
    sum_volumn_today = df_fte[today].sum()
    sum_active_case_this_week = df_fte["active case this week"].sum()
    sum_active_task_this_week = df_fte["active task this week"].sum()
    df_fte.loc["Total"] = ""
    df_fte.loc["Total", "case today"] = sum_case_today
    df_fte.loc["Total", "task today"] = sum_task_today
    df_fte.loc["Total", today] = sum_volumn_today
    df_fte.loc["Total", "active case this week"] = sum_active_case_this_week
    df_fte.loc["Total", "active task this week"] = sum_active_task_this_week
    df_fte = df_fte[["name","Role",today]]
    return df_fte

if __name__=="__main__":
    print(get_markdown4excel())


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
