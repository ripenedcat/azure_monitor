# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import json
from collections import defaultdict

import pandas as pd
import requests
from enum import Enum
from dateutil import relativedelta
from datetime import date, datetime, timedelta
import tabulate
#################Parameters######################
try:
    gl_fw
except:
    #gl_fw=7
    # 计算fiscal week
    NEXT_MONDAY = relativedelta.relativedelta(weekday=relativedelta.MO)
    LAST_MONDAY = relativedelta.relativedelta(weekday=relativedelta.MO(-1))
    ONE_WEEK = timedelta(weeks=1)
    def week_in_fiscal_year(d: date, fiscal_year_start: date) -> int:
        fy_week_2_monday = fiscal_year_start + NEXT_MONDAY
        if d < fy_week_2_monday:
            return 1
        else:
            cur_week_monday = d + LAST_MONDAY
            return int((cur_week_monday - fy_week_2_monday) / ONE_WEEK) + 2


    today = date.today()
    fiscal_year_start = date(today.year, 7, 1)
    #自动计算财年开始日期
    if fiscal_year_start > today:
        fiscal_year_start = fiscal_year_start.replace(year=today.year - 1)
    gl_fw =week_in_fiscal_year(today, fiscal_year_start)

    print(f"path of gl_fw not set. Using defalut {gl_fw}")

try:
    backlog_excel
except:
    backlog_excel = r'C:\Users\lucashuang.FAREAST\pods_data\backlog-fw'+str(gl_fw)+'.xlsx'
    print(f"path of backlog_excel not set. Using defalut {backlog_excel}")

#case_excel = r'C:\Users\v-zixinh\OneDrive - Microsoft\Desktop\FY21_CaseAssignment.xlsx'
try:
    generate_excel
except:
    generate_excel = r'C:\Users\lucashuang.FAREAST\pods_data\output-fw'+str(gl_fw)+'.xlsx'
    print(f"path of generate_excel not set. Using defalut {generate_excel}")

#################################################


monitoring_vendor_se = [
                 "Jack", "Jerome", "Jiaqi",  "Wan", "Aristo","Victor","Allen","Tony","Adelaide","Cici","Jack Zhou","Alen","Ivan"]

monitoring_fte_se = ['Arthur', 'Anna',   "Junsen", "Kelly",
                 "Niki", "Nina",  "Qianqian",  "Wuhao","Hugh","Sophia","Howard","Jimmy","Lucas","Jason","Wenru","Jingjing","Chunyan"]

monitoring_tw_se = ["Jeff","Tina","Cheryl"]

monitoring_au_se = ["Chris", "Nicky"]

integration_se = ["yuzhang6","yuaf","jiecao","zhangz","qili7","v-haoshe","jiaqitong","wenbi","wjzhang","yanden","yinshi","huidongliu","beixiao"
                  ]

#all_se = monitoring_fte_se + monitoring_vendor_se
all_se = monitoring_fte_se + monitoring_vendor_se + monitoring_tw_se+monitoring_au_se

possible_names = {"Andy":["Andy Wu","Hao Wu","Andy W"],"Anna":["Xue Gao","Xue G"],"Bruno":["Bruno L","Bruno Liu"],"Hugh":["Hui C","Hui Chao","Hugh C","Hugh Chao"],"Junsen":["Junsen C","Junsen Chen"],
                 "Kelly":["Yinan Zhou","Yinan Z","Kelly Zhou","Kelly Z"],"Qianqian":["Qianqian L","Qianqian l","Qianqian Liu"],"Maggie":["Meijiao Dong","Maggie D"],
                 "Mark":["Xiaowei He","Xiaowei H","Mark He","Mark H"],"Nina":["Na L","Nina Li"],"Qi":["Qi C","Qi Chen"],"Sophia":["Sophia Z","Sophia Zhang"],"Arthur":["Arthur Huang","Arthur H"],
                 "Jack":["Jack Bian","Jack B"],"Jeremy":["Jeremy Liang"],"Jerome":["Junhao Guan","Junhao G","Jerome G","Jerome Guan"],"Jiaqi":["Jiaqi Deng","Jiaqi D"],"Li":["Li Zhang"],
                 "Wan":["Treasure Huang","Treasure H","Wan Huang","Wan H"],"Edwin":["Edwin Mei","Edwin M"],"Lucas":["Zixin H","Lucas H","Lucas Huang","Zixin Huang"],"Wuhao":["Wuhao Chen","Wuhao C"],
                 "Xuanyi":["Xuanyi L"],"Niki":["Yan J","Yan Jin","Niki Jin","Niki J"],"Howard":["Howard P","Howard Pei"],"Jimmy":["Ji B","Ji Bian"],"Chener":["Chener Zhang","Chener Z"],
                  "Aristo":["Fang L","Fang Liao"],"Victor":["Guangyu Zhang","Guangyu Z"],"Jason":["Jason Zhou","Jason Z","Sheng Zhou","Sheng Z"],"Allen":["Zhaonian Liang","Zhaonian L"],"Tony":["Tianyu Li","Tianyu L"],
                  "yuzhang6":["Yu Zhang","Yu Z"],"v-haoshe":["Haozhou Shen","Haozhou S"],"jiaqitong":["Jiaqi T"],"jiecao":["Jie Cao","Jie C"],"qili7":["Qing L"],
                  "wjzhang":["Wen-Jun Zhang","Wen-Jun Z"],"wenbi":["Wenzhe Bi","Wenzhe B"],"yanden":["Yanbo Deng","Yanbo D"],"yinshi":["Yingjie Shi","Yingjie S"],
                  "yuaf":["Yuanchang F"],"zhangz":["Ziyu Z"],"huidongliu":["Huidong Liu","Huidong L"],"beixiao":["Bei Xiao","Bei X"],"Adelaide":["Lingjie Wu","Lingjie W"],"Cici":["Cici Xue","Cici X"],
                  "Jack Zhou":["Jack Zhou","Jack Z"],"Alen":["Zheyi Zheng","Zheyi Z","Alen Zheng","Alen Z"],"Chunyan":["Chunyan Liu","Chunyan L"],"Jingjing":["Jingjing Cai","Jingjing C"],"Phoebe":["Phoebe Wan","Phoebe W"],
                  "Wenru":["Wenru Huang","Wenru H"],"Ivan":["Ivan Tong","Ivan T","Chen Tong","Chen T"],
                    "Jeff":["Jeff Lee","Jeff L"],"Tina":["Tina Su","Tina S"],"Cheryl":["Cheryl Huang","Cheryl H"],
                  "Chris":["Chris B","Chris Butrymowicz"],"Nicky":["Nicky L","Nicky Lian"]
                  }
#to do: 不要自动获取alias,手写算了
name_alias_mapping = { 'Arthur':"arthurhuang", 'Anna':"xuegao",   "Junsen":"junsche", "Kelly":"yinazhou","Niki":"yanj", "Nina":"nali2",
                       "Qianqian":"liqianqi",  "Wuhao":"wuhchen","Hugh":"huichao","Sophia":"yiqianzhang","Howard":"howardpei","Jimmy":"bianji","Lucas":"lucashuang",
                       "Jason":"shengzhou","Wenru":"v-wenruhuang","Jingjing":"jingjingcai","Chunyan":"chunyanliu" ,

                        "Jeff":"leejeff","Tina":"tins","Cheryl":"yahua",

                        "Chris":"cbutrymowicz", "Nicky":"nickylian",

                        "Jack":"v-jacbia", "Jerome":"v-junhaoguan", "Jiaqi":"v-jiaqde",  "Wan":"v-trhuan", "Aristo":"v-fangliao","Victor":"v-guanzhang",
                       "Allen":"v-zhliang","Tony":"v-litiany","Adelaide":"v-lingjiewu","Cici":"v-xichenxue","Jack Zhou":"v-jingkzhou","Alen":"v-zheyizheng","Ivan":"v-chentong"
                       }

QueryforAllSEOpeningCases_jsonMapper = Enum('QueryforAllSEOpeningCases_jsonMapper', ( 'CaseNumber','AgentId','AgentIdUpdatedOn','AssignedTo','Severity','CaseType','State','taskId','StateAnnotation','StateAnnotationLastUpdatedOn','CreatedOn','CaseCreatedOn','CreationChannel','CustomerProgramPriority','CustomerProgramType','UpdatedOn','LastEmailInteractionCreatedOn','StateLastUpdatedOn','AgentIdAssignedCount','supportTimeZone','supportLanguage','supportCountry','Is24x7optedin','Customers','Title','IncidentId','isIncidentServiceImpactingEvent','IsCritSit','RestrictedAccess','InternalTitle','AssignmentPending','UpdatedBy','Description','IssueContext','IssueDescription','ActiveSystem','CaseActiveSystem','CaseTaskUri','CaseUri','BlockedBy','taskStartsOn','taskEndsOn','caseTaskSAP','Causes','Path','QueueName','CaseAge','Duration','SlaExpiresOn','SlaCompletedOn','SlaState','FCRTarget','FCRState','FCRKpiType','FDRTarget','FDRState','FDRKpiType','servicelevel','EntitlementDescription','EntitlementId','ServiceName','IsPublicSector','PolicyCaseType','unreadEmailsCount','LastInboundCustomerEmail'))
case_dict = defaultdict(int)
task_dict = defaultdict(int)
total_dict = defaultdict(int)

local_debug = False
downloaded_excel_path = "./CaseAssignment.xlsx" if local_debug else '/tmp/a.xlsx'
import os
if os.path.exists(downloaded_excel_path):
    os.remove(downloaded_excel_path)
# 显示所有行
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

def get_json_result():
    data = get_markdown4excel()
    success = True
    message = f'This is weekly credit for FW{gl_fw}.'
    return json.dumps({"data":data,"success":success,"message":message})

def get_markdown4excel():
    print_hi('Script is running, please wait until finish')
    df_general,df_all = get_excel_data(gl_fw)



    return df_general.to_markdown(stralign="center",numalign="center")+"place_to_split"+ df_all.to_markdown(stralign="center",numalign="center")

def getBacklog():
    url = "https://api.support.microsoft.com/v0/queryresult"
    payload = '''{"query":"WITH \r\ncommonFields AS (SELECT\r\n        Core.[CaseNumber],\r\n        'Case' AS CaseType,\r\n        '' AS AssignedTo,\r\n        NULL AS taskId,\r\n        NULL AS taskStartsOn,\r\n        NULL AS taskEndsOn,\r\n        createdOn AS caseTaskCreatedOn,\r\n        createdBy AS caseTaskCreatedBy,\r\n        CASE WHEN updatedOn > Core.LastClientModifiedOn \r\n            THEN updatedOn ELSE Core.LastClientModifiedOn\r\n        END AS caseTaskUpdatedOn,\r\n        RoutingContextId AS RCID,\r\n        SupportAreaPath AS caseTaskSAP,\r\n        '' AS BlockedBy,\r\n        '' AS Description,\r\n        MSEG.QueryService.V1.Case_v1.AssignmentPending AS caseTaskAssignmentPending,\r\n        AssignmentPendingSince AS caseTaskAssignmentPendingSince,\r\n        PreviousTimeInQueues AS PreviousTimeInQ,\r\n        State AS caseTaskState,\r\n        StateAnnotation AS caseTaskStatus,\r\n        StateAnnotationLastUpdatedOn AS caseTaskStatusUpdatedOn,\r\n        ActiveSystem AS CaseTaskActiveSystem,\r\n        CaseUri AS ExtUri,\r\n        agentIdLastUpdatedOn AS caseTaskAgentIdUpdatedOn\r\n    FROM MSEG.QueryService.V1.Case_v1\r\n    INNER JOIN MSEG.QueryService.V1.CoreSystemAttributes AS Core ON Core.CaseNumber = MSEG.QueryService.V1.Case_v1.CaseNumber\r\n    WHERE MSEG.QueryService.V1.Case_v1.State = 'Open' AND AgentId IN (PlaceToReplace)\r\n    UNION SELECT\r\n        CaseNumber,\r\n        'CollaborationTasks' AS CaseType,\r\n        AssignedTo,\r\n        Id AS taskId,\r\n        NULL AS taskStartsOn,\r\n        NULL AS taskEndsOn,\r\n        createdOn AS caseTaskCreatedOn,\r\n        createdBy AS caseTaskCreatedBy,\r\n        updatedOn AS caseTaskUpdatedOn,\r\n        RoutingContextId AS RCID,\r\n        SupportAreaPath AS caseTaskSAP,\r\n        '' AS BlockedBy,\r\n        Description,\r\n        AssignmentPending AS caseTaskAssignmentPending,\r\n        AssignmentPendingSince AS caseTaskAssignmentPendingSince,\r\n        PreviousTimeInQueues AS PreviousTimeInQ,\r\n        State AS caseTaskState,\r\n        StateAnnotation AS caseTaskStatus,\r\n        StateLastUpdatedOn AS caseTaskStatusUpdatedOn,\r\n        ActiveSystem AS CaseTaskActiveSystem,\r\n        ExternalUri as ExtUri,\r\n        agentIdLastUpdatedOn AS caseTaskAgentIdUpdatedOn\r\n    FROM MSEG.QueryService.V1.CollaborationTask\r\n    WHERE  (State = 'Open' OR State = 'InProgress') AND AssignedTo IN (PlaceToReplace)\r\n    UNION SELECT\r\n        CaseNumber,\r\n        'EscalationTasks' AS CaseType,\r\n        AssignedTo,\r\n        Id AS taskId,\r\n        NULL AS taskStartsOn,\r\n        NULL AS taskEndsOn,\r\n        createdOn AS caseTaskCreatedOn,\r\n        createdBy AS caseTaskCreatedBy,\r\n        updatedOn AS caseTaskUpdatedOn,\r\n        RoutingContextId AS RCID,\r\n        SupportAreaPath AS caseTaskSAP,\r\n        '' AS BlockedBy,\r\n        Description,\r\n        AssignmentPending AS caseTaskAssignmentPending,\r\n        AssignmentPendingSince AS caseTaskAssignmentPendingSince,\r\n        PreviousTimeInQueues AS PreviousTimeInQ,\r\n        State AS caseTaskState,\r\n        StateAnnotation AS caseTaskStatus,\r\n        StateLastUpdatedOn AS caseTaskStatusUpdatedOn,\r\n        'MSaaS' AS CaseTaskActiveSystem,\r\n        '' as ExtUri,\r\n        agentIdLastUpdatedOn AS caseTaskAgentIdUpdatedOn\r\n    FROM MSEG.QueryService.V1.EscalationTask\r\n    WHERE  (State = 'Open' OR State = 'InProgress') AND AssignedTo IN (PlaceToReplace)\r\n    UNION SELECT\r\n        CaseNumber,\r\n        'FollowupTasks' AS CaseType,\r\n        AssignedTo,\r\n        Id AS taskId,\r\n        StartsOn AS taskStartsOn,\r\n        EndsOn AS taskEndsOn,\r\n        createdOn AS caseTaskCreatedOn,\r\n        createdBy AS caseTaskCreatedBy,\r\n        updatedOn AS caseTaskUpdatedOn,\r\n        RoutingContextId AS RCID,\r\n        SupportAreaPath AS caseTaskSAP,\r\n        BlockedBy,\r\n        Description,\r\n        AssignmentPending AS caseTaskAssignmentPending,\r\n        AssignmentPendingSince AS caseTaskAssignmentPendingSince,\r\n        PreviousTimeInQueues AS PreviousTimeInQ,\r\n        State AS caseTaskState,\r\n        StateAnnotation AS caseTaskStatus,\r\n        StateLastUpdatedOn AS caseTaskStatusUpdatedOn,\r\n        'MSaaS' AS CaseTaskActiveSystem,\r\n        '' AS ExtUri,\r\n        agentIdLastUpdatedOn AS caseTaskAgentIdUpdatedOn\r\n    FROM MSEG.QueryService.V1.FollowUpTask\r\n    WHERE  (State = 'Open' OR State = 'InProgress') AND AssignedTo IN (PlaceToReplace)),CaseList AS (SELECT \r\n        C.CaseNumber,\r\n        C.SupportTimeZone,\r\n        C.SupportLanguage,\r\n        C.SupportCountry,\r\n        C.Is24x7optedin,\r\n        C.LastEmailInteractionCreatedOn,\r\n        C.StateLastUpdatedOn,\r\n        C.AgentIdAssignedCount,\r\n        C.AgentId,\r\n        C.Severity,\r\n        C.Customers,\r\n        C.Title,\r\n        C.IncidentId,\r\n        C.isIncidentServiceImpactingEvent,\r\n        C.IsCritSit,\r\n        C.RestrictedAccess,\r\n        C.InternalTitle,\r\n        C.CreatedOn AS CaseCreatedOn,\r\n        C.UpdatedOn AS CaseUpdatedOn,\r\n        C.CreationChannel,\r\n        C.CustomerProgramPriority,\r\n        C.CustomerProgramType,\r\n        C.UpdatedBy,\r\n        C.IssueDescription,\r\n        C.IssueContext,\r\n        C.Causes,\r\n        C.ActiveSystem AS CaseActiveSystem,\r\n        C.CaseUri,\r\n        commonFields.AssignedTo,\r\n        commonFields.CaseType,\r\n        commonFields.taskId,\r\n        commonFields.RCID,\r\n        commonFields.taskStartsOn,\r\n        commonFields.taskEndsOn,\r\n        commonFields.caseTaskUpdatedOn AS UpdatedOn,\r\n        commonFields.caseTaskCreatedOn AS CreatedOn,\r\n        commonFields.caseTaskCreatedBy AS CreatedBy,\r\n        commonFields.caseTaskSAP,\r\n        commonFields.BlockedBy,\r\n        commonFields.Description,\r\n        commonFields.caseTaskAssignmentPending AS AssignmentPending,\r\n        commonFields.caseTaskAssignmentPendingSince AS AssignmentPendingSince,\r\n        commonFields.caseTaskState AS State,\r\n        commonFields.caseTaskStatus AS StateAnnotation,\r\n        commonFields.caseTaskStatusUpdatedOn AS StateAnnotationLastUpdatedOn,\r\n        commonFields.CaseTaskActiveSystem AS ActiveSystem,\r\n        commonFields.ExtUri,\r\n        caseTaskAgentIdUpdatedOn AS AgentIdUpdatedOn\r\n    FROM MSEG.QueryService.V1.Case_v1 AS C\r\n    INNER JOIN commonFields ON C.CaseNumber = commonFields.CaseNumber\r\n    WHERE C.Severity IN (1,2,3,4) AND C.State = 'Open')\r\n,emailinteractions AS (SELECT DISTINCT\r\n        CaseList.caseNumber,\r\n        LastMail.CreatedOn AS LastInboundCustomerEmail\r\n      FROM CaseList\r\n      CROSS APPLY (SELECT TOP 1 Emails.CreatedOn\r\n        FROM MSEG.QueryService.V1.EmailInteraction AS Emails\r\n        WHERE Emails.Direction = 'Inbound' AND Emails.CaseNumber = CaseList.caseNumber AND Emails.iscustomerviewable = 1\r\n        ORDER BY EMails.CreatedOn DESC) LastMail)\r\n,caseLabors AS (SELECT\r\n        CaseList.casenumber,\r\n        sum(CAST (Duration AS INT)) AS Duration\r\n    FROM CaseList\r\n    INNER JOIN MSEG.QueryService.V1.Labor AS lbr ON lbr.CaseNumber = CaseList.CaseNumber\r\n    WHERE taskId IS NULL\r\n    GROUP BY CaseList.caseNumber)\r\n,caseFcrKpis AS (SELECT \r\n        DISTINCT CaseList.caseNumber,\r\n        MSEG.QueryService.V1.Kpi.KpiType AS FCRKpiType,\r\n        MSEG.QueryService.V1.Kpi.Name AS FCRName,\r\n        TRY_CONVERT (DATETIME2, MSEG.QueryService.V1.Kpi.[Target]) AS [FCRTarget],\r\n        MSEG.QueryService.V1.Kpi.[State] AS FCRState\r\n    FROM  CaseList\r\n    INNER JOIN MSEG.QueryService.V1.Kpi ON MSEG.QueryService.V1.Kpi.CaseNumber = CaseList.CaseNumber\r\n    WHERE  MSEG.QueryService.V1.Kpi.Name = 'FCR' AND MSEG.QueryService.V1.Kpi.[State] != 'Cancelled')\r\n,caseFdrKpis AS (SELECT \r\n        DISTINCT CaseList.caseNumber,\r\n        MSEG.QueryService.V1.Kpi.KpiType AS FDRKpiType,\r\n        MSEG.QueryService.V1.Kpi.Name AS FDRName,\r\n        TRY_CONVERT (DATETIME2, MSEG.QueryService.V1.Kpi.[Target]) AS [FDRTarget],\r\n        MSEG.QueryService.V1.Kpi.[State] AS FDRState\r\n    FROM  CaseList\r\n    INNER JOIN MSEG.QueryService.V1.Kpi ON MSEG.QueryService.V1.Kpi.CaseNumber = CaseList.CaseNumber\r\n    WHERE MSEG.QueryService.V1.Kpi.Name = 'FDR' AND MSEG.QueryService.V1.Kpi.[State] != 'Cancelled')\r\n,caseServiceLevel AS (SELECT \r\n        DISTINCT CaseList.caseNumber,\r\n        MSEG.QueryService.V1.Redemption.servicelevel,\r\n        MSEG.QueryService.V1.Redemption.friendlyname,\r\n        MSEG.QueryService.V1.Redemption.ServiceName,\r\n        MSEG.QueryService.V1.Redemption.IsPublicSector,\r\n        MSEG.QueryService.V1.Redemption.EntitlementId\r\n    FROM CaseList\r\n    INNER JOIN MSEG.QueryService.V1.Redemption ON CaseList.casenumber = MSEG.QueryService.V1.Redemption.caseNumber)\r\nSELECT CaseList.CaseNumber,\r\n    CaseList.AgentId,\r\n    CaseList.AgentIdUpdatedOn,\r\n    CaseList.AssignedTo,\r\n    CaseList.Severity,\r\n    CaseList.CaseType,\r\n    CaseList.State,\r\n    CaseList.taskId,\r\n    CaseList.StateAnnotation,\r\n    CaseList.StateAnnotationLastUpdatedOn,\r\n    CaseList.CreatedOn, \r\n    CaseList.CaseCreatedOn,\r\n    CaseList.CreationChannel,\r\n    CaseList.CustomerProgramPriority,\r\n    CaseList.CustomerProgramType,\r\n    CaseList.UpdatedOn,\r\n    CaseList.LastEmailInteractionCreatedOn,\r\n    CaseList.StateLastUpdatedOn,\r\n    CaseList.AgentIdAssignedCount,\r\n    CaseList.supportTimeZone,\r\n    CaseList.supportLanguage,\r\n    CaseList.supportCountry,\r\n    CaseList.Is24x7optedin,\r\n    CaseList.Customers,\r\n    CaseList.Title,\r\n    CaseList.IncidentId,\r\n    CaseList.isIncidentServiceImpactingEvent,\r\n    CaseList.IsCritSit,\r\n    CaseList.RestrictedAccess,\r\n    CaseList.InternalTitle,\r\n    CaseList.AssignmentPending,\r\n    CaseList.UpdatedBy,\r\n    CaseList.Description,\r\n    CaseList.IssueContext,\r\n    CaseList.IssueDescription,\r\n    CaseList.ActiveSystem,\r\n    CaseList.CaseActiveSystem,\r\n    CaseList.ExtUri AS CaseTaskUri,\r\n    CaseList.CaseUri,\r\n    CaseList.BlockedBy,\r\n    CaseList.taskStartsOn,\r\n    CaseList.taskEndsOn,\r\n    CaseList.caseTaskSAP,\r\n    CaseList.Causes,\r\n    MSEG.QueryService.V1.SAP.Path,\r\n    MSEG.QueryService.V1.Queue.Name as QueueName,\r\n    CASE WHEN CaseList.State = 'Closed' THEN \r\n        DATEDIFF(minute, CreatedOn, StateLastUpdatedOn) ELSE DATEDIFF(minute, CreatedOn, GETUTCDATE())\r\n    END AS CaseAge,\r\n    caseLabors.Duration,\r\n    caseSlas.ExpiresOn AS SlaExpiresOn,\r\n    caseSlas.CompletedOn AS SlaCompletedOn,\r\n    caseSlas.State AS SlaState,\r\n    caseFcrKpis.FCRTarget,\r\n    caseFcrKpis.FCRState,\r\n    caseFcrKpis.FCRKpiType,\r\n    caseFdrKpis.FDRTarget,\r\n    caseFdrKpis.FDRState,\r\n    caseFdrKpis.FDRKpiType,\r\n    caseServiceLevel.servicelevel,\r\n    caseServiceLevel.friendlyname AS EntitlementDescription,\r\n    caseServiceLevel.EntitlementId,\r\n    caseServiceLevel.ServiceName,\r\n    caseServiceLevel.IsPublicSector,\r\n    CaseList.caseType PolicyCaseType,\r\n    0 AS unreadEmailsCount,\r\n    LastInboundCustomerEmail\r\nFROM CaseList\r\nLEFT OUTER JOIN emailinteractions ON emailinteractions.caseNumber = CaseList.CaseNumber\r\nLEFT OUTER JOIN caseLabors ON CaseList.casenumber = caseLabors.casenumber\r\nLEFT OUTER JOIN MSEG.QueryService.V1.SAP ON CaseList.caseTaskSAP = MSEG.QueryService.V1.SAP.SapId\r\nLEFT OUTER JOIN MSEG.QueryService.V1.SlaItem AS caseSlas ON CaseList.casenumber = caseSlas.casenumber AND caseSlas.SlaType = 'InitialResponse'\r\nLEFT OUTER JOIN caseFcrKpis ON CaseList.casenumber = caseFcrKpis.casenumber\r\nLEFT OUTER JOIN caseServiceLevel ON CaseList.casenumber = caseServiceLevel.casenumber\r\nLEFT OUTER JOIN caseFdrKpis ON CaseList.casenumber = caseFdrKpis.casenumber\r\nLEFT OUTER JOIN MSEG.QueryService.V1.Queue ON CaseList.RCID = MSEG.QueryService.V1.Queue.Id\r\noption (maxdop 1);"}'''
    header = {"Authorization": getSDToken(), "Content-Type": "application/json; charset=utf-8"}
    email_list = []
    for k,v in name_alias_mapping.items():
        email_list.append(v+"@microsoft.com")
    SEsFormartedEmail = "'" + "','".join(email_list) + "'"
    payload = payload.replace( "PlaceToReplace",SEsFormartedEmail )
    r = requests.post(url,
                      data=payload,
                      headers=header).text
    return json.loads(r)



def analyzeReport(json1):
    table_parameter_result = json1["table_parameters"][0]["table_parameter_result"]
    #print(table_parameter_result)
    for record in table_parameter_result:
        if record[QueryforAllSEOpeningCases_jsonMapper.AgentId.value - 1].split("@")[0].lower() in name_alias_mapping.values() or \
            record[QueryforAllSEOpeningCases_jsonMapper.AssignedTo.value - 1].split("@")[0].lower() in name_alias_mapping.values():
            if "CollaborationTasks" == record[QueryforAllSEOpeningCases_jsonMapper.CaseType.value - 1]:
                task_dict[record[QueryforAllSEOpeningCases_jsonMapper.AssignedTo.value - 1].split("@")[0].lower()] +=1
            elif "Case"== record[QueryforAllSEOpeningCases_jsonMapper.CaseType.value - 1]:
                case_dict[record[QueryforAllSEOpeningCases_jsonMapper.AgentId.value - 1].split("@")[0].lower()] += 1
    #临时修一下Adelaide的bug case
    if case_dict["v-lingjiewu"]-36 >0:
        case_dict["v-lingjiewu"]-=36
    for k,v in name_alias_mapping.items():
        total_dict[v] = case_dict[v]+task_dict[v]

def get_df_backlog():
    df_backlog = pd.DataFrame(
        columns=["Names", "Cases", "All Items"])
    for k,v in name_alias_mapping.items():
        row = {'Names': k, 'Cases': case_dict[v], 'All Items': total_dict[v]}
        df_backlog= df_backlog.append(row,ignore_index=True)
    return df_backlog

def getSDToken():
    url = "https://lucasstorageaccount.blob.core.windows.net/token/SDToken.txt"
    r = requests.get(url).text
    return r

def getNameMapping():
    url = "https://api.support.microsoft.com/v0/queryresult"
    payload = '''{"query":"SELECT Email, Name, FirstName, LastName, AadUpn, DomainAlias FROM MSEG.QueryService.V1.Agent_v1 WHERE AadUpn in (PlaceToReplace)"}'''
    header = {"Authorization":getSDToken(),"Content-Type": "application/json; charset=utf-8"}

    r = requests.post(url,
                      data=payload,
                      headers=header).text

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
def get_excel_data(fw):

    # 读取工作簿和工作簿中的工作表
    # df_backlog = pd.read_excel(backlog_excel,engine='openpyxl')
    # df_backlog = df_backlog.dropna(axis=1, how='all')
    # df_backlog = df_backlog[["Names","Cases","All Items"]]
    analyzeReport(getBacklog())
    df_backlog = get_df_backlog()
    print("------------backlog=-------------")
    print(df_backlog)
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
    df_monitoring_case = df_monitoring_case.loc[df_monitoring_case['FW'] == fw]
    print("------------df_monitoring_case=-------------")

    print(df_monitoring_case)
    # df_integration_case = pd.read_excel(case_excel,sheet_name='Integration',engine='openpyxl')
    # df_integration_case = df_integration_case.dropna(axis=1, how='all')
    # df_integration_case = df_integration_case.loc[df_integration_case['FW'] == fw]
    print("------------df_integration_case=-------------")

    # print(df_integration_case)
    # df_case = pd.concat([df_monitoring_case,df_integration_case])
    # df_case = df_case.dropna(axis=1, how='all')
    #
    # print(df_case)
    # 新建一个工作簿
    ret = pd.DataFrame(columns=['Case Volumn', 'Collaboration Task Volume', 'Escalation Task Volume',
                                "Follow-up Task Volume", "Rave AR",
                                "Weekly Total", "Engineers Backlog (No Rave)", "Engineers Backlog Case Volume",
                                "Engineers Backlog Task Volume"],
                       index=all_se)
    ret.loc[:, :] = 0

    ret_general = pd.DataFrame(columns=["PoD Name","Case Volume","CritSit Case Volume","ARR Case Volume","S500 Case Volume",
                                        "Collaboration Task Volume","Escalation Task Volume","Follow-up Task Volume",
                                        "Rave AR","Weekly Total","Engineers Backlog (No Rave)","Engineers Backlog Case Volume",
                                        "Engineers Backlog Task Volume"])
    ret_general.loc[0,"PoD Name"]="Monitoring"
    ret_general.loc[:, :] = 0
    ret_general.loc[0,"PoD Name"]="Monitoring"

    #ret_general.loc[1,"PoD Name"]="Integration"
    #general crisit case volumn
    temp_df_monitoring_crisit_case_this_week = df_monitoring_case[df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_monitoring_crisit_case_this_week = temp_df_monitoring_crisit_case_this_week[temp_df_monitoring_crisit_case_this_week['Severity'].str.lower().str.contains("a")]
    ret_general.loc[0, "CritSit Case Volume"] = temp_df_monitoring_crisit_case_this_week.shape[0]

    # temp_df_integration_crisit_case_this_week = df_integration_case[df_integration_case['Case/Task'].str.lower().str.contains("case")]
    # temp_df_integration_crisit_case_this_week = temp_df_integration_crisit_case_this_week[temp_df_integration_crisit_case_this_week['Severity'].str.lower().str.contains("a")]
    # ret_general.loc[1, "CritSit Case Volume"] = temp_df_integration_crisit_case_this_week.shape[0]
    # general ARR case volumn
    temp_df_monitoring_arr_case_this_week = df_monitoring_case[df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_monitoring_arr_case_this_week = temp_df_monitoring_arr_case_this_week[temp_df_monitoring_arr_case_this_week['ARR/Unified/Premier/Pro?'].str.lower().str.contains("arr")]
    ret_general.loc[0, "ARR Case Volume"] = temp_df_monitoring_arr_case_this_week.shape[0]

    # temp_df_integration_arr_case_this_week = df_integration_case[df_integration_case['Case/Task'].str.lower().str.contains("case")]
    # temp_df_integration_arr_case_this_week = temp_df_integration_arr_case_this_week[temp_df_integration_arr_case_this_week['ARR/Unified/Premier/Pro?'].str.lower().str.contains("arr")]
    # ret_general.loc[1, "ARR Case Volume"] = temp_df_integration_arr_case_this_week.shape[0]

    # general S500 case volumn
    temp_df_monitoring_s500_case_this_week = df_monitoring_case[df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    temp_df_monitoring_s500_case_this_week = temp_df_monitoring_s500_case_this_week[temp_df_monitoring_s500_case_this_week['Is S500?'].str.lower().str.contains("y",na=False)]
    ret_general.loc[0, "S500 Case Volume"] = temp_df_monitoring_s500_case_this_week.shape[0]






    # general case volumn
    temp_df_monitoring_case_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("case")]
    ret_general.loc[0, "Case Volume"] = temp_df_monitoring_case_this_week.shape[0]

    # temp_df_integration_case_this_week = df_integration_case[
    #     df_integration_case['Case/Task'].str.lower().str.contains("case")]
    # ret_general.loc[1, "Case Volume"] = temp_df_integration_case_this_week.shape[0]

    # general collab volumn
    temp_df_monitoring_collab_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("collab")]
    ret_general.loc[0, "Collaboration Task Volume"] = temp_df_monitoring_collab_this_week.shape[0]

    # temp_df_integration_collab_this_week = df_integration_case[
    #     df_integration_case['Case/Task'].str.lower().str.contains("collab")]
    # ret_general.loc[1, "Collaboration Task Volume"] = temp_df_integration_collab_this_week.shape[0]

    # general follow up volumn
    temp_df_monitoring_follow_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("follow")]
    ret_general.loc[0, "Follow-up Task Volume"] = temp_df_monitoring_follow_this_week.shape[0]

    # temp_df_integration_follow_this_week = df_integration_case[
    #     df_integration_case['Case/Task'].str.lower().str.contains("follow")]
    # ret_general.loc[1, "Follow-up Task Volume"] = temp_df_integration_follow_this_week.shape[0]
    # general rave volumn
    temp_df_monitoring_rave_this_week = df_monitoring_case[
        df_monitoring_case['Case/Task'].str.lower().str.contains("rave")]
    ret_general.loc[0, "Rave AR"] = temp_df_monitoring_rave_this_week.shape[0]

    # temp_df_integration_rave_this_week = df_integration_case[
    #     df_integration_case['Case/Task'].str.lower().str.contains("rave")]
    # ret_general.loc[1, "Rave AR"] = temp_df_integration_rave_this_week.shape[0]

    # general weekly total volumn
    ret_general.loc[0, "Weekly Total"] = ret_general.loc[0, "Case Volume"] + ret_general.loc[0, "Collaboration Task Volume"] \
                                        +ret_general.loc[0, "Follow-up Task Volume"]+ret_general.loc[0, "Rave AR"]

    # ret_general.loc[1, "Weekly Total"] = ret_general.loc[1, "Case Volume"] + ret_general.loc[1, "Collaboration Task Volume"] \
    #                                      + ret_general.loc[1, "Follow-up Task Volume"] + ret_general.loc[1, "Rave AR"]

    for se in all_se:
        print(" ------------- ",se,'-----------------')
        # if se in integration_se:
        #     df_case = df_integration_case
        # else:
        df_case = df_monitoring_case

        # case this week
        temp_df_case_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("case")]
        temp_df_case_this_week = temp_df_case_this_week[temp_df_case_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Case Volumn"] = temp_df_case_this_week.shape[0]
        # collab this week
        temp_df_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("collab")]
        temp_df_task_this_week = temp_df_task_this_week[temp_df_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Collaboration Task Volume"] = temp_df_task_this_week.shape[0]

        # Follow-up this week
        temp_df_followup_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("follow")]
        temp_df_followup_task_this_week = temp_df_followup_task_this_week[temp_df_followup_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Follow-up Task Volume"] = temp_df_followup_task_this_week.shape[0]
        # Rave this week
        temp_df_rave_task_this_week = df_case[df_case['Case/Task'].str.lower().str.contains("rave")]
        temp_df_rave_task_this_week = temp_df_rave_task_this_week[temp_df_rave_task_this_week["Case owner"].str.strip() == se]
        ret.loc[se, "Rave AR"] = temp_df_rave_task_this_week.shape[0]
        # weekly total
        ret.loc[se, "Weekly Total"] = ret.loc[se, "Case Volumn"]+ret.loc[se, "Follow-up Task Volume"] \
                                    +ret.loc[se, "Rave AR"] + ret.loc[se, "Collaboration Task Volume"]


        print(" ------------- ",se,'-----------------')
        # Engineers Backlog Case Volume
        temp_df_backlog_case = df_backlog[df_backlog["Names"].str.strip() == se]
        print(temp_df_backlog_case)
        try:
            ret.loc[se, "Engineers Backlog Case Volume"] += temp_df_backlog_case["Cases"].values[0]
        except IndexError:
            ret.loc[se, "Engineers Backlog Case Volume"] = 0

        # Engineers Backlog (No Rave)
        temp_df_backlog_all = df_backlog[df_backlog["Names"].str.strip() == se]
        try:
            ret.loc[se, "Engineers Backlog (No Rave)"] += temp_df_backlog_all["All Items"].values[0]
        except IndexError:
            ret.loc[se, "Engineers Backlog (No Rave)"] = 0
        #Engineers Backlog Task Volume
        ret.loc[se, "Engineers Backlog Task Volume"] =  ret.loc[se, "Engineers Backlog (No Rave)"] - ret.loc[se, "Engineers Backlog Case Volume"]

    ret_monitoring_vendor = ret.loc[monitoring_vendor_se]
    ret_monitoring_vendor.sort_values(["Weekly Total", "Engineers Backlog (No Rave)"], ascending=(False, False),
                                   inplace=True)

    ret_monitoring_fte = ret.loc[monitoring_fte_se]
    ret_monitoring_fte.sort_values(["Weekly Total","Engineers Backlog (No Rave)"], ascending=(False,False), inplace=True)

    ret_monitoring_tw = ret.loc[monitoring_tw_se]
    ret_monitoring_tw.sort_values(["Weekly Total", "Engineers Backlog (No Rave)"], ascending=(False, False),
                                   inplace=True)

    ret_monitoring_au = ret.loc[monitoring_au_se]
    ret_monitoring_au.sort_values(["Weekly Total", "Engineers Backlog (No Rave)"], ascending=(False, False),
                                   inplace=True)

    # ret_integration = ret.loc[integration_se]
    # ret_integration.sort_values("Engineers Backlog (No Rave)", ascending=False, inplace=True)
    # ret_integration.sort_values("Weekly Total", ascending=False, inplace=True)

    # general all backlog volumn
    ret_general.loc[0, "Engineers Backlog (No Rave)"] = ret_monitoring_fte["Engineers Backlog (No Rave)"].sum()+ \
                                                        ret_monitoring_vendor["Engineers Backlog (No Rave)"].sum()+ \
                                                        ret_monitoring_tw["Engineers Backlog (No Rave)"].sum()+ \
                                                        ret_monitoring_au["Engineers Backlog (No Rave)"].sum()
    # ret_general.loc[1, "Engineers Backlog (No Rave)"] = ret_integration["Engineers Backlog (No Rave)"].sum()

    # general backlog case volumn
    ret_general.loc[0, "Engineers Backlog Case Volume"] = ret_monitoring_fte["Engineers Backlog Case Volume"].sum() + \
                                                        ret_monitoring_vendor["Engineers Backlog Case Volume"].sum() + \
                                                        ret_monitoring_tw["Engineers Backlog Case Volume"].sum() + \
                                                        ret_monitoring_au["Engineers Backlog Case Volume"].sum()

    # ret_general.loc[1, "Engineers Backlog Case Volume"] = ret_integration["Engineers Backlog Case Volume"].sum()
    # general backlog task volumn
    ret_general.loc[0, "Engineers Backlog Task Volume"] = ret_monitoring_fte["Engineers Backlog Task Volume"].sum() + \
                                                        ret_monitoring_vendor["Engineers Backlog Task Volume"].sum()+ \
                                                        ret_monitoring_tw["Engineers Backlog Task Volume"].sum()+ \
                                                        ret_monitoring_au["Engineers Backlog Task Volume"].sum()
    # ret_general.loc[1, "Engineers Backlog Task Volume"] = ret_integration["Engineers Backlog Task Volume"].sum()

    return ret_general,pd.concat([ret_monitoring_fte,ret_monitoring_tw,ret_monitoring_au,ret_monitoring_vendor])




if __name__=="__main__":
    print(get_markdown4excel())

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
