import pandas as pd
import datetime
import os
from pymongo import MongoClient
from dotenv import load_dotenv
load_dotenv('opi.env')
def getdata():
    collect_opi = connect_to_mongo_db(str_db_name='Simulation',str_tb_name='OPITaskCtrl')
    col=[
        "_id",
        "form_id",
        "applicant",
        "boardNumber",
        "boardStage",
        "customer",
        "busItem",
        "product",
        "project_code",
        "project_name",
        "platform",
        "finished_dt"
    ]
    df_opi = pd.DataFrame(collect_opi.find({'status':'Finished'})).reset_index(drop=True)
    df_excel=pd.DataFrame()
    df_dashboard_detail=pd.DataFrame()
    unique_check=[]
    collect_requestform = connect_to_mongo_db(str_db_name='Simulation',str_tb_name='RequestForm')
    
    for index,row in df_opi.iterrows():
        df_request = pd.DataFrame(collect_requestform.find({
            '_id':row['form_id']
                            })).reset_index(drop=True)
        if 'Bus' not in df_request.keys():
            continue
        for re_index in range(len(df_request['Bus'])):
            # 要判斷在RequestForm中，對應 form_id 裡的 Status 是否為'Finished' 且 bus Item 為相同的(PDN or DDR CCT)
            if row['busItem']==df_request['Bus'][re_index][0]['Item'] and df_request['Bus'][re_index][0]['Status'] == 'Finished':
            # 確認目前 row['boardNumber']+'-'+row['boardStage'] 是否已經使用過
            # if row['boardNumber']+'-'+row['boardStage'] not in unique_check and row['_id']== ObjectId('659b5dcaad4012096a889c2f'):
                if row['boardNumber']+'-'+row['boardStage'] not in unique_check :  
                    one_excel=[]
                    #抓出 df_OPI 中，一樣版號和Stage的資料
                    df_opi_part=df_opi.loc[(df_opi['boardNumber']==row['boardNumber'])&(df_opi['boardStage']==row['boardStage'])]
                    #找到最新完成時間的資料
                    max_finished_dt=df_opi_part['finished_dt'].max()
                    #抓出最新完成DB資料的'tcl'資訊
                    df_lastest_pj = df_opi_part.loc[(df_opi_part['finished_dt']==max_finished_dt)]
                    dict_date = df_lastest_pj['projectSchedule'].iloc[0]
                    series_tcl = df_lastest_pj['tcl']
                    #抓出G/O日期
                    go_date = dict_date['gerberDate'].split('-')
                    go_yearly =  str(go_date[0])
                    go_monthly =  str(go_date[1])
                    go_day =  str(go_date[2])
                    go_season=[]
                    if go_monthly:
                        if go_monthly in ['01','02','03']:
                            go_season = 'Q1'
                        elif go_monthly in ['04','05','06']:
                            go_season = 'Q2'
                        elif go_monthly in ['07','08','09']:
                            go_season = 'Q3'
                        elif go_monthly in ['10','11','12']:
                            go_season = 'Q4'
                    #最新完成DB資料的版號和Stage的資料放進dashboard_list
                    one_excel=df_opi_part.loc[df_opi_part['finished_dt'] == max_finished_dt, col].reset_index(drop=True)
                    #計算Optimize_efficiency放進dashboard_list
                    list_tcl=series_tcl.iloc[0]
                    total_orig_value = sum(item['Cost']['Org_total_cost'] for item in list_tcl)
                    total_opt_value = sum(item['Cost']['Opt_total_cost'] for item in list_tcl)
                    if total_orig_value !=0:
                        efficiency = (total_orig_value-total_opt_value)*100/total_orig_value
                    else:
                        efficiency = 0
                    one_excel['yearly'] = str(go_yearly)
                    one_excel['quarterly'] = str(go_season)
                    one_excel['monthly'] = str(go_monthly)
                    one_excel['G/O Date'] = str(go_day)
                    cost_type_name = 'Cost type'
                    if efficiency == 0:
                        one_excel[cost_type_name] = 'Remain'
                    elif efficiency > 0:
                        one_excel[cost_type_name] = 'Saving'
                    else :
                        one_excel[cost_type_name] = 'Increase'
                    one_excel['Optimize_efficiency'] = efficiency
                    # ! 加入判斷opt_report/org_report status
                    opt_result = any('Fail' in d['report_result'] for d in list_tcl)
                    ori_result = any('Fail' in d['ori_report_result'] for d in list_tcl)
                    opt_status_name = 'Optimize Status'
                    if opt_result is True: 
                        one_excel[opt_status_name]='Fail'
                    else: one_excel[opt_status_name]='Pass'
                    sim_status_name = 'Simulation Status'
                    if ori_result is True: 
                        one_excel[sim_status_name]='Fail'
                    else: one_excel[sim_status_name]='Pass'
                    for item in one_excel.keys():
                        if one_excel[item][0] == '':
                            one_excel.at[0, item] = ''
                        else:
                            one_excel.at[0, item] = one_excel[item][0]
                    df_excel = pd.concat([df_excel,one_excel])
                    one_cost=[]
                    for index in range(len(list_tcl)):
                        one_cost= list_tcl[index]['Cost']
                        if pd.isna(one_cost['Org_total_cap']):
                            one_cost['Org_total_cap'] = 0
                        elif pd.isna(one_cost['Opt_total_cap']):
                            one_cost['Opt_total_cap'] = 0
                        elif pd.isna(one_cost['Org_total_cost']):
                            one_cost['Org_total_cost'] = 0
                        elif pd.isna(one_cost['Opt_total_cost']):
                            one_cost['Opt_total_cost'] = 0
                        elif pd.isna(one_cost['Cost_saving']):
                            one_cost['Cost_saving'] = 0
                        elif pd.isna(one_cost['Opt_efficiency']):
                            one_cost['Opt_efficiency'] = 0



                        one_cost['boardNumber']= str(row['boardNumber'])
                        one_cost['boardStage']= str(row['boardStage'])
                        one_cost['yearly'] = str(go_yearly)
                        one_cost['quarterly'] = str(go_season)
                        one_cost['monthly'] = str(go_monthly)
                        one_cost['G/O Date'] = str(go_day)
                        one_cost[sim_status_name] = str(list_tcl[index]['ori_report_result'])
                        one_cost[opt_status_name] = str(list_tcl[index]['report_result'])
                        #最新完成DB資料的版號和Stage的Cost資料放進dashboard_detail
                        df_dashboard_detail = pd.concat([df_dashboard_detail, pd.DataFrame.from_dict(one_cost, orient='index').T], ignore_index=True).reset_index(drop=True)
                    unique_check+=[row['boardNumber']+'-'+row['boardStage']]


          
    relist=[
        "_id",
        "form_id",
        "yearly",
        "quarterly",
        "monthly",
        "G/O Date",
        "applicant",
        "customer",
        "busItem",
        "product",
        "project_code",
        "project_name",
        "boardNumber",
        "boardStage",
        "platform",
        "Optimize_efficiency",
        "Cost type",
        'Optimize Status',
        'Simulation Status'
    ]

    df_excel=df_excel[relist]
    df_excel=df_excel.reset_index(drop=True)
    return df_excel,df_dashboard_detail
def connect_to_mongo_db(str_db_addr=os.getenv('DB_URL'),str_db_name="",str_tb_name="") :
    conn=MongoClient(str_db_addr)
    db = conn[str_db_name]
    collect = db[str_tb_name]
    return(collect)

# ---------------- read_cpu_info------------------------------
# Input: CPU_info (ex. Intel MTL-H)
# Output: 1.Vendor = 'Intel'
#         2.CPU_name = 'MTL'
#         3.Platform = 'H' (有可能不存在)

def read_cpu_info(cpu_info):
    cpu_info_split =  cpu_info.split(" ")
    cpu_target = ''
    cpu_type = ''
    platform = ''
    cpu_name = ''
    vendor = ''
    for split_name in  cpu_info_split:
        if split_name !='' and split_name in ['Intel','AMD','Qualcomm']:
            vendor = split_name
        elif split_name !='' and len(split_name.split("-")) >1 and 'TYPE' not in split_name.upper() and 'TARGET' not in split_name.upper():
            cpu_name = split_name.split("-")[0]
            platform = split_name.split("-")[1]
        elif split_name !='' and len(split_name.split("-")) ==1 and 'TYPE' not in split_name.upper() and 'TARGET' not in split_name.upper():
            cpu_name = split_name.split("-")[0]
        elif 'TYPE' in split_name.upper():
            cpu_type = split_name
        elif 'TARGET' in split_name.upper():
            cpu_target = split_name

    return vendor,cpu_name,platform,cpu_type,cpu_target

def get_ganttdata():
    tool_list=['OPI1','OPI2','OPI3','SPEED2000 #1','SPEED2000 #2']
    unfinished_hours=20 # ! 預設未完成使用時數
    collect_opi = connect_to_mongo_db(str_db_name='Simulation',str_tb_name='OPITaskCtrl')
    # ! 已完成/timeout/模擬中的資料都要抓進來
    df_opi = pd.DataFrame(collect_opi.find({'$or': [{'status': {'$in':['Finished','Unfinished']} },{'order':{'$gt':0},'status':{'$in':['Running']}}]})).reset_index(drop=True)
    df_data=pd.DataFrame()
    for index, row in df_opi.iterrows():
        vendor, cpu_name, platform, _, _ = read_cpu_info(row['platform'])
        df_data.at[index, 'Tools'] = str(row['license'])
        df_data.at[index, 'Task_Name'] = str(row['customer']) + ' ' + str(row['boardNumber'])
        df_data.at[index, 'Project_Code'] = str(row['project_code'])
        df_data.at[index, 'Project_Name'] = str(row['project_name'])
        df_data.at[index, 'PCBNO'] = str(row['project_code'])
        df_data.at[index, 'GO_Date'] = str(row['projectSchedule']['gerberDate'])
        df_data.at[index, 'ver'] = str(row['boardStage'])
        df_data.at[index, 'Item'] = str(row['busItem']) if str(row['busItem']) !='nan' else ''
        df_data.at[index, 'Vendor'] = vendor
        df_data.at[index, 'Platform'] = str(cpu_name + ' ' + platform) if cpu_name != '' else str(platform)
        df_data.at[index, 'Brand'] = str(row['customer'])
        df_data.at[index, 'ProductLine'] = str(row['product'])
        df_data.at[index, 'Status'] = str(row['status'])
        timestamp = row['current_opi_start_dt']
        date = datetime.datetime.fromtimestamp(timestamp)
        format_date=date.strftime("%Y/%m/%d %H:%M")
        df_data.at[index, 'Start_Date'] = format_date
        df_data.at[index, 'Sum_of_Duration'] = 0.00
        if 'finished_dt' in row.keys():
            sun_duration=round((row['finished_dt'] - row['current_opi_start_dt'])/86400,2)
            df_data.at[index, 'Sum_of_Duration'] = float(sun_duration) if sun_duration >=0 else 0.00
        df_data.loc[df_data['Status'] =='Unfinished', 'Sum_of_Duration'] = round(unfinished_hours/24,2)
        # ! 新增未有資料tool預設值
        cur_tools_list= df_data['Tools'].unique()
        for tool in tool_list:
            if tool not in cur_tools_list:
                new_row = {
                    'Tools': tool,
                    'Task_Name': '', 
                    'Project_Code': '',
                    'Project_Name':'',
                    'PCBNO':'',
                    'GO_Date':'',
                    'ver':'',
                    'Item':'',
                    'Vendor':'',
                    'Platform':'',
                    'Brand':'',
                    'ProductLine':'',
                    'Status':'',
                    'Sum_of_Duration':0,
                    'Start_Date':datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
                }  
                dfnew=pd.DataFrame(new_row, index=[0])
                df_data = pd.concat([df_data, dfnew], ignore_index=True).reset_index(drop=True)

    return df_data
if __name__ == '__main__' :

    df_excel,df_dashboard_detail= getdata()
    df_data = get_ganttdata()
    # 給 POWER BI 不需要輸出成 excel
    df_excel.to_excel(os.path.join(os.getcwd(),'Dashboard_list_data.xlsx'),index=False)
    df_dashboard_detail.to_excel(os.path.join(os.getcwd(),'Dashboard_detail_data.xlsx'),index=False)