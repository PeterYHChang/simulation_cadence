import subprocess
import time
import shutil
import re
import os
import bson
import json
import traceback
import psutil
import socket
import pandas as pd
import datetime
from package import utils, gentcl, DDRGentcl, mailNotify
from package import export_report as report
# ----------------------------- Send OSF mail ------------------------------ #
def osf_warning():
    """
    Sending OSF mail at setting time everyday
    """    
    if config['check_osf'].upper() != 'ON':
        return
    ct=int(time.time())
    warning_times = config['osf_warning_time']
    #! 10:00 everyday, Sending upload file remind 
    y = str(datetime.datetime.now().year)
    m = str(datetime.datetime.now().month).zfill(2)
    d = str(datetime.datetime.now().day).zfill(2)
    today = f"{y}-{m}-{d} {warning_times}"
    everydary = datetime.datetime.strptime(today,'%Y-%m-%d %H:%M').timestamp()
    if abs(everydary-ct) >= 5*60 :
        save_to_log(logPathName,f"Not At {warning_times} everyday,Do Not Sending OSF_Warning Mail")
        return True
    save_to_log(logPathName,"osf_warning...")
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName = DB_sheet)
    df_osf_task = pd.DataFrame(collect_tc.find({'status':'Initial_Fail','initialResult':{'$in':['NG_placement','NG_placement_pin','NG_Short_Net','NG_Net']}})).reset_index(drop=True)
    for _,row in  df_osf_task.iterrows():
        ini_date = datetime.datetime.strptime(row['initial_date'],'%Y-%m-%d').timestamp()
        if ct < ini_date:
            mailNotify.send_osf(row,row['initial_date'],2)
            continue
        mailNotify.send_osf(row,row['initial_date'],1)
    return True
# ----------------------------- Send Non Common Cap Used mail ------------------------------ #
def non_common_cap_used_warning():
    """
    Sending non_common_cap_used mail at setting time everyday
    """    
    if config['check_common_cap'].upper() != 'ON':
        return
    ct=int(time.time())
    warning_times = config['non_common_cap_warning_time']
    #! 10:00 everyday, Sending upload file remind 
    y = str(datetime.datetime.now().year)
    m = str(datetime.datetime.now().month).zfill(2)
    d = str(datetime.datetime.now().day).zfill(2)
    today = f"{y}-{m}-{d} {warning_times}"
    everydary = datetime.datetime.strptime(today,'%Y-%m-%d %H:%M').timestamp()
    if abs(everydary-ct) >= 5*60 :
        save_to_log(logPathName,f"Not At {warning_times} everyday,Do Not non_common_cap_used_warning Mail")
        return True
    save_to_log(logPathName,"non_common_cap_used_warning...")
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName = DB_sheet)
    df_osf_task = pd.DataFrame(collect_tc.find({'status':'Initial_Fail','initialResult':'Non_Common_Cap_Used'})).reset_index(drop=True)
    for _,row in  df_osf_task.iterrows():
        ini_date = datetime.datetime.strptime(row['initial_date'],'%Y-%m-%d').timestamp()
        if ct < ini_date:
            mailNotify.send_non_common_cap_used(row,row['initial_date'],2)
            continue
        mailNotify.send_non_common_cap_used(row,row['initial_date'],1)
    return True
# ----------------------- Over_ini_date --------------------- #(OK)
# PDN/ DDR CCT 共用
#  Function : 
#  1. check initial_date of task is over. -> Cancel
#  2. The Status of over-schedule task changes to Cancel 
def over_ini_date():
    try:
        save_to_log(logPathName,"Start Running Over_ini_date()...")
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        # 以$regex 查詢包含有 前方任意文字 + busItem 結尾之單據 
        list_unspec = list(collect_tc.find({'license':'Unspecified','busItem':{'$regex':'.*'+busItem+'$'},'order':{'$gt':0}}).sort([("order", 1)]))
        if len(list_unspec) == 0:
            save_to_log(logPathName,'No Scheduled Request')
            return True
        for idx in range(0,len(list_unspec)):
            item = list_unspec[idx]
            save_to_log(logPathName,f"check form_id : {item['form_id']}...")
            ct=int(time.time())
            ini_run_timestamp = datetime.datetime.strptime(item['initial_date'],'%Y-%m-%d').timestamp()
            if  ct > ini_run_timestamp :
                if check_ini_status(item):
                    save_to_log(logPathName,f"form_id : {item['form_id']} is initial successfully.")
                    continue
                save_to_log(logPathName,f"form_id : {item['form_id']} Schedule is over Initial date.")
                collect_tc.update_one(
                                    {'form_id': item['form_id']},
                                    {'$set': {'status': 'Cancel', 'order': 0}})
                notify_type = 'Cancel'
                mailNotify.opiMailNotify (notify_type, item, config_dict)
                save_to_log(logPathName,"Cancel Mail SENT" )
                for index in range(idx+1,len(list_unspec)):
                    list_unspec[index]['order'] = list_unspec[index]['order']-1
                    collect_tc.update_one(
                                    {'form_id': list_unspec[index]['form_id']},
                                    {'$set': {'order': list_unspec[index]['order']}})                                                                
            else:
                save_to_log(logPathName,f"form_id : {item['form_id']} Schedule is before Initial date - 7days.")
        save_to_log(logPathName,"Running Over_ini_date Done")
        return True
    except Exception as _ :
        save_to_log(logPathName,f"Function Over_ini_date() Error:/n {traceback.format_exc()}")
        return False
# -------------------------- reduce complexity -------------------------- #
# 只有 GPU PDN 上傳檔案後會直接 item['initialResult'] == 'Succeed'，且檔案只存在 SIM2
# 其他 Auto/Manual 都會丟到 Sim2 and Sim3/Sim5，照樣執行 Initial 流程
def check_ini_status(item)->bool:
    if 'initialResult' in item.keys():
        return item['initialResult'] == 'Succeed'
    return False

# ======================================================================= #
    
# ----------------------- Over_Sim_date --------------------- #(OK)
# PDN/ DDR CCT 共用
#  Function : 
#  1. check sim_start_date of task is over. -> Cancel
#  2. The Status of over-schedule task changes to Cancel 
def over_sim_date():
    try:
        save_to_log(logPathName,"Start Running Over_sim_date()...")
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        list_unspec = list(collect_tc.find({'license':'Unspecified','busItem':{'$regex':'.*'+busItem+'$'},'order':{'$gt':0}}).sort([("order", 1)]))
        if len(list_unspec) == 0:
            save_to_log(logPathName,'No Scheduled Request')
            return True
        for idx in range(0,len(list_unspec)):
            item = list_unspec[idx]
            save_to_log(logPathName,f"check form_id : {item['form_id']}...")
            ct=int(time.time())
            sim_run_timestamp = datetime.datetime.strptime(item['sim_start_date'],'%Y-%m-%d').timestamp()
            # 請注意此處，若為 PDN 則超過設定的 'sim_start_date' 時間則 Cancel; 若為 'GPU PDN','Manual PDN','Manual DDR CCT' 則是超過 targetDate
            if  (ct > sim_run_timestamp and 'AUTO' in item['busItem'].upper()) or \
            (ct > datetime.datetime.strptime(item['projectSchedule']['targetDate'],'%Y-%m-%d').timestamp() and 'MANUAL' in item['busItem'].upper()):
                save_to_log(logPathName,f"form_id : {item['form_id']} Schedule is over sim_start_date.")
                collect_tc.update_one(
                                    {'form_id': item['form_id']},
                                    {'$set': {'status': 'Cancel', 'order': 0}})
                notify_type = 'Cancel'
                mailNotify.opiMailNotify (notify_type, item, config_dict)
                save_to_log(logPathName,"Cancel Mail SENT" )
                for index in range(idx+1,len(list_unspec)):
                    list_unspec[index]['order'] = list_unspec[index]['order']-1
                    collect_tc.update_one(
                                    {'form_id': list_unspec[index]['form_id']},
                                    {'$set': {'order': list_unspec[index]['order']}})
            else:
                save_to_log(logPathName,f"form_id : {item['form_id']} Schedule is before Simulation date.")
        save_to_log(logPathName,"Running Over_sim_date Done")
        return True
    except Exception as _:
        save_to_log(logPathName,f"Function Over_sim_date() Error:/n {traceback.format_exc()}")
        return False
    
# ----------------------- Conflict remind --------------------- #(OK)
# PDN/ DDR CCT 共用
#  Function : 
#  1. capture all conflict tasks
#  2. send mail to remind admin dealing
def conflict_remind():
    try:
        save_to_log(logPathName,"Start Conflict_remind...")
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        df_conflict = pd.DataFrame(collect_tc.find({'status':'Conflict','busItem':{'$regex':'.*'+busItem+'$'}})).reset_index(drop=True)
        if df_conflict.shape[0] != 0:
            mailNotify.conflict_mail(df_conflict, config_dict)
            save_to_log(logPathName,"Conflict_mail SENT")
            for _,row in df_conflict.iterrows():
                notify_type = 'Conflict'
                mailNotify.opiMailNotify (notify_type, row, config_dict)
                save_to_log(logPathName,"Conflict Mail SENT" )
        else:
            save_to_log(logPathName,"No Conflict Task!")
        return True
    except Exception as _:
        save_to_log(logPathName,f"Error : \n {traceback.format_exc()}")
        return False

# ----------------------- Schedule change check --------------------- #(OK)
# PDN/ DDR CCT 共用
#  Function : 
#  1. Scan Schedule_Notify
#  2. send mail
#  3. delet Schedule_Notify content if mail sent
def schedule_change():
    try:
        save_to_log(logPathName,"Schedule change check...")
        collect_sn = utils.ConnectToMongoDB(strDBName="Simulation",strTableName = "Schedule_Notify")
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName = DB_sheet)
        df_sn = pd.DataFrame(collect_sn.find()).reset_index(drop=True)
        if df_sn.shape[0] != 0:
            save_to_log(logPathName,f"Get Tasks changed schedule :\n {df_sn}")
            for _,row in df_sn.iterrows():
                one_task=collect_tc.find_one({"form_id":row['form_id']})
                mailNotify.opiMailNotify('simulation_change',one_task, config_dict)
                save_to_log(logPathName,f"Task changed scheduled : {row['form_id']}")
            collect_sn.delete_many({})
        else:
            save_to_log(logPathName,"No Task changed scheduled.")
        return True
    except Exception as _:
        save_to_log(logPathName,f"Error :\n {traceback.format_exc()}")
        return False

# ======================================== 以下為執行initial相關 Function ============================= #
# ----------------------- initial check_old --------------------- #(OK)
# PDN/ DDR CCT(SPEED2000) 放一起
#  Function : 
#  1.   capture all unspecified scheduled tasks
#  2-1. check if all unspecified scheduled tasks exist files
#  2-2. check if layout API meets config_dict['Layout_threshold'] progress 
#  3-1. send mail to remind uploading files when item 2-1 is false on "initial_date" before "sim_start_date"
#  3-2. send mail to remind layout API progress when item 2-2 is false on "initial_date" before "sim_start_date"
#  4.   when 2-1/2-2 are True, Do Initial
# def Initial_check_old():
#     try:
#         SaveToLog(logPathName,"Start Running Initial_check...")
#         collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#         list_unspec = list(collect_tc.find({'license':'Unspecified','busItem':busItem,'order':{'$gt':0}}).sort([("order", 1)]))
#         if list_unspec:
#             ini_idx=[]
#             for idx in range(0,len(list_unspec)):
#                 item = list_unspec[idx]
#                 SaveToLog(logPathName,f"check form_id : {item['form_id']}...")
#                 ct=int(time.time())
#                 initial_timestamp = datetime.datetime.strptime(item['initial_date'],'%Y-%m-%d').timestamp()
#                 daylimit = config_dict['notify_mail']*24*60*60
#                 if (initial_timestamp  >= ct and ct > initial_timestamp - daylimit) or (item['debug'] == True):
#                     SaveToLog(logPathName,f"form_id : {item['form_id']} Schedule OK.")
#                     if layout_api_check(item):
#                         SaveToLog(logPathName,f"form_id : {item['form_id']} Power check OK.")
#                         if upload_check(item):
#                             ini_idx.append(idx)
#                             SaveToLog(logPathName,f"form_id : {item['form_id']} Document Prepared.")
#                         else:
#                             SaveToLog(logPathName,f"XXXXX form_id : {item['form_id']} Document Doesn't Upload XXXXX.")
#                             #! 10:00 everyday, Sending upload file remind 
#                             y = str(datetime.datetime.now().year)
#                             m = str(datetime.datetime.now().month).zfill(2)
#                             d = str(datetime.datetime.now().day).zfill(2)
#                             if busItem == auto_bus[0]:
#                                 today = f"{y}-{m}-{d} 10:00"
#                             elif busItem == auto_bus[1]:
#                                 today = f"{y}-{m}-{d} 10:15"
#                             everydary_10 = datetime.datetime.strptime(today,'%Y-%m-%d %H:%M').timestamp()
#                             if abs(everydary_10-ct) <= 5*60 and ((initial_timestamp  >= ct and ct > initial_timestamp - daylimit) or item['debug'] == True):
#                                 SaveToLog(logPathName,f"10:00 everyday, Sending Upload_Remind mail")
#                                 #! Send upload file remind
#                                 notify_type = 'Upload_Remind'
#                                 mailNotify.opiMailNotify (notify_type, item, config_dict)
#                                 SaveToLog(logPathName,"Upload_Remind SENT")
#                             else:
#                                 SaveToLog(logPathName,f"It's not 10:00 everyday... Skip Upload_Remind Mail...")
#                     else:
#                         # SaveToLog(logPathName,f"XXXXX form_id : {item['form_id']} Layout Info Doesn't Meet {config_dict['Layout_threshold']} XXXXX")
#                         #! Send LayoutAPI  remind
#                         notify_type = 'POWER_API_ERROR'
#                         # notify_type = 'Layout_API_ERROR'
#                         mailNotify.opiMailNotify (notify_type, item, config_dict)
#                         SaveToLog(logPathName,"Layout_API_ERROR SENT" )
#                 else:
#                     SaveToLog(logPathName,f"Waiting for initial day of form_id : {item['form_id']}.")

#             for idx in range(0,len(list_unspec)):
#                 # 檢查 Auto DDR CCT 是否是 SPEED2000 #1 還是 SPEED2000 #2
#                 if busItem == auto_bus[1]:
#                     if LisenseChoose(busItem) != license_name:
#                         SaveToLog(logPathName,f"LisenseChoose doesn't match {license_name}")
#                         break
#                     else:
#                         SaveToLog(logPathName,f"LisenseChoose matches {license_name}")
#                 item = list_unspec[idx]
#                 if idx in ini_idx :
#                     if item['initialResult'] != 'Succeed':
#                         SaveToLog(logPathName,f"form_id : {item['form_id']} Initial Result: {item['initialResult']}")
#                         SaveToLog(logPathName,f"do initial form_id : {item['form_id']}...")
#                         if initial_pdn(item):
#                             SaveToLog(logPathName,f"form_id : {item['form_id']} Initial Successed")
#                         else:
#                             SaveToLog(logPathName,f"XXXXX form_id : {item['form_id']} Initial Fail XXXXX")
#                     else:
#                         SaveToLog(logPathName,f"form_id : {item['form_id']} is Waiting for Starting Simulation...")
#                 else:
#                     SaveToLog(logPathName,f"XXXXX Blocked at form_id : {item['form_id']} to Do Initial XXXXX")
#         else:
#             SaveToLog(logPathName,"No Scheduled Request")
#         return True
#     except:
#         SaveToLog(logPathName,f"Error message: \n {traceback.format_exc()}")
#         return False

# ----------------------- initial check_new_20241217 --------------------- #
# DDR changed to use Power SI
# Runtime 30分鐘執行一次
def initial_check_new():
    try:
        save_to_log(logPathName,"Start Running Initial_check...")
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        list_unspec = list(collect_tc.find({'license':'Unspecified','busItem':{'$regex':'.*'+busItem+'$'},'order':{'$gt':0},'debug':False,'status':{'$in':['Scheduled','Initial_Fail']}}).sort([("order", 1)]))
        # debug form 的狀態為 order 0, 僅能靠 'license':'Unspecified',debug':True,'status':{'$in':['Scheduled','Initial_Fail']} 查詢
        list_debug = list(collect_tc.find({'license':'Unspecified','busItem':{'$regex':'.*'+busItem+'$'},'debug':True,'status':{'$in':['Scheduled','Initial_Fail']}})) 
        list_run = list(collect_tc.find({'busItem':{'$regex':'.*'+busItem+'$'},'status':'Running'}))
        if len(list_unspec) == 0 and  len(list_debug) == 0:
            save_to_log(logPathName,"No Scheduled Request")
            return True
        # 先處理 debug task
        order = len(list_run)
        for idx in range(len(list_debug)):
            item = list_debug[idx]
            save_to_log(logPathName,f"get debug form_id : {item['form_id']}")
            if False in create_check_list(item):
                save_to_log(logPathName,f"XXX debug form_id : {item['form_id']} Did not Prepare XXX")
                continue
            order +=1
            handle_initial(item)
            item['order']=order
            collect_tc.update_one({'form_id':item['form_id']},
                                {'$set':item})
        do_initial = True
        ini_idx=[]
        for idx in range(0,len(list_unspec)):
            item = list_unspec[idx]
            save_to_log(logPathName,f"check form_id : {item['form_id']}...")
            check_every_list = create_check_list(item)
            if False not in check_every_list:
                ini_idx.append(idx)
                save_to_log(logPathName,f"form_id : {item['form_id']} was Well-Prepared.")
            else:
                save_to_log(logPathName,f"Not Ready of form_id : {item['form_id']}.")
                # SaveToLog(logPathName,f"Not Ready Item- 早於 initial time/ 介於 initial time 前 daylimit 天/ PowerAPI_check/ 是否上傳資料 : {check_every_list}.")

            # 會一同執行 Initial 直到被某單 Task 卡住後停止
            if idx in ini_idx and do_initial:
                handle_initial(item)
            else:
                do_initial = False
            if not do_initial:
                save_to_log(logPathName,f"XXXXX Blocked form_id : {item['form_id']} to Do Initial XXXXX")
                break
        for idx in range(0,len(list_unspec)):
            item = list_unspec[idx]
            order +=1
            item['order']=order
            collect_tc.update_one({'form_id':item['form_id']},
                                {'$set':item})
        save_to_log(logPathName,"Running Initial_check Done")
        return do_initial
    except Exception as _:
        save_to_log(logPathName,f"Error message: \n {traceback.format_exc()}")
        return False
# ---------------------------- create_check_list -------------------------#
# Function:
# 1. 是否符合日期範圍 initial_timestamp 前 daylimit 天到 initial_timestamp 期間
# 2. LayoutAPI_check == True
# 3. upload_check == True
# 或"是否是 debug request"
# Input:
# 1. Task dict.
# Output:
# 1. True/ False
def create_check_list(item)-> list:
    # 確認是否室 Succeed 是則返回 [True]
    if 'initialResult' in item:
        if item['initialResult'] == 'Succeed':
            return [True]
    ct=int(time.time())
    initial_timestamp = datetime.datetime.strptime(item['initial_date'],'%Y-%m-%d').timestamp()
    daylimit = config_dict['notify_mail']*24*60*60
    ch_list = [initial_timestamp  >= ct, ct > initial_timestamp - daylimit, layout_api_check(item), upload_check(item)]
    # Send Uploaded Reminded Mail
    if not ch_list[-1]:
        #! 10:00 everyday, Sending upload file remind 
        y = str(datetime.datetime.now().year)
        m = str(datetime.datetime.now().month).zfill(2)
        d = str(datetime.datetime.now().day).zfill(2)
        mail_time_set ={
                        "PDN":f"{y}-{m}-{d} 10:00",
                        "DDR CCT":f"{y}-{m}-{d} 10:15"
                        }
        today = mail_time_set[busItem]
        everydary_10 = datetime.datetime.strptime(today,'%Y-%m-%d %H:%M').timestamp()
        if abs(everydary_10-ct) <= 5*60 and (initial_timestamp  >= ct and ct > initial_timestamp - daylimit):
            save_to_log(logPathName,"10:00 everyday, Sending Upload_Remind mail")
            #! Send upload file remind
            notify_type = 'Upload_Remind'
            mailNotify.opiMailNotify (notify_type, item, config_dict)
            save_to_log(logPathName,"Upload_Remind SENT")
        else:
            save_to_log(logPathName,"It's not 10:00 everyday... Skip Upload_Remind Mail...")
    if item['debug']:
        return [True]
    else:
        return ch_list
# Redueced Cognitive Complexity
def handle_initial(item) -> bool:
    if 'initialResult' not in item:
        save_to_log(logPathName,f"form_id : {item['form_id']} is Waiting for Upload Files...")
        return False
    if item['initialResult'] == 'Succeed':
        save_to_log(logPathName,f"form_id : {item['form_id']} is Waiting for Starting Simulation...")
        return True
    save_to_log(logPathName,f"do initial form_id : {item['form_id']}...")
    if busItem =='PDN':
        if initial_pdn(item):
            save_to_log(logPathName,f"form_id : {item['form_id']} Initial Successed")
            return True
        else:
            save_to_log(logPathName,f"XXXXX form_id : {item['form_id']} Initial Fail XXXXX")
    else:
        if initial_psi(item):
            save_to_log(logPathName,f"form_id : {item['form_id']} Initial Successed")
            return True
        else:
            save_to_log(logPathName,f"XXXXX form_id : {item['form_id']} Initial Fail XXXXX")
    return False
# ---------------------------- upload check ----------------------------- #
# Function:
# 1. 檢查Key中是否有 'filePath'
# Input:
# 1. Task dict.
# Output:
# 1. True/ False
def upload_check(item)-> bool:
    return 'filePath' in item.keys()

# ---------------------------- LayoutAPI check ---------------------------- #
# PDN/ DDR CCT 一樣(已經移除，所以只會回傳 True)
# Function:
# 1.Vendor == Intel -> 檢查 power API > 空 Power API data
# 2. Vendor == Intel -> 檢查 power check progress > config_dict['Layout_threshold']
# 3. Vendor == Qualcomm / DDR CCT-> skip
def layout_api_check(item : dict) -> bool :
    return True
    # Vendor,CPU_name,Platform,CPU_type,CPU_Target=utils.read_CPU_info(item['platform'])
    # if Vendor.upper() == "INTEL" and item['busItem'] =='PDN':
    #     board_number = item['boardNumber']
    #     board_version = item['boardStage']
    #     ### get power information (TODO: Intel)
    #     try :
    #         SaveToLog(logPathName,"Check Power API")
    #         power_info = utils.getPowerInfo(board_number, board_version)
    #         df_power = pd.DataFrame(power_info)
    #         if len(df_power)==0:
    #             SaveToLog(logPathName,"POWER API NO DATA")
    #         else:
    #             SaveToLog(logPathName,"Get POWER API DATA")
    #         if ('OutputNet2' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns):
    #             raise NameError('POWER_API_NO_DATA')
    #         return True
    #     except Exception as e:
            #! Need Insert read Layout API Info !!!!!!!!!!!!!
            # SaveToLog(logPathName,"Check Layout API")
            # Layout_info = utils.LayoutAPIInfo(board_number, board_version)
            # # All Power check process >0.85 ---> Pass, if No, Sending mail
            # if len(Layout_info)!=0:
            #     if all(value >= config_dict['Layout_threshold'] for value in Layout_info.values()):
            #         SaveToLog(logPathName,f"Power check Pass")
            #         return True
            #     else:
            #         SaveToLog(logPathName,f"XXXXX Power check fail XXXXX")
            #         return False
            # else:
            #     SaveToLog(logPathName,f"XXXXX No LayoutAPI Info XXXXX")
            #     return False
            # SaveToLog(logPathName,f"LayoutAPI_check not ready ...")
            # return False
    # else:
    #     SaveToLog(logPathName,f"{Vendor} {item['busItem']} Skips LayoutAPI_check")
    #     return True

# ------------------------------- initial ---------------------------------- #
# PDN/ DDR CCT (SPEED2000) 放一起
# Function:
# 1. do Initial
# Input:
# 1. Task dict.
# Output:
# 1. True/ False
def initial_pdn(insert_data) -> bool:
    try:
        save_to_log(logPathName,"Call Function Initial_pdn")
        board_number = insert_data['boardNumber']
        board_version = insert_data['boardStage']
        path_dict=insert_data['filePath']
        # 讀取 CPU Platform
        vendor,cpu_name,platform,cpu_type,cpu_target=utils.read_CPU_info(insert_data['platform'])
        save_to_log(logPathName,f"parse Platform result: {vendor}-{cpu_name}-{platform}-{cpu_type}-{cpu_target}")
        save_to_log(logPathName,"Read %s_Config.json."%vendor)
        cpu_config_dir=os.path.join(os.getcwd(), 'CPU_config',vendor+'_'+cpu_name+'.json')
        cpu_config =utils.__read_config_file(cpu_config_dir)
        save_to_log(logPathName,"Reading %s_Config.json Done."%vendor)
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        # ! ------------------------------------
        license_info_dict=__find_license_info()
        if not license_info_dict :
            save_to_log(logPathName,'read license info failed.')
            return False
        save_to_log(logPathName,'Read license info succeeded.')
        # ------------------------- Choose License and run Cadence with print_list ---------------------------- #
        license_name_list=list(license_info_dict.keys())
        check_license=[]
        for ini_license_name in license_name_list:
            check_license.append(__check_license(ini_license_name))
        # ! 先挑第一個license
        first_license_name = next(iter(license_info_dict))
        use_license=license_info_dict[first_license_name]
        true_indexes = [index for index, value in enumerate(check_license) if value]
        b_license=False
        if len(true_indexes) == 0:# ! 目前沒有license
            save_to_log(logPathName," **** no available_license for initial ****")
            raise NameError('NO_LICENSE')
        use_license=license_info_dict[license_name_list[true_indexes[0]]]
        b_license=True
        save_to_log(logPathName," Use license %s" %(license_name_list[true_indexes[0]]))

        #* =============================================== *#
        save_to_log(logPathName," running gentcl.stackup() to create print_list" )
        check_print_list = gentcl.stackup(path_dict)
        if not check_print_list :
            raise NameError('PARSE_ERROR')
        #* =============================================== *#
        if not b_license:
            save_to_log(logPathName," **** no available_license for initial ****")
            raise NameError('NO_LICENSE')
        save_to_log(logPathName,"available_license for initial: "+use_license)
        save_to_log(logPathName,"print_list: "+path_dict['print_list'])
        result=subprocess.run([config_dict['Cad_site'],use_license ,'-tcl', path_dict['print_list']], capture_output=True)
        save_to_log(logPathName,"subprocess result: "+str(result.returncode))
        
        #===========================================  Layout API/ Power API 跳過 外來更改為 DDE 資料===================================================================#
        save_to_log(logPathName,"Check API")
        # df_power = pd.DataFrame({})
        # Layout_info = {}
        # if Vendor.upper() == 'INTEL' and busItem.upper() == 'PDN':
        #     ### get power information (TODO: Intel)
        #     try :
        #         SaveToLog(logPathName,"Check Power API")
        #         power_info = utils.getPowerInfo(board_number, board_version)
        #         df_power = pd.DataFrame(power_info)
        #         if len(df_power)==0:
        #             SaveToLog(logPathName,"POWER API NO DATA")
        #         else:
        #             API_choose = 'POWER_API'
        #             SaveToLog(logPathName,"Get POWER API DATA")
        #         if ('OutputNet2' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns):
        #             raise NameError('POWER_API_NO_DATA')
        #     except Exception as e:
        #         notify_type = 'POWER_API_ERROR'
        #         mailNotify.opiMailNotify (notify_type, insert_data,config_dict)
        #         SaveToLog(logPathName," POWER_API_ERROR MAIL SEND" )
        #         raise NameError('POWER_API_ERROR')
        #         #! Need Insert read Layout API Info !!!!!!!!!!!!!
        #         # SaveToLog(logPathName,"Check Layout API")
        #         # Layout_info = utils.LayoutAPIInfo(board_number, board_version)
        #         # # All Power check process >0.85 ---> Pass, if No, Sending mail
        #         # if all(value >= config_dict['Layout_threshold'] for value in Layout_info.values()):
        #         #     API_choose = 'Layout_API'
        #         #     SaveToLog(logPathName,f"Power check Pass")
        #         # else:
        #         #     SaveToLog(logPathName,"XXXXX Layout_API_ERROR XXXXX")
        #         #     SaveToLog(logPathName,f"Layout Info :\n {Layout_info}")
        #         #     notify_type = 'Layout_API_ERROR'
        #         #     mailNotify.opiMailNotify (notify_type, insert_data,config_dict)
        #         #     SaveToLog(logPathName,"Layout_API_ERROR SENT" )
        #         #     raise NameError('Layout_API_ERROR')
        # else:
        #     API_choose = 'SKIP'
        #     SaveToLog(logPathName,f"{Vendor} {busItem} Skips LayoutAPI_check")
        # SaveToLog(logPathName,"Check API END")
        #==============================================================================================================#
        # -------------------------------- Parse DKDF.csv ------------------------------ #
        try:
            data_dkdf = utils.parseStackup(path_dict['dkdf'])
            save_to_log(logPathName,"parseStackup Done")
        except Exception as _:
            save_to_log(logPathName,f"{traceback.format_exc()}")
            raise NameError('PARSE_ERROR_DKDF')
        #==============================================================================================================#
        # -------------------------------- gen material cmx ------------------------------ #
        try :
            tree = utils.modifyMaterial(data_dkdf['layer'], path_dict['material_org'], data_dkdf['data'])
            tree.write(path_dict['material'])
            save_to_log(logPathName,"modifyMaterial Done")
        except Exception as _:
            raise NameError('GENCMX_ERROR')
        #==============================================================================================================#
        # -------------------------------- gen stackup csv ------------------------------ #
        try :
            modified_stackup = utils.modifyStackup(data_dkdf['layer'], path_dict['stackup_org'])
            modified_stackup.to_csv(path_dict['stackup'], index=False)
            save_to_log(logPathName,"modifyStackup Done")
        except Exception as e:
            save_to_log(logPathName,f"GENSTACKUP_ERROR Error: {e}" )
            raise NameError('GENSTACKUP_ERROR')
        #==============================================================================================================#
        #  判斷 PDN 是否在目前的 busItem中
        # if busItem in insert_data['busItem'].upper():
        # -------------------------------- Parse VRM.txt ------------------------------ #
        try :
            if cpu_config['vrm'] != '':
                df_vrm = utils.parseVRM(path_dict['vrm'])
                save_to_log(logPathName,"parseVRM Done")
            else:
                df_vrm = pd.DataFrame()
                save_to_log(logPathName,"VRM NO DATA")
        except Exception as _:
            raise NameError('VRM_ERROR')
    #==============================================================================================================#
    # ------------------------------------ OPI Gen TCL ---------------------------------- #
        try :
            tcl_list=[]
            save_to_log(logPathName,"Gentcl Start")
            if vendor.upper() =='INTEL':
                save_to_log(logPathName,"Choose Intel gentcl")
                tcl_list,list_net,dict_bom_cap = gentcl.opi(path_dict,df_vrm,vendor,cpu_name,platform,cpu_type,cpu_target,config_dict)
            elif vendor.upper() == 'QUALCOMM':
                save_to_log(logPathName,"Choose Qualcomm gentcl")
                tcl_list,list_net,dict_bom_cap = gentcl.opi(path_dict,df_vrm,vendor,cpu_name,platform,cpu_type,cpu_target,config_dict)
            elif vendor.upper() == 'AMD':
                tcl_list,list_net,dict_bom_cap = gentcl.opi(path_dict,df_vrm,vendor,cpu_name,platform,cpu_type,cpu_target,config_dict)
            save_to_log(logPathName,"Gentcl End")
        except Exception as _:
            raise NameError('GEN TCL_ERROR')
        #==============================================================================================================#
        # else:
        #==============================================================================================================#
        # # ------------------------------------- DDR Gen TCL ---------------------------------- # (PSI 不會用到)
        #     try :
        #         SaveToLog(logPathName,"Gentcl Start")
        #         tcl_list = DDRGentcl._DDRGentcl(path_dict,board_number,board_version,vendor,insert_data['DDRModule'],insert_data['Mapping'])
        #         SaveToLog(logPathName,"Gentcl End")
        #     except Exception as _:
        #         notify_type = 'GEN TCL_ERROR'
        #         mailNotify.opiMailNotify (notify_type, insert_data,config_dict)
        #         SaveToLog(logPathName,"GEN TCL_ERROR MAIL SENT" )
        #         raise NameError('GEN TCL_ERROR')
        #==============================================================================================================#
        # ---------------------------- check osf *20250430 add----------------------- #
        # 由產生tcl 時，抓取 all nets and all components on the net
        dict_ng = {}
        insert_data['NG_placement']=''
        insert_data['NG_placement_pin']=''
        insert_data['NG_Short_Net']=''
        insert_data['NG_Net']=''
        if config_dict["check_osf"].upper() =="ON":
            dict_ng = utils.get_dde_ng_data(board_number,board_version)
        for key,lis_bom_cap in dict_bom_cap.items():
            for bom_cap in lis_bom_cap:
                if utils.check_in_ng_placement(bom_cap,dict_ng):
                    insert_data['NG_placement'] = key
                    save_to_log(logPathName,f"Check NG_placement: Net {key} NG Cap: {bom_cap}")
                    raise NameError('NG_placement')
                if utils.check_in_ng_placement_pin(bom_cap,dict_ng):
                    insert_data['NG_placement_pin'] = key
                    save_to_log(logPathName,f"Check NG_placement_pin: Net {key} NG Cap: {bom_cap}")
                    raise NameError('NG_placement_pin')

        for net in list_net:
            if utils.check_in_short_net(net,dict_ng):
                insert_data['NG_Short_Net'] = net
                raise NameError('NG_Short_Net')
            if utils.check_in_ng_net(net,dict_ng):
                insert_data['NG_Net'] = net
                raise NameError('NG_Net')
        #==============================================================================================================#
        # ---------------------------- check common cap *20250430 add----------------------- #
        df_common = pd.DataFrame()
        if config_dict["check_common_cap"].upper() =="ON" and config_dict["path_common_cap"]:
            df_common = utils.parse_common_cap(config_dict["path_common_cap"])
        save_to_log(logPathName,f"Check_Common_Cap: {config_dict['check_common_cap']}")
        save_to_log(logPathName,f"Path_Common_Cap: {config_dict['path_common_cap']}")
        # 讀取 real bom(.prt)
        df_real_bom = utils.parse_real_bom(path_dict['prt'])
        dict_non_common,missing_pn = utils.check_common_cap(dict_bom_cap,df_common,df_real_bom)
        if dict_non_common :
            insert_data['Non_Common_Cap_Used'] = dict_non_common
            save_to_log(logPathName,f"XXXXX Non_Common_Cap_Used: {dict_non_common}")
            save_to_log(logPathName,f"XXXXX Part numbers of Non_Common_Cap_Used: {missing_pn}")
            raise NameError('Non_Common_Cap_Used')

        #==============================================================================================================#
        # # ---------------------------------- Giving License -------------------------------- #
        # unfinshed_task = pd.DataFrame(collect_tc.find({'busItem':busItem,'status':{'$in':['Running','Scheduled']}}))
        # if len(unfinshed_task) !=0:
        #     # 設定 order 數
        #     ser=unfinshed_task[unfinshed_task['license']==license_choose].reset_index(drop=True)
        #     if  len(ser) !=0:
        #         insert_data['order'] = int(ser['order'].max() + 1)
        #     else:
        #         insert_data['order']=1
        # else:
        #     insert_data['order']=1
        # license_choose = LisenseChoose(busItem)
        # insert_data['license'] = license_choose
        # SaveToLog(logPathName,"assign schedule for %s"%license_choose)
        to_db = 0
        save_to_log(logPathName,f"Length of TCL list : {str(len(tcl_list))} , Type : {type(tcl_list)}" )
        #  判斷 'Auto PDN' or 'Auto DDR'
        # if busItem == 'PDN':
        if len(tcl_list) <= insert_data['net_count']:
            to_db = 1
        # elif busItem =='DDR CCT':
        #     if len(tcl_list) == 1 :
        #         to_db = 1
        #         insert_data['license']= license_name
        # else:
        #     SaveToLog(logPathName,"XXXXX bustItem isn't PDN or DDR CCT XXXXX" )
        if to_db == 1:
            insert_data['current_opi_start_dt'] = int(time.time())
            insert_data['initialTime'] = int(time.time())
            insert_data['initialResult'] = 'Succeed'
            insert_data['status'] = 'Scheduled'
            if tcl_list != []:
                insert_data['tcl'] = tcl_list
            collect_tc.update_one({'form_id':insert_data['form_id']},
                                {'$set':insert_data})
            notify_type = 'initial_success'
            mailNotify.opiMailNotify (notify_type, insert_data,config_dict)
            save_to_log(logPathName,"initial_success MAIL SENT" )
            return True
        else:
            save_to_log(logPathName,f"XXXXX Power Rail number is over {insert_data['net_count']} XXXXX" )
            notify_type = 'Over_Net'
            mailNotify.opiMailNotify (notify_type, insert_data,config_dict)
            save_to_log(logPathName,"Over_Net MAIL SENT" )
            collect_tc.update_one({'form_id':insert_data['form_id']},
                              {'$set':{'status':'Cancel','order':0}})
            df_schedule = pd.DataFrame(collect_tc.find({'license': 'Unspecified','order': {'$gt':0},'busItem':{'$regex':'.*'+busItem+'$'},'debug':False}).sort("order"))
            order=1
            for _,row in df_schedule.iterrows():
                collect_tc.update_one({'form_id': row['form_id']},
                                                {'$set':{'order':order}})
                order += 1
        return False
    except NameError as e:
        if (insert_data['fail_times'] == 0 or initial_fail_sending_mail_time(config)) and str(e) not in ['NG_placement','NG_placement_pin','NG_Short_Net','NG_Net','Non_Common_Cap_Used']:
            notify_type = str(e)
            mailNotify.opiMailNotify(notify_type, insert_data, config_dict)
            save_to_log(logPathName,f"{notify_type} MAIL SENT" )
        insert_data['status'] = 'Initial_Fail'
        insert_data['initialResult'] = str(e)
        insert_data['fail_times'] = 1
        collect_tc.update_one({'form_id':insert_data['form_id']},
                              {'$set':insert_data})
        save_to_log(logPathName,f"***** {busItem} ERROR: {e} ******")
        return False
    except Exception as _:
        collect_tc.update_one({'form_id':insert_data['form_id']},
                              {'$set':{'status':'Initial_Fail',
                               'initialResult':'Error'}})
        save_to_log(logPathName,f"XXXXXX Python ERROR:\n {traceback.format_exc()} XXXXX")
        mailNotify.error_mail(insert_data,logPathName,config_dict)
        return False
def initial_fail_sending_mail_time(config:dict)->bool:
    """Checking if the moment to send any error mail involved to initial fail.

    Args:
        config (dict): data in config.json

    Returns:
        bool: True for sending remind mail now.
    """    
    ct=int(time.time())
    warning_times = config['initial_fail_warning_time']
    #! 10:00 everyday, Sending upload file remind 
    y = str(datetime.datetime.now().year)
    m = str(datetime.datetime.now().month).zfill(2)
    d = str(datetime.datetime.now().day).zfill(2)
    #? Week day 星期一~ 天 為 0~6，故以下要加 1 (drop)
    week_day =datetime.datetime.now().weekday() + 1
    if week_day in [0]:
        save_to_log(logPathName,"Weekend Now, Do Not Sending Initial Fail Mail")
        return False
    for wt in warning_times:
        today = f"{y}-{m}-{d} {wt}"
        wt_stamp = datetime.datetime.strptime(today,'%Y-%m-%d %H:%M').timestamp()
        if abs(wt_stamp-ct) <= 5*60 :
            save_to_log(logPathName,f"At {wt} ,Do Send Initial Fail Mail Now")
            return True
    return False
# --------------------------- initial_psi 20241218 ------------------------- #
# PDN/ DDR CCT (Power Si) 放一起
# Function:
# 1. DDR CCT(Power SI) Initial 要產生 Stackup.csv 的 export_stackup.tcl
# 2. SIMetrics_Automation.exe 需要使用的 simetrics_settings.json
# Input:
# 1. Task dict.
# Output:
# 1. True/ False
def initial_psi(insert_data) -> bool:
    # --------------------- Reduce Congitive Complexity ---------------------- #
    def check_platform_in_database(cpu_name,platform,configuration,rate):
        save_to_log(logPathName,"Check check_platform_in_database...")
        database_path = os.path.dirname(config_dict['SI_auto_site'])
        df_database = pd.read_excel(os.path.join(database_path,"Wistron_PSI_database.xlsx"),sheet_name='Model')
        save_to_log(logPathName,"Get database...")
        base_key = df_database.keys()
        check_platform = df_database[(df_database[base_key[0]] == f"{cpu_name}_{platform}")&
                                    (df_database[base_key[1]] == configuration)&
                                    (df_database[base_key[2]] == rate)]
        if check_platform.shape[0] != 0:
            save_to_log(logPathName,"Database Check OK")
            return True
        else: 
            save_to_log(logPathName,f"Database Not Support {cpu_name}_{platform}_{configuration}_{rate}")
            raise NameError(f"Database Not Support {cpu_name}_{platform}_{configuration}_{rate}")
    # ======================================================================== #
    try:
        save_to_log(logPathName,"Call Function Initial_Power SI")
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        #* =============================================== *#
        save_to_log(logPathName," running gentcl.stackup() to create print_list" )
        path_dict = insert_data['filePath']
        temp_path = os.path.dirname(path_dict['print_list'])
        output_path = path_dict['output_path']
        # 讀取 CPU Platform
        vendor,cpu_name,platform,cpu_type,cpu_target=utils.read_CPU_info(insert_data['platform'])
        save_to_log(logPathName,f"parse Platform result: {vendor}-{cpu_name}-{platform}-{cpu_type}-{cpu_target}")
        # ---------------------------- export_print_list_tcl ----------------------- #
        check_print_list = DDRGentcl.export_print_list_tcl(path_dict)
        if not check_print_list :
            raise NameError('PARSE_ERROR')
        save_to_log(logPathName,"export_stackup.tcl Created")
        # ---------------------------- Create_simetrics_settings -------------------- #
        rank_speed_rate = f"{insert_data['Rank']}_{insert_data['DataRate']}"
        si_setting = [{
                    "DATABASE": "Wistron_PSI_database.xlsx",
                    "SIG_VER": config_dict['SIG_VER'],
                    "PLATFORM": f"{cpu_name}_{platform}",
                    "CONFIGURATION": insert_data['RamType'],
                    "RATE": rank_speed_rate,
                    "SIM_FOLDER": str.replace(os.path.join(output_path,rank_speed_rate),'\\', '/'),
                    "BOARDS": [str.replace(path_dict['brd'],'\\', '/')],
                    "STACKUP": [str.replace(path_dict['stackup'],'\\', '/')]
                    }]
        
        if check_platform_in_database(cpu_name,platform,insert_data['RamType'],rank_speed_rate):
            with open(os.path.join(temp_path,'simetrics_settings.json'), 'w', encoding='utf-8') as f:
                json.dump(si_setting, f, ensure_ascii=False, indent=4)
            save_to_log(logPathName,"simetrics_settings.json Created")
            # =========================================================================== #
            cost={
                'Item':'',
                'Org_total_cap':0,
                'Opt_total_cap':0,
                'Org_total_cost':0,
                'Opt_total_cost':0,
                'Cost_saving':0,
                'Opt_efficiency':0
                }
            tcl_list=[{'report_result':'','ori_report_result':'','Cost':cost,'net': 'DDR','Modelname':'','NetName1':'','NetName2':'','status': 'unfinished', 'filepath': "", 'report_path': ""}]
            insert_data['current_opi_start_dt'] = int(time.time())
            insert_data['initialTime'] = int(time.time())
            insert_data['initialResult'] = 'Succeed'
            insert_data['status'] = 'Scheduled'
            insert_data['tcl'] = tcl_list
            collect_tc.update_one({'form_id':insert_data['form_id']},
                                        {'$set':insert_data})
            notify_type = 'initial_success'
            mailNotify.opiMailNotify (notify_type, insert_data,config_dict)
            save_to_log(logPathName,"initial_success MAIL SENT" )
            return True
        return False
    except Exception as _:
        if insert_data['fail_times'] == 0 or initial_fail_sending_mail_time(config):
            mailNotify.error_mail(insert_data,logPathName,config_dict)
        insert_data['status'] = 'Initial_Fail'
        insert_data['initialResult'] = 'Error'
        insert_data['fail_times'] = 1
        collect_tc.update_one({'form_id':insert_data['form_id']},
                              {'$set':insert_data})
        save_to_log(logPathName,f"XXXXXX Python ERROR:\n {traceback.format_exc()} XXXXX")
        return False
    
# ============================================= 以下為 Checking 排程狀態並進行模擬 ============================================ #  
def check_pdn_status(license_name) :
    save_to_log(logPathName,"Call check_pdn_status")
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    running_task = collect_tc.find_one({'status': 'Running', 'license': license_name})
    # 要同時抓取 debug : False or True 的 'Unspecified' 'Scheduled' 'initialResult':'Succeed' Task 
    # debug: True Task 的 order 為0 所以不可以用 order $gt 0 索取資料
    df_schedule = pd.DataFrame(collect_tc.find({'license': 'Unspecified','busItem':{'$regex':'.*PDN$'},'status':'Scheduled','initialResult':'Succeed'}).sort("order"))
    try:
        if not running_task:
            if df_schedule.empty:
                return 'Inprogress: 0'
            if not __license_limit():
                save_to_log(logPathName,f"XXXXX Up to PDN Simulation limit: {config_dict['pdn_sim_limit_count']} XXXXX")
                return 'Inprogress:%d'%len(df_schedule)
            __start_next_project_v2(df_schedule,license_name)
            return 'Running'
        task_id = running_task['_id']
        form_id = running_task['form_id']
        # ! 新增license判斷
        task_license = running_task['license']
        save_to_log(logPathName,'form_id: '+form_id)
        for idx, tcl in enumerate(running_task['tcl']):
            save_to_log(logPathName,"tcl net: "+tcl['net'])
            save_to_log(logPathName,"tcl path: "+tcl['filepath'])
            save_to_log(logPathName,"tcl status: "+tcl['status'])
            save_to_log(logPathName,"----------")
            net = tcl['net']
            org_report_path = os.path.join(os.path.dirname(tcl['filepath']),f"Original_Simulation_Report_{net}.htm")
            save_to_log(logPathName,f"Original report :{org_report_path}")
            opt_tcl_result = __check_tcl_result__(tcl['report_path'])
            ori_tcl_result = __check_tcl_result__(org_report_path)
            tcl_status = (tcl['status'] == 'unfinished')
            # 目前tcl 結果有四種 condition
            check_con1 = (tcl_status and opt_tcl_result)
            check_con2 = (tcl['report_path'] == 'No_report' and tcl_status and ori_tcl_result)
            if  check_con1 or check_con2:
                save_to_log(logPathName,"enter process export report")
                try:
                    # ! 轉出報告
                    report.__export_report(str(task_id),tcl,running_task,idx,'auto')
                    running_task['tcl'][idx]['status'] = 'finished'
                    save_to_log(logPathName,running_task['tcl'][idx]['status']+' --> update status : finished')
                    res=collect_tc.update_one({'form_id': form_id},{'$set':running_task})
                    __check_update_db_result(res)
                except Exception as _:
                    running_task['tcl'][idx]['status'] = 'timeout'
                    save_to_log(logPathName,running_task['tcl'][idx]['status']+' --> update status : timeout')
                    res=collect_tc.update_one(  {'form_id': form_id},
                                            {'$set':running_task})
                    if res.acknowledged:
                        mailNotify.opiMailNotify ('timeout', running_task,config_dict)
                    save_to_log(logPathName, 'XXXXX Create Report Error: \n%s' % (traceback.format_exc()))
            elif (tcl['status'] == 'unfinished' and not __check_tcl_result__(tcl['report_path'])) or (tcl['report_path'] == 'No_report' and tcl['status'] == 'unfinished' and not __check_tcl_result__(org_report_path)):
                #* 沒有 pid，代表還沒有被執行
                if 'pid' not in tcl:
                    save_to_log(logPathName,f"**** Process Not Running, Call It. **** idx ={idx}")
                    bLicense=__check_license(task_license)
                    license_trans=license_info_dict[task_license]
                    if not bLicense:
                        save_to_log(logPathName,f"License Blocked :{task_license}")
                        save_to_log(logPathName,'Stop :%s'%tcl['filepath'])
                        return 'Pending-Inprogress:%d'%len(df_schedule)
                    # 檢查上一條 net 是否正在執行模擬
                    if not check_task_by_pid(running_task['tcl'][idx-1]['pid']):
                        save_to_log(logPathName,f"{task_license} Blocked at Net:{running_task['tcl'][idx-1]['net']}")
                        return 'Running'
                    save_to_log(logPathName,"Call Subprocess - Next Power Rail")
                    res=subprocess.Popen([config_dict['Cad_site'], license_trans,'-b','-tcl', tcl['filepath']], start_new_session=True, close_fds=True)
                    save_to_log(logPathName,f"Run: {tcl['filepath']}")
                    running_task['tcl'][idx]['TCL_start_time'] = int(time.time())
                    running_task['status'] = 'Running'
                    running_task['tcl'][idx]['pid'] = res.pid
                    res = collect_tc.update_one({'form_id': form_id},
                                                {'$set':running_task})
                    __check_update_db_result(res)
                    return 'Running'
                #* 已經開始跑，但還沒跑完
                Vendor,_,_,_,_=utils.read_CPU_info(running_task['platform'])
                if busItem == 'PDN':
                    if Vendor.upper() == 'INTEL':
                        overtime_lim = config['intel_pdn_sim_days']
                    elif Vendor.upper() == 'QUALCOMM':
                        overtime_lim = config['qcm_pdn_sim_days']
                    else:
                        overtime_lim = config['amd_pdn_sim_days']
                elif busItem == 'DDR CCT':
                    overtime_lim = config['ddr_sim_days']
                # 目前 net 還未 timeout 可以停止迴圈
                if (int(time.time())- tcl['TCL_start_time']) < overtime_lim or not check_task_by_pid(tcl['pid']):
                    save_to_log(logPathName,f"form_id :{form_id} Net: {tcl['net']} isn't Timeout or PID exists({tcl['pid']})")
                    break
                running_task['tcl'][idx]['status']='timeout'
                save_to_log(logPathName,f"[Warning] form_id: {form_id} Net: {running_task['tcl'][idx]['net']} Time Over {overtime_lim/3600} hours.")
                res=collect_tc.update_one(  {'form_id': form_id},
                                            {'$set':running_task})
                if res.acknowledged:
                    save_to_log(logPathName,"Update Status for Running Overtime ok.")
                    mailNotify.opiMailNotify ('timeout', running_task,config_dict)
                else:
                    save_to_log(logPathName,"Update Status for Running Overtime fail.")           
            # 最後一個 tcl 才執行以下檢查
            if idx != len(running_task['tcl'])-1:
                save_to_log(logPathName,f"form_id {form_id} Progress... {idx+1}/{len(running_task['tcl'])}")
                # 如果目前 Net 還判定為 Timeout，則跳到下個迴圈以相同 license 名稱繼續模擬
                continue 
            save_to_log(logPathName,f"Check Last TCL :{net}")
            # ! 檢查所有線路是否都完成    
            check_pj_result,report_result = __check_project_result__(running_task['tcl'],task_id)
            if check_pj_result == 'Unfinished' :
                res=collect_tc.update_one({'_id': task_id},{'$set':{'status': 'Unfinished', 'finished_dt': int(time.time()), 'order': 0 ,'report_result':report_result}})
                if res.acknowledged:
                    notify_type='simulation_fail'
                    mailNotify.opiMailNotify (notify_type, running_task,config_dict)
                    save_to_log(logPathName,'Update Overall Status : Finished (FAIL)')
                else:
                    save_to_log(logPathName,'XXXXX Update to DB FAIL XXXXX')
                save_to_log(logPathName,f"**** form_id {form_id} Unfinished ****")
            elif check_pj_result:
                res=collect_tc.update_one({'_id': task_id},{'$set':{'status': 'Finished', 'finished_dt': int(time.time()), 'order': -1,'report_result':report_result}})
                if res.acknowledged:
                    notify_type='Finished'
                    mailNotify.opiMailNotify (notify_type, running_task,config_dict)
                    save_to_log(logPathName,'Update Overall Status : Finished (OK)')
                else:
                    save_to_log(logPathName,'XXXXX Update to DB FAIL XXXXX')
                save_to_log(logPathName,f"**** form_id {form_id} Finished ****")
            else :
                # 再度全部檢查是否有遺漏 Net 未執行模擬
                for idx1, tcl1 in enumerate(running_task['tcl'] ):
                    if tcl1['status'] == 'unfinished':
                        save_to_log(logPathName,"Call Next Power Rail : "+ tcl1['filepath'])
                        bLicense=__check_license(task_license)
                        license_trans= license_info_dict[task_license]
                        if not bLicense:
                            save_to_log(logPathName,f"When Checking All Nets(license blocked {task_license})")
                            save_to_log(logPathName,'Stop :%s'%tcl['filepath'])
                            return 'Pending-Inprogress:%d'%len(df_schedule)
                        res = subprocess.Popen([config_dict['Cad_site'], license_trans,'-b','-tcl', tcl1['filepath']], start_new_session=True, close_fds=True)
                        save_to_log(logPathName,'Run:%s'%tcl['filepath'])
                        now_timestamp = int(time.time())
                        running_task['tcl'][idx1]['TCL_start_time'] = now_timestamp
                        running_task['tcl'][idx1]['pid'] = res.pid
                        running_task['license'] = task_license
                        res=collect_tc.update_one({'form_id': running_task['form_id']},
                                                    {'$set':running_task})
                        return 'Running'
            # 確認是否有下一個任務
            if df_schedule.empty:
                return 'Inprogress: 0'
            # 確認是否有超出模擬license數量限制
            if not __license_limit() :
                save_to_log(logPathName,f"XXXXX Up to PDN Simulation limit: {config_dict['pdn_sim_limit_count']} XXXXX")
                return 'Inprogress:%d'%len(df_schedule)
            # 都沒問題執行下一個任務
            __start_next_project_v2(df_schedule,license_name)
        return 'Running'
    except Exception as _:
        save_to_log(logPathName, 'Error Message: \n%s' % (traceback.format_exc()))
        return 'Error'
    
# def checkDDRStatus_old(license_name) :#For SPEED2000, 此function 一定有錯，請確認後再重啟
#     try:
#         SaveToLog(logPathName,"Call checkDDRStatus")
#         collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#         running_task = collect_tc.find_one({'status': 'Running', 'license': license_name})
#         df_schedule = pd.DataFrame(collect_tc.find({'license': license_name,'order': {'$gt':0},'busItem':busItem}).sort("order"))
#         if running_task is not None :
#             id=running_task['_id']
#             SaveToLog(logPathName,"_id: "+str(id))
#             for idx, tcl in enumerate(running_task['tcl']):
#                 SaveToLog(logPathName,"tcl net: "+tcl['net'])
#                 SaveToLog(logPathName,"Batch path: "+tcl['filepath'])
#                 SaveToLog(logPathName,"Simulation status: "+tcl['status'])
#                 SaveToLog(logPathName,"----------")
#                 if tcl['status'] == 'unfinished' and __check_tcl_result__(tcl['report_path'])==True :
#                     SaveToLog(logPathName,"enter process : "+running_task['form_id'])
#                     running_task['tcl'][idx]['status'] = 'finished'
#                     SaveToLog(logPathName,running_task['tcl'][idx]['status']+" --> update status : finished")
#                     res=collect_tc.update_one({'_id': id},{'$set':running_task})
#                     if res.acknowledged==True:
#                         SaveToLog(logPathName,"update status : ok")
#                     else:
#                         SaveToLog(logPathName,"update status : fail")
#                     # ! 轉出報告
#                     # report.__export_report(str(id),tcl,running_task,idx,'auto')
#                     # ! 檢查所有線路是否都完成    
#                     check_pj_result,report_result = __check_project_result__(running_task['tcl'],id)
#                     if check_pj_result == True:                
#                         res=collect_tc.update_one({'_id': id},{'$set':{'status': 'Finished', 'finished_dt': int(time.time()), 'order': -1,,'report_result':report_result}})
#                         if res.acknowledged==True:
#                             notify_type='Finished'
#                             mailNotify.opiMailNotify (notify_type, running_task,config_dict)
#                             SaveToLog(logPathName,"update overall status : Finished (OK)")
#                         else: SaveToLog(logPathName,"XXXXX update to DB FAIL XXXXX")
#                         SaveToLog(logPathName,f"**** form_id {running_task['form_id']} Finished ****")
#                         # 以下設定 Simulation limit 為限制模擬使用 license 總數
#                         if len(df_schedule):
#                             if __license_limit():
#                                 __start_next_project_v2(df_schedule,license_name)
#                             else:
#                                 SaveToLog(logPathName,f"XXXXX Up to DDR Simulation limit: {config_dict['ddr_sim_limit_count']} XXXXX")
#                             return f"Inprogress:{len(df_schedule)}"
#                         else:
#                             return 'Inprogress: 0'
#                     elif check_pj_result == 'Unfinished':
#                         res=collect_tc.update_one({'_id': id},{'$set':{'status': 'Unfinished', 'finished_dt': int(time.time()), 'order': 0}})
#                         if res.acknowledged==True:
#                             notify_type='simulation_fail'
#                             mailNotify.opiMailNotify (notify_type, running_task,config_dict)
#                             SaveToLog(logPathName,"update overall status : Finished (FAIL)")
#                         else: SaveToLog(logPathName,"XXXXX update to DB FAIL XXXXX")
#                         SaveToLog(logPathName,f"**** form_id {running_task['form_id']} Unfinished ****")
#                         # 以下設定 Simulation limit 為限制模擬使用 license 總數 
#                         if len(df_schedule):
#                             if __license_limit():
#                                 __start_next_project_v2(df_schedule,license_name)
#                             else:
#                                 SaveToLog(logPathName,f"XXXXX Up to DDR Simulation limit: {config_dict['ddr_sim_limit_count']} XXXXX")
#                             return f"Inprogress:{len(df_schedule)}"
#                         else:
#                             return 'Inprogress: 0'
#                     else :
#                         for idx, tcl in enumerate(running_task['tcl'] ):
#                             if tcl['status'] == 'unfinished' :
#                                 bLicense=__check_DDR_license(license_name)
#                                 if bLicense:
#                                     SaveToLog(logPathName,"call DDR Bat : "+ tcl['filepath'])
#                                     subprocess.Popen(tcl['filepath'], shell=True)
#                                     SaveToLog(logPathName,f"Run:{tcl['filepath']}")
#                                     print(f"Run:{tcl['filepath']}")
#                                     return f"INprogress:{len(df_schedule)}"
#                                 else:
#                                     SaveToLog(logPathName,"call subprocess -  (license blocked)")
#                                     print(f"Pending:{tcl['filepath']}")
#                                     return f"Pending-Inprogress:{len(df_schedule)}"
#                 elif tcl['status'] == 'unfinished' and __check_tcl_result__(tcl['report_path'])==False:
#                     # ! make sure tcl is running
#                     # DDR 要使用os.path.dirname(tcl['filepath']) 來判斷，OPI 是 tcl['report_path']
#                     files = os.listdir(os.path.dirname(tcl['filepath']))
#                     has_log_file=any(file.endswith(('temp_netGroup.txt'))for file in files)
#                     if not has_log_file: # * 沒有產出其他檔案，代表還沒有被執行
#                         SaveToLog(logPathName,"**** process not running, call it. ****")
#                         # ! not in running mode
#                         bLicense=__check_DDR_license(license_name)
#                         if bLicense:
#                             SaveToLog(logPathName,"call subprocess - next power rail")
#                             now_timestamp = int(time.time())
#                             running_task['tcl'][idx]['TCL_start_time'] = now_timestamp
#                             res=collect_tc.update_one({'form_id': running_task['form_id']},
#                                                         {'$set':{'order':1,
#                                                                 'status':'Running',
#                                                                 'license':license_name,
#                                                                 'tcl': running_task['tcl'][idx]}})
#                             bat_dir = os.path.dirname(tcl['filepath'])
#                             # 切換至指定目錄
#                             os.chdir(bat_dir)
#                             subprocess.Popen(['start', 'cmd', '/K', tcl['filepath']], shell=True)
#                             SaveToLog(logPathName,f"Run:{tcl['filepath']}")
#                             print(f"Run:{tcl['filepath']}")
#                             return f"Inprogress:{len(df_schedule)}"
#                         else:
#                             SaveToLog(logPathName,"call subprocess -  (license blocked)")
#                             print(f"Pending:{tcl['filepath']}")
#                             return f"Pending-Inprogress:{len(df_schedule)}"
#                     else: # * 已經開始跑，但還沒跑完
#                         Vendor,CPU_name,Platform,CPU_type,CPU_Target=utils.read_CPU_info(running_task['platform'])
#                         if busItem.upper() == 'PDN':
#                             if Vendor.upper() == 'INTEL':
#                                 overtime_lim = config['intel_pdn_sim_days']
#                             else:
#                                 overtime_lim = config['qcm_pdn_sim_days']
#                         else:
#                             overtime_lim = config['ddr_sim_days']
#                         if (int(time.time())- tcl['TCL_start_time']) > overtime_lim:
#                             running_task['tcl'][idx]['status']='timeout'
#                             SaveToLog(logPathName,f"[warning] form_id: {running_task['form_id']} Net: {running_task['tcl'][idx]['net']} time over {overtime_lim} hours.")
#                             res=collect_tc.update_one({'form_id': running_task['form_id']},
#                                                         {'$set':{'status': 'Uninished','tcl': running_task['tcl'][idx]}})
#                             if res.acknowledged==True:
#                                 SaveToLog(logPathName,"update status for running overtime ok.")
#                                 mailNotify.opiMailNotify ('timeout', running_task,config_dict)
#                             else:
#                                 SaveToLog(logPathName,"update status for running overtime fail.")
#                         else:
#                             SaveToLog(logPathName,f"Form_id :{running_task['form_id']} Net: {tcl['net']} isn't Timeout.")
#                         if idx == len(running_task['tcl'])-1:
#                             SaveToLog(logPathName,f"check last TCL :{tcl['net']}")
#                             # ! 檢查所有線路是否都完成    
#                             check_pj_result,report_result = __check_project_result__(running_task['tcl'],id)
#                             if check_pj_result == 'Unfinished':
#                                 res=collect_tc.update_one({'_id': id},{'$set':{'status': 'Unfinished', 'finished_dt': int(time.time()), 'order': 0,,'report_result':report_result}})
#                                 if res.acknowledged==True:
#                                     notify_type='simulation_fail'
#                                     mailNotify.opiMailNotify (notify_type, running_task,config_dict)
#                                     SaveToLog(logPathName,"update overall status : Finished (FAIL)")
#                                 else: SaveToLog(logPathName,"XXXXX update to DB FAIL XXXXX")
#                                 SaveToLog(logPathName,f"**** form_id {running_task['form_id']} Unfinished ****")
#                                 # 以下設定 Simulation limit 為限制模擬使用 license 總數 
#                                 if len(df_schedule):
#                                     if __license_limit():
#                                         __start_next_project_v2(df_schedule,license_name)
#                                     else:
#                                         SaveToLog(logPathName,f"XXXXX Up to DDR Simulation limit: {config_dict['ddr_sim_limit_count']} XXXXX")
#                                     return f"Inprogress:{len(df_schedule)}"
#                                 else:
#                                     return 'Inprogress: 0'
#                             else:
#                                 SaveToLog(logPathName,f"form_id {running_task['form_id']} isn't Unfinished.")
#                         else:
#                             SaveToLog(logPathName,f"form_id {running_task['form_id']} progress... {idx+1}/{len(running_task['tcl'])}")
#             return 'Running'
#         elif len(df_schedule) :
#             # 以下設定 Simulation limit 為限制模擬使用 license 總數 
#             if __license_limit():
#                 __start_next_project_v2(df_schedule,license_name)
#             else:
#                 SaveToLog(logPathName,f"XXXXX Up to DDR Simulation limit: {config_dict['ddr_sim_limit_count']} XXXXX")
#             return f"Inprogress:{len(df_schedule)}"
#         else:
#             return 'Inprogress: 0'
#     except Exception as e:
#         SaveToLog(logPathName,f"Error : {traceback.format_exc()}")

# ----------------------------------------- 導入 Power SI ------------------------------------ #
def check_psi_status_new(license_name) :
    save_to_log(logPathName,"Call check_psi_status_new")
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    running_task = collect_tc.find_one({'status': 'Running', 'license': license_name})
    # 要同時抓取 debug : False or True 的 'Unspecified' 'Scheduled' 'initialResult':'Succeed' Task 
    # debug: True Task 的 order 為0 所以不可以用 order $gt 0 索取資料
    df_schedule = pd.DataFrame(collect_tc.find({'license': 'Unspecified','busItem':{'$regex':'.*DDR CCT$'},'status':'Scheduled','initialResult':'Succeed'}).sort("order"))
    try:
        if not running_task:
            if len(df_schedule) == 0:
                return 'Inprogress: 0'
            __start_next_project_psi(df_schedule)
            return 'Running'
        task_id = running_task['_id']
        save_to_log(logPathName,"_id: "+str(task_id))
        path_dict = running_task['filePath']
        output_path = path_dict['output_path']
        rank_speed_rate = f"{running_task['Rank']}_{running_task['DataRate']}"
        check_done = os.path.exists(os.path.join(output_path,rank_speed_rate,'board0','Simulation.done'))
        if check_done:
            running_task['status'] = 'Finished'
            running_task['tcl'][0]['status'] = 'finished'
            running_task['order'] = -1
            running_task['finished_dt'] = int(time.time())
            # ! 等待可以讀取報告結果: 要新增 running_task['report_result'] = 'Fail' or 'Pass'
            save_to_log(logPathName,running_task['status']+" --> Update Status : Finished")
            res=collect_tc.update_one({'_id': task_id},{'$set':running_task})
            __check_update_db_result(res)
        else:
            overtime_lim = config_dict['ddr_sim_days']
            if (int(time.time())- running_task['tcl'][0]['TCL_start_time']) < overtime_lim or not check_task_by_pid(running_task['tcl'][0]['pid']):
                save_to_log(logPathName,f"form_id :{running_task['form_id']} Net: {running_task['tcl'][0]['net']} isn't Timeout or PID ({running_task['tcl'][0]['pid']}) Exists.")
                return 'Running'
            running_task['tcl'][0]['status']='timeout'
            running_task['status'] = 'Unfinished'
            running_task['report_result'] = 'Fail'
            running_task['finished_dt'] = int(time.time())
            running_task['order'] = 0
            save_to_log(logPathName,f"[Warning] form_id: {running_task['form_id']} Net: {running_task['tcl'][0]['net']} Time Over {overtime_lim/3600} hours.")
            res=collect_tc.update_one({'form_id': running_task['form_id']},
                                        {'$set': running_task})
            if res.acknowledged:
                save_to_log(logPathName,"Update Status for Running Overtime ok.")
                mailNotify.opiMailNotify ('timeout', running_task,config_dict)
            else:
                save_to_log(logPathName,"Update Status for Running Overtime fail.")
        if len(df_schedule) == 0:
            return 'Inprogress: 0'
        __start_next_project_psi(df_schedule)
        return 'Running'
    except Exception as _:
        save_to_log(logPathName, 'Error Message: \n%s' % (traceback.format_exc()))
        return 'Error'
def __start_next_project_v2(df_schedule,license_name) : # PDN/ DDR CCT(SPEED2000) merge in V2 #(OK)
    save_to_log(logPathName,"Call startNextProject_v2")
    get_task =  df_schedule.iloc[0].to_dict()
    bus_item =get_task['busItem']
    form_id = get_task['form_id']
    now_timestamp = int(time.time())
    if not check_task_status(get_task):
        save_to_log(logPathName,f"Waiting for form_id: {form_id} Well-Prepared ....")
        return False
    # if bus_item.upper() == 'DDR CCT':
    #     if os.path.exists(get_task['tcl'][0]['filepath']) == False:
    #         SaveToLog(logPathName,f"form_id: {form_id} Prepared in Another Server")
    #         return
    if get_task['tcl'][0]['status'] != 'unfinished':
        return False
    save_to_log(logPathName,"next scheduled project : "+ get_task['tcl'][0]['filepath'])
    if 'PDN' in bus_item.upper():
        b_license=__check_license(license_name)
        license_trans=license_info_dict[license_name]
    elif 'DDR CCT' in bus_item.upper():
        b_license=__check_DDR_license(license_name)
    if not b_license:
        save_to_log(logPathName,"XXXXX Run next project (license blocked) XXXXX")
        save_to_log(logPathName,'Stop: %s'%get_task['tcl'][0]['filepath'])
        return False
    # 以下有區分 PDN /DDR CCT(SPEED2000) 執行模擬方式
    # if bus_item.upper() == 'PDN':
    res=subprocess.Popen([config_dict['Cad_site'], license_trans,'-b','-tcl', get_task['tcl'][0]['filepath']],start_new_session=True, close_fds=True)
    # elif bus_item.upper() == 'DDR CCT':
    #     bat_dir = os.path.dirname(get_task['tcl'][0]['filepath'])
    #     # 切換至指定目錄
    #     os.chdir(bat_dir)
    #     subprocess.Popen(['start', 'cmd', '/K', get_task['tcl'][0]['filepath']], shell=True)
    save_to_log(logPathName,f"Run:{get_task['tcl'][0]['filepath']}")
    print(f"Run:{get_task['tcl'][0]['filepath']}")
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    save_to_log(logPathName,"next project form_id : "+ str(form_id))
    get_task['status'] = 'Running'
    get_task['license'] = license_name
    get_task['tcl'][0]['TCL_start_time'] = now_timestamp
    get_task['tcl'][0]['pid'] = res.pid
    res=collect_tc.update_one({'form_id': form_id},
                            {'$set':get_task})
    save_to_log(logPathName,"%s assinged for simulation!" %license_name)
    __check_update_db_result(res)
    notify_type='simulation_start'
    mailNotify.opiMailNotify (notify_type, get_task,config_dict)
    return True
# DDR CCT only(Power SI)
# 1.Execute PowerSI.exe to create Stackup.csv
# 2.Create Stackup_{board_number}.csv
# 3.Copy "simetrics_settings.json" to config_dict['SI_auto_site'] folder
# 4.Execute Simulation
# Input: 
# 1.all scheduled Tasks
# 2.liscense name (SPEED2000 #1) or change to PSI #1
# Output:
# If Step1 to 4 are Successful -> True
# any way -> False
def __start_next_project_psi(df_schedule) ->bool:
    save_to_log(logPathName,"Call __start_next_project_psi")
    get_task =  df_schedule.iloc[0].to_dict()
    form_id = get_task['form_id']
    save_to_log(logPathName,f"Get Task -> form_id:{form_id}")
    now_timestamp = int(time.time())
    path_dict = get_task['filePath']
    temp_path = os.path.dirname(path_dict['brd'])
    si_auto_folder = os.path.dirname(config_dict['SI_auto_site'])
    if not check_task_status(get_task):
        save_to_log(logPathName,f"Waiting for form_id: {form_id} Well-Prepared ....")
        return False
    if not power_si_license_check():
        save_to_log(logPathName,"XXXXX Run Next Project (License Blocked) XXXXX")
        return False
    # step 1.
    use_license = config_dict['PowerSI_license']
    result=subprocess.run([config_dict['PowerSI_Path'],use_license ,'-tcl', path_dict['print_list']], capture_output=True)
    save_to_log(logPathName,"Subprocess Result: "+str(result.returncode))
    # step 2. and step 3.
    file_prepare = [utils.create_psi_stackup_num_csv(path_dict) , copy_si_json(temp_path,si_auto_folder)]
    if False in file_prepare:
        save_to_log(logPathName,f"XXXXX Created stackup_boardnum / copy si json result: {file_prepare} XXXXX")
        return False
    # step 4.
    log_file = os.path.join(os.getcwd(),'SI_Auto_exe_running.log')
    with open(log_file, 'w', encoding='utf-8') as auto_log:
        # Use subprocess.Popen to run the executable without waiting for the result
        res=subprocess.Popen([config_dict['SI_auto_site']],stdout=auto_log,stderr=auto_log,creationflags=subprocess.CREATE_NEW_CONSOLE)
    save_to_log(logPathName,"Execute SIMetrics_Automation.exe... Created Log: "+ log_file)
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    get_task['status'] = 'Running'
    get_task['license'] = license_name
    get_task['tcl'][0]['TCL_start_time'] = now_timestamp
    get_task['tcl'][0]['pid'] = res.pid
    collect_tc.update_one(  {'form_id': form_id},
                            {'$set':get_task})
    save_to_log(logPathName,f"{license_name} Assinged for Simulation!")
    return True

# ---------------- Reduce Congitive Complexity -------------------- #
def copy_si_json(temp_path:str,si_folder:str) -> bool:
    # ------------------ Copy simetrics_settings.json to config_dict['SI_auto_site'] folder ---------------------- #
    si_json='simetrics_settings.json'
    shutil.copyfile(os.path.join(temp_path,si_json), os.path.join(si_folder,si_json))
    return os.path.exists(os.path.join(si_folder,si_json))
def check_task_status(task_data:dict) -> bool:
    # Check1 為檢查Task已上傳(已上傳才會有 'initialResult' key)
    check1 = 'initialResult'  in task_data
    if not check1:
        save_to_log(logPathName,f"XXXXX form_id: {task_data['form_id']} didn't upload files XXXXX")
        return False
    now_timestamp = int(time.time())
    sim_run_timestamp = datetime.datetime.strptime(task_data['sim_start_date'],'%Y-%m-%d').timestamp()
    # Check2 為檢查: 1. 是否是 debug 2.initial 成功 3. 是否是Task預定的模擬時間
    check2 = [task_data['debug'],task_data['initialResult'] == 'Succeed',now_timestamp >= sim_run_timestamp]
    # 兩種組合 -> 1.(debug 且 initial 成功) 2.(initial 成功且是Task預定的模擬時間)
    if check2[0] and check2[1]:
        return True
    if check2[1] and check2[2]:
        return True
    save_to_log(logPathName,f"XXXXX form_id: {task_data['form_id']} didn't initial successfully or be in simulation time XXXXX")
    return False
def power_si_license_check()-> bool:
    license_in_use = 0
    psi_license = config_dict['check_PSI']
    exe_dir= os.path.join(config_dict['check_license_exe_path'],config_dict['check_license_exe_name'])
    res=subprocess.run([exe_dir, 'lmstat', '-c' ,config_dict['check_license_cmd'],'-f', psi_license],shell=True, capture_output=True, text=True)
    stdoutstr=res.stdout
    # * 解析console畫面內的文字
    save_to_log(logPathName,"check PowerSI license console output : "+stdoutstr)
    if len(stdoutstr)==0:   # 執行檔路徑有錯誤 or 參數錯誤
        mailNotify.query_license_error(config_dict)
        save_to_log(logPathName,'Sent Query_license_error mail.')
        raise NameError('XXXXX check license execution failed XXXXX')
    start_index = stdoutstr.find(f"Users of {psi_license}:")
    end_index = stdoutstr.find(")", start_index) + 1
    check_string = stdoutstr[start_index:end_index]
    numbers = re.findall(r'\d+', check_string)
    license_in_use += int(numbers[-1])
    save_to_log(logPathName,f"Total License In Use : {license_in_use}")
    return 1>=license_in_use
# ======================================================================================================== #
def __startNextProject__old(df_schedule) : # 非常舊版 Function 只能執行 PDN
    save_to_log(logPathName,"Call startNextProject")
    config=utils.__read_config_file()
    #! 讀取時間限制 OPI1 00:00/ OPI2 12:00/ OPI3 20:00 才開始 run SIM.
    #抓取現在 time stamp
    now_timestamp = int(time.time())
    # 讀出 License 開始 模擬時間
    license_name=df_schedule.loc[0, 'license']
    OPI_time = config[license_name+'_time']
    time_format = '%H:%M'

    time_datetime = datetime.datetime.strptime(OPI_time, time_format)
    current_date =datetime.datetime.now().date() 
    # 產生 License 限制的時間 Stamp
    OPI_timestamp = datetime.datetime.combine(current_date, time_datetime.time()).timestamp()
    # 抓 limititation time 前後 5 分鐘可以跑模擬
    if int(OPI_timestamp+5*60)>=int(now_timestamp) and int(now_timestamp)>= int(OPI_timestamp-5*60) :
        save_to_log(logPathName,"%s is on duty for simulation!" %df_schedule.loc[0, 'license'])
        df_schedule.sort_values(by='order', inplace=True, ignore_index=True)
        # df_schedule['order'] = range(1, len(df_schedule)+1)
        # df_schedule.loc[0, 'status'] = 'Running'
        license=df_schedule.loc[0, 'license']
        _id=df_schedule.loc[0, '_id']
        for tcl in df_schedule.loc[0, 'tcl'] :
            if tcl['status'] == 'unfinished' :
                # df_schedule.loc[0,'current_opi_start_dt'] = int(time.time())
                # collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName="OPITaskCtrl")
                # for doc in df_schedule.to_dict('records') :
                #     # doc.pop('_id')
                #     SaveToLog(logPathName,"next project id : "+ str(doc['_id']))
                #     res=collect_tc.update_one({'_id': doc['_id']},{'$set':doc})
                #     if res.acknowledged==True:
                #         SaveToLog(logPathName,"update next project status : Finished (OK)")
                #     else:
                #         SaveToLog(logPathName,"update next project status : Finished (FAIL)")
                save_to_log(logPathName,"call next scheduled project : "+ tcl['filepath'])
                bLicense=__check_license(license)
                # ! license轉換
                # license_trans='-OptimizePI_20'
                # if license == 'OPI1' or license == 'OPI2':
                #     license_trans='-OptimizePI_20'
                # elif license == 'OPI3':
                #     license_trans='-AdvancedPI_TI_20'
                license_trans=license_info_dict[license]
                if bLicense:
                    df_schedule['order'] = range(1, len(df_schedule)+1)
                    df_schedule.loc[0, 'status'] = 'Running'
                    df_schedule.loc[0,'current_opi_start_dt'] = int(time.time())
                    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName="OPITaskCtrl")
                    save_to_log(logPathName,"next project id : "+ str(_id))
                    res=collect_tc.update_one({'_id': _id},{'$set':df_schedule.to_dict('records')[0]})
                    if res.acknowledged:
                        save_to_log(logPathName,"update next project status : Finished (OK)")
                    else:
                        save_to_log(logPathName,"update next project status : Finished (FAIL)")
                    save_to_log(logPathName,"call subprocess - next project")
                    subprocess.Popen([config_dict['Cad_site'], license_trans,'-tcl', tcl['filepath']],start_new_session=True, close_fds=True)
                    print('Run:%s'%tcl['filepath'])
                    notify_type='simulation_start'
                    mailNotify.opiMailNotify (notify_type, df_schedule.iloc[0],config_dict)
                    break
                else:
                    save_to_log(logPathName,"call subprocess - next project (license blocked)")
                    print('Not Run:%s'%tcl['filepath'])
                    break
    else:
        save_to_log(logPathName,"%s isn't on duty for simulation!" %df_schedule.loc[0, 'license'])
def __DDRstartNextProject__old(df_schedule) : # 舊版 Function 只能執行 SPEED2000
    save_to_log(logPathName,"Call startNextProject")
    df_schedule.sort_values(by='order', inplace=True, ignore_index=True)
    license=df_schedule.loc[0, 'license']
    _id=df_schedule.loc[0, '_id']
    for tcl in df_schedule.loc[0, 'tcl'] :
        if tcl['status'] == 'unfinished' :
            save_to_log(logPathName,"call next scheduled project : "+ tcl['filepath'])
            bLicense=__check_license(license)
            if bLicense:
                df_schedule['order'] = range(1, len(df_schedule)+1)
                df_schedule.loc[0, 'status'] = 'Running'
                df_schedule.loc[0,'current_opi_start_dt'] = int(time.time())
                collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName="OPITaskCtrl")
                save_to_log(logPathName,"next project id : "+ str(_id))
                res=collect_tc.update_one({'_id': _id},{'$set':df_schedule.to_dict('records')[0]})
                if res.acknowledged==True:
                    save_to_log(logPathName,"update next project status : Finished (OK)")
                else:
                    save_to_log(logPathName,"update next project status : Finished (FAIL)")
                save_to_log(logPathName,"call subprocess - next project")
                bat_dir = os.path.dirname(tcl['filepath'])
                # 切換至指定目錄
                os.chdir(bat_dir)
                subprocess.Popen(['start', 'cmd', '/K', tcl['filepath']], shell=True)
                save_to_log(logPathName,f"Run:{tcl['filepath']}")
                print(f"Run:{tcl['filepath']}")
                notify_type='simulation_start'
                mailNotify.opiMailNotify (notify_type, df_schedule.iloc[0],config_dict)
                break
            else:
                save_to_log(logPathName,"call subprocess - next project (license blocked)")
                save_to_log(logPathName,f"Not Run:{tcl['filepath']}")
                break
# ======================================================================================================== #
# ============================================= 以下為 Checking 排程狀態使用到的雜七雜八 Function ============================================ #  
def __check_tcl_result__(report_path) :#(OK)
    return os.path.exists(report_path)
# ------------------ checkProjectResult ------------------------ #(OK)
# 檢查模擬是否完成。
# 分成確認 finished/ time out 兩種 TCL 模擬結果。
# report 結果分為有跑 Optimize 或沒跑 ---> 沒跑 tcl['report_path'] = 'No_report'
#! TODO 還會檢查 report_result ='Fail'or'Pass' 如果tcl result 有 Timeout or Fail 就是 Fail
# Output:
# True/ Unfinished/ False
def __check_project_result__(tcl_list,id) :
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    save_to_log(logPathName,"Call checkProjectResult")
    tcl_count = len(tcl_list)
    finished_count = 0
    timeout_count = 0
    report_result ='Pass'
    for idx,tcl in enumerate(tcl_list) :
        if tcl['status'] == 'timeout':
            timeout_count += 1
            report_result ='Fail'
        else:
            # No Report 的情況為沒做 Optimize Simulation 就會為 No Report(No OPT report)，但仍有 Origin report
            if tcl['report_path'] == 'No_report':
                org_report_path = os.path.join(os.path.dirname(tcl['filepath']),f"Original_Simulation_Report_{tcl['net']}.htm")
                if __check_tcl_result__(org_report_path) :
                    finished_count += 1
                    tcl_list[idx]['status']='finished'
                if tcl['report_result'] == 'Fail':
                    report_result ='Fail'
            else:
                if __check_tcl_result__(tcl['report_path']) :
                    finished_count += 1
                    tcl_list[idx]['status']='finished'
                if tcl['ori_report_result'] == 'Fail':
                    report_result ='Fail'
    save_to_log(logPathName,f"_id {id} -> TCL Finished : {finished_count}/ Time Out : {timeout_count}/ Total : {tcl_count}")
    res=collect_tc.update_one({'_id': id},{'$set':{'tcl':tcl_list}})
    __check_update_db_result(res)
    return give_result(finished_count,tcl_count,timeout_count),report_result
# Reduced Congitive Complexity in __checkProjectResult__
def give_result(finished_count,tcl_count,timeout_count):
    if finished_count == tcl_count :
        return True
    elif timeout_count != 0 and finished_count + timeout_count == tcl_count :
        return 'Unfinished'
    else:
        return False
# ============================================================== #
def __check_license(license): #(OK)
    config=config_dict
    check_license_dict=dict()
    if len(config)==0:
        save_to_log(logPathName,"read config.json file failed.")
        return False # ! config.json 檔案有問題
    if not('check_license_exe_path' in config and  'check_license_exe_name' in  config and 'check_license_cmd' in config and 'pdn_license_count' in config):
        save_to_log(logPathName,"config.json with no params for check license.")
        return False
    license_count=config['pdn_license_count']        
    if license_count<=0:
        save_to_log(logPathName,"read config.json file - license_count is 0.")
        return False # ! 抓不到license資訊
    
    for i in range(license_count):
        check_license_dict['OPI%d'%(i+1)]=config['check_OPI%d'%(i+1)]
    # ! license轉換
    license_trans=check_license_dict[license]
    exe_dir= os.path.join(config['check_license_exe_path'],config['check_license_exe_name'])
    res=subprocess.run([exe_dir, 'lmstat', '-c' ,config['check_license_cmd'],'-f', license_trans],shell=True, capture_output=True, text=True)
    stdoutstr=res.stdout
    save_to_log(logPathName,"check license exe: " + exe_dir)
    save_to_log(logPathName,"check license cmd: " + config['check_license_cmd'])
    save_to_log(logPathName,"license: "+license)
    if len(stdoutstr)==0:   # 執行檔路徑有錯誤 or 參數錯誤
        save_to_log(logPathName,"check license execution failed")
        mailNotify.query_license_error(config)
        save_to_log(logPathName,"Sent Query_license_error mail.")
        return False
    # * 解析console畫面內的文字 
    save_to_log(logPathName,"check license console output : "+stdoutstr)
    # 搜索以下 : Users of AdvancedPI:  (Total of 1 license issued;  Total of 0 licenses in use) ## 1 為滿 license for AdvancedPI_TI_20
    # 搜索以下 : Users of OptimizePI:  (Total of 2 licenses issued; Total of 2 licenses in use) ## 2 為滿 license for OptimizePI_20
    start_index = stdoutstr.find(f"Users of {license_trans}:")
    end_index = stdoutstr.find(")", start_index) + 1
    check_string = stdoutstr[start_index:end_index]
    # 找到其中數字 (3,2) or (2,2)
    numbers = re.findall(r'\d+', check_string)
    if int(numbers[-2]) != int(numbers[-1]):
        save_to_log(logPathName,"license available")
        return True
    save_to_log(logPathName,"XXXXX license not available XXXXX")
    return False

# --------------------------- check_task_by_pid --------------------------- #
def check_task_by_pid(pid:int)->bool:
    """
    Args:
        pid (int): pid by power rail
        check_duration (int): check cpu usage of pid for "check_duration" seconds(times)

    Returns:
        bool: False for tool executing/ True for the pid simualtion is closed.  
    """
    try:
        psutil.Process(pid)
        return False
    except psutil.NoSuchProcess:
        # pid is noSuchProcess -> pid is closed
        return True

# ====================================================================== #
def __check_DDR_license(license): #(未使用)
    config=config_dict
    check_license_dict=dict()
    if len(config)==0:
        save_to_log(logPathName,"read config.json file failed.")
        return False # ! config.json 檔案有問題
    if not('check_license_exe_path' in config and  'check_license_exe_name' in  config and 'check_license_cmd' in config and 'ddr_license_count' in config):
        save_to_log(logPathName,"config.json with no params for check license.")
        return False
    # Users of AdvancedPI_TI_20:  (Total of 1 license issued;  Total of 0 licenses in use) # console畫面內的使用字串
    license_count=config['ddr_license_count']        
    if license_count<=0:
        save_to_log(logPathName,"read config.json file - license_count is 0.")
        return False # ! 抓不到license資訊
    for i in range(license_count):
        check_license_dict['SPEED2000 #%d'%(i+1)]=config['check_SPEED2000 #%d'%(i+1)]

    # ! license轉換
    license_trans=check_license_dict[license]
    exe_dir= os.path.join(config['check_license_exe_path'],config['check_license_exe_name'])
    save_to_log(logPathName,"[check_license] call subprocess.")
    res=subprocess.run([exe_dir, 'lmstat', '-c' ,config['check_license_cmd'],'-f', license_trans],shell=True, capture_output=True, text=True)
    stdoutstr=res.stdout
    save_to_log(logPathName,"check license exe: " + exe_dir)
    save_to_log(logPathName,"check license cmd: " + config['check_license_cmd'])
    save_to_log(logPathName,"license: "+license)
    if len(stdoutstr)==0:   # 執行檔路徑有錯誤 or 參數錯誤
        save_to_log(logPathName,"XXXXX check license execution failed XXXXX")
        mailNotify.query_license_error(config)
        save_to_log(logPathName,"Sent Query_license_error mail.")
        return False
    # * 解析console畫面內的文字
    save_to_log(logPathName,"check license console output : "+stdoutstr)
    # 搜索以下 : Users of OptimizePI_20:  (Total of 2 licenses issued;  Total of 2 licenses in use)
    start_index = stdoutstr.find(f"Users of {license_trans}:")
    end_index = stdoutstr.find(")", start_index) + 1
    check_string = stdoutstr[start_index:end_index]
    numbers = re.findall(r'\d+', check_string)
    # 查找 TPER90115562 (Sim3)\ TPER90115563 (Sim4) 是否有在console畫面內的文字中
    server_check=stdoutstr.find(config[license+'_server'])
    # ! True: license available
    if int(numbers[-1])<license_count and server_check == -1: 
        save_to_log(logPathName,"license available")
        return True
    else: 
        save_to_log(logPathName,"license not available")
        return False
def __find_license_info(): #(OK)
    config=config_dict
    license_dict=dict()
    if len(config)==0:
        return False # ! config.json 檔案有問題
    # ! 讀取license個數 & 名稱
    if 'pdn_license_count' in config:
        license_count=config['pdn_license_count']        
        if license_count>0:
            for i in range(license_count):
                license_dict['OPI%d'%(i+1)]=config['OPI%d'%(i+1)]
        else:
            return False # ! 抓不到license資訊
    else:
        return False # ! 抓不到license資訊
    return license_dict
def __find_ddr_license_info(): #(未使用，Configue.json 也沒有相關key，若要使用要再新增)
    config=config_dict
    license_dict=dict()
    if len(config)==0:
        save_to_log(logPathName,f"{os.getcwd()}")
        save_to_log(logPathName,"read config.json file failed.")
        return False # ! config.json 檔案有問題
    # ! 讀取license個數 & 名稱
    if 'ddr_license_count' in config:
        license_count=config['ddr_license_count']        
        if license_count>0:
            for i in range(license_count):
                license_dict['SPEED2000 #%d'%(i+1)]=config['SPEED2000 #%d'%(i+1)]
        else:
            save_to_log(logPathName,"read config.json file - license_count is 0.")
            return False # ! 抓不到license資訊
    else:
        save_to_log(logPathName,"read config.json file - license_count failed.")
        return False # ! 抓不到license資訊
    return license_dict
# ----------------------------- License_limit ------------------------------ #
def __license_limit(): #(OK，僅PDN再使用)
    save_to_log(logPathName,"Running __count_using_license ...")
    config=config_dict
    check_license_dict=dict()
    if  busItem == 'PDN':
        license_info_dict=__find_license_info()
        license_count=config['pdn_license_count']
        lim_count = config_dict['pdn_sim_limit_count']
        
    elif busItem == 'DDR CCT':
        license_info_dict=__find_ddr_license_info()
        license_count=config['ddr_license_count']
        lim_count = config_dict['ddr_sim_limit_count']
    if license_count<=0:
        raise NameError('XXXXX read config.json file - license_count is 0. XXXXX') # ! config.json 檔案有問題
    for i in range(license_count):
        if busItem == 'PDN':
            check_license_dict['OPI%d'%(i+1)]=config['check_OPI%d'%(i+1)]
        elif busItem == 'DDR CCT':
            check_license_dict['SPEED2000 #%d'%(i+1)]=config['check_SPEED2000 #%d'%(i+1)]

    if not ('check_license_exe_path' in config and  'check_license_exe_name' in  config and 'check_license_cmd' in config and 'pdn_license_count' in config):
        raise NameError('read config.json failed : check_license_exe_path/ check_license_exe_name/ check_license_cmd/ pdn_license_count')
    # Users of AdvancedPI_TI_20:  (Total of 1 license issued;  Total of 0 licenses in use) # console畫面內的使用字串   
    license_name_list=list(license_info_dict.keys())
    license_in_use = 0
    license_memo = []
    for license_name in license_name_list:
        license_trans=check_license_dict[license_name]
        if license_trans not in license_memo:
            license_memo.append(license_trans)
            exe_dir= os.path.join(config['check_license_exe_path'],config['check_license_exe_name'])
            save_to_log(logPathName,f"[check_license] call subprocess license {license_trans}.")
            res=subprocess.run([exe_dir, 'lmstat', '-c' ,config['check_license_cmd'],'-f', license_trans],shell=True, capture_output=True, text=True)
            stdoutstr=res.stdout
            if len(stdoutstr)==0:   # 執行檔路徑有錯誤 or 參數錯誤
                mailNotify.query_license_error(config)
                save_to_log(logPathName,"Sent Query_license_error mail.")
                raise NameError('XXXXX check license execution failed XXXXX')
            # 搜索以下 : Users of OptimizePI_20:  (Total of 2 licenses issued;  Total of 2 licenses in use)
            start_index = stdoutstr.find(f"Users of {license_trans}:")
            end_index = stdoutstr.find(")", start_index) + 1
            check_string = stdoutstr[start_index:end_index]
            numbers = re.findall(r'\d+', check_string)
            license_in_use += int(numbers[-1])
    save_to_log(logPathName,f"Total License In Use : {license_in_use}")
    return lim_count > license_in_use

def __check_update_db_result(res):
    if res.acknowledged:
        save_to_log(logPathName,'update status : ok')
    else:
        save_to_log(logPathName,'update status : fail')
# ------------------ update_setting ------------------------ # (OK)
# Function :
# 1. update some config.json info to DB
# Output :
# 1. True/ False
def update_setting():
    try:
        save_to_log(logPathName,"Running update_setting() ...")
        collect_set = utils.ConnectToMongoDB(strDBName="Simulation",strTableName="Setting")
        save_to_log(logPathName,"Clear Setting ...")
        collect_set.update_one({"_id":bson.ObjectId("66836ce4f83b57bbb93074cf")},
                                   {'$set':{}})
        save_to_log(logPathName,"Done")
        set_list = config_dict.keys()
        save_to_log(logPathName,f"Update list : [{set_list}] ")
        save_to_log(logPathName,"Update Setting ...")
        for slist in set_list:
            collect_set.update_one({"_id":bson.ObjectId("66836ce4f83b57bbb93074cf")},
                                   {'$set':{slist:config_dict[slist]}})
        save_to_log(logPathName,"Update List Done")
        return True
    except Exception as _:
        save_to_log(logPathName,f"Update Setting Error:\n {traceback.format_exc()}")
        return False
# ------------------ LisenseChoose ------------------------ # (SPEED 2000 掛 Sim3/Sim4 才會使用到)
# Function :
# 1. check which license finishes soon
# 2. giving the license for scheduled
# Output :
# 1.Lisense name
def lisense_choose():# (SPEED 2000 掛 Sim3/Sim4 才會使用到)
    #  ------------------------------------ #
    if busItem == 'DDR CCT':
        license_info_dict=__find_ddr_license_info()
    elif busItem == 'PDN':
        license_info_dict=__find_license_info()
    if not license_info_dict:
        save_to_log(logPathName,'read license info failed.')
        return False
    save_to_log(logPathName,'Read license info succeeded.')
    # ------------------------- Choose License and run Cadence with print_list ---------------------------- #
    license_name_list=list(license_info_dict.keys())
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    unfinshed_task = pd.DataFrame(collect_tc.find({'license':{'$in': license_name_list},'status':{'$in':['Running','Scheduled']}}))
    if len(unfinshed_task) !=0:
        ser_opi={}
        for opi in license_name_list:
            ser_opi[opi] = len(unfinshed_task[unfinshed_task['license']==opi].reset_index(drop=True))
        # * 任務數為0就assign給他
        min_opi = min(ser_opi, key=lambda k: ser_opi[k])
        if ser_opi[min_opi] == 0 :
            return min_opi
        else:
            # * 三個license都有任務的時候，比誰的任務 Target_date 最先到就assign給他
            for opi in license_name_list:
                df = (unfinshed_task[unfinshed_task['license']==opi].reset_index(drop=True))
                last_df = df.iloc[-1]
                ser_opi[opi] = datetime.datetime.strptime(last_df['projectSchedule']['targetDate'],'%Y-%m-%d').timestamp()
                min_opi = min(ser_opi, key=lambda k: ser_opi[k])
            save_to_log(logPathName,f"License choose : {min_opi}")
            return min_opi
    else:
        # ! 先挑第一個licenseg
        first_license_name = next(iter(license_info_dict))
        save_to_log(logPathName,f"License choose : {first_license_name}")
        return first_license_name
# ------------------SaveToLog--------------------------------
# Input: 1. log path name 
#        2. text that need to put in file
# Output: none
# -----------------------------------------------------------
def save_to_log(log_path_name,msg):
    curr_time=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    log_file = open(log_path_name, 'a', encoding='utf-8')
    write_msg="[ %s ]  %s"%(str(curr_time),str(msg))
    log_file.write(write_msg+'\n')
    log_file.close()  
# ------------------CreateLogFolder------------------------
# Input: current time
# Output: log file path and name ex. 2019-02-21_12-00-00.log
# ---------------------------------------------------------
def create_log_folder(cur_time): # "D:\\MEUpdateLog"
    # check log dir 
    curr_dir=os.getcwd()
    today = datetime.date.today()
    log_path=curr_dir+ "\\" +str(today)
    if not os.path.isdir(log_path):
        os.mkdir(log_path)
    # create log file
    current_datetime=datetime.datetime.fromtimestamp(cur_time/1000).strftime('%Y-%m-%d_%H-%M-%S.%f')
    log_file_path_name=log_path+"\\"+current_datetime+".log"
    log_file = open(log_file_path_name, 'w+',) 
    log_file.close()
    return log_file_path_name,log_path

if __name__ == "__main__" :
    curTime=int(time.time()*1000)
    computer_name = socket.gethostname().upper()
    dict_computer = {
                    "TPER90115562":"DDR CCT",
                    "TPEO54012809":"PDN"
                    }
    busItem = dict_computer[computer_name]
    logPathName,logPath=create_log_folder(curTime)
    # auto_bus 的內容為 Auto PDN/ DDR 目前不support 其他自動模擬，執行中會不斷使用 auto_bus[0] 或 auto_bus[1] 是否於 busItem 中，判斷不同設定
    auto_bus = ['Auto PDN','Auto DDR CCT']
    manual_bus = ['GPU PDN','Manual PDN','Manual DDR CCT']
    config_name = 'config.json'
    #!!! Set DB sheet !!!!!!!!!!!!!!!!!!!!!!!!!! #
    if 'dev' in os.getcwd().split('\\'):
        DB_sheet = "OPITaskCtrl_Debug"
    elif 'prd' in os.getcwd().split('\\'):
        DB_sheet = "OPITaskCtrl"
    config_dict=utils.__read_config_info(config_name)
    if busItem =='PDN':
        license_info_dict=__find_license_info()
        if not license_info_dict:
            save_to_log(logPathName,"read license info failed.")
            exit()
        do_US = update_setting()
        save_to_log(logPathName,"Read license info succeeded.")
        config = config_dict
        if 'intel_pdn_sim_days' in config:
            config['intel_pdn_sim_days']=config['intel_pdn_sim_days']*60*60*24
        if 'qcm_pdn_sim_days' in config:
            config['qcm_pdn_sim_days']=config['qcm_pdn_sim_days']*60*60*24
        if 'amd_pdn_sim_days' in config:
            config['amd_pdn_sim_days']=config['amd_pdn_sim_days']*60*60*24
        config_dict = config
        osf_warning()
        non_common_cap_used_warning()
        do_OID = over_ini_date()
        do_CR = conflict_remind()
        do_SC = schedule_change()
        do_IC = initial_check_new()
        # ! 將 OPI checkOPIStatus 進 loop
        for license_name in license_info_dict.keys() :

            save_to_log(logPathName,"Check %s"%license_name)
            OPI_status = check_pdn_status(license_name)
            print(OPI_status)
        do_OSD = over_sim_date()
    elif busItem =='DDR CCT':
        # 更改為 Power SI 執行DDR CCT 模擬，但license name 沿用 SPEED 2000
        license_name = 'SPEED2000 #1'
        # --- 不是用 Speed2000# 所以不用讀了 --- #
        # license_info_dict=__find_DDR_license_info()
        # if license_info_dict == False :
        #     SaveToLog(logPathName,"read license info failed.")
        #     exit()
        # ====================================== #
        do_US = update_setting()
        config = config_dict
        if 'ddr_sim_days' in config:
            config['ddr_sim_days']=config['ddr_sim_days']*60*60*24
        config_dict = config
        if license_name == 'SPEED2000 #1':
            do_OID = over_ini_date()
            do_CR = conflict_remind()
            do_SC = schedule_change()

        do_IC = initial_check_new()
        save_to_log(logPathName,"Read license info succeeded.")
        DDR_status = check_psi_status_new(license_name)
        print(DDR_status)
        # Sim3 排程時間5分開始 10分/ XXXXXX Sim4 25分 Over_Sim_date() 只再 Sim4 執行 XXXXX
        do_OSD = over_sim_date()
        
