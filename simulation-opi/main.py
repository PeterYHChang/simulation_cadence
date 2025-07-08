import os
import logging
import json
import bson
import time
import shutil
import datetime
import traceback
import socket
from flask import Flask,request
from flask_cors import CORS
from flask import jsonify
from package import utils, mailNotify
from functools import wraps
from dotenv import load_dotenv

# load opi.env
computer_name = socket.gethostname().upper()
dict_env = {
            "TPER90115562":r'H:\simulation-opi\opi.env',
            "TPEO54012809":r'G:\simulation-opi\opi.env'
            }
load_dotenv(dict_env[computer_name])
class JSONEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, bson.ObjectId):
            return str(o)
        return json.JSONEncoder.default(self, o)

logging.basicConfig(filename='record.log', level=logging.DEBUG)
app=Flask(__name__)
# Configure session cookie attributes
app.config['SESSION_COOKIE_SECURE'] = True       # Ensure cookies are only sent over HTTPS
app.config['SESSION_COOKIE_HTTPONLY'] = True     # Prevent JavaScript from accessing cookies
app.config['SESSION_COOKIE_SAMESITE'] = 'Strict' # Restrict cookies to same-site requests

CORS(app, resources={r'/*': {"origins": [
    "https://eSIDM-Sim2.Wistron.com",
    "https://erfdm.wistron.com"
]}})
CORS(app, resources={r'/*'})

error_sign = {
    "MISSING_FIELD":'x001',
    "MISSING_CHECKLIST_FILE":'x002',
    "MISSING_BRD_FILE":'x003',
    "MISSING_dkdf_file":'x004',
    "Cannot Store Data to DB":'x005',
    "MISSING_Form_id":'x006',
    "INSERT_FAIL":'x007'
}

@app.route('/')
def main():
  app.logger.debug("Debug log level")
  app.logger.info("Program running correctly")
  app.logger.warning("Warning; low disk space!")
  app.logger.error("Error!")
  app.logger.critical("Program halt!")
  return "logger levels!"
def csrf_protect(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        form_token = request.headers.get('X-Csrf-Token')
        if   form_token == os.getenv('CSRF_TOKEN'):
            logging.info("Token Pass")
            return f(*args, **kwargs)
        else:
            return jsonify(success=False, message="CSRF token validation failed"), 400

    return decorated_function


@app.route('/data',methods = ['POST'])
def data_endpoint():
    data = {"key": error_sign}
    return jsonify(data)

# #! 目前改版不會用到
# @app.route('/opi/initial',methods=["POST"])
# def opiInitial_old():
#     curTime=int(time.time()*1000)
#     logPathName,logPath=CreateLogFolder(curTime)
#     SaveToLog(logPathName,"opiInitial")
#     try:
#         request_data = request.form
#         require_fields = ['busItem','platform','form_id','applicant', 'boardNumber', 'boardStage', 'startDate', 'targetDate', 'gerberDate', 'smtDate','reason','customer','project_name','project_code']
#         for rf in require_fields :
#             if rf not in request_data :
#                 return JSONEncoder().encode("MISSING_FIELD"), 200
        
#         # ! 加入查form_id是否有重複送件 (在未開始執行前，只能重送一次)
#         # ! keep_order!=-1 & resend=True 才可執行正常的inital動作 
#         SaveToLog(logPathName,"check resend")
#         resend,keep_order,keep_license=__check_resend_available(request_data['form_id'],logPathName)
#         # resend,keep_order,keep_license=__check_resend_available("20231017_1697521028")
#         if resend == False:
#             return JSONEncoder().encode("Can't re-send simulation, please contact with simulation team."), 200
#         # ! --------- 讀取license資訊 ----------
#         # { 'OPI1': '-PSOptimizePI_20', 
#         #   'OPI2': '-PSOptimizePI_20', 
#         #   'OPI3': '-PSAdvancedPI_TI_20'
#         # }
#         # ! ------------------------------------
#         license_info_dict=__find_license_info()
#         license_key=list(license_info_dict.keys())
#         if license_info_dict==False :
#             SaveToLog(logPathName,"read license info failed.")
#             exit()
#         SaveToLog(logPathName,"Read license info succeeded.")

#         # 讀取 CPU Platform
#         Vendor,CPU_name,Platform,CPU_type,CPU_Target=utils.read_CPU_info(request_data['platform'],logPathName)
#         # Qualcomm 需要
#         try :
#             checklist_file = request.files['checklist']
#         except :
#             return JSONEncoder().encode("MISSING_FILE_CHECKLIST"), 400
#         # platform = pd.read_excel(checklist_file, sheet_name='PI', header=None).iloc[16,5]
#         # if len(platform) <= 0 or type(platform) != str :
#         #     return JSONEncoder().encode("MISSING_FIELD_PLATFORM"), 400
#         board_number = request_data['boardNumber']
#         board_stage = request_data['boardStage']
        
#         ct = int(time.time())
#         try :
#             brd_file = request.files['brd']
#         except :
#             return JSONEncoder().encode("MISSING_FILE_BRD"), 400
        
        
#         # LiB_PATH 指向: Lib\\Vendor\\CPU_name
#         LIB_PATH = os.path.join(os.getcwd(), 'Lib',Vendor,CPU_name)
#         SaveToLog(logPathName,"Lib path: %s."%LIB_PATH)
#         output_path = os.path.join(os.getcwd(), 'output', '%s-%s_%d'%(board_number, board_stage, ct))
#         utils.createDirectory(output_path)
#         SaveToLog(logPathName,"output file created.")
#         temp_path = os.path.join(os.getcwd(), 'temp', '%s-%s_%d'%(board_number, board_stage, ct))
#         utils.createDirectory(temp_path)
#         SaveToLog(logPathName,"temp file created.")
#         # ! 新增讀取config.json檔案內容 Erin+ 2023.10.23
#         SaveToLog(logPathName,"Read Config.json.")
#         config_dict=utils.__read_config_file('config.json')
#         if len(config_dict)==0: return JSONEncoder().encode("MISSING_FILE_CONFIG"), 400
#         SaveToLog(logPathName,"Reading Config.json done.")
#         # 讀取 Platform 的 CPU_conig
#         SaveToLog(logPathName,"Read %s_Config.json."%Vendor)
#         CPU_config_dir=os.path.join(os.getcwd(), 'CPU_config',Vendor+'_'+CPU_name+'.json')
#         with open(CPU_config_dir) as f:
#             CPU_config = json.load(f)
#         if len(CPU_config)==0: return JSONEncoder().encode("MISSING_CPU_CONFIG"), 400
#         SaveToLog(logPathName,"Reading %s_Config.json Done."%Vendor)

#         path_dict = {
#             'output_path': output_path,
#             'spd_filename': '%s.spd'%brd_file.filename.split('.')[0],
#             'print_list': os.path.join(temp_path, 'print_list.tcl'),
#             'dkdf': os.path.join(temp_path, 'stackup_dkdf.xlsx'),
#             'material': os.path.join(temp_path, 'Stackup_material_%s.cmx'%board_number),
#             'stackup_org': os.path.join(temp_path, 'stackup.csv'),
#             'stackup': os.path.join(temp_path, 'stackup_%s.csv'%board_number),
#             'brd': os.path.join(temp_path, brd_file.filename),
#             'checklist' : os.path.join(temp_path, checklist_file.filename),
#             'component': os.path.join(temp_path, 'component.csv'),
#             'component_pin': os.path.join(temp_path, 'component_pin.csv')
#         }
#         # 讀取 CPU_config 中 Path 資料
#         for key_name in CPU_config.keys():
#             if all(keyword not in key_name for keyword in ['vrm','sheet', 'model', 'Freq','report','net']):
#                 path_dict[key_name] = os.path.join(LIB_PATH, CPU_config[key_name])
#             elif 'model' in key_name:
#                 if  Vendor.upper() in ['INTEL','QUALCOMM'] and Platform:
#                     path_dict['model_path'] = os.path.join(LIB_PATH,Platform,CPU_type,CPU_config[___join_str(CPU_name,Platform,CPU_type,link='_')+'_model_path'])
#                 elif Vendor.upper() in ['INTEL','QUALCOMM'] and not Platform :
#                     path_dict['model_path'] = os.path.join(LIB_PATH,'original',CPU_type,CPU_config[___join_str(CPU_name,Platform,CPU_type,link='_')+'_model_path'])
#             elif 'vrm' in key_name:
#                 path_dict[key_name] = os.path.join(LIB_PATH,Platform,CPU_type,CPU_config[key_name])
#                 if  Vendor.upper() in ['INTEL','QUALCOMM'] and Platform:
#                     path_dict[key_name] = os.path.join(LIB_PATH,Platform,CPU_type,CPU_config[key_name])
#                 elif Vendor.upper() in ['INTEL','QUALCOMM'] and not Platform :
#                     path_dict[key_name] = os.path.join(LIB_PATH,'original',CPU_type,CPU_config[key_name])

#         # * 確認檔案是否存在，避免人工輸入檔名錯誤
#         # material_org
#         if CPU_config['material_org'] !='':
#             if not os.path.exists(path_dict['material_org']):
#                 raise NameError("MISSING_FILE : "+CPU_config['material_org'])
#         # opi_options
#         if CPU_config['opi_options'] !='':
#             if not os.path.exists(path_dict['opi_options']):
#                 raise NameError("MISSING_FILE : "+CPU_config['opi_options'])
#         # vrm
#         if CPU_config['vrm'] !='':
#             if not os.path.exists(path_dict['vrm']):
#                 raise NameError("MISSING_FILE : "+CPU_config['vrm'])
#         # main_lib_autopi
#         if CPU_config['main_lib_autopi'] !='':
#             if not os.path.exists(path_dict['main_lib_autopi']):
#                 raise NameError("MISSING_FILE : "+CPU_config['main_lib_autopi'])
#         # model_path
#         if CPU_config[___join_str(CPU_name,Platform,CPU_type,link='_')+'_model_path'] !='':
#             if not os.path.exists(path_dict['model_path']):
#                 raise NameError("MISSING_FILE : model_path")
#         # cap_lib
#         if CPU_config['cap_lib'] !='':
#             if not os.path.exists(path_dict['cap_lib']):
#                 raise NameError("MISSING_FILE : "+CPU_config['cap_lib'])
#         # Save BRD
#         try:
#             brd_file.save(path_dict['brd'])
#             SaveToLog(logPathName,'Saved BRD FILED!')
#         except :
#             SaveToLog(logPathName,'Missing BRD FILED!')
#             return JSONEncoder().encode("MISSING_BRD_FILED"), 400
#         # Save stackup.dkdf
#         try :
#             request.files['dkdf'].save(path_dict['dkdf'])
#             SaveToLog(logPathName,'Saved stackup.dkdf FILED!')
#         except :
#             SaveToLog(logPathName,'Missing stackup.dkdf FILED!')
#             return JSONEncoder().encode("MISSING_FILE_DKDF"), 400
#         # Save PDN Checklist
#         try :
#             request.files['checklist'].save(path_dict['checklist'])
#             SaveToLog(logPathName,'Saved PDN Checklist!')
#         except :
#             SaveToLog(logPathName,'Missing PDN Checklist!')
#             return JSONEncoder().encode("MISSING_FILE_checklist"), 400
        
#         SaveToLog(logPathName,"read config.json ok")
#         try:
#             insert_data = {
#                 'form_id': request_data['form_id'],
#                 'applicant': request_data['applicant'],
#                 'busItem' : request_data['busItem'],  
#                 'boardNumber': board_number,
#                 'boardStage': board_stage,
#                 'customer' : request_data['customer'],
#                 'project_code' : request_data['project_code'],
#                 'project_name' : request_data['project_name'],
#                 'product': request_data['product'],
#                 'other_file':request.files['other_file'] if 'other_file' in request.files else '',
#                 'schematic': request.files['schematic'].filename if 'schematic' in request.files else '',
#                 'createTime': ct,
#                 'initialTime': ct,
#                 'report_result':'',
#                 'stackup_no':request_data['stackup_no'],
#                 'platform': request_data['platform'],
#                 'projectSchedule': {
#                     'startDate': request_data['startDate'],
#                     'targetDate': request_data['targetDate'],
#                     'gerberDate': request_data['gerberDate'],
#                     'smtDate':request_data['smtDate']
#                 },
#                 'order': 0,
#                 'status': 'Initial_Fail',
#                 'initialResult': 'ERROR',
#                 'filePath': path_dict,
#                 'reason': request_data['reason'],
#             }
#             SaveToLog(logPathName,"Insert data create")
#         except Exception as e:
#             SaveToLog(logPathName,f"Print insert data :\n{insert_data}")
#             SaveToLog(logPathName,"set data created error %s"%(str(e)))
#         collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#         SaveToLog(logPathName,"Check Board in Task")
#         # Check first send 的"版號" 是否 In Schedules/Running 或 Over 3 個 Finished
#         if ___CheckBoardinTask(board_number,board_stage,request_data['busItem'],logPathName)=='In Schedules/Running':
#             collect_tc.insert_one(insert_data)
#             # TODO send mail : In Schedules/Running
#             SaveToLog(logPathName,"Send 'In Schedules/Running'")
#             notify_type = 'In Schedules/Running'
#             mailNotify.opiMailNotify (notify_type, insert_data)
#             return JSONEncoder().encode("INITIAL_FAIL Task Has Been In Schedules/Running"), 200
#         elif ___CheckBoardinTask(board_number,board_stage,request_data['busItem'],logPathName)=='Over 3 Finished':
#             collect_tc.insert_one(insert_data)
#             # TODO send mail : Over 3 Finished
#             SaveToLog(logPathName,"Send 'Over 3 Finished'")
#             notify_type = 'Over 3 Finished'
#             mailNotify.opiMailNotify (notify_type, insert_data)
#             return JSONEncoder().encode("INITIAL_FAIL Over 3 Finished Task"), 200
#         # ! 加入選擇license
#         try:
#             license_name_list=list(license_info_dict.keys())
#             check_license=[]
#             for license_name in license_name_list:
#                 check_license.append(__check_license(license_name,logPathName))
#             # ! 先挑第一個license
#             first_license_name = next(iter(license_info_dict))
#             use_license=license_info_dict[first_license_name]
#             # license=['-PSOptimizePI_20','-PSOptimizePI_20','-PSAdvancedPI_TI_20']
#             # check_license=[__check_license('OPI1'),__check_license('OPI2'),__check_license('OPI3')]
#             true_indexes = [index for index, value in enumerate(check_license) if value]
#         except:
#             SaveToLog(logPathName,traceback.format_exc())
#         try :
#             bLicense=False
#             if len(true_indexes):
#                 use_license=license_info_dict[license_name_list[true_indexes[0]]]
#                 bLicense=True
#                 SaveToLog(logPathName," Use license %s" %(license_name_list[true_indexes[0]]))
#             else: # ! 目前沒有license
#                 SaveToLog(logPathName," **** no available_license for initial ****")
#             if bLicense==False:
#                 notify_type = 'NO_LICENSE'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 raise NameError('NO_LICENSE')
#             try :
#                 SaveToLog(logPathName," running gentcl.stackup()" )
#                 gentcl.stackup(path_dict)
#                 # subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe', '-PSAdvancedPI_TI_20' ,'-tcl', path_dict['print_list']])
#                 # subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe', use_license ,'-tcl', path_dict['print_list']])
#                 # ! 加入選擇license
#                 # license=['-OptimizePI_20','-OptimizePI_20','-PSAdvancedPI_TI_20']
#                 # check_license=[__check_license('OPI1'),__check_license('OPI2'),__check_license('OPI3')]
#                 # true_indexes = [index for index, value in enumerate(check_license) if value]
#                 if bLicense:
#                     SaveToLog(logPathName,"available_license for initial: "+use_license)
#                     SaveToLog(logPathName,"print_list: "+path_dict['print_list'])
#                     # subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe', use_license ,'-tcl', path_dict['print_list']])
#                     result=subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe',use_license ,'-tcl', path_dict['print_list']], capture_output=True)
#                     SaveToLog(logPathName,"subprocess result: "+str(result.returncode))

#                 else:
#                     SaveToLog(logPathName," **** no available_license for initial ****")
#                     # TODO send mail : NO_LICENSE
#                     notify_type = 'NO_LICENSE'
#                     mailNotify.opiMailNotify (notify_type, insert_data)
#                     raise NameError('NO_LICENSE')
#             except Exception as e:
#                 # TODO send mail : PARSE_ERROR
#                 notify_type = 'PARSE_ERROR'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 raise NameError('PARSE_ERROR')
#             ### get power information (TODO: Intel)
#             try :
#                 power_info = utils.getPowerInfo(board_number, board_stage)
#                 df_power = pd.DataFrame(power_info)
#                 if len(df_power)==0:
#                     SaveToLog(logPathName,"POWER API NO DATA")
#                 else:
#                     SaveToLog(logPathName,"Get POWER API DATA")
#                 if Vendor=='Intel' and ('OutputNet2' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns):
#                     raise NameError('POWER_API_NO_DATA')
#             except Exception as e:
#                 # TODO send mail : POWER_API_ERROR
#                 notify_type = 'POWER_API_ERROR'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 SaveToLog(logPathName," POWER_API_ERROR MAIL SEND" )
#                 raise NameError('POWER_API_ERROR')
#             ###parse dk df data (TODO: Intel/ Qualcomm)
#             try :
#                 data_DKDF = utils.parseStackup(path_dict['dkdf'])
#                 SaveToLog(logPathName,"parseStackup Done")
#             except Exception as e:
#                     #TODO send mail : PARSE_ERROR_DKDF
#                     notify_type = 'PARSE_ERROR_DKDF'
#                     mailNotify.opiMailNotify (notify_type, insert_data)
#                     SaveToLog(logPathName,"PARSE_ERROR_DKDF MAIL SEND" )
#                     raise NameError('PARSE_ERROR_DKDF')
#             if data_DKDF :
#                 ### gen material cmx (TODO: Intel/ Qualcomm)
#                 try :
#                     tree = utils.modifyMaterial(data_DKDF['layer'], path_dict['material_org'], data_DKDF['data'])
#                     tree.write(path_dict['material'])
#                     SaveToLog(logPathName,"modifyMaterial Done")
#                 except :
#                     #TODO send mail : GENCMX_ERROR
#                     notify_type = 'GENCMX_ERROR'
#                     mailNotify.opiMailNotify (notify_type,insert_data)
#                     SaveToLog(logPathName,"GENCMX_ERROR MAIL SEND" )
#                     raise NameError('GENCMX_ERROR')
#                 ### gen stackup csv (TODO: Intel/ Qualcomm)
#                 try :
#                     modified_stackup = utils.modifyStackup(data_DKDF['layer'], path_dict['stackup_org'])
#                     modified_stackup.to_csv(path_dict['stackup'], index=False)
#                     SaveToLog(logPathName,"modifyStackup Done")
#                 except :
#                     #TODO send mail : GENSTACKUP_ERROR
#                     notify_type = 'GENSTACKUP_ERROR'
#                     mailNotify.opiMailNotify (notify_type, insert_data)
#                     SaveToLog(logPathName,"GENSTACKUP_ERROR MAIL SEND" )
#                     raise NameError('GENSTACKUP_ERROR')
#                     ### Parse VRM (TODO: Intel)
#                 try :
#                     if CPU_config['vrm'] != '':
#                         df_vrm = utils.parseVRM(path_dict['vrm'])
#                         SaveToLog(logPathName,"parseVRM Done")
#                     else:
#                         df_vrm = pd.DataFrame()
#                         SaveToLog(logPathName,"VRM NO DATA")
#                 except:
#                     #TODO send mail : VRM_ERROR
#                     notify_type = 'VRM_ERROR'
#                     mailNotify.opiMailNotify (notify_type, insert_data)
#                     SaveToLog(logPathName,"VRM_ERROR MAIL SEND" )
#                     raise NameError('VRM_ERROR')
#                 ### gen final tcl
#                 try :
#                     SaveToLog(logPathName,"Gentcl Start")
#                     tcl_list = gentcl.OPI(path_dict, df_power,df_vrm,Vendor,CPU_name,Platform,CPU_type,CPU_Target)
#                     SaveToLog(logPathName,"Gentcl End")
#                 except :
#                     #TODO send mail : GEN TCL_ERROR
#                     notify_type = 'GEN TCL_ERROR'
#                     mailNotify.opiMailNotify (notify_type, insert_data)
#                     SaveToLog(logPathName,"GEN TCL_ERROR MAIL SEND" )
#                     raise NameError('GEN TCL_ERROR')
#             else :
#                 #TODO send mail : DKDF_FORMAT_ERROR
#                 notify_type = 'DKDF_FORMAT_ERROR'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 SaveToLog(logPathName,"DKDF_FORMAT_ERROR MAIL SEND" )
#                 raise NameError('DKDF_FORMAT_ERROR')
#             unfinshed_task = pd.DataFrame(collect_tc.find({'status':{'$in':['Running','Scheduled']}}))
#             insert_data['tcl'] = tcl_list
#             if len(unfinshed_task) !=0:
#                 ser_opi={}
#                 for opi in license_key:
#                     ser_opi[opi] = len(unfinshed_task[unfinshed_task['license']==opi].reset_index(drop=True))
#                 # * 三個license都有任務的時候，比誰的任務少就assign給他
#                 min_opi = min(ser_opi, key=lambda k: ser_opi[k])
#                 insert_data['license'] = min_opi
#                 SaveToLog(logPathName,"assign schedule for %s"%min_opi)
#                 # 設定 order 數
#                 ser=unfinshed_task[unfinshed_task['license']==min_opi].reset_index(drop=True)
#                 if  ser_opi[min_opi] !=0:
#                     insert_data['order'] = int(ser['order'].max() + 1)
#                     insert_data['current_opi_start_dt'] = int(ser.loc[0, 'current_opi_start_dt'])
#                 else:
#                     insert_data['order']=1
#                     insert_data['current_opi_start_dt'] = int(time.time())
#                 insert_data['initialResult'] = 'Succeed'
#                 insert_data['status'] = 'Scheduled'
#                 insert_data['tcl'] = tcl_list
#             else:
#                 insert_data['license'] = license_key[0]
#                 SaveToLog(logPathName,"assign schedule for %s"%license_key[0])
#                 insert_data['order']=1
#                 insert_data['current_opi_start_dt'] = int(time.time()) #BUG: need to check # int(unfinshed_task.loc[0, 'current_opi_start_dt'])
#                 insert_data['initialResult'] = 'Succeed'
#                 insert_data['status'] = 'Scheduled'
#                 insert_data['tcl'] = tcl_list
#             try :
#                 if resend and keep_order!=-1: # ! 如果為重送的，要維持一樣的license, order
#                     insert_data['license'] = keep_license
#                     insert_data['order']=keep_order
#                     insert_data['resent_count']=1
#                 collect_tc.insert_one(insert_data)
#                 # TODO send mail : initial_success
#                 notify_type = 'initial_success'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 return JSONEncoder().encode("OK"), 200
#             except:
#                 return JSONEncoder().encode("INSERT_FAIL"), 400
#         except Exception as e :
#             try :
#                 SaveToLog(logPathName,traceback.format_exc())
#                 insert_data['initialResult'] = str(e)
#                 collect_tc.insert_one(insert_data)
#                 return JSONEncoder().encode("INITIAL_FAIL"), 200
#             except :
#                 return JSONEncoder().encode("INSERT_FAIL for initial fail"), 400
#     except Exception as e :
#         SaveToLog(logPathName,f"ERROR : {e} \n {traceback.format_exc()}")
#         return JSONEncoder().encode("Initial ERROR"), 400

# #! 目前改版不會用到
# @app.route('/DDR/ddr_initial',methods=["POST"])
# def ddr_initial():
#     curTime=int(time.time()*1000)
#     logPathName,logPath=CreateLogFolder(curTime)
#     request_data = request.form
#     require_fields = ['busItem','platform','form_id','applicant', 'boardNumber', 'boardStage',
#                        'startDate', 'targetDate', 'gerberDate', 'smtDate','reason','customer','project_name','project_code',
#                       'Mapping','PCBType','DDRModule','RamType','Rank','DataRate']
#     for rf in require_fields :
#         if rf not in request_data :
#             return JSONEncoder().encode("MISSING_FIELD"), 200
#     #******* 是否上傳 DB Ture/False
#     ToMongo=True
#     SaveToLog(logPathName,f"If Updated to DB : {ToMongo}")
#     # ! 加入查form_id是否有重複送件 (在未開始執行前，只能重送一次)
#     # ! keep_order!=-1 & resend=True 才可執行正常的inital動作 
#     SaveToLog(logPathName,"check resend")
#     try:
#         resend,keep_order,keep_license=__check_resend_available(request_data['form_id'],logPathName)
#         # resend,keep_order,keep_license=__check_resend_available("20231017_1697521028")
#         if resend == False:
#             return JSONEncoder().encode("Can't re-send simulation, please contact with simulation team."), 200
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")

#     license_info_dict=__find_license_info()
#     # license_key=list(license_info_dict.keys())
#     if license_info_dict==False :
#         SaveToLog(logPathName,"read license info failed.")
#         exit()
#     SaveToLog(logPathName,"Read license info succeeded.")

    
#     try :
#         checklist_file = request.files['checklist']
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#         return JSONEncoder().encode("MISSING_FILE_CHECKLIST"), 400
#     board_number = request_data['boardNumber']
#     board_stage = request_data['boardStage']
#     ct = int(time.time())
#     try :
#         brd_file = request.files['brd']
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#         return JSONEncoder().encode("MISSING_FILE_BRD"), 400
    
#     # 讀取 CPU Platform
#     Vendor,CPU_name,platform,CPU_type,CPU_Target=utils.read_CPU_info(request_data['platform'],logPathName)
#     SaveToLog(logPathName,f"Parse CPU name : {Vendor} - {CPU_name} -{platform}")
#     # LiB_PATH 指向: Lib\\Vendor\\CPU_name
#     LIB_PATH = os.path.join(os.getcwd(), 'Lib',Vendor,CPU_name)
#     SaveToLog(logPathName,f"Lib path: {LIB_PATH}")
#     output_path = os.path.join(os.getcwd(), 'output', f"{board_number}-{board_stage}_DDR_{ct}")
#     utils.createDirectory(output_path)
#     SaveToLog(logPathName,"output file created.")
#     temp_path = os.path.join(os.getcwd(), 'temp', f"{board_number}-{board_stage}_{ct}")
#     utils.createDirectory(temp_path)
#     SaveToLog(logPathName,"temp file created.")
#     with open(os.path.join(os.getcwd(),'power_pin.json'), encoding='utf-8') as f:
#         power_pin=json.load(f)
#     dict_power_pin=power_pin['Memory_type']
#     if len(dict_power_pin)==0: return JSONEncoder().encode("MISSING_Power.Pin.json"), 400
#     SaveToLog(logPathName,"Reading power_pin.json done.")

#     # ======================== end ========================== #
#     # 讀取 Platform 的 CPU_conig
#     SaveToLog(logPathName,"Read %s_Config.json."%Vendor)
#     CPU_config_dir=os.path.join(os.getcwd(), 'CPU_config',Vendor+'_'+CPU_name+'.json')
#     with open(CPU_config_dir) as f:
#         CPU_config = json.load(f)
#     if len(CPU_config)==0: return JSONEncoder().encode("MISSING_CPU_CONFIG"), 400
#     #!-----------------DDR path_dict 建立--------------------#
#     path_dict = {
#         'output_path': output_path,
#         'spd_filename': f"{brd_file.filename.split('.')[0]}.spd",
#         'print_list': os.path.join(temp_path, 'print_list.tcl'),
#         'dkdf': os.path.join(temp_path, 'stackup_dkdf.xlsx'),
#         'material': os.path.join(temp_path, f"Stackup_material_{board_number}.cmx"),
#         'stackup_org': os.path.join(temp_path, 'stackup.csv'),
#         'stackup': os.path.join(temp_path, f"stackup_{board_number}.csv"),
#         'brd': os.path.join(temp_path, brd_file.filename),
#         'checklist' : os.path.join(temp_path, checklist_file.filename),
#         'component': os.path.join(temp_path, 'component.csv'),
#         'component_pin': os.path.join(temp_path, 'component_pin.csv'),
#         'opi_options': os.path.join(LIB_PATH,'DDRCCT_Options.xml')
#     }
#     # 讀取 CPU_config 中 Path 資料
#     for key_name in CPU_config.keys():
#         if all(keyword not in key_name for keyword in ['vrm','sheet', 'model', 'Freq','report','net']):
#             path_dict[key_name] = os.path.join(LIB_PATH, CPU_config[key_name])
#     # * 確認檔案是否存在，避免人工輸入檔名錯誤
#     # material_org
#     if CPU_config['material_org'] !='':
#         if not os.path.exists(path_dict['material_org']):
#             raise NameError("MISSING_FILE : "+CPU_config['material_org'])
#     # opi_options
#     if CPU_config['opi_options'] !='':
#         if not os.path.exists(path_dict['opi_options']):
#             raise NameError("MISSING_FILE : "+CPU_config['opi_options'])
#     SaveToLog(logPathName,f"Path_dict Created :\t{path_dict}")
#     #!----------------- 建立 output file ---------------- #
#     if not os.path.exists(path_dict['output_path']) :
#         utils.createDirectory(path_dict['output_path'])
#         SaveToLog(logPathName,'Output path file created!!!')
#     tcl_script_path=os.path.join(temp_path,'script')
#     if not os.path.exists(tcl_script_path) :
#         utils.createDirectory(tcl_script_path)
#         SaveToLog(logPathName,'script folder created!!!')
#     # ======================== end ========================== #
#     # Save BRD
#     try:
#         brd_file.save(path_dict['brd'])
#         SaveToLog(logPathName,'Saved BRD FILED!')
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#         return JSONEncoder().encode("MISSING_BRD_FILED"), 400
#     # Save stackup.dkdf
#     try :
#         request.files['dkdf'].save(path_dict['dkdf'])
#         SaveToLog(logPathName,'Saved stackup.dkdf FILED!')
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#         return JSONEncoder().encode("MISSING_STACKUP_DKDF"), 400
#     # Save PDN Checklist
#     try :
#         request.files['checklist'].save(path_dict['checklist'])
#         SaveToLog(logPathName,'Saved PDN Checklist!')
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#         return JSONEncoder().encode("MISSING_FILE_checklist"), 400
#     try:
#         insert_data = {
#             'form_id': request_data['form_id'],
#             'applicant': request_data['applicant'],
#             'busItem' : request_data['busItem'],  
#             'boardNumber': board_number,
#             'boardStage': board_stage,
#             'customer' : request_data['customer'],
#             'project_code' : request_data['project_code'],
#             'project_name' : request_data['project_name'],
#             'product': request_data['product'],
#             'other_file':request.files['other_file'] if 'other_file' in request.files else '',
#             'schematic': request.files['schematic'].filename if 'schematic' in request.files else '',
#             'createTime': ct,
#             'initialTime': ct,
#             'report_result':'',
#             'stackup_no':request_data['stackup_no'],
#             'platform': request_data['platform'],
#             'projectSchedule': {
#                 'startDate': request_data['startDate'],
#                 'targetDate': request_data['targetDate'],
#                 'gerberDate': request_data['gerberDate'],
#                 'smtDate':request_data['smtDate']
#             },
#             'order': 0,
#             'status': 'Initial_Fail',
#             'initialResult': 'ERROR',
#             'filePath': path_dict,
#             'reason': request_data['reason'],
#             'Mapping':request_data['Mapping'],
#             'license': server_name,
#             'PCBType':request_data['PCBType'],
#             'DDRModule':request_data['DDRModule'],
#             'RamType':request_data['RamType'],
#             'Rank':request_data['Rank'],
#             'DataRate':request_data['DataRate']
#         }
#         SaveToLog(logPathName,f"assign schedule for {server_name}")
#         SaveToLog(logPathName,f"set data : \t {insert_data}")
#     except Exception as e:
#         SaveToLog(logPathName,"XXXXX set data error %s XXXXX"%(str(e)))
#     collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName="OPITaskCtrl")
#     SaveToLog(logPathName,"Check Board in Task")
#     # Check first send 的"版號" 是否 In Schedules/Running 或 Over 3 個 Finished
#     if ___CheckBoardinTask(board_number,board_stage,request_data['busItem'],logPathName)=='In Schedules/Running':
#         collect_tc.insert_one(insert_data)
#         # TODO send mail : In Schedules/Running
#         SaveToLog(logPathName,"Send 'In Schedules/Running'")
#         notify_type = 'In Schedules/Running'
#         mailNotify.opiMailNotify (notify_type, insert_data)
#         return JSONEncoder().encode("INITIAL_FAIL Task Has Been In Schedules/Running"), 200
#     elif ___CheckBoardinTask(board_number,board_stage,request_data['busItem'],logPathName)=='Over 3 Finished':
#         collect_tc.insert_one(insert_data)
#         # TODO send mail : Over 3 Finished
#         SaveToLog(logPathName,"Send 'Over 3 Finished'")
#         notify_type = 'Over 3 Finished'
#         mailNotify.opiMailNotify (notify_type, insert_data)
#         return JSONEncoder().encode("INITIAL_FAIL Over 3 Finished Task"), 200
#     SaveToLog(logPathName,"Check Board OK")
#     # ! 加入選擇license
#     try:
#         license_name_list=list(license_info_dict.keys())
#         check_license=[]
#         for license_name in license_name_list:
#             check_license.append(__check_license(license_name,logPathName))
#         # ! 先挑第一個license
#         first_license_name = next(iter(license_info_dict))
#         use_license=license_info_dict[first_license_name] 
#         true_indexes = [index for index, value in enumerate(check_license) if value]
#     except Exception as e:
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#     try :
#         bLicense=False
#         if len(true_indexes):
#             use_license=license_info_dict[license_name_list[true_indexes[0]]]
#             bLicense=True
#             SaveToLog(logPathName,f"Use license {license_name_list[true_indexes[0]]}")
#         else: # ! 目前沒有license
#             SaveToLog(logPathName," **** no available_license for initial ****")
#         if bLicense==False:
#             notify_type = 'NO_LICENSE'
#             mailNotify.opiMailNotify (notify_type, insert_data)
#             raise NameError('NO_LICENSE')
#         try :
#             SaveToLog(logPathName," running gentcl.stackup()" )
#             DDRGentcl.stackup(path_dict)
#             if bLicense:
#                 SaveToLog(logPathName,"available_license for initial: "+use_license)
#                 SaveToLog(logPathName,"print_list: "+path_dict['print_list'])
#                 result=subprocess.run([r'C:\CADENCE\Sigrity2023.1\tools\bin\OptimizePI.exe',use_license ,'-tcl', path_dict['print_list']], capture_output=True)
#                 SaveToLog(logPathName,"subprocess result: "+str(result.returncode))
#             else:
#                 SaveToLog(logPathName," **** no available_license for initial ****")
#                  # TODO send mail : NO_LICENSE
#                 notify_type = 'NO_LICENSE'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 raise NameError('NO_LICENSE')
#         except :
#              # TODO send mail : PARSE_ERROR
#             notify_type = 'PARSE_ERROR'
#             mailNotify.opiMailNotify (notify_type, insert_data)
#             raise NameError('PARSE_ERROR')
#         ### parse dk df data (TODO: Intel/ Qualcomm)
#         try :
#             data_DKDF = utils.parseStackup(path_dict['dkdf'])
#             SaveToLog(logPathName,"parseStackup Done")
#         except Exception as e:
#                 #TODO send mail : PARSE_ERROR_DKDF
#                 notify_type = 'PARSE_ERROR_DKDF'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 SaveToLog(logPathName,"PARSE_ERROR_DKDF MAIL SEND" )
#                 SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#                 raise NameError('PARSE_ERROR_DKDF')
#         if data_DKDF :
#             try :
#                 tree = utils.modifyMaterial(data_DKDF['layer'], path_dict['material_org'], data_DKDF['data'])
#                 tree.write(path_dict['material'])
#                 SaveToLog(logPathName,"modifyMaterial Done")
#             except Exception as e:
#                 #TODO send mail : GENCMX_ERROR
#                 notify_type = 'GENCMX_ERROR'
#                 mailNotify.opiMailNotify (notify_type,insert_data)
#                 SaveToLog(logPathName,"GENCMX_ERROR MAIL SEND" )
#                 SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#                 raise NameError('GENCMX_ERROR')
#             ## gen stackup csv 
#             try :
#                 modified_stackup = utils.modifyStackup(data_DKDF['layer'], path_dict['stackup_org'])
#                 modified_stackup.to_csv(path_dict['stackup'], index=False)
#                 SaveToLog(logPathName,"modifyStackup Done")
#             except Exception as e:
#                 #TODO send mail : GENSTACKUP_ERROR
#                 notify_type = 'GENSTACKUP_ERROR'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 SaveToLog(logPathName,"GENSTACKUP_ERROR MAIL SEND" )
#                 SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#                 raise NameError('GENSTACKUP_ERROR')
#             ### gen final tcl
#             try :
#                 SaveToLog(logPathName,"Gentcl Start")
#                 tcl_list = DDRGentcl._DDRGentcl(path_dict,board_number,board_stage,Vendor,insert_data['DDRModule'],insert_data['Mapping'])
#                 SaveToLog(logPathName,"Gentcl End")
#             except Exception as e:
#                 #TODO send mail : GEN TCL_ERROR
#                 notify_type = 'GEN TCL_ERROR'
#                 mailNotify.opiMailNotify (notify_type, insert_data)
#                 SaveToLog(logPathName,"GEN TCL_ERROR MAIL SEND" )
#                 raise NameError('GEN TCL_ERROR')
#         else :
#             #TODO send mail : DKDF_FORMAT_ERROR
#             notify_type = 'DKDF_FORMAT_ERROR'
#             mailNotify.opiMailNotify (notify_type, insert_data)
#             SaveToLog(logPathName,"DKDF_FORMAT_ERROR MAIL SEND" )
#             raise NameError('DKDF_FORMAT_ERROR')

#         unfinshed_task = pd.DataFrame(collect_tc.find({'$and': [{'status': {'$in': ['Running', 'Scheduled']}}, {'license': server_name}]}))
#         insert_data['tcl'] = tcl_list
#         if len(unfinshed_task) != 0:
#             SaveToLog(logPathName,f"assign schedule for {server_name}")
#             # 設定 order 數
#             insert_data['order'] = int(unfinshed_task['order'].max() + 1)
#             insert_data['current_opi_start_dt'] = int(unfinshed_task.loc[0, 'current_opi_start_dt'])
#             insert_data['initialResult'] = 'Succeed'
#             insert_data['status'] = 'Scheduled'
#             insert_data['tcl'] = tcl_list
#         else:
#             SaveToLog(logPathName,f"assign schedule for {server_name}")
#             insert_data['order']=1
#             insert_data['current_opi_start_dt'] = int(time.time()) #BUG: need to check # int(unfinshed_task.loc[0, 'current_opi_start_dt'])
#             insert_data['initialResult'] = 'Succeed'
#             insert_data['status'] = 'Scheduled'
#             insert_data['tcl'] = tcl_list
#         # =================================================== end ===================================================== #
#         try :
#             if resend and keep_order!=-1: # ! 如果為重送的，要維持一樣的license, order
#                 insert_data['license'] = keep_license
#                 insert_data['order']=keep_order
#                 insert_data['resent_count']=1
#             if ToMongo:
#                 collect_tc.insert_one(insert_data)
#                 SaveToLog(logPathName,f"Create DB")
#             # TODO send mail : initial_success
#             notify_type = 'initial_success'
#             mailNotify.opiMailNotify (notify_type, insert_data)
#             return JSONEncoder().encode("Initial_success"), 200
#         except Exception as e:
#             SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#             return JSONEncoder().encode("INSERT_FAIL"), 400
#     except Exception as e :
#         insert_data['initialResult'] = str(e)
#         if ToMongo:
#             collect_tc.insert_one(insert_data)
#             SaveToLog(logPathName,f"Create DB")
#         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#         return JSONEncoder().encode("INITIAL_FAIL"), 200

#! initial/v2 才將 PDN/DDR CCT 整合
@app.route('/opi/initial/v2',methods=["POST"])
@csrf_protect
def opiInitial():
    curTime=int(time.time()*1000)
    logPathName,_=create_log_folder(curTime)
    save_to_log(logPathName,"opiInitial")
    do=True
    ct = int(time.time())
    request_data = request.form
    if 'busItem' not in request_data :
        return JSONEncoder().encode(error_sign["MISSING_FIELD"]), 400
    try:
        if request_data['busItem'] in ['GPU PDN']:
            insert_data = {
            'product': request_data['product'],
            'other_file':request.files['other_file'] if 'other_file' in request.files else '',
            'schematic': request.files['schematic'].filename if 'schematic' in request.files else '',
            'createTime': ct,
            'initialTime': ct,
            'report_result':'',
            'stackup_no':request_data['stackup_no'],
            'initialResult': 'Succeed',
            'filePath': '',
            'reason': request_data['reason'],
            'current_opi_start_dt':ct,
            'fail_times':0
            }
            collect_tc = utils.ConnectToMongoDB(str_db_name="Simulation",str_tb_name=DB_sheet)
            collect_tc.update_one(
                                {'form_id': request_data['form_id']},
                                {'$set': insert_data})
            save_to_log(logPathName,f"Form ID : {request_data['form_id']}, Get {request_data['busItem']} Mission")
            return JSONEncoder().encode(f"Get {request_data['busItem']} Mission"), 200
        if 'PDN' in request_data['busItem']:
            require_fields = ['busItem','platform','form_id','applicant', 'boardNumber',
                            'boardStage', 'startDate', 'targetDate', 'gerberDate','smtDate',
                            'reason','customer','project_name','project_code','product']
        elif 'DDR CCT' in request_data['busItem']:
            require_fields = ['busItem','platform','form_id','applicant', 'boardNumber', 'boardStage','product',
                        'startDate', 'targetDate', 'gerberDate', 'smtDate','reason','customer','project_name','project_code',
                        'Mapping','PCBType','DDRModule','RamType','Rank','DataRate']
        else:
            save_to_log(logPathName,f"Form ID : {request_data['form_id']}")
            return JSONEncoder().encode(f"Get {request_data['busItem']} Mission"), 200
        for rf in require_fields :
            if rf not in request_data :
                save_to_log(logPathName,f"XXXXX MISSING_FIELD :{rf} XXXXX")
                return JSONEncoder().encode(error_sign["MISSING_FIELD"]), 400
        save_to_log(logPathName,f"Form ID : {request_data['form_id']}")
        # 讀取 CPU Platform
        Vendor,CPU_name,Platform,CPU_type,CPU_Target=utils.read_CPU_info(request_data['platform'])
        save_to_log(logPathName,f"Read Platform Result :{Vendor},{CPU_name},{Platform},{CPU_type},{CPU_Target}")
        # Qualcomm 需要
        try :
            checklist_file = request.files['checklist']
        except Exception as _:
            save_to_log(logPathName,"XXXXX MISSING_CHECKLIST_FILE XXXXX")
            return JSONEncoder().encode(error_sign ["MISSING_CHECKLIST_FILE"]), 400
        try :
            brd_file = request.files['brd']
        except Exception as _:
            save_to_log(logPathName,"XXXXX MISSING_BRD_FILE XXXXX")
            return JSONEncoder().encode(error_sign["MISSING_BRD_FILE"]), 400
        try :
            dkdf_file = request.files['dkdf']
        except Exception as _:
            save_to_log(logPathName,"XXXXX MISSING_dkdf_file XXXXX")
            return JSONEncoder().encode(error_sign["MISSING_dkdf_file"]), 400
        
        board_number = request_data['boardNumber']
        form_id = request_data['form_id']
        ct = int(time.time())
        # LiB_PATH 指向: Lib\\Vendor\\CPU_name
        LIB_PATH = os.path.join(os.getcwd(), 'Lib',Vendor,CPU_name)
        save_to_log(logPathName,"Lib path: %s."%LIB_PATH)
        temp_path = os.path.join(os.getcwd(), 'temp', f"{form_id}")
        if os.path.exists(temp_path):
            shutil.rmtree(temp_path)
            save_to_log(logPathName,f"remove temp file : {temp_path}")
        utils.createDirectory(temp_path)
        save_to_log(logPathName,"temp file created.")

        if  'PDN' in request_data['busItem']:
            # check prt file only for auto pdn
            try :
                prt_file = request.files['prt']
            except Exception as _:
                save_to_log(logPathName,"XXXXX MISSING_prt_file XXXXX")
                return JSONEncoder().encode(error_sign["MISSING_prt_file"]), 400
            
            output_path = os.path.join(os.getcwd(), 'output', f"{form_id}")
            if os.path.exists(output_path):
                shutil.rmtree(output_path)
                save_to_log(logPathName,f"remove temp file : {output_path}")
            utils.createDirectory(output_path)
            save_to_log(logPathName,"output file created.")
            # ! 新增讀取config.json檔案內容 Erin+ 2023.10.23
            save_to_log(logPathName,"Read Config.json.")
            config_dict=utils.__read_config_file('config.json')
            if len(config_dict)==0:
                do=False
                save_to_log(logPathName,"XXXXX MISSING_CONFIG XXXXX")
            else: save_to_log(logPathName,"Reading Config.json done.")
        elif 'DDR CCT' in request_data['busItem']:
            output_path = os.path.join(os.getcwd(), 'output', f"{form_id}")
            if os.path.exists(output_path):
                shutil.rmtree(output_path)
                save_to_log(logPathName,f"remove temp file : {output_path}")
            utils.createDirectory(output_path)
            save_to_log(logPathName,"output file created.")
            tcl_script_path=os.path.join(temp_path,'script')
            if not os.path.exists(tcl_script_path) :
                utils.createDirectory(tcl_script_path)
                save_to_log(logPathName,'script folder created!!!')
            with open(os.path.join(os.getcwd(),'power_pin.json'), encoding='utf-8') as f:
                power_pin=json.load(f)
            dict_power_pin=power_pin['Memory_type']
            if len(dict_power_pin)==0: 
                do=False
                save_to_log(logPathName,"XXXXX MISSING_Power.Pin.json XXXXX")
            else:save_to_log(logPathName,"Reading power_pin.json done.")

        #! 讀取 Platform 的 CPU_conig
        save_to_log(logPathName,f"Read {Vendor}_{CPU_name}.json.")
        CPU_config_dir=os.path.join(os.getcwd(), 'CPU_config',Vendor+'_'+CPU_name+'.json')
        CPU_config =utils.__read_config_file(CPU_config_dir)
        if len(CPU_config)==0: 
            do=False
            save_to_log(logPathName,f"XXXXX MISSING_{Vendor}_{CPU_name}.json XXXXX")
        else:  save_to_log(logPathName,f"Reading {Vendor}_{CPU_name}.json Done.")

        if  'PDN' in request_data['busItem']:
            path_dict = {
                'output_path': output_path,
                'spd_filename': '%s.spd'%brd_file.filename.split('.')[0],
                'print_list': os.path.join(temp_path, 'print_list.tcl'),
                'dkdf': os.path.join(temp_path, dkdf_file.filename),
                'prt':os.path.join(temp_path, prt_file.filename),
                'material': os.path.join(temp_path, 'Stackup_material_%s.cmx'%board_number),
                'stackup_org': os.path.join(temp_path, 'stackup.csv'),
                'stackup': os.path.join(temp_path, 'stackup_%s.csv'%board_number),
                'brd': os.path.join(temp_path, brd_file.filename),
                'checklist' : os.path.join(temp_path, checklist_file.filename),
                'component': os.path.join(temp_path, 'component.csv'),
                'component_pin': os.path.join(temp_path, 'component_pin.csv')
                }
            # 讀取 CPU_config 中 Path 資料
            for key_name in CPU_config.keys():
                if all(keyword not in key_name for keyword in ['vrm','sheet', 'model', 'Freq','report','net']):
                    path_dict[key_name] = os.path.join(LIB_PATH, CPU_config[key_name])
                elif 'model' in key_name:
                    if Platform:
                        path_dict['model_path'] = os.path.join(LIB_PATH,Platform,CPU_type,CPU_config[___join_str(CPU_name,Platform,CPU_type,link='_')+'_model_path'])
                    else:
                        path_dict['model_path'] = os.path.join(LIB_PATH,'original',CPU_type,CPU_config[___join_str(CPU_name,Platform,CPU_type,link='_')+'_model_path'])
                elif 'vrm' in key_name:
                    if  Platform:
                        path_dict[key_name] = os.path.join(LIB_PATH,Platform,CPU_type,CPU_config[key_name])
                    else :
                        path_dict[key_name] = os.path.join(LIB_PATH,'original',CPU_type,CPU_config[key_name])
            # Save prt file only for auto pdn
            try :
                request.files['prt'].save(path_dict['prt'])
                save_to_log(logPathName,'Saved PDN prt File!')
            except Exception as _:
                do=False
                save_to_log(logPathName,'XXXXX Missing prt File! XXXXX')
        elif 'DDR CCT' in request_data['busItem']:
            #!-----------------DDR path_dict 建立--------------------#
            path_dict = {
                'output_path': output_path,
                'spd_filename': f"{brd_file.filename.split('.')[0]}.spd",
                'print_list': os.path.join(temp_path, 'print_list.tcl'),
                'dkdf': os.path.join(temp_path, 'stackup_dkdf.xlsx'),
                'material': os.path.join(temp_path, f"Stackup_material_{board_number}.cmx"),
                'stackup_org': os.path.join(temp_path, 'stackup.csv'),
                'stackup': os.path.join(temp_path, f"stackup_{board_number}.csv"),
                'brd': os.path.join(temp_path, brd_file.filename),
                'checklist' : os.path.join(temp_path, checklist_file.filename),
                'component': os.path.join(temp_path, 'component.csv'),
                'component_pin': os.path.join(temp_path, 'component_pin.csv'),
                'opi_options': os.path.join(LIB_PATH,'DDRCCT_Options.xml')
                }
            # 讀取 CPU_config 中 Path 資料
            for key_name in CPU_config.keys():
                if all(keyword not in key_name for keyword in ['vrm','sheet', 'model', 'Freq','report','net']):
                    path_dict[key_name] = os.path.join(LIB_PATH, CPU_config[key_name])
        # * 確認檔案是否存在，避免人工輸入檔名錯誤
        # material_org
        if CPU_config['material_org'] !='':
            if not os.path.exists(path_dict['material_org']):
                do=False
                save_to_log(logPathName,f"XXXXX MISSING_FILE : {CPU_config['material_org']} XXXXX")
        # opi_options
        if CPU_config['opi_options'] !='':
            if not os.path.exists(path_dict['opi_options']):
                do=False
                save_to_log(logPathName,f"XXXXX MISSING_FILE : {CPU_config['opi_options']} XXXXX")
        if 'Auto PDN' in request_data['busItem']:
            # vrm
            if CPU_config['vrm'] !='':
                if not os.path.exists(path_dict['vrm']):
                    do=False
                    save_to_log(logPathName,f"XXXXX MISSING_FILE : {CPU_config['vrm']} XXXXX")
            # main_lib_autopi
            if CPU_config['main_lib_autopi'] !='':
                if not os.path.exists(path_dict['main_lib_autopi']):
                    do=False
                    save_to_log(logPathName,f"XXXXX MISSING_FILE : {CPU_config['main_lib_autopi']} XXXXX")
            # model_path
            if CPU_config[___join_str(CPU_name,Platform,CPU_type,link='_')+'_model_path'] !='':
                if not os.path.exists(path_dict['model_path']):
                    do=False
                    save_to_log(logPathName,"XXXXX MISSING_FILE : model_path XXXXX")
            # cap_lib
            if CPU_config['cap_lib'] !='':
                if not os.path.exists(path_dict['cap_lib']):
                    do=False
                    save_to_log(logPathName,f"XXXXX MISSING_FILE : {CPU_config['cap_lib']} XXXXX")
            save_to_log(logPathName,"read config.json ok")
        # Save BRD
        try:
            brd_file.save(path_dict['brd'])
            save_to_log(logPathName,'Saved BRD FILED!')
        except Exception as _:
            do=False
            save_to_log(logPathName,'XXXXX Missing BRD FILED! XXXXX')

        # Save stackup.dkdf
        try :
            request.files['dkdf'].save(path_dict['dkdf'])
            save_to_log(logPathName,'Saved stackup.dkdf FILED!')
        except Exception as _:
            do=False
            save_to_log(logPathName,'XXXXX Missing stackup.dkdf FILED! XXXXX')

        # Save PDN Checklist
        try :
            request.files['checklist'].save(path_dict['checklist'])
            save_to_log(logPathName,'Saved PDN Checklist!')
        except Exception as _:
            do=False
            save_to_log(logPathName,'XXXXX Missing PDN Checklist! XXXXX')
        
        try:
            insert_data = {
                'product': request_data['product'],
                'other_file':request.files['other_file'] if 'other_file' in request.files else '',
                'schematic': request.files['schematic'].filename if 'schematic' in request.files else '',
                'createTime': ct,
                'report_result':'',
                'stackup_no':request_data['stackup_no'],
                'initialResult': '',
                'filePath': path_dict,
                'reason': request_data['reason'],
                'current_opi_start_dt':ct,
                'fail_times':0
                }
            if 'DDR CCT' in request_data['busItem']:
                insert_data['Mapping'] = request_data['Mapping']
                insert_data['PCBType'] = request_data['PCBType']
                insert_data['DDRModule'] = request_data['DDRModule']
                insert_data['RamType'] = request_data['RamType']
                insert_data['Rank'] = request_data['Rank']
                insert_data['DataRate'] = request_data['DataRate']

            save_to_log(logPathName,"Insert data create")
        except Exception as e:
            do=False
            save_to_log(logPathName,f"Print insert data :\n{insert_data}")
            save_to_log(logPathName,"set data created error %s"%(str(e)))
        #! 如果 do 為True 表示資料正常更新進DB，反之，Lib/ date create 有問題，Initial fail(LIBRARY_ERROR)
        if do:
            collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
            collect_tc.update_one(
                                {'form_id': request_data['form_id']},
                                {'$set': insert_data})
            save_to_log(logPathName,"Updated Board To DB")
            # SaveToLog(logPathName,"Start to Check Board in Task")
            # # Check first send 的"版號" 是否 In Schedules/Running 或 Over 3 個 Finished
            # if ___CheckBoardinTask(board_number,board_stage,request_data['busItem'],logPathName)=='In Schedules/Running':
            #     # TODO send mail : In Schedules/Running
            #     SaveToLog(logPathName,"Send 'In Schedules/Running'")
            #     notify_type = 'In Schedules/Running'
            #     mailNotify.opiMailNotify (notify_type, request_data,config_dict)
            #     return JSONEncoder().encode("The Task Has Been In Schedules/Running"), 200
            # elif ___CheckBoardinTask(board_number,board_stage,request_data['busItem'],logPathName)=='Over 3 Finished':
            #     # TODO send mail : Over 3 Finished
            #     SaveToLog(logPathName,"Send 'Over 3 Finished'")
            #     notify_type = 'Over 3 Finished'
            #     mailNotify.opiMailNotify (notify_type, request_data,config_dict)
            #     return JSONEncoder().encode("The Task is Over 3 Finished Task"), 200
            return JSONEncoder().encode("OK"), 200
        else:
            insert_data = {
                'status':'Initial_Fail',
                'initialResult':'LIBRARY_ERROR'
            }
            collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
            collect_tc.update_one(
                                {'form_id': request_data['form_id']},
                                {'$set': insert_data})
            save_to_log(logPathName,"Updated Board To DB")
            config_dict=utils.__read_config_file('config.json')
            mailNotify.lib_error (request_data,logPathName,config_dict)
            save_to_log(logPathName,"Send Library error mail!")
            return JSONEncoder().encode("STORED ERROR! Call Admin please!"), 400
    except Exception as e :
        save_to_log(logPathName,f"ERROR : {e} \n {traceback.format_exc()}")
        config_dict=utils.__read_config_file('config.json')
        mailNotify.error_mail(request_data,logPathName,config_dict)
        return JSONEncoder().encode("ERROR! Call Admin please!"), 400

# #! 目前改版不會用到    
# @app.route('/opi/getscheduledlist',methods=["GET"])
# def getOPISchedule_old():
#     collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#     docs = list(collect_tc.find({'order': {'$gt': 0}, 'status':{'$ne': 'Cancel'}}).sort("order"))
#     uf_count = 0
#     DDR_count = 0
#     for idx, row in enumerate(docs) :
#         if row['order']==1:
#             uf_count = 0
#             DDR_count = 0
#         if 'tcl' in row :
#             if row['status'] == 'Running' :
#                 docs[idx]['estStart'] = int(row['current_opi_start_dt'])
#                 if row['busItem'] == 'PDN':
#                     for tcl in row['tcl'] :
#                         if tcl['status'] == 'unfinished' :
#                             uf_count += 1
#                         docs[idx]['estFinish'] =  int(row['current_opi_start_dt']) +9000*uf_count
#                 else:
#                     for tcl in row['tcl'] :
#                         if tcl['status'] == 'unfinished' :
#                             DDR_count += 1
#                     docs[idx]['estFinish'] = int(row['current_opi_start_dt']) +172800*DDR_count
#             else :
#                 if row['busItem'] == 'PDN':
#                      docs[idx]['estStart'] = int(row['current_opi_start_dt']) +9000*uf_count
#                     for tcl in row['tcl'] :
#                         if tcl['status'] == 'unfinished' :
#                             uf_count += 1
#                     docs[idx]['estFinish'] =  int(row['current_opi_start_dt']) +9000*uf_count
#                 else:
#                     docs[idx]['estStart'] = int(row['current_opi_start_dt']) +172800*DDR_count
#                     for tcl in row['tcl'] :
#                         if tcl['status'] == 'unfinished' :
#                             DDR_count += 1
#                     docs[idx]['estFinish'] = int(row['current_opi_start_dt']) +172800*DDR_count
#     return JSONEncoder().encode(docs), 200

@app.route('/opi/getscheduledlist/v2',methods=["GET"])
def getOPISchedule():
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    docs = list(collect_tc.find({'order': {'$gt': 0}}).sort("order"))
    for idx, row in enumerate(docs) :
        if row['status'] =='Running':
            docs[idx]['estStart'] = int(row['current_opi_start_dt'])
        else:
            docs[idx]['estStart'] = int(datetime.datetime.strptime(row['sim_start_date'],'%Y-%m-%d').timestamp())

        docs[idx]['estFinish'] =  docs[idx]['estStart'] +int(row['run_sim_days']*24*60*60)


    return JSONEncoder().encode(docs), 200

@app.route('/opi/getFaillist',methods=["GET"])
def getOPIFailList():
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    docs = list(collect_tc.find({'order': 0, 'status':{'$nin': ['Cancel']}}).sort([("_id", -1)]))
    for idx, row in enumerate(docs) :
        docs[idx]['estStart'] = '--'
        docs[idx]['estFinish'] = '--'
        docs[idx]['tcl'] = [{"net" : "----", "filepath" : ""}]
    return JSONEncoder().encode(docs), 200

@app.route('/opi/getFinishedlist',methods=["GET"])
def getFinishedlist():
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    docs = list(collect_tc.find({'status':'Finished'}).sort([("_id", -1)]))
    for idx, row in enumerate(docs) :
        docs[idx]['estStart'] = row['current_opi_start_dt']
        docs[idx]['estFinish'] = row['finished_dt']
    return JSONEncoder().encode(docs), 200

@app.route('/opi/getConflict',methods=["GET"])
def getConflict():
    collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
    docs = list(collect_tc.find({'order': 0, 'status':'Conflict'}).sort([("_id", -1)]))
    for idx, row in enumerate(docs) :
        docs[idx]['estStart'] = '--'
        docs[idx]['estFinish'] = '--'
    return JSONEncoder().encode(docs), 200

# #! 目前改版不會用到
# @app.route('/opi/setorder',methods=["POST"])
# def setOPIOrder():
#     request_json = request.get_json(force=True)
#     if 'orderList' not in request_json :
#         return JSONEncoder().encode("MISSING_FILED"), 400
#     # require_fileds = ['applicant', 'boardNumber', 'boardStage', 'createTime', 'order']
#     require_fileds = ['_id', 'order']
#     for row in request_json['orderList'] :
#         for rf in require_fileds :
#             if rf not in row :
#                 return JSONEncoder().encode("MISSING_FILED"), 400

#     collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#     try :
#         for row in request_json['orderList'] :
#             dict_onesheet = collect_tc.find_one({'_id':bson.ObjectId(row['_id'])})
#             if dict_onesheet['order'] != row['order']:
#                 collect_tc.update_one({'_id': bson.ObjectId(row['_id'])},
#                     {'$set':{'order': row['order']}}
#                 )
#                 # TODO send mail : simulation_change
#                 # notify_type = 'simulation_change'
#                 # mailNotify.opiMailNotify (notify_type, dict_onesheet,config_dict)
#         return JSONEncoder().encode("OK"), 200
#     except :
#         return JSONEncoder().encode("INSERT_FAIL"), 400

@app.route('/opi/cancel',methods=["POST"])
@csrf_protect
def cancel():
    request_json = request.get_json(force=True)
    require_fileds = ['form_id']
    for rf in require_fileds :
        if rf not in request_json :
            return JSONEncoder().encode(error_sign["MISSING_FIELD"]), 400
    try : 
        collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
        item = collect_tc.find_one({'form_id':request_json['form_id']})
        config_dict=utils.__read_config_file('config.json')
        if not item:
            return JSONEncoder().encode(error_sign["MISSING_Form_id"]), 200
        res=collect_tc.update_one(
                    {'form_id':item['form_id']},
                    {'$set':{'status': 'Cancel', 'order': 0}})
        if not res.acknowledged:
            app.logger.error(f"Acknowledged: {res.acknowledged}")
            app.logger.error(f"Matched count:{res.matched_count}")
            app.logger.error(f"Modified count:{res.modified_count}")
            app.logger.error(f"Upserted id:{res.upserted_id}")
            app.logger.error(f"Raw result:{res.raw_result}")
            return JSONEncoder().encode(error_sign["Cannot Store Data to DB"]), 400
        notify_type = 'Cancel'
        mailNotify.opiMailNotify (notify_type, item,config_dict)
        list_unspec = list(collect_tc.find({'license':'Unspecified','busItem':{'$regex':'.*'+busitem+'$'},'order':{'$gt':0},'debug':False,'status':{'$in':['Scheduled','Initial_Fail']}}).sort([("order", 1)]))
        # debug form 的狀態為 order 0, 僅能靠 'license':'Unspecified',debug':True,'status':{'$in':['Scheduled','Initial_Fail']} 查詢
        list_debug = list(collect_tc.find({'license':'Unspecified','busItem':{'$regex':'.*'+busitem+'$'},'debug':True,'status':{'$in':['Scheduled','Initial_Fail']}})) 
        list_run = list(collect_tc.find({'busItem':{'$regex':'.*'+busitem+'$'},'status':'Running'}))
        # 先處理 debug task
        order = len(list_run)
        for idx in range(len(list_debug)):
            item = list_debug[idx]
            order +=1
            item['order']=order
            res=collect_tc.update_one({'form_id':item['form_id']},
                                {'$set':item})
            if not res.acknowledged:
                app.logger.error(f"Acknowledged: {res.acknowledged}")
                app.logger.error(f"Matched count:{res.matched_count}")
                app.logger.error(f"Modified count:{res.modified_count}")
                app.logger.error(f"Upserted id:{res.upserted_id}")
                app.logger.error(f"Raw result:{res.raw_result}")
                return JSONEncoder().encode(error_sign["Cannot Store Data to DB"]), 400
        for idx in range(len(list_unspec)):
            item = list_unspec[idx]
            order +=1
            item['order']=order
            collect_tc.update_one({'form_id':item['form_id']},
                                {'$set':item})
            res=collect_tc.update_one({'form_id':item['form_id']},
                                {'$set':item})
            if not res.acknowledged:
                app.logger.error(f"Acknowledged: {res.acknowledged}")
                app.logger.error(f"Matched count:{res.matched_count}")
                app.logger.error(f"Modified count:{res.modified_count}")
                app.logger.error(f"Upserted id:{res.upserted_id}")
                app.logger.error(f"Raw result:{res.raw_result}")
                return JSONEncoder().encode(error_sign["Cannot Store Data to DB"]), 400
        return JSONEncoder().encode("OK"), 200
    except Exception as e:
        app.logger.error(f"Call Cancel API ERROR : {e}")
        return JSONEncoder().encode(error_sign["INSERT_FAIL"]), 400

# #! 目前改版不會用到
# @app.route('/opi/restart',methods=["POST"])
# def restartOPI():
#     curTime=int(time.time()*1000)
#     logPathName,logPath=CreateLogFolder(curTime)
#     SaveToLog(logPathName,"/opi/restart")
#     try:
#         request_json = request.get_json(force=True)
#         # require_fileds = ['applicant', 'boardNumber', 'boardStage', 'createTime']
#         require_fileds = ['_id']
#         for rf in require_fileds :
#             if rf not in request_json :
#                 return JSONEncoder().encode("MISSING_FILED"), 400
        
#         collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#         # data = collect_tc.find_one({'applicant': request_json['applicant'], 'boardNumber': request_json['boardNumber'], 'boardStage': request_json['boardStage'], 'createTime': request_json['createTime']})
#         data = collect_tc.find_one({'_id': bson.ObjectId(request_json['_id'])})
#         # ! ------------------------------------------------------------------------------------------- !#
#         # 確認是否為 DDR_CCT 的 restart，若"是"則送request to Sim3/4
#         if data.get('license'):
#             url = ''
#             if data['license'] == 'SPEED2000 #1':
#                 url = 'http://10.34.24.128:66//DDR//ddr_restart'
#                 SaveToLog(logPathName, f"Call Sim3 API:{url}")
#             elif data['license'] == 'SPEED2000 #2':
#                 url = 'http://10.34.24.129:66//DDR//ddr_restart'
#                 SaveToLog(logPathName, f"Call Sim4 API:{url}")

#             if url:
#                 try:
#                     response = requests.post(url, json=request_json)
#                     SaveToLog(logPathName, f"Print DDR restart result :\n {response.json()}")
#                     if response.status_code == 200:
#                         return jsonify({response.json()}), 200
#                     elif response.status_code == 400:
#                         return jsonify({response.json()}), 400
#                 except Exception as e:
#                     SaveToLog(logPathName, f"Error occurred during calling DDR CCT restart: {str(e)}")
#                     return jsonify("Error occurred during calling DDR CCT restart"),400
#             else:
#                 SaveToLog(logPathName, "XXXXX No url for DDR CCT restart XXXXX")
#                 return jsonify("Missing url for DDR CCT restart"), 400
#         else:
#             SaveToLog(logPathName, "No license key........Turn To Do OPI restart.......")
#         # ============================================================================================= #
#         license_info_dict=__find_license_info()
#         license_key=list(license_info_dict.keys())
#         if license_info_dict==False :
#             SaveToLog(logPathName,"read license info failed.")
#             exit()
#         SaveToLog(logPathName,"Read license info succeeded.")
#         # 讀取 CPU Platform
#         Vendor,CPU_name,Platform,CPU_type,CPU_Target=utils.read_CPU_info(data['platform'],logPathName)
#         # 讀取 Platform 的 CPU_conig
#         CPU_config_dir=os.path.join(os.getcwd(), 'CPU_config',Vendor+'_'+CPU_name+'.json')
#         with open(CPU_config_dir) as f:
#             CPU_config = json.load(f)
#         if len(CPU_config)==0: return JSONEncoder().encode("MISSING_CPU_CONFIG"), 400
#         if data :
#             try:
#                 path_dict = data['filePath']
#                 board_number = data['boardNumber']
#                 board_stage = data['boardStage']
#                 ct = data['createTime']

#                 update_data = {
#                     'form_id': data['form_id'],
#                     'applicant': data['applicant'], 
#                     'busItem' : data['busItem'],
#                     'boardNumber': board_number,
#                     'boardStage': board_stage,
#                     'customer' : data['customer'],
#                     'project_code' : data['project_code'],
#                     'project_name' : data['project_name'],
#                     'createTime': ct,
#                     'initialTime': ct,
#                     'report_result':'',
#                     'stackup_no':data.get('stackup_no', ''),
#                     'platform': data['platform'],
#                     'projectSchedule': {
#                         'startDate': data['projectSchedule']['startDate'],
#                         'targetDate': data['projectSchedule']['targetDate'],
#                         'gerberDate': data['projectSchedule']['gerberDate'],
#                         'smtDate':data['projectSchedule']['smtDate']
#                     },
#                     'order': 0,
#                     'status': 'Initial_Fail',
#                     'initialResult': 'ERROR',
#                     'filePath': path_dict,
#                     'reason': data['reason'],
#                 }
#             except Exception as e:
#                 SaveToLog(logPathName,"set update data error %s"%(str(e)))
#             # ! 加入選擇license
#             license_name_list=list(license_info_dict.keys())
#             check_license=[]
#             for license_name in license_name_list:
#                 check_license.append(__check_license(license_name,logPathName))
#             # ! 先挑第一個license
#             first_license_name = next(iter(license_info_dict))
#             use_license=license_info_dict[first_license_name] #'-PSAdvancedPI_TI_20'
#             # license=['-PSOptimizePI_20','-PSOptimizePI_20','-PSAdvancedPI_TI_20']
#             # check_license=[__check_license('OPI1'),__check_license('OPI2'),__check_license('OPI3')]
#             true_indexes = [index for index, value in enumerate(check_license) if value]
#             try :
#                 bLicense=False
#                 if len(true_indexes):
#                     use_license=license_info_dict[license_name_list[true_indexes[0]]]
#                     bLicense=True
#                     SaveToLog(logPathName," Use license %s" %(license_name_list[true_indexes[0]]))
#                 else: # ! 目前沒有license
#                     SaveToLog(logPathName," **** no available_license for initial ****")
#                     # TODO send mail : NO_LICENSE
#                 if bLicense==False:
#                     notify_type = 'NO_LICENSE'
#                     mailNotify.opiMailNotify (notify_type, update_data,config_dict)
#                     raise NameError('NO_LICENSE')
#                 # 產生 print_list.tcl
#                 try :
#                     SaveToLog(logPathName," running gentcl.stackup()" )
#                     gentcl.stackup(path_dict)
#                     # subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe', '-PSAdvancedPI_TI_20' ,'-tcl', path_dict['print_list']])
#                     # subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe', use_license ,'-tcl', path_dict['print_list']])
#                     # ! 加入選擇license
#                     # license=['-OptimizePI_20','-OptimizePI_20','-PSAdvancedPI_TI_20']
#                     # check_license=[__check_license('OPI1'),__check_license('OPI2'),__check_license('OPI3')]
#                     # true_indexes = [index for index, value in enumerate(check_license) if value]
#                     if bLicense:
#                         SaveToLog(logPathName,"available_license for initial: "+use_license)
#                         SaveToLog(logPathName,"print_list: "+path_dict['print_list'])
#                         # subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe', use_license ,'-tcl', path_dict['print_list']])
#                         result=subprocess.run([r'C:\Cadence\Sigrity2023.1\tools\bin\OptimizePI.exe',use_license ,'-tcl', path_dict['print_list']], capture_output=True)
#                         SaveToLog(logPathName,"subprocess result: "+str(result.returncode))

#                     else:
#                         SaveToLog(logPathName," **** no available_license for initial ****")
#                         # TODO send mail : NO_LICENSE
#                         notify_type = 'NO_LICENSE'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('NO_LICENSE')
#                 except Exception as e:
#                     # TODO send mail : PARSE_ERROR
#                     notify_type = 'PARSE_ERROR'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     raise NameError('PARSE_ERROR')
#                 ### get power information (TODO: Intel)
#                 try :
#                     power_info = utils.getPowerInfo(board_number, board_stage)
#                     df_power = pd.DataFrame(power_info)
#                     if len(df_power)==0:
#                         SaveToLog(logPathName,"POWER API NO DATA")
#                     else:
#                         SaveToLog(logPathName,"Get POWER API DATA")
#                     if Vendor=='Intel' and ('OutputNet2' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns or 'OutputNet1' not in df_power.columns):
#                         raise NameError('POWER_API_NO_DATA')
#                 except Exception as e:
#                     # TODO send mail : POWER_API_ERROR
#                     notify_type = 'POWER_API_ERROR'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     raise NameError('POWER_API_ERROR')
#                 ###parse dk df data (TODO: Intel/ Qualcomm)
#                 try :
#                     data_DKDF = utils.parseStackup(path_dict['dkdf'])
#                     SaveToLog(logPathName,"parseStackup Done")
#                 except Exception as e:
#                         #TODO send mail : PARSE_ERROR_DKDF
#                         notify_type = 'PARSE_ERROR_DKDF'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('PARSE_ERROR_DKDF')
#                 if data_DKDF :
#                     ### gen material cmx (TODO: Intel/ Qualcomm)
#                     try :
#                         tree = utils.modifyMaterial(data_DKDF['layer'], path_dict['material_org'], data_DKDF['data'])
#                         tree.write(path_dict['material'])
#                         SaveToLog(logPathName,"modifyMaterial Done")
#                     except :
#                         #TODO send mail : GENCMX_ERROR
#                         notify_type = 'GENCMX_ERROR'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('GENCMX_ERROR')
#                     ### gen stackup csv (TODO: Intel/ Qualcomm)
#                     try :
#                         modified_stackup = utils.modifyStackup(data_DKDF['layer'], path_dict['stackup_org'])
#                         modified_stackup.to_csv(path_dict['stackup'], index=False)
#                         SaveToLog(logPathName,"modifyStackup Done")
#                     except :
#                         #TODO send mail : GENSTACKUP_ERROR
#                         notify_type = 'GENSTACKUP_ERROR'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('GENSTACKUP_ERROR')
#                         ### Parse VRM (TODO: Intel)
#                     try :
#                         if CPU_config['vrm'] != '':
#                             df_vrm = utils.parseVRM(path_dict['vrm'])
#                             SaveToLog(logPathName,"parseVRM Done")
#                         else:
#                             df_vrm = pd.DataFrame()
#                             SaveToLog(logPathName,"VRM NO DATA")
#                     except:
#                         #TODO send mail : VRM_ERROR
#                         notify_type = 'VRM_ERROR'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('VRM_ERROR')
#                     ### gen final tcl
#                     try :
#                         SaveToLog(logPathName,"Gentcl Start")
#                         tcl_list = gentcl.OPI(path_dict, df_power,df_vrm,Vendor,CPU_name,Platform,CPU_type,CPU_Target)
#                         SaveToLog(logPathName,"Gentcl End")
#                     except :
#                         #TODO send mail : GEN TCL_ERROR
#                         notify_type = 'GEN TCL_ERROR'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('GEN TCL_ERROR')
#                 else :
#                     #TODO send mail : DKDF_FORMAT_ERROR
#                     notify_type = 'DKDF_FORMAT_ERROR'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     raise NameError('DKDF_FORMAT_ERROR')
                
#                 unfinished_task = pd.DataFrame(collect_tc.find({'status':{'$in':['Running','Scheduled']}}))
#                 update_data['tcl'] = tcl_list
#                 if len(unfinished_task) !=0:
#                     ser_opi={}
#                     for opi in license_key:
#                         ser_opi[opi] = len(unfinished_task[unfinished_task['license']==opi].reset_index(drop=True))
#                     # * 三個license都有任務的時候，比誰的任務少就assign給他
#                     min_opi = min(ser_opi, key=lambda k: ser_opi[k])
#                     update_data['license'] = min_opi
#                     SaveToLog(logPathName,"assign schedule for %s"%min_opi)
#                     # 設定 order 數
#                     ser=unfinished_task[unfinished_task['license']==min_opi].reset_index(drop=True)
#                     if  ser_opi[min_opi] !=0:
#                         update_data['order'] = int(ser['order'].max() + 1)
#                         update_data['current_opi_start_dt'] = int(ser.loc[0, 'current_opi_start_dt'])
#                     else:
#                         update_data['order']=1
#                         update_data['current_opi_start_dt'] = int(time.time())
#                     update_data['initialResult'] = 'Succeed'
#                     update_data['status'] = 'Scheduled'
#                     update_data['tcl'] = tcl_list
#                 else:
#                     update_data['license'] = license_key[0]
#                     SaveToLog(logPathName,"assign schedule for %s"%license_key[0])
#                     update_data['order']=1
#                     update_data['current_opi_start_dt'] = int(time.time()) #BUG: need to check # int(unfinshed_task.loc[0, 'current_opi_start_dt'])
#                     update_data['initialResult'] = 'Succeed'
#                     update_data['status'] = 'Scheduled'
#                     update_data['tcl'] = tcl_list
#                 try :
                    
#                     collect_tc.update_one(
#                                     {'_id': bson.ObjectId(request_json['_id'])},
#                                     # {'applicant': request_json['applicant'], 'boardNumber': request_json['boardNumber'], 'boardStage': request_json['boardStage'], 'createTime': request_json['createTime']},
#                                     {'$set': update_data})
#                     # TODO send mail : initial_success
#                     notify_type = 'initial_success'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     return JSONEncoder().encode("OK"), 200
#                 except:
#                     return JSONEncoder().encode("UPDATE_FAIL_2"), 400
#             except Exception as e :
#                 update_data['initialResult'] = str(e)
#                 try :
#                     collect_tc.update_one(
#                         {'_id': bson.ObjectId(request_json['_id'])},
#                         # {'applicant': request_json['applicant'], 'boardNumber': request_json['boardNumber'], 'boardStage': request_json['boardStage'], 'createTime': request_json['createTime']},
#                         {'$set': update_data}
#                     )
#                     return JSONEncoder().encode("INITIAL_FAIL_in restart"), 200
#                 except :
#                     return JSONEncoder().encode("UPDATE_FAIL_3"), 400
#         else :
#             return JSONEncoder().encode("DATA_NOT_FOUND"), 400
#     except Exception as e :
#         SaveToLog(logPathName,f"ERROR : {e} \n {traceback.format_exc()}")
#         return JSONEncoder().encode("Initial ERROR"), 400

# #! 目前改版不會用到
# @app.route('/DDR/ddr_restart',methods=["POST"])
# def ddr_restart():
#     curTime=int(time.time()*1000)
#     logPathName,logPath=CreateLogFolder(curTime)
#     SaveToLog(logPathName,"Start ddr_restart")
#     # 是否上傳 DB Ture/False
#     ToMongo=True
#     SaveToLog(logPathName,f"If Updated to DB : {ToMongo}")
#     try:
#         request_json = request.get_json(force=True)
#         # require_fileds = ['applicant', 'boardNumber', 'boardStage', 'createTime']
#         require_fileds = ['_id']
#         for rf in require_fileds :
#             if rf not in request_json :
#                 SaveToLog(logPathName,f"Missing parameter:{rf}")
#                 return JSONEncoder().encode("MISSING_FILED"), 400
        
#         collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName="OPITaskCtrl")
#         data = collect_tc.find_one({'_id': bson.ObjectId(request_json['_id'])})
#         # ! ------------------------------------
        
#         license_info_dict=__find_license_info()
#         license_key=list(license_info_dict.keys())
#         if license_info_dict==False :
#             SaveToLog(logPathName,"read license info failed.")
#             exit()
#         SaveToLog(logPathName,"Read license info succeeded.")
#         # 讀取 CPU Platform
#         Vendor,CPU_name,Platform,CPU_type,CPU_Target=utils.read_CPU_info(data['platform'])
#         # 讀取 Platform 的 CPU_conig
#         CPU_config_dir=os.path.join(os.getcwd(), 'CPU_config',Vendor+'_'+CPU_name+'.json')
#         with open(CPU_config_dir) as f:
#             CPU_config = json.load(f)
#         if len(CPU_config)==0: return JSONEncoder().encode("MISSING_CPU_CONFIG"), 400
#         # ======================== end ========================== #
#         if data :
#             try:
#                 server_name=data['license']
#                 path_dict = data['filePath']
#                 board_number = data['boardNumber']
#                 board_stage = data['boardStage']
#                 ct = data['createTime']
#                 update_data = {
#                     'form_id': data['form_id'],
#                     'applicant': data['applicant'], 
#                     'busItem' : data['busItem'],
#                     'boardNumber': board_number,
#                     'boardStage': board_stage,
#                     'customer' : data['customer'],
#                     'project_code' : data['project_code'],
#                     'project_name' : data['project_name'],
#                     'createTime': ct,
#                     'initialTime': ct,
#                     'report_result':'',
#                     'stackup_no':data.get('stackup_no', ''),
#                     'platform': data['platform'],
#                     'projectSchedule': {
#                         'startDate': data['projectSchedule']['startDate'],
#                         'targetDate': data['projectSchedule']['targetDate'],
#                         'gerberDate': data['projectSchedule']['gerberDate'],
#                         'smtDate':data['projectSchedule']['smtDate']
#                     },
#                     'order': 0,
#                     'status': 'Initial_Fail',
#                     'initialResult': 'ERROR',
#                     'filePath': path_dict,
#                     'reason': data['reason'],
#                     'reason': data['reason'],
#                     'Mapping':data['Mapping'],
#                     'license': server_name,
#                     'PCBType':data['PCBType'],
#                     'DDRModule':data['DDRModule'],
#                     'RamType':data['RamType'],
#                     'Rank':data['Rank'],
#                     'DataRate':data['DataRate']
#                 }
#                 SaveToLog(logPathName,f"assign schedule for {data['license']}")
#                 SaveToLog(logPathName,f"set update data : \t {update_data}")
#             except Exception as e:
#                 SaveToLog(logPathName,"XXXXX set update data error %s XXXXX"%(str(e)))
#             # ! 加入選擇license
#             try:
#                 license_name_list=list(license_info_dict.keys())
#                 check_license=[]
#                 for license_name in license_name_list:
#                     check_license.append(__check_license(license_name,logPathName))
#                 # ! 先挑第一個license
#                 first_license_name = next(iter(license_info_dict))
#                 use_license=license_info_dict[first_license_name] 
#                 true_indexes = [index for index, value in enumerate(check_license) if value]
#             except Exception as e:
#                 SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#             try :
#                 bLicense=False
#                 if len(true_indexes):
#                     use_license=license_info_dict[license_name_list[true_indexes[0]]]
#                     bLicense=True
#                     SaveToLog(logPathName,f"Use license {license_name_list[true_indexes[0]]}")
#                 else: # ! 目前沒有license
#                     SaveToLog(logPathName," **** no available_license for initial ****")          
#                 if bLicense==False:
#                     notify_type = 'NO_LICENSE'
#                     mailNotify.opiMailNotify (notify_type,update_data)
#                     raise NameError('NO_LICENSE')
#                 try :
#                     SaveToLog(logPathName," running gentcl.stackup()" )
#                     DDRGentcl.stackup(path_dict)
#                     if bLicense:
#                         SaveToLog(logPathName,"available_license for initial: "+use_license)
#                         SaveToLog(logPathName,"print_list: "+path_dict['print_list'])
#                         result=subprocess.run([r'C:\CADENCE\Sigrity2023.1\tools\bin\OptimizePI.exe',use_license ,'-tcl', path_dict['print_list']], capture_output=True)
#                         SaveToLog(logPathName,"subprocess result: "+str(result.returncode))
#                     else:
#                         SaveToLog(logPathName," **** no available_license for initial ****")
#                         # TODO send mail : NO_LICENSE
#                         notify_type = 'NO_LICENSE'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         raise NameError('NO_LICENSE')
#                 except :
#                     # TODO send mail : PARSE_ERROR
#                     notify_type = 'PARSE_ERROR'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     raise NameError('PARSE_ERROR')
#                 ### parse dk df data (TODO: Intel/ Qualcomm)
#                 try :
#                     data_DKDF = utils.parseStackup(path_dict['dkdf'])
#                     SaveToLog(logPathName,"parseStackup Done")
#                 except Exception as e:
#                         #TODO send mail : PARSE_ERROR_DKDF
#                         notify_type = 'PARSE_ERROR_DKDF'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         SaveToLog(logPathName,"PARSE_ERROR_DKDF MAIL SEND" )
#                         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#                         raise NameError('PARSE_ERROR_DKDF')
#                 if data_DKDF :
#                     ### gen material cmx (TODO: Intel/ Qualcomm)
#                     try :
#                         tree = utils.modifyMaterial(data_DKDF['layer'], path_dict['material_org'], data_DKDF['data'])
#                         tree.write(path_dict['material'])
#                         SaveToLog(logPathName,"modifyMaterial Done")
#                     except Exception as e:
#                         #TODO send mail : GENCMX_ERROR
#                         notify_type = 'GENCMX_ERROR'
#                         mailNotify.opiMailNotify (notify_type,update_data)
#                         SaveToLog(logPathName,"GENCMX_ERROR MAIL SEND" )
#                         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#                         raise NameError('GENCMX_ERROR')
#                     ### gen stackup csv (TODO: Intel/ Qualcomm)
#                     try :
#                         modified_stackup = utils.modifyStackup(data_DKDF['layer'], path_dict['stackup_org'])
#                         modified_stackup.to_csv(path_dict['stackup'], index=False)
#                         SaveToLog(logPathName,"modifyStackup Done")
#                     except Exception as e:
#                         #TODO send mail : GENSTACKUP_ERROR
#                         notify_type = 'GENSTACKUP_ERROR'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         SaveToLog(logPathName,"GENSTACKUP_ERROR MAIL SEND" )
#                         SaveToLog(logPathName,f"Exception : {e} \nError : \n{traceback.format_exc()}")
#                         raise NameError('GENSTACKUP_ERROR')
#                     ### gen final tcl
#                     try :
#                         SaveToLog(logPathName,"Gentcl Start")
#                         tcl_list = DDRGentcl._DDRGentcl(path_dict,board_number,board_stage,Vendor,update_data['DDRModule'],update_data['Mapping'])
#                         SaveToLog(logPathName,"Gentcl End")
#                     except Exception as e:
#                         #TODO send mail : GEN TCL_ERROR
#                         notify_type = 'GEN TCL_ERROR'
#                         mailNotify.opiMailNotify (notify_type, update_data)
#                         SaveToLog(logPathName,"GEN TCL_ERROR MAIL SEND" )
#                         raise NameError('GEN TCL_ERROR')
#                 else :
#                     #TODO send mail : DKDF_FORMAT_ERROR
#                     notify_type = 'DKDF_FORMAT_ERROR'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     SaveToLog(logPathName,"DKDF_FORMAT_ERROR MAIL SEND" )
#                     raise NameError('DKDF_FORMAT_ERROR')
#                 unfinshed_task = pd.DataFrame(collect_tc.find({'$and': [{'status': {'$in': ['Running', 'Scheduled']}}, {'license': server_name}]}))
#                 update_data['tcl'] = tcl_list
#                 if len(unfinshed_task) != 0:
#                     # 設定 order 數
#                     update_data['order'] = int(unfinshed_task['order'].max() + 1)
#                     update_data['current_opi_start_dt'] = int(unfinshed_task.loc[0, 'current_opi_start_dt'])
#                     update_data['initialResult'] = 'Succeed'
#                     update_data['status'] = 'Scheduled'
#                     update_data['tcl'] = tcl_list
#                 else:
#                     update_data['order']=1
#                     update_data['current_opi_start_dt'] = int(time.time()) #BUG: need to check # int(unfinshed_task.loc[0, 'current_opi_start_dt'])
#                     update_data['initialResult'] = 'Succeed'
#                     update_data['status'] = 'Scheduled'
#                     update_data['tcl'] = tcl_list
#                 try :
#                     if ToMongo:
#                         collect_tc.update_one(
#                                         {'_id': bson.ObjectId(request_json['_id'])},
#                                         {'$set': update_data})
#                     # TODO send mail : initial_success
#                     notify_type = 'initial_success'
#                     mailNotify.opiMailNotify (notify_type, update_data)
#                     return JSONEncoder().encode("Initial_success"), 200
#                 except:
#                     return JSONEncoder().encode("UPDATE_FAIL_2"), 400
#             except Exception as e :
#                 update_data['initialResult'] = str(e)
#                 if ToMongo:
#                     try :
#                         collect_tc.update_one(
#                             {'_id': bson.ObjectId(request_json['_id'])},
#                             {'$set': update_data}
#                         )
#                         return JSONEncoder().encode("INITIAL_FAIL_in restart"), 200
#                     except :
#                         return JSONEncoder().encode("UPDATE_FAIL_3"), 400
#         else :
#             return JSONEncoder().encode("DATA_NOT_FOUND"), 400
#     except Exception as e :
#         SaveToLog(logPathName,f"ERROR : {e} \n {traceback.format_exc()}")
#         return JSONEncoder().encode("Initial ERROR"), 400

# #! 目前改版不會用到
# def __checkTaskStatus(form_id):
#     # ! 狀態為 Scheduled 才可以進行重送的檢查 (只能重送一次)
#     collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#     task=collect_tc.find_one({'form_id':form_id})
#     if task is not None:
#         if task.get('status') in ['Running']: # ! 執行中不可變更
#             app.logger.debug("status: %s (不可進行重送).", task.get('status'))
#             return False
#         if task.get('status')=='Scheduled':
#             if 'request_count' in task.keys():
#                 if task.get('request_count')>=2:    return False
#                 else:
#                     result = collect_tc.update_one({'form_id':form_id},  {'$set': {'request_count': task.get('request_count')+1}})
#                     # 檢查更新操作是否成功
#                     if result.modified_count > 0:
#                         app.logger.debug("set request count: %d successfully updated", task.get('request_count')+1)
#                         return True
#                     else:
#                         app.logger.debug("set request count: %d fail.", task.get('request_count')+1)
#                         return False
#             else:
#                 result = collect_tc.update_one({'form_id':form_id},  {'$set': {'request_count': 1}})
#                 # 檢查更新操作是否成功
#                 if result.modified_count > 0:
#                     app.logger.debug("set request count: 1 successfully updated")
#                     return True
#                 else:
#                     app.logger.debug("set request count: 1 fail")
#                     return False

# ------------------SaveToLog--------------------------------
# Input: 1. log path name 
#        2. text that need to put in file
# Output: none
# -----------------------------------------------------------
def save_to_log(logPathName,Msg):
    currTime=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    logFile = open(logPathName, 'a')
    WriteMsg="[ %s ]  %s"%(str(currTime),str(Msg))
    logFile.write(WriteMsg+'\n')
    logFile.close()  
# ------------------CreateLogFolder------------------------
# Input: current time
# Output: log file path and name ex. 2019-02-21_12-00-00.log
# ---------------------------------------------------------
def create_log_folder(curTime): # "D:\\MEUpdateLog"
    # check log dir 
    currDir=os.getcwd()
    today = datetime.date.today()
    logPath=currDir+ "\\" +str(today)
    if not os.path.isdir(logPath):
        os.mkdir(logPath)
    logPath+="\\Main_Log"
    if not os.path.isdir(logPath):
        os.mkdir(logPath)
    # create log file
    currentDateTime=datetime.datetime.fromtimestamp(curTime/1000).strftime('%Y-%m-%d_%H-%M-%S.%f')
    logFilePathName=logPath+"\\"+currentDateTime+"(API).log"
    logFile = open(logFilePathName, 'w+') 
    logFile.close()
    return logFilePathName,logPath

# #! 目前改版不會用到
# def __check_license(license,logPathName):
#     config=utils.__read_config_file('config.json')
#     check_license_dict=dict()
#     if len(config)==0:
#         SaveToLog(logPathName,"read config.json file failed.")
#         return False # ! config.json 檔案有問題
#     if 'check_license_exe_path' in config and  'check_license_exe_name' in  config and 'check_license_cmd' in config and 'license_count' in config:
#         # Users of AdvancedPI_TI_20:  (Total of 1 license issued;  Total of 0 licenses in use) # console畫面內的使用字串
#         license_count=config['license_count']        
#         if license_count>0:
#             for i in range(license_count):
#                 check_license_dict['OPI%d'%(i+1)]=config['check_OPI%d'%(i+1)]
#         else:
#             SaveToLog(logPathName,"read config.json file - license_count is 0.")
#             return False # ! 抓不到license資訊
#         # ! license轉換
#         license_trans=check_license_dict[license]
#         # if license == 'OPI1' or license == 'OPI2':
#         #     license_trans='OptimizePI_20'
#         # elif license == 'OPI3':
#         #     license_trans='AdvancedPI_TI_20'
#         exe_dir= os.path.join(config['check_license_exe_path'],config['check_license_exe_name'])
#         res=subprocess.run([exe_dir, 'lmstat', '-c' ,config['check_license_cmd'],'-f', license_trans],shell=True, capture_output=True, text=True)
#         stdoutstr=res.stdout
#         SaveToLog(logPathName,"check license exe: " + exe_dir)
#         SaveToLog(logPathName,"check license cmd: " + config['check_license_cmd'])
#         SaveToLog(logPathName,"license: "+license)
#         if len(stdoutstr)==0:   # 執行檔路徑有錯誤 or 參數錯誤
#             SaveToLog(logPathName,"check license execution failed")
#             return False
#         # * 解析console畫面內的文字
#         SaveToLog(logPathName,"check license console output : "+stdoutstr)
#         start_index = stdoutstr.find(f"Users of {license_trans}:")
#         end_index = stdoutstr.find(")", start_index) + 1
#         check_string = stdoutstr[start_index:end_index]
#         numbers = re.findall(r'\d+', check_string)
#         # ! True: license available
#         if numbers[2] != numbers[1]: 
#             SaveToLog(logPathName,"license available")
#             return True
#         else: 
#             SaveToLog(logPathName,"license not available")
#             return False
#     else:
#         SaveToLog(logPathName,"config.json with no params for check license.")
#         return False

# #! 目前改版不會用到
# def __find_license_info():
#     config=utils.__read_config_file('config.json')
#     license_dict=dict()
#     if len(config)==0:
#         return False # ! config.json 檔案有問題
#     # ! 讀取license個數 & 名稱
#     if 'license_count' in config:
#         license_count=config['license_count']        
#         if license_count>0:
#             for i in range(license_count):
#                 license_dict['OPI%d'%(i+1)]=config['OPI%d'%(i+1)]
#         else:
#             return False # ! 抓不到license資訊
#     else:
#         return False # ! 抓不到license資訊
#     return license_dict

# #! 目前改版不會用到
# def __check_resend_available(form_id,logPathName):
#     license=''
#     collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#     res=collect_tc.find_one({'form_id':form_id})
#     if res is None:
#         return True,-1,license
#     dfTask=pd.DataFrame(list(collect_tc.find({'form_id':form_id}))).sort_values(by='createTime', ascending=False)
    
#     if len(dfTask)==0:
#         SaveToLog(logPathName,"check resend available: first send")
#         return True,-1,license
#     if dfTask.iloc[0].status == 'Scheduled':
#         _id=dfTask.iloc[0]._id
#         order=dfTask.iloc[0].order
#         license=dfTask.iloc[0].license
#         resent=True
        
#         SaveToLog(logPathName,"check resend available: %s"%(str(_id)))
#         if 'resent_count' in dfTask.iloc[0]:
#             resent=False
#             SaveToLog(logPathName," ---- resend count: %d ----"%(dfTask.iloc[0]['resent_count']))
#         if 'current_opi_start_dt' in dfTask.iloc[0]:
#             if abs(dfTask.iloc[0].current_opi_start_dt-int(time.time())) <= (24*60*60) or resent==False: # 模擬開始的24小時內不能再更新
#                 SaveToLog(logPathName,"check resend available: false")
#                 return False,order,license
            
#             else:
#                 res=collect_tc.delete_one({'_id':_id})
#                 if res.deleted_count!=0:
#                     SaveToLog(logPathName,"check resend available: ok")
#                     return True,order,license
#         else:
#             return True,-1,license
#     elif dfTask.iloc[0].status == 'Running':
#         return False,-1,license
#     elif dfTask.iloc[0].status  not in ['Scheduled','Running']:
#         return True,-1,license
# #! 目前改版不會用到
# ----------------CheckBoardinTask------------------------------
# Input: board_number,board_stage
# Output: 1. 該版號無 'Running','Scheduled' and 'Finished'>=3  --->True
#         2. 該版號 'Running','Scheduled'  --->'In Schedules/Running'
#         3. 該版號 'Finished'>=3  --->'Over 3 Finished'

# ==================def ___CheckBoardinTask(board_number,board_stage,busItem,logPathName):
#     collect_tc = utils.ConnectToMongoDB(strDBName="Simulation",strTableName=DB_sheet)
#     check_boardnumber=list(collect_tc.find({'boardNumber':board_number,'boardStage':board_stage,'busItem':busItem,'status':{'$in':['Running','Scheduled']}}))
#     check_boardnumber_Finished=list(collect_tc.find({'boardNumber':board_number,'boardStage':board_stage,'status':'Finished','busItem':busItem}))
#     if len(check_boardnumber)==0 and len(check_boardnumber_Finished)==0:
#         SaveToLog(logPathName,"Check Board in Task: Pass!")
#         return True
#     elif len(check_boardnumber)>0:
#         SaveToLog(logPathName,"Check Board in Task: In Schedules/Running!")
#         return 'In Schedules/Running'
#     elif len(check_boardnumber_Finished)>=3:
#         SaveToLog(logPathName,"Check Board in Task: Over 3 Finished!")
#         return 'Over 3 Finished'

# ----------------___join_str------------------------------
# Input: *Args, 要重組的所有字樣會以link給的符號連上
# Output: result (ex. Peter_YH_Chang)
def ___join_str(*args,link='_'):
    result = args[0]
    for arg in args:
        if arg == args[0]:
            result = args[0]
        elif arg != '':
            result = result+link+arg
    return result


if __name__ == '__main__' :
    #!!! Set DB sheet !!!!!!!!!!!!!!!!!!!!!!!!!! #
    computer_name = socket.gethostname().upper()
    dict_computer = {
                    "TPER90115562":"DDR CCT",
                    "TPEO54012809":"PDN"
                    }
    busitem = dict_computer[computer_name]
    if 'dev' in os.getcwd().split('\\'):
        port_number = 726
        DB_sheet = "OPITaskCtrl_Debug"
    elif 'prd' in os.getcwd().split('\\'):
        port_number = 66
        DB_sheet = "OPITaskCtrl"
    app.debug = False
    if busitem == 'DDR CCT':
        server_name = "SPEED2000 #1"
    cert_file = os.path.join(os.getcwd(),'STAR_wistron.com.crt')
    key_file = os.path.join(os.getcwd(),'STAR_wistron.com.key')
    app.run(
        host='0.0.0.0',
        port=port_number,
        threaded=True,
        ssl_context=(cert_file, key_file)
    )