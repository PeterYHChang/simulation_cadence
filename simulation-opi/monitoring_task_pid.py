import psutil
import os
import logging
import socket
from package import utils
from bson import ObjectId
import pandas as pd
from pymongo import MongoClient
from dotenv import load_dotenv


# load opi.env
computer_name = socket.gethostname().upper()
dict_env = {
            "TPER90115562":r'H:\simulation-opi\opi.env',
            "TPEO54012809":r'G:\simulation-opi\opi.env'
            }
load_dotenv(dict_env[computer_name])
log_path =os.getcwd()
log_format = '%(asctime)s [%(levelname)s] <%(funcName)s> : %(message)s'
date_format = '%Y-%m-%d %H:%M:%S'
logging.basicConfig(filename=os.path.join(log_path, 'Task_PID_Monitor.log'),format=log_format,datefmt=date_format,level=logging.INFO)
logging.info('Task_PID_Monitor')
logging.info("[Start]")

def connect_to_mongodb(strdbaddr=os.getenv('DB_URL'),strdbname="",str_table="") ->tuple:
    conn=MongoClient(strdbaddr)
    db = conn[strdbname]
    collect = db[str_table]
    return collect

def check_si_task_by_pid(check_duration:int)-> bool:
    """
    [Warnning]只能滿足一個license,有多個必須詢問 Cadence 如何抓到Automation.exe 執行 PowerSI.exe 的 PID
    Args:
        check_duration (int): check cpu usage of pid for "check_duration" seconds(times)

    Returns:
        bool: False for tool executing/ True for the pid simualtion is closed.  
    """
    target_name = "SIMetrics_Automation.exe"
    power_si_pid = None
    for proc in psutil.process_iter(['pid', 'name', 'cmdline','username', 'cpu_percent']):
        # match 'name' and user is Erin
        if proc.info['name'] == target_name and proc.info['username'] == 'Erin':
            power_si_pid=proc.info['pid']
    if not power_si_pid:
        logging.info(f"Process with PID {power_si_pid} does not exist.")
        return False
    try:
        process = psutil.Process(power_si_pid)
        n_cores = psutil.cpu_count()
        for _ in range(check_duration):
            cpu_usage = process.cpu_percent(interval=1) / n_cores
            logging.info(f"{power_si_pid=}, average {cpu_usage=} %")
            if cpu_usage != float(0):
                logging.info("PowerSI.exe Is Executing...")
                return False  #  tool is executing
    except psutil.NoSuchProcess:
        logging.info(f"Process with PID {power_si_pid} does not exist.")
        return False
    except psutil.AccessDenied:
        logging.error(f"XXXXX Access denied to terminate process with PID {power_si_pid}. XXXXXX")
        return True
    except Exception as e:
        logging.error(f"XXXXX Error when check_task_by_pid: {e}")
        return True

# --------------------------- check_task_by_pid --------------------------- #
def check_task_by_pid(pid:int,check_duration:int)->bool:
    """
    Args:
        pid (int): pid by power rail
        check_duration (int): check cpu usage of pid for "check_duration" seconds(times)

    Returns:
        bool: False for tool executing/ True for the pid simualtion is closed.  
    """
    try:
        process = psutil.Process(pid)
        n_cores = psutil.cpu_count()
        for _ in range(check_duration):
            cpu_usage = process.cpu_percent(interval=1) / n_cores
            logging.info(f"{pid=}, average {cpu_usage=} %")
            if cpu_usage != float(0):
                logging.info(f"{pid=} Is Executing!!!")
                return False  #  tool is executing
    except psutil.NoSuchProcess:
        logging.info(f"Process with PID {pid} does not exist.")
        return False
    except psutil.AccessDenied:
        logging.error(f"XXXXX Access denied to terminate process with PID {pid}. XXXXXX")
        return True
    except Exception as e:
        logging.error(f"XXXXX Error when check_task_by_pid: {e}")
        return True

def check_run_time_log(path:str)->bool:
    if not os.path.exists(path):
        return False
    file = open(path,'r')
    content = file.read()
    file.close()
    logging.info(f"check RunTimeErrorLog_path:{path}")
    logging.info(f"check RunTimeErrorLog:\n {content}")
    if 'ERROR' in content.upper() or len(content)==0:
        return False
    return True
def find_and_print_warning_or_error(file_content):
    lines = file_content.split('\n')
    message = ""
    for index, line in enumerate(lines):
        if "Warning/Error message of the simulation:" in line:
            message = lines[index + 1].strip() if index + 1 < len(lines) else ""
    return message
def main():
    config_name = 'config.json'
    config_dict = utils.__read_config_info(config_name)
    monitor_lim = config_dict["Monitor_limit"]
    computer_name = socket.gethostname().upper()
    dict_computer = {
                    "TPER90115562":"Auto DDR CCT",
                    "TPEO54012809":"Auto PDN"
                    }
    dict_db = {
            "dev":'67f60a74b489d66cb6263390',
            "prd":'682593e55f3ca14e4425115a'
            }

    busitem = dict_computer[computer_name]
    if 'dev' in os.getcwd().split('\\'):
        db_id = dict_db['dev']
        db_sheet = "OPITaskCtrl_Debug"
    elif 'prd' in os.getcwd().split('\\'):
        db_sheet = "OPITaskCtrl"
        db_id = dict_db['prd']
    collect_run = connect_to_mongodb(strdbname="Simulation", str_table=db_sheet)
    df_running_task = pd.DataFrame(collect_run.find({'status': 'Running','busItem':busitem}))
    collect_check = connect_to_mongodb(strdbname="Simulation", str_table='Check_Stuck')
    dict_check = collect_check.find_one({'_id':ObjectId(db_id)})
    if df_running_task.empty:
        logging.info("No Task Now")
        return
    for _,row in df_running_task.iterrows():
        r_license = row.get('license')
        logging.info(f"Get Form_id : {row.form_id}, license : {r_license}")
        count_now = dict_check.get(r_license,0)
        for tcl in row.tcl:
            if 'pid' not in tcl:
                break
            pid = tcl['pid']
        logging.info(f"get PID : {pid}")
        logging.info(f"get DB(Check_Stuck) for {r_license} count now: {count_now}")
        if 'PDN' in row.busItem:
            count_now += 1 if check_task_by_pid(pid,3) else 0
            tcl_base_path = os.path.dirname(tcl['filepath'])
            error_log_path = os.path.join(tcl_base_path,"RunTimeError.log")
        else:
            count_now += 1 if check_si_task_by_pid(3) else 0
            error_log_path = os.path.join(row.filePath['output_path'],row.Rank+'_'+row.DataRate,'board0')
        logging.info(f"Update Count Now: {count_now}")
        if count_now <= monitor_lim:
            result = collect_check.update_one(
                                            {'_id': ObjectId(db_id)},
                                            {'$set': {r_license:count_now}}
                                        )
            logging.info("Update to DB")
            logging.info(f"Matched count: {result.matched_count}, Modified count: {result.modified_count}")
            continue
        logging.info(f"Over limits: {monitor_lim}")
        # logging.info("Check RunTimeErrorLog")
        # if check_run_time_log(error_log_path):
        #     logging.info("Check No Error in RunTimeErrorLog")
        #     continue
        # tool is stuck, so terminated
        try:
            process = psutil.Process(pid)
            # process.terminate()  # Attempt to terminate the process
            process.wait(timeout=5)  # Wait for the process to terminate, timeout in 5 seconds
            logging.info(f"Terminate PID: {pid}")
            if 'DDR CCT' in row.busItem:
                target_name = "SIMetrics_Automation.exe"
                power_si_pid = None
                for proc in psutil.process_iter(['pid', 'name', 'cmdline','username', 'cpu_percent']):
                    if proc.info['name'] == target_name and proc.info['username'] == 'Erin':
                        power_si_pid=proc.info['pid']
                        process = psutil.Process(power_si_pid)
                        # process.terminate()
                        logging.info("Terminate SIMetrics_Automation.exe")
                        break
            # check again if pid is closed. 
            process = psutil.Process(pid)
        except psutil.NoSuchProcess:
            result = collect_check.update_one(
                                            {'_id': ObjectId(db_id)},
                                            {'$set': {r_license:0}}
                                        )
            logging.info(f"Process with PID {pid} does not exist.")
        except psutil.AccessDenied:
            logging.error(f"XXXXX Access denied to terminate process with PID {pid}. XXXXXX")
        except Exception as e:
            logging.error(f"XXXXX Error when check_task_by_pid: {e}")
if __name__ == '__main__':
    main()
    