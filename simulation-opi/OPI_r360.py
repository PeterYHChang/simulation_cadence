import traceback
import secrets
import logging,os
import string
import pandas as pd
from dotenv import load_dotenv
from pymongo import MongoClient
from datetime import datetime,timedelta
from kafka.errors import KafkaError
from confluent_kafka import Producer
from confluent_kafka import avro
from confluent_kafka.avro import AvroProducer
from pymongo import DESCENDING
from os.path import basename

# load opi.env

load_dotenv(r'G:\simulation-opi_20241216\opi.env')

program = 'OPI Data to R360'
logPath = os.getcwd()+r'\Log_OPItoR360'
if not os.path.exists(logPath):
    os.makedirs(logPath)
logFormatter = '%(asctime)s [%(levelname)s] <%(funcName)s> : %(message)s'
dateFormat = '%Y-%m-%d %H:%M:%S'
today = datetime.today()
logging.basicConfig(filename=os.path.join(logPath, 'Log_OPItoR360.log'),
                    format=logFormatter, datefmt=dateFormat, level=logging.INFO)

logging.info(program)
logging.info("[Start]")

def connect_to_mongodb(strdbaddr=os.getenv('DB_URL'),strdbname="",str_table="") :
    conn=MongoClient(strdbaddr)
    db = conn[strdbname]
    collect = db[str_table]
    return(collect, conn)

# ================================== __create_random_string ==================================== #
# 產生 Kafka id -> 24 碼
# 不能有任何重複
def __create_random_string(field_name) -> str:
    col_request, client_request = connect_to_mongodb(strdbname="Simulation",str_table="RequestForm")
    random_str = ''.join(secrets.choice(string.ascii_letters + string.digits) for _ in range(24))
    while col_request.find_one({field_name:random_str}) is not None:
        random_str = ''.join(secrets.choice(string.ascii_letters + string.digits) for _ in range(24))
    client_request.close()
    return random_str

# ================================== get_plan_actual_data ==================================== #
# get sending data
def get_plan_actual_data() -> pd.DataFrame:
    logging.info("Run filter_request_data ...")
    col_request, _ = connect_to_mongodb(strdbname="Simulation", str_table=site_request)
    df_request = pd.DataFrame(list(col_request.find()))
    logging.info("Choose Rosa NB Intel pj successfully")
    #  3. C2/C3 Schedule (EVT/DVT), 發送 Plan PJ/ Actual PJ - Finished
    # df_plan = df_p1[df_p1['RequestStatus']=='Assignment'].reset_index(drop = True)
    # logging.info("Get Assignment pj(plan pj) successfully")
    df_p3 = drop_repeat(df_request)
    logging.info("Get Finished pj(plan/actual pj) successfully")
    return_list =[]
    for _,row in df_p3.iterrows():
        if check_str360(row):
            return_list.append(row.to_dict())
            logging.info(f"Get _id :{row['_id']}")
        else:logging.info(f"XXX Skip _id :{row['_id']} XXX")
    return pd.DataFrame(return_list)
#  reduce Complexity for get_plan_actual_data #
def drop_repeat(df_request)-> pd.DataFrame:
    # 篩選 project 符合 1. Rosa NB; 2. Intel 3. Stage in 'EVT' , Start date 為 2024/11/11 開始找 PJ
    # df_p1 = df_request[ (df_request['Product']=='NB') &
    #                     (df_request['Customer']=='Rosa') &
    #                     (df_request['createTime'] > start_d) &
    #                     (df_request['Stage'].isin(['RFQ','EVT'])) &
    #                     (df_request['Platform'].str.contains('Intel'))&
    #                     (df_request['_id'].isin(['20250324_1742827401','20250317_1742190300','20250217_1739791263']))
    #                     ]
    # df_p1 = df_request[ (df_request['Product']=='NB') &
    #                     (df_request['Customer']=='Rosa') &
    #                     (df_request['createTime'] > start_d) &
    #                     (df_request['Stage'].isin(['RFQ','EVT'])) &
    #                     (df_request['Platform'].str.contains('Intel'))
    #                     ]
    df_p1 = df_request[ (df_request['Product']=='NB') &
                    (df_request['Customer']=='Rosa') &
                    (df_request['createTime'] > start_d) &
                    (df_request['editTime'] >= pre_t) &
                    (df_request['editTime'] <= ct) &
                    (df_request['Stage'].isin(['RFQ','EVT'])) &
                    (df_request['Platform'].str.contains('Intel'))
                    ]
                    
    df_p2 = df_p1[df_p1['RequestStatus']=='Finished'].reset_index(drop = True)
    df_p2 = df_p2.sort_values('createTime').reset_index(drop = True)
    df_return = pd.DataFrame(columns = df_p2.columns)
    for index,row in df_p2.iterrows():
        if df_p2['Bus'][index][0]['Item'] !='PDN': continue
        ser = df_return[(df_return['PCBNO']==row['PCBNO'])&(df_return['PCBVer']==row['PCBVer'])]
        if ser.shape[0] == 0:
            df_return = pd.concat([df_return,pd.DataFrame([row])], ignore_index=True)
    return df_p2
def check_str360(ch_row) -> bool:
    col_str360, _ = connect_to_mongodb(strdbname="Simulation", str_table=site_SendtoR360)
    record = col_str360.find_one({'PCBNO': ch_row['PCBNO'], 'PCBVer': ch_row['PCBVer']})
    if record is None:return True
    elif record and record['_id'] == ch_row['_id']:return True
    else:return False
# =============================================================================================== #

def add_kr_data(df):
    logging.info("Run add_kr_actual_data ...")
    col_request, _ = connect_to_mongodb(strdbname="Simulation", str_table=site_request)
    # ---------------------------------------------------------------------------------------- #
    # 為符合 Sonarlint 寫法 建立 def update_row
    def update_row(dict_row, col_request):
        dict_row['Plan_KR1_id'] = __create_random_string('Plan_KR1_id')
        dict_row['Plan_KR2_id'] = __create_random_string('Plan_KR2_id')
        dict_row['KR1_Plan_Deliver_Date'] = ''
        dict_row['KR2_Plan_Deliver_Date'] = ''
        dict_row['Actual_KR1_id'] = __create_random_string('Actual_KR1_id')
        dict_row['Actual_KR2_id'] = __create_random_string('Actual_KR2_id')
        dict_row['Actual_Status'] = 'Finished'
        dict_row['KR1_Actual_Deliver_Date'] = ''
        dict_row['KR2_Actual_Deliver_Date'] = ''
        res = col_request.update_one(
            {'_id': dict_row['_id']},
            {
                '$set': {
                    'Plan_KR1_id':dict_row['Plan_KR1_id'],
                    'Plan_KR2_id':dict_row['Plan_KR2_id'],
                    'KR1_Plan_Deliver_Date':dict_row['KR1_Plan_Deliver_Date'],
                    'KR2_Plan_Deliver_Date':dict_row['KR2_Plan_Deliver_Date'],
                    'Actual_KR1_id': dict_row['Actual_KR1_id'],
                    'Actual_KR2_id': dict_row['Actual_KR2_id'],
                    'Actual_Status': dict_row['Actual_Status'],
                    'KR1_Actual_Deliver_Date': dict_row['KR1_Actual_Deliver_Date'],
                    'KR2_Actual_Deliver_Date': dict_row['KR2_Actual_Deliver_Date']
                }
            }
        )
        if res.acknowledged:
            logging.info(f"_id {dict_row['_id']} created plan/actual data successfully")
            return dict_row
        else:
            raise NameError(f"XXX _id {dict_row['_id']} created plan/actual data Fail XXX")
    def check_new_a(ch_row):
        if 'Actual_Status' not in ch_row.keys():
            return True
        return pd.isna(ch_row['Actual_Status'])

    # ---------------------------------------------------------------------------------------- #
    list_new = []
    df_sorted = df.sort_values(by='createTime')
    for _,row in df_sorted.iterrows():
        dict_row = row.to_dict()
        if check_new_a(row):
            dict_new = update_row(dict_row, col_request)
            list_new.append(dict_new)
        else:
            list_new.append(dict_row)
            logging.info(f"_id {row['_id']} is old plan/actual data")
    df_a_new = pd.DataFrame(list_new).reset_index(drop=True)

    # ----------------remove 過渡參數----------------#
    if df_a_new.shape[0] != 0:
        return df_a_new.copy()
    logging.info("No actual data")
    return pd.DataFrame({})

def delivery_report(err, msg):
    if err is not None:
        print(f'Message delivery failed: {err}')
    else:
        print(f'Message delivered to {msg.topic()} [{msg.partition()}] at offset {msg.offset()}')

def __update_deliver_status(doc,output_dict,mode,kr_id): # * mode: plan/actual 
    col_str360, _ = connect_to_mongodb(strdbname="Simulation", str_table=site_SendtoR360)
    col_request, _ = connect_to_mongodb(strdbname="Simulation", str_table=site_request)
    def save_str360_record(fieldname,doc,mode,kr_id):
        record = col_str360.find_one({'_id': doc['_id']})
        if mode =='plan' and kr_id ==1 and pd.isna(record):
            col_str360.insert_one({
                                '_id':doc['_id'],
                                'PCBNO':doc['PCBNO'],
                                'PCBVer':doc['PCBVer'],
                                'ProjectCode':doc['ProjectCode'],
                                'ProjectName':doc['ProjectName'],
                                'send_time': int(datetime.now().timestamp()),
                                fieldname: 'ok'})
        else:
            col_str360.update_one({'_id': doc['_id']}, {'$set': {fieldname: 'ok','send_time': int(datetime.now().timestamp())}})
    result=False
    # 更新文檔中的部分欄位資料
    if mode=='plan':        
        if kr_id==1:
            res=col_request.update_one({'_id': doc['_id']}, {'$set': {'KR1_Plan_Deliver_Date': output_dict['sync_ts']}})
            save_str360_record('KR1_Plan_Status',doc,mode,kr_id)
        if kr_id==2:
            res=col_request.update_one({'_id': doc['_id']}, {'$set': {'KR2_Plan_Deliver_Date': output_dict['sync_ts']}})
            save_str360_record('KR2_Plan_Status',doc,mode,kr_id)
    elif mode=='actual':        
        if kr_id==1:
            res=col_request.update_one({'_id': doc['_id']}, {'$set': {'KR1_Actual_Deliver_Date': output_dict['sync_ts']}})
            save_str360_record('KR1_Actual_Status',doc,mode,kr_id)
        if kr_id==2:
            res=col_request.update_one({'_id': doc['_id']}, {'$set': {'KR2_Actual_Deliver_Date': output_dict['sync_ts']}})
            save_str360_record('KR2_Actual_Status',doc,mode,kr_id)
    if res.modified_count>0:
        result=True
    return result

def __send_kafka_plan(dfplan,op): # op: update(U) 
    try:
        # * ---- kafka settings --------------------------------
        kafka_account=os.getenv('KAFKA_ACCOUNT')
        kafka_pwd={
            'dev':os.getenv('KAFKA_PASSWORD_DEV'),
            'qas':os.getenv('KAFKA_PASSWORD_QAS'),
            'prd':os.getenv('KAFKA_PASSWORD_PRD')
        }
        kafka_servers_addr = {
            'dev':os.getenv('KAFKA_ADDRESS_DEV'),
            'qas':os.getenv('KAFKA_ADDRESS_QAS'),
            'prd':os.getenv('KAFKA_ADDRESS_PRD')
        }
        schema_registry_url = {
            'dev':os.getenv('KAFKA_URL_DEV'),
            'qas':os.getenv('KAFKA_URL_QAS'),
            'prd':os.getenv('KAFKA_URL_PRD')
        }
        plan_topic='whq.r360dc.com.r360plan'
        plan_schema='''
            {
                "type": "record",
                "name": "ConnectDefault",
                "namespace": "io.confluent.connect.avro",
                "fields": [
                    {
                    "name": "bg",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "function",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "project_code",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "pcb_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "sw_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "sku_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "part_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "emdm_project_code",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "qt_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "sub_id",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "kr_id",
                    "type": [
                        "null",
                        "int"
                    ],
                    "default": null
                    },
                    {
                    "name": "sub_name",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "kr_name",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "saving_type",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "execution_count",
                    "type": [
                        "null",
                        "double"
                    ]
                    },
                    {
                    "name": "saving_hours_plan",
                    "type": [
                        "null",
                        "double"
                    ],
                    "default": null
                    },
                    {
                    "name": "cost_saving_plan",
                    "type": [
                        "null",
                        "double"
                    ],
                    "default": null
                    },
                    {
                    "name": "execution_datetime",
                    "type": [
                        "null",
                        {
                        "type": "long",
                        "connect.version": 1,
                        "connect.name": "org.apache.kafka.connect.data.Timestamp",
                        "logicalType": "timestamp-millis"
                        }
                    ],
                    "default": null
                    },
                    {
                    "name": "duration_datetime",
                    "type": [
                        "null",
                        {
                        "type": "long",
                        "connect.version": 1,
                        "connect.name": "org.apache.kafka.connect.data.Timestamp",
                        "logicalType": "timestamp-millis"
                        }
                    ],
                    "default": null
                    },
                    {
                    "name": "c_stage",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "sync_id",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "sync_op",
                    "type": [
                        "null",
                        "string"
                    ],
                    "default": null
                    },
                    {
                    "name": "sync_ts",
                    "type": [
                        "null",
                        {
                        "type": "long",
                        "connect.version": 1,
                        "connect.name": "org.apache.kafka.connect.data.Timestamp",
                        "logicalType": "timestamp-millis"
                        }
                    ],
                    "default": null
                    }
                ]
                }
        '''
        producer_config = {
            'bootstrap.servers': kafka_servers_addr[mode],
            'sasl.mechanisms': 'PLAIN',
            'security.protocol': 'SASL_PLAINTEXT',
            'schema.registry.url': schema_registry_url[mode],
            'sasl.username': kafka_account,
            'sasl.password': kafka_pwd[mode],
        }
        # ! 將 Avro 模式轉換為 AvroSchema 物件
        avro_schema_plan = avro.loads(plan_schema)        
        # * 傳送plan資料
        try:
            avro_producer = AvroProducer(producer_config, default_value_schema=avro_schema_plan)
        except Exception as e:
            print(traceback.format_exc())
            logging.error("(plan)[Error creating avro producer] %s" % traceback.format_exc())
        for _,row in dfplan.iterrows():
            # ! KR1
            row_dict_kr1 = {'bg': 'CPBG',
                'function': 'SIV',
                'project_code': '',
                'pcb_no': row['PCBNO'],
                'sw_no': '',
                'sku_no': '',
                'part_no': '',
                'emdm_project_code': '',
                'qt_no':'',
                'sub_id': os.getenv('KAFKA_SUB_ID'),
                'kr_id': 1,
                'sub_name': mission_name,
                'kr_name': mission_name,
                'saving_type': 'R',
                'execution_count': 1,
                'saving_hours_plan': 6,
                'cost_saving_plan': None,
                'execution_datetime': int(row['createTime']),
                'duration_datetime': int(row['createTime'] + 30*24*3600),
                'c_stage': 'C3',
                'sync_id': str(row['Plan_KR1_id']),
                'sync_op': op,
                'sync_ts': int(datetime.now().timestamp())
                }
            if send_data :
                avro_producer.produce(topic=plan_topic, value=row_dict_kr1, callback=delivery_report)
            # ! 更新傳送過去的狀態
            res=__update_deliver_status(row,row_dict_kr1,'plan',1)
            if res: logging.info('(plan)(%s)(KR1) %s update successfully' %(op,row['_id']))
            else: logging.error('(plan)(%s)(KR1) %s update fail' %(op,row['_id']))
            
            # ! KR2
            row_dict_kr2 = {'bg': 'CPBG',
                    'function': 'EE',
                    'project_code': row['ProjectCode'],
                    'pcb_no': row['PCBNO'],
                    'sw_no': '',
                    'sku_no': '',
                    'part_no': '',
                    'emdm_project_code': '',
                    'qt_no':'',
                    'sub_id': os.getenv('KAFKA_SUB_ID'),
                    'kr_id': 2,
                    'sub_name': mission_name,
                    'kr_name': mission_name,
                    'saving_type': 'C',
                    'execution_count': 1,
                    'saving_hours_plan': 0.1,
                    'cost_saving_plan': 0.1,
                    'execution_datetime': int(row['createTime']),
                    'duration_datetime': int(row['createTime'] + 30*24*3600),
                    'c_stage': 'C3',
                    'sync_id': str(row['Plan_KR2_id']),
                    'sync_op': op,
                    'sync_ts': int(datetime.now().timestamp())
                    }
            if send_data :
                avro_producer.produce(topic=plan_topic, value=row_dict_kr2, callback=delivery_report)
            # ! 更新傳送過去的狀態
            res=__update_deliver_status(row,row_dict_kr2,'plan',2)
            if res: logging.info('(plan)(%s)(KR2) %s update successfully' %(op,row['_id']))
            else: logging.error('(plan)(%s)(KR2) %s update fail' %(op,row['_id']))             
    except KafkaError as e:
        logging.error('(plan)(%s) kafka error %s' %(op,str(e)))
    finally:
        avro_producer.flush()

def __send_kafka_actual(dfactual,op): # op: update(U)  
    try:
        # * ---- kafka settings --------------------------------
        kafka_account=os.getenv('KAFKA_ACCOUNT')
        kafka_pwd={
            'dev':os.getenv('KAFKA_PASSWORD_DEV'),
            'qas':os.getenv('KAFKA_PASSWORD_QAS'),
            'prd':os.getenv('KAFKA_PASSWORD_PRD')
        }
        kafka_servers_addr = {
            'dev':os.getenv('KAFKA_ADDRESS_DEV'),
            'qas':os.getenv('KAFKA_ADDRESS_QAS'),
            'prd':os.getenv('KAFKA_ADDRESS_PRD')
        }
        schema_registry_url = {
            'dev':os.getenv('KAFKA_URL_DEV'),
            'qas':os.getenv('KAFKA_URL_QAS'),
            'prd':os.getenv('KAFKA_URL_PRD')
        }
        saving_topic='whq.r360dc.com.savingraw'
        saving_schema='''
            {
                "type": "record",
                "name": "savingraw",
                "namespace": "com.wistron.r360dc.com",
                "fields": [
                    {
                    "name": "bg",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "BG代碼, EX:CPBG, EBG",
                    "default": null
                    },
                    {
                    "name": "function",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": " EE, ME, SW, QT",
                    "default": null
                    },
                    {
                    "name": "project_code",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": " PLM project code, ME Project code",
                    "default": null
                    },
                    {
                    "name": "pcb_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "PCB no., GreenFormID",
                    "default": null
                    },
                    {
                    "name": "sw_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "DDE sw_no, CTW id",
                    "default": null
                    },
                    {
                    "name": "sku_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "SKU number. SKU P/N, RF SKU…",
                    "default": null
                    },
                    {
                    "name": "sub_id",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "R360 sub-module ID",
                    "default": null
                    },
                    {
                    "name": "kr_id",
                    "type": [
                        "null",
                        "int"
                    ],
                    "doc": "R360 sub-module KR ID. (1,2,3,4……)",
                    "default": null
                    },
                    {
                    "name": "sub_name",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "R360 sub-module name",
                    "default": null
                    },
                    {
                    "name": "kr_name",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "R360 sub-module KR name",
                    "default": null
                    },
                    {
                    "name": "saving_type",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "R or C. R: RD expense. C: Material Cost ",
                    "default": null
                    },
                    {
                    "name": "execution_count",
                    "type": [
                        "null",
                        "float"
                    ],
                    "doc": "執行次數. (Saving Type為R時, 此欄位需要有值)",
                    "default": null
                    },
                    {
                    "name": "saving_hours",
                    "type": [
                        "null",
                        "float"
                    ],
                    "doc": "每次節省工時. 單位: hours",
                    "default": null
                    },
                    {
                    "name": "cost_saving_plan",
                    "type": [
                        "null",
                        "float"
                    ],
                    "doc": "Cost Saving預計金額. 單位: USD$",
                    "default": null
                    },
                    {
                    "name": "cost_saving_actual",
                    "type": [
                        "null",
                        "float"
                    ],
                    "doc": "Cost Saving實際金額. 單位: USD$.",
                    "default": null
                    },
                    {
                    "name": "execution_datetime",
                    "type": [
                        "null",
                        {
                        "type": "long",
                        "logicalType": "timestamp-millis"
                        }
                    ],
                    "doc": "執行日期時間",
                    "default": null
                    },
                    {
                    "name": "duration_datetime",
                    "type": [
                        "null",
                        {
                        "type": "long",
                        "logicalType": "timestamp-millis"
                        }
                    ],
                    "doc": "結束執行日期時間",
                    "default": null
                    },
                    {
                    "name": "sync_id",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "UUID. 避免不然來源Primary key重覆",
                    "default": null
                    },
                    {
                    "name": "sync_op",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "I-Insert U-Update"
                    },
                    {
                    "name": "sync_ts",
                    "type": [
                        "null",
                        {
                        "type": "long",
                        "logicalType": "timestamp-millis"
                        }
                    ],
                    "doc": "Keep track of time when inserting the table",
                    "default": null
                    },
                    {
                    "name": "part_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "part no"
                    },
                    {
                    "name": "emdm_project_code",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "emdm_project_code"
                    },
                    {
                    "name": "qt_no",
                    "type": [
                        "null",
                        "string"
                    ],
                    "doc": "qt_no"
                    }
                ]
                }
        '''
        producer_config = {
            'bootstrap.servers': kafka_servers_addr[mode],
            'sasl.mechanisms': 'PLAIN',
            'security.protocol': 'SASL_PLAINTEXT',
            'schema.registry.url': schema_registry_url[mode],
            'sasl.username': kafka_account,
            'sasl.password': kafka_pwd[mode],
        }
        # ! 傳送saving資料
        avro_schema_saving = avro.loads(saving_schema)
        try:
           avro_producer = AvroProducer(producer_config, default_value_schema=avro_schema_saving)
        except Exception as e:
            print(traceback.format_exc())
            logging.error("(actual)[Error creating avro producer] %s" % traceback.format_exc())
        for _,row in dfactual.iterrows():
            # ! KR1
            row_dict_kr1 = {'bg': 'CPBG',
                    'function': 'SIV',
                    'project_code': '',
                    'pcb_no': row['PCBNO'],
                    'sw_no': '',
                    'sku_no': '',
                    'sub_id': os.getenv('KAFKA_SUB_ID'),
                    'kr_id': 1,
                    'sub_name': mission_name,
                    'kr_name': mission_name,
                    'saving_type': 'R',
                    'execution_count': 1,
                    'saving_hours': 6,
                    'cost_saving_plan': None,
                    'cost_saving_actual': None,
                    'execution_datetime': int(row['createTime'] + 30*24*3600),
                    'duration_datetime': None,
                    'sync_id': str(row['Actual_KR1_id']),
                    'sync_op': op,
                    'sync_ts': int(datetime.now().timestamp()),                    
                    'part_no': '',
                    'emdm_project_code': '',
                    'qt_no':''                  
                    }

            if send_data :
                avro_producer.produce(topic=saving_topic, value=row_dict_kr1, callback=delivery_report)
            # ! 更新傳送過去的狀態
            res=__update_deliver_status(row,row_dict_kr1,'actual',1)
            if res: logging.info('(actual)(%s)(KR1) %s update successfully' %(op,row['_id']))
            else: logging.error('(actual)(%s)(KR1) %s update fail' %(op,row['_id']))

            # ! KR2
            row_dict_kr2 = {'bg': 'CPBG',
                    'function': 'EE',
                    'project_code': row['ProjectCode'],
                    'pcb_no':row['PCBNO'],
                    'sw_no': '',
                    'sku_no': '',
                    'sub_id': os.getenv('KAFKA_SUB_ID'),
                    'kr_id': 2,
                    'sub_name': mission_name,
                    'kr_name': mission_name,
                    'saving_type': 'C',
                    'execution_count': 1,
                    'saving_hours': None,
                    'cost_saving_plan': 0.1,
                    'cost_saving_actual': 0.1,
                    'execution_datetime': int(row['createTime'] + 30*24*3600),
                    'duration_datetime': None,
                    'sync_id': str(row['Actual_KR2_id']),
                    'sync_op': op,
                    'sync_ts': int(datetime.now().timestamp()),                    
                    'part_no': '',
                    'emdm_project_code': '',
                    'qt_no':''                  
                    }
            if send_data :
                avro_producer.produce(topic=saving_topic, value=row_dict_kr2, callback=delivery_report)
            # ! 更新傳送過去的狀態
            res=__update_deliver_status(row,row_dict_kr2,'actual',2)
            if res: logging.info('(actual)(%s)(KR2) %s update successfully' %(op,row['_id']))
            else: logging.error('(actual)(%s)(KR2) %s update fail' %(op,row['_id']))
                
    except KafkaError as e:
        print(f"Error: {e}")
        logging.error('(actual)(%s) kafka error %s' %(op,str(e)))
    finally:
        # producer.flush()
        avro_producer.flush()

if __name__ == '__main__':
    # try:
    site_request = "RequestForm"
    site_opi_taskctrl = "OPITaskCtrl"
    site_SendtoR360 = "SendtoR360"
    mission_name="Auto OPI simulation"
    start_d = int(datetime(2024,11,11).timestamp())
    ct = int(today.timestamp())
    pre_date = datetime(today.year,today.month,today.day) - timedelta(days=1)
    pre_t = int(pre_date.timestamp())
    df_plan_actual = get_plan_actual_data()
    if df_plan_actual.shape[0] != 0:
        df_actual_kr = add_kr_data(df_plan_actual)
        mode='prd'
        send_data = True
        if df_actual_kr.shape[0] !=0:
            __send_kafka_plan(df_actual_kr,'U')
            __send_kafka_actual(df_actual_kr,'U')
        else: logging.info("PCB No. for New request today Has Been Sent to R360.")
    else: logging.info("No Any New Request Form Today.")
