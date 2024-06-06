import pymysql
import variables as v
import my_function as my
from datetime import datetime

def create_db():
    # 建立DB
    db = pymysql.connect(host=v.mysql_host,port=v.mysql_port,user=v.mysql_user,password=v.mysql_password)
    cursor = db.cursor()
    cursor.execute("CREATE DATABASE IF NOT EXISTS {}".format(v.db_name))
    db.close()

def create_info_table():
    # 建立 Info
    db = pymysql.connect(host=v.mysql_host,port=v.mysql_port,user=v.mysql_user,password=v.mysql_password,database=v.db_name)
    cursor = db.cursor()
    sql = """CREATE TABLE IF NOT EXISTS {}_info (
                     `Tilte` VARCHAR(500),
                     `Requirements Introduction` LONGTEXT,
                     `Common Columns` LONGTEXT,
                     `Individual Columns` LONGTEXT,
                     `Function` VARCHAR(3000),
                     `Verification Code` LONGTEXT,
                     `Internal Mail Receipients` VARCHAR(3000),
                     `Supplier Mail Setting` VARCHAR(3000),
                     `Photo` VARCHAR(3000),
                     `Check` LONGTEXT,                 
                     `Update DateTime` DateTime,
                     PRIMARY KEY (Tilte))""".format(v.program_name)
    cursor.execute(sql)
    db.close()

def create_temp_table():
    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    cursor = db.cursor()
    sql = """CREATE TABLE IF NOT EXISTS {}_temp (
                         `Name` VARCHAR(500),
                         `Data` LONGTEXT,
                         PRIMARY KEY (Name))""".format(v.program_name)
    cursor.execute(sql)
    db.close()


def update_info_table():
    data_df_list = my.load_info_by_excel()
    data_df_list[2].fillna("",inplace=True)
    data_df_list[3].fillna("",inplace=True)
    data_str_list = [df.to_json(orient="split") for df in data_df_list] + [datetime.now()]

    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    cursor = db.cursor()
    sql = "REPLACE INTO {}_info (`Tilte`,`Requirements Introduction`,`Common Columns`,`Individual Columns`,`Function`,`Verification Code`," \
          "`Internal Mail Receipients`,`Supplier Mail Setting`,`Photo`,`Check`,`Update DateTime`) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)".format(v.program_name)
    k = cursor.executemany(sql, [data_str_list])
    db.commit()
    db.close()

def create_results_table():
    title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df = my.load_info_by_excel()

    columns_list = v.first_col +\
                   [[common_col_df.loc[i, "Column Name"], common_col_df.loc[i, "Data Type"]] for i in range(0, len(common_col_df))] + \
                   [[individual_col_df.loc[i, "Column Name"], individual_col_df.loc[i, "Data Type"]] for i in range(0, len(individual_col_df))] + \
                   v.last_col

    db_col_str_list = []
    for col in columns_list:
        name = col[0]
        col_type = col[1]
        col_str = "`" + name + "`" + " "
        if col_type.lower() == "text" or col_type.lower() == "radio":
            if name == "RowID":
                col_str = col_str + "VARCHAR(300)"
            else:
                # col_str = col_str + "VARCHAR(500)"
                col_str = col_str + "TEXT"
        elif col_type.lower() == "integer":
            col_str = col_str + "INT"
        elif col_type.lower() == "decimal":
            col_str = col_str + "DOUBLE"
        elif col_type.lower() == "date":
            col_str = col_str + "Date"
        elif col_type.lower() == "datetime":
            col_str = col_str + "DateTime"
        db_col_str_list.append(col_str)

    db_col_str = ", ".join(db_col_str_list)

    # 建立 Results
    db = pymysql.connect(host=v.mysql_host, port=v.mysql_port, user=v.mysql_user, password=v.mysql_password, database=v.db_name)
    cursor = db.cursor()
    sql = """CREATE TABLE IF NOT EXISTS {}_results ({}, PRIMARY KEY (RowID))""".format(v.program_name,db_col_str)
    cursor.execute(sql)
    db.close()

create_info_table()
create_temp_table()
update_info_table()
create_results_table()