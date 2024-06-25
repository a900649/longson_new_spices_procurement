
# 資料存放位置
# data_source = "Excel"
data_source = "Azure SQL"

# 專案名稱
program_name = "longson_new_seasoning_n_premix_procurement_20240606"

# 存在可下載的專案名稱
program_name_list = ["longson_new_seasoning_n_premix_chicken_proc_20240611",
                     "longson_fruit_flavours_procurement_20240611",
                     "longson_new_spices_procurement_20240606",
                     "longson_new_seasoning_n_premix_procurement_20240606",
                     "longson_new_food_ingredient_procurement_20240606",
                     "longson_new_flavours_procurement_20240606",
                     "longson_wheat_flour_procurement_20240604",
                     "longson_tomato_paste_procurement_20240604",
                     "longson_soybean_oil_procurement_20240604",
                     "longson_refined_white_sugar_procurement_20240604",
                     "longson_pepper_procurement_20240604",
                     "longson_liquid_egg_procurement_20240604",
                     "longson_dried_chilli_procurement_20240604",
                     "longson_acid_procurement_20240604"]

# 網頁Title
page_title = "Longson Procurement System"

system_photo_path = "System Photo/{}"

logo_filename = 'Longson.jpg'
icon_filename = 'Smile.webp'
form_tail_filename = 'Thanks.jpg'

# 網頁Title
page_title = "Longson Procurement System"

# Excel的Info檔案名稱
excel_info_filename = "Longson Purchase Requirements Info.xlsx"

# Excel的結果檔案放置位置
results_file_path = "Procurement Results {}.xlsx".format(program_name)

# 其他資料存放區
temp_path = "Temp" + "/" + program_name
temp_data_path = temp_path + "//{}.json"
attachment_path = "Attachment" + "/" + program_name + "/{}" + "/{}"
system_photo_path = "System Photo/{}"


# DB 資訊
mysql_host = "longson.mysql.database.azure.com"
mysql_user='paul'
mysql_password='Yunxuan123'
mysql_port = 3306
db_name = "longson_procurement"

# Results Other Columns
first_col = [["RowID","Text"],["Product","Text"]]
last_col = [["Attachment","Text"],["Update DateTime","DateTime"],["Verification Code","Text"],["Verification Code Name","Text"]]

# BLOB
blob_connection_string = "DefaultEndpointsProtocol=https;AccountName=kenso;AccountKey=Wto5Ig361Z/aVQuxEfvM7b9MnKi3IctRB70fq5X53CCLlQ84BpFaS9T5HWVLcFwOVSEcljz0Aa40+AStQgHifw==;EndpointSuffix=core.windows.net"
blob_container = "longson"

# 寄送Mail資訊
# smtp_server = "smtp.office365.com"
# port = 587
# sender = "lfp-procure@outlook.com"
# mail_user = "lfp-procure@outlook.com"
# mail_password = "Power365*"

# Aaron給的Mail
smtp_server = "smtp.office365.com"
port = 587
sender = "longsonprocure@outlook.com"
mail_user = "longsonprocure@outlook.com"
mail_password = "Yunxuan123"

version = "Version 1.0625"

