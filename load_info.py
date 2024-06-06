import variables as v
import my_function as my
from PIL import Image

# 讀取Info Data
if v.data_source == "Excel":
    title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df = my.load_info_by_excel()
elif v.data_source == "Azure SQL":
    title_df, introduction_df, common_col_df, individual_col_df, function_df, verification_code_df, internal_mail_receipients_df, supplier_mail_setting_df, photo_df, check_df = my.load_info_by_sql(v.program_name)

# 確認功能啟用
use_attachment = function_df[function_df["Item"] == "Attachment"]["Use"].values[0]
use_internal_mail = function_df[function_df["Item"] == "Send Internal Mail"]["Use"].values[0]
use_supplyer_mail = function_df[function_df["Item"] == "Send Supplier Mail"]["Use"].values[0]
use_verification_code = function_df[function_df["Item"] == "Verification Code"]["Use"].values[0]


# 讀取圖片檔
logo_filename = ""
icon_filename = ""
form_tail_filename = ""
if len(list(photo_df[photo_df["Item"] == "Logo"]["Filename"])) != 0:
    logo_filename = photo_df[photo_df["Item"] == "Logo"]["Filename"].values[0]
if len(list(photo_df[photo_df["Item"] == "Icon"]["Filename"])) != 0:
    icon_filename = photo_df[photo_df["Item"] == "Icon"]["Filename"].values[0]
if len(list(photo_df[photo_df["Item"] == "Form Tail"]["Filename"])) != 0:
    form_tail_filename = photo_df[photo_df["Item"] == "Form Tail"]["Filename"].values[0]
icon_image = ""
if icon_filename != "":
    icon_image = Image.open(v.system_photo_path.format(icon_filename))
