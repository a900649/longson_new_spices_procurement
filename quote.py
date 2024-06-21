from datetime import datetime,timezone,timedelta
import pandas as pd
import streamlit as st
import variables as v
import numpy as np
import my_function as my
import warnings
import load_info
import json
import os
from azure.storage.blob import BlobServiceClient

warnings.filterwarnings("ignore")

st.set_page_config(page_title=v.page_title,page_icon=v.system_photo_path.format(v.icon_filename),layout='wide')

st.markdown("""
            <style>
            .big-font {
                font-size:40px;
                line-height: 0.7;
                font-weight:bold;
                text-align: center;
            }
            </style>
            """, unsafe_allow_html=True)

st.markdown('''
   <style>
       .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
       font-size:27px;
       font-weight:bold;
       }
   </style>
   ''', unsafe_allow_html=True)

st.markdown("""
            <style>
            .medium-font {
                font-size:25px;
                line-height: 0.7;
                font-weight:bold;
            }
            </style>
            """, unsafe_allow_html=True)

st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] 
    {
    gap: 40px;
    }
    .stTabs [data-baseweb="tab"] {
    height: 80px;
    white-space: pre-wrap;
    background-color: #FFFFFF;
    border-radius: 4px 4px 0px 0px;
    gap: 1px;
    padding-top: 10px;
    padding-bottom: 10px;
    }

    .stTabs [aria-selected="true"] {
    background-color: #D9D9D9;
    }

</style>""", unsafe_allow_html=True)





def verification_confirm(keyin_verification_code,now_datetime):

    verification_code_list = list(load_info.verification_code_df["Code"])

    if keyin_verification_code not in verification_code_list:
        st.session_state.verification = False
        st.session_state.name = now_datetime
        st.error("Verification Code Error")

    else:
        name = load_info.verification_code_df.loc[verification_code_list.index(keyin_verification_code), "Name"]
        st.session_state.verification = True
        st.session_state.name = name
        st.session_state.verification_code = keyin_verification_code
        st.success("Verification Code OK")

def create_df(name, product_list, col_name_list):

    if "df" not in st.session_state:

        record_dict = get_temp_data(name)

        if record_dict == {}:
            df = pd.DataFrame([[product] + [""] * (len(col_name_list) - 1) for product in product_list], columns=col_name_list)
        else:
            df = pd.DataFrame(record_dict)

        # index改成數字
        df.index = [i for i in range(0,len(df))]
        st.session_state["df"] = df

    st.dataframe(st.session_state["df"].set_index("Product"))
    # st.dataframe(st.session_state["df"])

    st.button("Submit Quote", key='submit')

    if "submit_status" in st.session_state:
        if st.session_state["submit_status"]:
            st.success("Submit Succeeded")

    # st.write(st.session_state)


def get_temp_data(name):
    if v.data_source == "Azure SQL":
        record = my.load_temp_row(name)
        if record == ():
            record_dict = {}
        else:
            record_dict = json.loads(record[0][0])
    elif v.data_source == "Excel":
        if os.path.isfile(v.temp_data_path.format(name)):
            with open(v.temp_data_path.format(name)) as f:
                record_dict = json.loads(json.load(f))
        else:
            record_dict = {}

    return record_dict

def create_data_input(col_info_list,results_dict,area):

    col = col_info_list[0]
    col_type = col_info_list[1]
    col_option = col_info_list[2]
    col_instructions = col_info_list[3]
    if col_instructions == "" or str(col_instructions) == "nan":
        show_col = col
    else:
        show_col = col + "(" + str(col_instructions) + ")"
    if area == "Common":
        ind = 0
    else:
        ind = int(area[-1]) - 1

    default_value = st.session_state['df'].loc[ind,col]

    if col_type.lower() == "text":
        results_dict[col] = st.text_input(show_col, key=col + " " + area, value=default_value)
    elif col_type.lower() == "date":
        if default_value == None or default_value == "" or str(default_value) == "nan":
            results_dict[col] = st.date_input(show_col, key=col + " " + area)
        else:
            results_dict[col] = st.date_input(show_col, key=col + " " + area, value=datetime.strptime(str(default_value), "%Y-%m-%d").date())
    elif col_type.lower() == "decimal":
        col_option = round(float(col_option))
        if default_value == None or default_value == "" or str(default_value) == "nan":
            default_value = None
        results_dict[col] = st.number_input(show_col, key=col + " " + area, step=10 ** -col_option, format="%.{}f".format(col_option), value=default_value)
    elif col_type.lower() == "integer":
        if default_value == None or default_value == "" or str(default_value) == "nan":
            default_value = None
        else:
            default_value = int(default_value)
        results_dict[col] = st.number_input(show_col, key=col + " " + area, step=1, value=default_value)
    elif col_type.lower() == "radio":
        col_option_list = col_option.replace("\n", "").split(";")
        if col_option_list[-1] == "":
            del col_option_list[-1]
        if default_value == None or default_value == "" or str(default_value) == "nan":
            results_dict[col] = st.radio(show_col, key=col + " " + area, options=col_option_list, index=0)
        else:
            results_dict[col] = st.radio(show_col, key=col + " " + area, options=col_option_list, index=col_option_list.index(default_value))
    else:
        results_dict[col] = st.text_input(show_col, key=col + " " + area, value=default_value)
    return results_dict

def form(now_datetime,name):

    product_list = list(load_info.introduction_df["Product"])
    product_num = len(product_list)
    df_col_name_list = ["Product"] + list(load_info.common_col_df["Column Name"]) + list(load_info.individual_col_df["Column Name"])
    save_col_name_list = [i[0] for i in v.first_col] + list(load_info.common_col_df["Column Name"]) + list(load_info.individual_col_df["Column Name"]) + [i[0] for i in v.last_col]

    # 前圖
    front_image1, front_image2, front_image3, front_image4, front_image5 = st.columns([1, 1, 1, 1, 1])
    with front_image3:
        st.image(v.system_photo_path.format(v.logo_filename))

    # Title
    title1, title2, title3 = st.columns([1, 1, 1])
    with title2:
        title_list = list(load_info.title_df["Title"])
        for text in title_list:
            st.markdown('<p class="big-font">{}<p>'.format(str(text)), unsafe_allow_html=True)



    # Content
    common_col_list = [[load_info.common_col_df.loc[i,"Column Name"],
                        load_info.common_col_df.loc[i,"Data Type"],
                        load_info.common_col_df.loc[i,"Option"],
                        load_info.common_col_df.loc[i,"Instructions"]] for i in range(0,len(load_info.common_col_df))]
    individual_col_list = [[load_info.individual_col_df.loc[i,"Column Name"],
                            load_info.individual_col_df.loc[i,"Data Type"],
                            load_info.individual_col_df.loc[i,"Option"],
                            load_info.individual_col_df.loc[i,"Instructions"]] for i in range(0, len(load_info.individual_col_df))]

    a1, a2, a3 = st.columns([1, 10, 1])
    with a2:
        st.markdown('---')

        create_df(name, product_list, df_col_name_list)

        st.markdown('---')

        common_results_dict = {}
        # common_col_area_1, common_col_area_2 = st.columns([1, 1])
        # with common_col_area_1:
        #     for i in range(0, len(common_col_list),2):
        #         common_results_dict = create_data_input(common_col_list[i], common_results_dict, "Common")
        #
        # with common_col_area_2:
        #     for i in range(1, len(common_col_list),2):
        #         common_results_dict = create_data_input(common_col_list[i], common_results_dict, "Common")

        for i in range(0, len(common_col_list),1):
            common_results_dict = create_data_input(common_col_list[i], common_results_dict, "Common")


        st.write("")  # 做個間隔
        st.write("")  # 做個間隔


        # Tab
        if product_num == 1:
            pass
            # tab1 = st.tabs(product_list)
        elif product_num == 2:
            tab1, tab2 = st.tabs(product_list)
        elif product_num == 3:
            tab1, tab2, tab3 = st.tabs(product_list)
        elif product_num == 4:
            tab1, tab2, tab3, tab4 = st.tabs(product_list)
        elif product_num == 5:
            tab1, tab2, tab3, tab4, tab5 = st.tabs(product_list)

        for ind in range(0,product_num):
            if ind == 0:
                if product_num == 1 :
                    st.write("")  # 做個間隔
                    st.markdown('<p class="medium-font">{}<p>'.format("Requirements introduction."), unsafe_allow_html=True)
                    st.write("")  # 做個間隔
                    introduction_str = load_info.introduction_df.loc[ind, "Item"]
                    st.text(introduction_str)
                    st.write("")  # 做個間隔

                    # individual_col_area_1, individual_col_area_2 = st.columns([1, 1])
                    tab1_individual_results_dict = {}
                    # with individual_col_area_1:
                    #     for i in range(0, len(individual_col_list), 2):
                    #         tab1_individual_results_dict = create_data_input(individual_col_list[i], tab1_individual_results_dict, "Tab1")
                    #
                    # with individual_col_area_2:
                    #     for i in range(1, len(individual_col_list), 2):
                    #         tab1_individual_results_dict = create_data_input(individual_col_list[i], tab1_individual_results_dict, "Tab1")

                    for i in range(0, len(individual_col_list)):
                        tab1_individual_results_dict = create_data_input(individual_col_list[i], tab1_individual_results_dict, "Tab1")

                    if str(load_info.use_attachment).lower() == "true":
                        tab1_uploaded_file_list = st.file_uploader(label="Upload Attachment",
                                                                   type=["xlsx", "xls", "jpg", "doc", "docx", "pdf", "ppt", "pptx"],
                                                                   accept_multiple_files=True,
                                                                   key="Attachment Tab1")
                else:
                    with tab1:
                        st.write("")  # 做個間隔
                        st.markdown('<p class="medium-font">{}<p>'.format("Requirements introduction."), unsafe_allow_html=True)
                        st.write("")  # 做個間隔
                        introduction_str = load_info.introduction_df.loc[ind,"Item"]
                        st.text(introduction_str)
                        st.write("")  # 做個間隔

                        # individual_col_area_1, individual_col_area_2 = st.columns([1, 1])
                        tab1_individual_results_dict = {}
                        # with individual_col_area_1:
                        #     for i in range(0, len(individual_col_list), 2):
                        #         tab1_individual_results_dict = create_data_input(individual_col_list[i], tab1_individual_results_dict, "Tab1")
                        #
                        # with individual_col_area_2:
                        #     for i in range(1, len(individual_col_list), 2):
                        #         tab1_individual_results_dict = create_data_input(individual_col_list[i], tab1_individual_results_dict, "Tab1")

                        for i in range(0, len(individual_col_list)):
                            tab1_individual_results_dict = create_data_input(individual_col_list[i], tab1_individual_results_dict, "Tab1")

                        if str(load_info.use_attachment).lower() == "true":
                            tab1_uploaded_file_list = st.file_uploader(label="Upload Attachment",
                                                                       type=["xlsx", "xls", "jpg", "doc", "docx", "pdf", "ppt", "pptx"],
                                                                       accept_multiple_files=True,
                                                                       key="Attachment Tab1")

            elif ind == 1:
                with tab2:
                    st.write("")  # 做個間隔
                    st.markdown('<p class="medium-font">{}<p>'.format("Requirements introduction."), unsafe_allow_html=True)
                    st.write("")  # 做個間隔
                    introduction_str = load_info.introduction_df.loc[ind, "Item"]
                    st.text(introduction_str)
                    st.write("")  # 做個間隔

                    # individual_col_area_1, individual_col_area_2 = st.columns([1, 1])
                    tab2_individual_results_dict = {}
                    # with individual_col_area_1:
                    #     for i in range(0, len(individual_col_list), 2):
                    #         tab2_individual_results_dict = create_data_input(individual_col_list[i], tab2_individual_results_dict, "Tab2")
                    #
                    # with individual_col_area_2:
                    #     for i in range(1, len(individual_col_list), 2):
                    #         tab2_individual_results_dict = create_data_input(individual_col_list[i], tab2_individual_results_dict, "Tab2")

                    for i in range(0, len(individual_col_list)):
                        tab2_individual_results_dict = create_data_input(individual_col_list[i], tab2_individual_results_dict, "Tab2")

                    if str(load_info.use_attachment).lower() == "true":
                        tab2_uploaded_file_list = st.file_uploader(label="Upload Attachment",
                                                                   type=["xlsx", "xls", "jpg", "doc", "docx", "pdf", "ppt", "pptx"],
                                                                   accept_multiple_files=True,
                                                                   key="Attachment Tab2")


            elif ind == 2:
                with tab3:
                    st.write("")  # 做個間隔
                    st.markdown('<p class="medium-font">{}<p>'.format("Requirements introduction."), unsafe_allow_html=True)
                    st.write("")  # 做個間隔
                    introduction_str = load_info.introduction_df.loc[ind, "Item"]
                    st.text(introduction_str)
                    st.write("")  # 做個間隔

                    # individual_col_area_1, individual_col_area_2 = st.columns([1, 1])
                    tab3_individual_results_dict = {}
                    # with individual_col_area_1:
                    #     for i in range(0, len(individual_col_list), 2):
                    #         tab3_individual_results_dict = create_data_input(individual_col_list[i], tab3_individual_results_dict, "Tab3")
                    #
                    # with individual_col_area_2:
                    #     for i in range(1, len(individual_col_list), 2):
                    #         tab3_individual_results_dict = create_data_input(individual_col_list[i], tab3_individual_results_dict, "Tab3")

                    for i in range(0, len(individual_col_list)):
                        tab3_individual_results_dict = create_data_input(individual_col_list[i], tab3_individual_results_dict, "Tab3")

                    if str(load_info.use_attachment).lower() == "true":
                        tab3_uploaded_file_list = st.file_uploader(label="Upload Attachment",
                                                                   type=["xlsx", "xls", "jpg", "doc", "docx", "pdf", "ppt", "pptx"],
                                                                   accept_multiple_files=True,
                                                                   key="Attachment Tab3")


            elif ind == 3:
                with tab4:
                    st.write("")  # 做個間隔
                    st.markdown('<p class="medium-font">{}<p>'.format("Requirements introduction."), unsafe_allow_html=True)
                    st.write("")  # 做個間隔
                    introduction_str = load_info.introduction_df.loc[ind, "Item"]
                    st.text(introduction_str)
                    st.write("")  # 做個間隔

                    # individual_col_area_1, individual_col_area_2 = st.columns([1, 1])
                    tab4_individual_results_dict = {}
                    # with individual_col_area_1:
                    #     for i in range(0, len(individual_col_list), 2):
                    #         tab4_individual_results_dict = create_data_input(individual_col_list[i], tab4_individual_results_dict, "Tab4")
                    #
                    # with individual_col_area_2:
                    #     for i in range(1, len(individual_col_list), 2):
                    #         tab4_individual_results_dict = create_data_input(individual_col_list[i], tab4_individual_results_dict, "Tab4")

                    for i in range(0, len(individual_col_list)):
                        tab4_individual_results_dict = create_data_input(individual_col_list[i], tab4_individual_results_dict, "Tab4")

                    if str(load_info.use_attachment).lower() == "true":
                        tab4_uploaded_file_list = st.file_uploader(label="Upload Attachment",
                                                                   type=["xlsx", "xls", "jpg", "doc", "docx", "pdf", "ppt", "pptx"],
                                                                   accept_multiple_files=True,
                                                                   key="Attachment Tab4")


            elif ind == 4:
                with tab5:
                    st.write("")  # 做個間隔
                    st.markdown('<p class="medium-font">{}<p>'.format("Requirements introduction."), unsafe_allow_html=True)
                    st.write("")  # 做個間隔
                    introduction_str = load_info.introduction_df.loc[ind, "Item"]
                    st.text(introduction_str)
                    st.write("")  # 做個間隔

                    # individual_col_area_1, individual_col_area_2 = st.columns([1, 1])
                    tab5_individual_results_dict = {}
                    # with individual_col_area_1:
                    #     for i in range(0, len(individual_col_list), 2):
                    #         tab5_individual_results_dict = create_data_input(individual_col_list[i], tab5_individual_results_dict, "Tab5")
                    #
                    # with individual_col_area_2:
                    #     for i in range(1, len(individual_col_list), 2):
                    #         tab5_individual_results_dict = create_data_input(individual_col_list[i], tab5_individual_results_dict, "Tab5")

                    for i in range(0, len(individual_col_list)):
                        tab5_individual_results_dict = create_data_input(individual_col_list[i], tab5_individual_results_dict, "Tab5")

                    if str(load_info.use_attachment).lower() == "true":
                        tab5_uploaded_file_list = st.file_uploader(label="Upload Attachment",
                                                                   type=["xlsx", "xls", "jpg", "doc", "docx", "pdf", "ppt", "pptx"],
                                                                   accept_multiple_files=True,
                                                                   key="Attachment Tab5")


        st.button("Confirm modification",key='alter_df')


    if st.session_state["alter_df"] == True:
        for ind in range(0, product_num):
            for key in common_results_dict.keys():
                value = common_results_dict[key]
                st.session_state["df"].loc[ind, key] = value

        for ind in range(0, product_num):
            if ind == 0:
                for key in tab1_individual_results_dict.keys():
                    value = tab1_individual_results_dict[key]
                    st.session_state["df"].loc[ind, key] = value
            elif ind == 1:
                for key in tab2_individual_results_dict.keys():
                    value = tab2_individual_results_dict[key]
                    st.session_state["df"].loc[ind, key] = value
            elif ind == 2:
                for key in tab3_individual_results_dict.keys():
                    value = tab3_individual_results_dict[key]
                    st.session_state["df"].loc[ind, key] = value
            elif ind == 3:
                for key in tab4_individual_results_dict.keys():
                    value = tab4_individual_results_dict[key]
                    st.session_state["df"].loc[ind, key] = value
            elif ind == 4:
                for key in tab5_individual_results_dict.keys():
                    value = tab5_individual_results_dict[key]
                    st.session_state["df"].loc[ind, key] = value

        # st.experimental_rerun()
        st.rerun()



    if st.session_state['submit'] == True:
        df = st.session_state["df"]

        # 判斷Verification Code
        v_code = st.session_state.verification_code
        if st.session_state.verification_code != now_datetime:
            v_name = load_info.verification_code_df[load_info.verification_code_df["Code"] == st.session_state.verification_code]["Name"].values[0]
        else:
            v_name = None


        # 暫存檔
        for col_list in common_col_list+individual_col_list:
            if col_list[1] == "Date":
                df[col_list[0]] = df[col_list[0]].apply(lambda x:str(x))


        if v.data_source == "Azure SQL":
            my.upload_temp_row(name, df.to_json())
        elif v.data_source == "Excel":
            if os.path.isdir(v.temp_path) == False:
                os.makedirs(v.temp_path)
            with open(v.temp_data_path.format(name), 'w') as f:
                json.dump(df.to_json(), f) # 暫存檔

        # 寫入附件
        attachment_filename_list = []
        if str(load_info.use_attachment).lower() == "true":

            if v.data_source == "Azure SQL":
                blob_service_client = BlobServiceClient.from_connection_string(v.blob_connection_string)
                container_client = blob_service_client.get_container_client(v.blob_container)
                for ind in range(0, len(product_list)):
                    product = product_list[ind]
                    attachment_data_path = v.attachment_path.format(product,name + " " + now_datetime)
                    if ind == 0:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab1_uploaded_file_list])
                        for uploaded_file in tab1_uploaded_file_list:
                            blob_client = container_client.get_blob_client(attachment_data_path + "\\" + uploaded_file.name)
                            blob_client.upload_blob(uploaded_file.getbuffer(), overwrite=True)
                    elif ind == 1:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab2_uploaded_file_list])
                        for uploaded_file in tab2_uploaded_file_list:
                            blob_client = container_client.get_blob_client(attachment_data_path + "\\" + uploaded_file.name)
                            blob_client.upload_blob(uploaded_file.getbuffer(), overwrite=True)
                    elif ind == 2:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab3_uploaded_file_list])
                        for uploaded_file in tab3_uploaded_file_list:
                            blob_client = container_client.get_blob_client(attachment_data_path + "\\" + uploaded_file.name)
                            blob_client.upload_blob(uploaded_file.getbuffer(), overwrite=True)
                    elif ind == 3:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab4_uploaded_file_list])
                        for uploaded_file in tab4_uploaded_file_list:
                            blob_client = container_client.get_blob_client(attachment_data_path + "\\" + uploaded_file.name)
                            blob_client.upload_blob(uploaded_file.getbuffer(), overwrite=True)
                    elif ind == 4:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab5_uploaded_file_list])
                        for uploaded_file in tab5_uploaded_file_list:
                            blob_client = container_client.get_blob_client(attachment_data_path + "\\" + uploaded_file.name)
                            blob_client.upload_blob(uploaded_file.getbuffer(), overwrite=True)

                    attachment_filename_list.append(attachment_filename)

            elif v.data_source == "Excel":
                for ind in range(0, len(product_list)):
                    product = product_list[ind]
                    attachment_data_path = v.attachment_path.format(product, name + " " + now_datetime)

                    # 檢查附件目錄
                    if os.path.isdir(attachment_data_path) == False:
                        os.makedirs(attachment_data_path)

                    if ind == 0:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab1_uploaded_file_list])
                        for uploaded_file in tab1_uploaded_file_list:
                            with open(attachment_data_path + "\\" + uploaded_file.name, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                    elif ind == 1:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab2_uploaded_file_list])
                        for uploaded_file in tab2_uploaded_file_list:
                            with open(attachment_data_path + "\\" + uploaded_file.name, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                    elif ind == 2:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab3_uploaded_file_list])
                        for uploaded_file in tab3_uploaded_file_list:
                            with open(attachment_data_path + "\\" + uploaded_file.name, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                    elif ind == 3:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab4_uploaded_file_list])
                        for uploaded_file in tab4_uploaded_file_list:
                            with open(attachment_data_path + "\\" + uploaded_file.name, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                    elif ind == 4:
                        attachment_filename = ",".join([uploaded_file.name for uploaded_file in tab5_uploaded_file_list])
                        for uploaded_file in tab5_uploaded_file_list:
                            with open(attachment_data_path + "\\" + uploaded_file.name, "wb") as f:
                                f.write(uploaded_file.getbuffer())

                    attachment_filename_list.append(attachment_filename)
        else:
            attachment_filename_list = [None] * len(product_list)


        # 存正式資料
        if v.data_source == "Azure SQL":
            sql_data_list_dict = {}
            for ind in range(0,len(product_list)):
                product = product_list[ind]
                sql_data_list = []
                sql_data_list.append(now_datetime + " " + product)  # RowID
                for col in df_col_name_list:
                    value = st.session_state['df'].loc[ind, col]
                    sql_data_list.append(value)

                sql_data_list.append(attachment_filename_list[ind])
                sql_data_list.append(datetime.strptime(now_datetime, "%Y-%m-%d %H%M%S"))
                sql_data_list.append(v_code)
                sql_data_list.append(v_name)
                sql_data_list_dict[product] = sql_data_list

            my.upload_results_table(sql_data_list_dict,save_col_name_list)

        elif v.data_source == "Excel":
            if os.path.isfile(v.results_file_path) == False:
                my.create_new_results_file(common_col_list, individual_col_list, product_list)

            excel_results_df_list = []
            for ind in range(0,len(product_list)):
                product = product_list[ind]
                sheetname = product.replace("/"," ").replace("\\"," ").replace("?"," ").replace("*"," ").replace("["," ").replace("]"," ") # Excel Sheetname 不接受
                excel_results_df = pd.read_excel(v.results_file_path, engine="openpyxl", sheet_name=sheetname)
                excel_results_df.set_index("RowID", inplace=True, drop=True)
                for col in df_col_name_list:

                    value = st.session_state['df'].loc[ind,col]
                    excel_results_df.loc[now_datetime, col] = value

                excel_results_df.loc[now_datetime, "Update DateTime"] = now_datetime
                excel_results_df.loc[now_datetime, "Verification Code"] = v_code
                excel_results_df.loc[now_datetime, "Verification Code Name"] = v_name
                excel_results_df.loc[now_datetime, "Attachment"] = attachment_filename_list[ind]

                excel_results_df_list.append(excel_results_df)

            writer = pd.ExcelWriter(v.results_file_path)
            for ind in range(0, len(product_list)):
                sheetname = product_list[ind].replace("/", " ").replace("\\", " ").replace("?", " ").replace("*", " ").replace("[", " ").replace("]", " ")  # Excel Sheetname 不接受
                excel_results_df_list[ind].to_excel(writer, sheet_name=sheetname)
            writer.close()

        # Send Mail
        if str(load_info.use_internal_mail).lower() == "true":
            my.send_internal_mail(st.session_state['df'])

        # Send Supplyer Mail
        if str(load_info.use_supplyer_mail).lower() == "true":
            my.send_supplyer_mail(st.session_state['df'])

        st.session_state["submit_status"] = True
        # st.experimental_rerun()
        st.rerun()




    # 尾圖
    tail_image1, tail_image2, tail_image3 = st.columns([1, 1, 1])
    with tail_image2:
        st.image(v.system_photo_path.format(v.form_tail_filename))



def form_page():

    now_datetime = (datetime.now(timezone.utc) + timedelta(hours=8)).strftime("%Y-%m-%d %H%M%S")
    if str(load_info.use_verification_code).lower() == "false":
        st.session_state.verification = True
        st.session_state.name = now_datetime
        st.session_state.verification_code = now_datetime

    if st.session_state.get("verification") == False or st.session_state.get("verification") == None:

        keyin_verification_code = st.text_input(label="Please Enter Verification Code")

        # if st.button('Check Verification Code', on_click=verification_confirm, args=[keyin_verification_code, now_datetime]):
        #     pass

        if st.button('Check Verification Code',key="Check Verification Code"):
            verification_code_list = list(load_info.verification_code_df["Code"])
            if keyin_verification_code not in verification_code_list:
                st.session_state.verification = False
                st.session_state.name = now_datetime
                st.error("Verification Code Error")

            else:
                name = load_info.verification_code_df.loc[verification_code_list.index(keyin_verification_code), "Name"]
                st.session_state.verification = True
                st.session_state.name = name
                st.session_state.verification_code = keyin_verification_code
                st.success("Verification Code OK")
                st.rerun()

    elif st.session_state.get("verification") == True:
        form(now_datetime, st.session_state.name)

    st.markdown('---')



form_page()
