import os
import pandas as pd
import numpy as np
# ADD metadata for check list
from object.metadata_for_checklist import Metadata_BC
import warnings
import xlwings as xw
from openpyxl import load_workbook
import datetime
import time
#Bỏ thông báo không cần thiết
warnings.filterwarnings("ignore", category=UserWarning, message="Data Validation extension is not supported and will be removed")
warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)
start_time = time.time()

project_name = "VN2025158_GOMME"
# project_name = "RGB_W17_M7"
# project_name = "VN2025024_EVE"


os.chdir("projects\\{}".format(project_name))

current_mdd_file = "data\\{}_EXPORT.mdd".format(project_name)
current_ddf_file = "data\\{}_EXPORT.ddf".format(project_name)

# current_mdd_file = "data\\VINCENZO_RGB_W17_M7_EXPORT_BATCH_2v1.mdd"
# current_ddf_file = "data\\VINCENZO_RGB_W17_M7_EXPORT_BATCH_2v1.ddf"

# current_mdd_file = "data\\{}_EXPORT_CLEAN_BOOSTER_2_25072025_v2.mdd".format(project_name)
# current_ddf_file = "data\\{}_EXPORT_CLEAN_BOOSTER_2_25072025_v2.ddf".format(project_name)

sql_query = "SELECT * FROM VDATA WHERE _LoaiPhieu = {_1} or _LoaiPhieu = {_2} or _LoaiPhieu = {_5}"

#Add các câu cần xem, or để null để show full
questions_list = []


# Mở workbook
wb = xw.Book("Check_list.xlsx")
# Lưu file mới
wb.save("Check_list_Temp.xlsx")
wb.close()

df_excel = pd.read_excel("Check_list_Temp.xlsx", engine="openpyxl",sheet_name="Checklist")
os.remove("Check_list_Temp.xlsx")

m = Metadata_BC(mdd_file=current_mdd_file, ddf_file=current_ddf_file, sql_query=sql_query)
df_datasource = m.convertToDataFrame(questions=questions_list)

header = pd.DataFrame(data=df_datasource.columns)

wb = xw.Book("Check_list.xlsx")

# Kiểm tra xem có sheet 'Variable' không, nếu không có thì thêm mới
if 'Variable' not in [sheet.name for sheet in wb.sheets]:
    sheet = wb.sheets.add('Variable', after='Checklist')
else:
    sheet = wb.sheets['Variable']
sheet.range('A:A').clear()  
sheet.range('B:B').clear() 

sheet.range('A1').value = header
sheet.range('1:1').delete()
sheet.visible = False

wb.sheets['Droplist'].visible = False

wb.save("Check_list.xlsx")
wb.close()

workbook = load_workbook("Check_list.xlsx")
wb = xw.Book("Check_list.xlsx")

df_datasource.set_index(["InstanceID"], inplace=True)

check = m.valcheck_checklist_import(df_excel)

today = datetime.datetime.now().strftime("%d%m%Y")
file_txt = f'Checklist_records_ErrorData_{today}.txt'

path = os.path.join('data', 'Checklist_DATA')

if not os.path.exists(path):
    os.makedirs(path, exist_ok=True)

file_path = os.path.join('data', 'Checklist_DATA', file_txt)

if os.path.exists(file_path):
    os.remove(file_path)

with open(file_path, 'a', encoding='utf-8') as f:
    f.write(f"Project: {project_name}\n")
    f.write(f"Date: {today}\n")


if check[1] == True:
    start_time_current = time.time()
    for index in range(len(df_excel)):
        func_name = df_excel["Funtion"][index]
        if func_name in ["Create_pushdata", "Create_diff", "Create_inter", "Create_get_iteration", "Create_union","Create_AnswerCount","Create_Compare_num"]:
            #Gọi funtion theo funtion name từ file metadata
            func = getattr(m, func_name, None)
            if callable(func):
                # Truyền các tham số cần thiết vào funtion được chọn
                create_dummy = func(df_excel, df_excel["Question Check"][index],index)
                print_KQ = m.save_results_to_file(create_dummy, file_path)
    print(f"Done - Create dummy data_({(time.time() - start_time_current)/60:.2f} phút)")

    check_functions = [
        ("askall", m.Valcheck_askall),
        ("Selected", m.Valcheck_Selected),
        ("num", m.Valcheck_num),
        ("Autocode Logic", m.Valcheck_Autocode_Logic),
        ("equal", m.Valcheck_equal),
        ("Not equal", m.Valcheck_Not_equal),
        ("sum", m.Valcheck_sum),
        ("initialize", m.Valcheck_initialize),
        ("filter by count", m.Valcheck_filterbycount),
        ("filter by cat", m.Valcheck_filterbycat),
        ("ask by route", m.Valcheck_askbyroute),
    ]

    for label, func in check_functions:
        start_time_current = time.time()
        result = func(df_excel)
        print_KQ = m.save_results_to_file(result, file_path)
        print(f"Done - Check {label}_({(time.time() - start_time_current)/60:.2f} phút)")

else:
    # print(check[0])
    print_KQ = m.save_results_to_file(check[0],file_path)

print(f"Complete check list_Total thời gian: ({(time.time() - start_time)/60:.2f} phút)")
