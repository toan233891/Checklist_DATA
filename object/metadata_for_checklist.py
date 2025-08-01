import win32com.client as w32
import pandas as pd
import numpy as np
import re

import tkinter as tk
from tkinter import filedialog, messagebox


class mrDataFileDsc:
    def __init__(self, **kwargs):
        if "mdd_file" in list(kwargs.keys()):
            self.mdd_file = kwargs.get("mdd_file")
        if "ddf_file" in list(kwargs.keys()):
            self.ddf_file = kwargs.get("ddf_file")
        if "sql_query" in list(kwargs.keys()):
            self.sql_query = kwargs.get("sql_query")
        if "dms_file" in list(kwargs.keys()):
            self.dms_file = kwargs.get("dms_file")
        
        self.MDM = w32.Dispatch(r'MDM.Document')
        self.adoConn = w32.Dispatch(r'ADODB.Connection')
        self.adoRS = w32.Dispatch(r'ADODB.Recordset')
        self.DMOMJob = w32.Dispatch(r'DMOM.Job')
        self.Directives = w32.Dispatch(r'DMOM.StringCollection')

    def openMDM(self):
        self.MDM.Open(self.mdd_file)

    def saveMDM(self):
        self.MDM.Save(self.mdd_file)

    def closeMDM(self):
        self.MDM.Close()

    def openDataSource(self):
        conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(self.ddf_file, self.mdd_file)

        self.adoConn.Open(conn)

        self.adoRS.ActiveConnection = conn
        self.adoRS.Open(self.sql_query)
        
    def closeDataSource(self):
        #Close and clean up
        if self.adoRS.State == 1:
            self.adoRS.Close()
            self.adoRS = None
        if self.adoConn is not None:
            self.adoConn.Close()
            self.adoConn = None

    def runDMS(self):
        self.Directives.Clear()
        self.Directives.add('#define InputDataFile ".\{}"'.format(self.mdd_file))

        self.Directives.add('#define OutputDataMDD ".\{}"'.format(self.mdd_file.replace('.mdd', '_EXPORT.mdd')))
        self.Directives.add('#define OutputDataDDF ".\{}"'.format(self.mdd_file.replace('.mdd', '_EXPORT.ddf')))

        self.DMOMJob.Load(self.dms_file, self.Directives)
        self.DMOMJob.Run()

class Metadata_BC(mrDataFileDsc):
    def __init__(self, **kwargs):
        try:
            match len(kwargs.keys()):
                case 1:
                    self.mdd_file = kwargs.get("mdd_file")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file)
                case 2:
                    self.mdd_file = kwargs.get("mdd_file")
                    self.dms_file = kwargs.get("dms_file")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file, dms_file=self.dms_file)
                case 3:
                    self.mdd_file = kwargs.get("mdd_file")
                    self.ddf_file = kwargs.get("ddf_file")
                    self.sql_query = kwargs.get("sql_query")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file, ddf_file=self.ddf_file, sql_query=self.sql_query)
                case 4:
                    self.mdd_file = kwargs.get("mdd_file")
                    self.ddf_file = kwargs.get("ddf_file")
                    self.sql_query = kwargs.get("sql_query")
                    self.dms_file = kwargs.get("dms_file")
                    mrDataFileDsc.__init__(self, mdd_file=self.mdd_file, ddf_file=self.ddf_file, sql_query=self.sql_query, dms_file=self.dms_file)
        except ValueError as ex:
            print("Error")

    def convertToDataFrame(self, questions):
        self.openMDM()
        self.openDataSource()
        
        data = self.adoRS.GetRows()
        
        columns = [f.Name for f in self.adoRS.Fields]
        
        self.df = pd.DataFrame(data=np.array(data).T, columns=columns)

        if len(questions)>0:
            self.df = self.df[questions]
        else:
            #Loại các câu có format date
            delete_questions = ['SOURCEPROJECTID','System_LocationID','NWB_STATUS','NWB_LAST_SAVE_ON_SERVER','NWB_LAST_SUBMIT','NWB_CANCEL_REASON','NWB_CANCEL_REASON.ZZZ','SHELL_QFA','SHELL_QFB','SHELL_QFC','SHELL_QFD','SHELL_QFE','SHELL_QFF','SHELL_INT_LENGTH','SHELL_BLOCK.SHELL_APPLICATION_ID','SHELL_BLOCK.SHELL_INTERVIEWER_LOGIN','SHELL_BLOCK.SHELL_SCH1','SHELL_BLOCK.SHELL_SCH2','SHELL_BLOCK.SHELL_SCH3','SHELL_BLOCK.SHELL_GEOLOCATION_OUTCOME','SHELL_BLOCK.SHELL_GEOLOCATION_LATITUDE','SHELL_BLOCK.SHELL_GEOLOCATION_LONGITUDE','SHELL_BLOCK.SHELL_GEOLOCATION_ACCURACY','SHELL_BLOCK.SHELL_GEOLOCATION_TIMESTAMP','SHELL_CHAINID','SHELL_COUNTRY','SHELL_LANGUAGE','SHELL_LANGUAGECODE','SHELL_INTRO_GDPR','SHELL_RECORDING_CONFIRMATION','SHELL_PRIVACY_NOTICE','SHELL_SIGNATURE','SHELL_GENDER','SHELL_AGE','SHELL_AGE._A1','SHELL_AGE_RECODED','_ResName','_ResAddress','_ResHouseNo','_ResStreet','_ResDistricts','_ResDistrictSelected','_ResWards','_ResWardSelected','_ResPhone','_ResPhone._1','_ResCellPhone','_ResCellPhone._1','_Email','_Email._1','_Interview_Year','_STARTTIME','_IntID','_IntName','_Signature','_ENDTIME','_SPANTIME','_TOTALTIME','_Info_Sup','SHELL_DDG_STATUS','SHELL_NAME','SHELL_BLOCK_TEL.SHELL_MOBTEL','SHELL_BLOCK_TEL.SHELL_HOMETEL','SHELL_BLOCK_TEL.SHELL_BC_EMAIL','SHELL_TEL','SHELL_TEXTTEL','SHELL_EMAIL','SHELL_BLOCK_ADDRESS.SHELL_HOUSENO','SHELL_BLOCK_ADDRESS.SHELL_STREET','SHELL_BLOCK_ADDRESS.SHELL_DISTRICT','SHELL_BLOCK_ADDRESS.SHELL_TOWN','SHELL_BLOCK_ADDRESS.SHELL_ZIP','SHELL_ADDRESS','SHELL_COMP_MODE','SHELL_IP_ADDRESS','_BHP','_PAQ','SHELL_SUP']
            non_date_questions = []
            for question in columns:
                if not pd.api.types.is_datetime64_any_dtype(self.df[question]) and question not in delete_questions:
                    non_date_questions.append(question)
            if len(non_date_questions) > 0:
                self.df = self.df[non_date_questions]    

        # self.df= self.df.replace({"_": "","{": "","}": ""}, regex=True)
        if "InstanceID" in self.df.columns:
            instance_col = self.df["InstanceID"].copy()
            df_replace = self.df.drop(columns=["InstanceID"])
            df_replace = df_replace.replace({"_": "","{": "","}": ""}, regex=True)
            df_replace["InstanceID"] = instance_col
            self.df = df_replace
        else:
            self.df = self.df.replace({"_": "","{": "","}": ""}, regex=True)

        self.closeMDM()
        self.closeDataSource()
        
        return self.df
       
    def valcheck_checklist_import(self,df_excel):
        results = []
        # Hiển thị popup thông báo
        root = tk.Tk()
        root.withdraw()
        check = True  
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Create_pushdata") | (df_excel["Funtion"] == "Create_diff") | (df_excel["Funtion"] == "Create_inter") | (df_excel["Funtion"] == "Valcheck_filterbycount") | (df_excel["Funtion"] == "Valcheck_filterbycat") | (df_excel["Funtion"] == "Create_get_iteration") | (df_excel["Funtion"] == "Create_union") | (df_excel["Funtion"] == "Create_AnswerCount") , "Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                if i in self.df.columns:
                    row_excel = self.get_row_index(df_excel,i)
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} tên đã có trong DATA MDD"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)
                    check = False
        
        #Kiểm tra all
        Question_Check = df_excel.loc[(df_excel["Funtion"] != "") ,"Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                if i.count("[..]") > 1:
                    row_excel = self.get_row_index(df_excel,i)
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} có hơn 2 loop '[..]'. Nếu hơn 2 loop thì chỉ cần loop cuối cùng, các loop trước khai báo cụ thể iteration. Ví dụ: Loop1[_1].Loop2[..]._Codes"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)

                for question_condition, value_condition, relation in conditions:
                    if pd.notnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() != ""):
                        if question_condition.count("[..]") > 1:
                            result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {question_condition} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} có hơn 2 loop '[..]'. Nếu hơn 2 loop thì chỉ cần loop cuối cùng, các loop trước khai báo cụ thể iteration. Ví dụ: Loop1[_1].Loop2[..]._Codes"
                            print(result_str)
                            results.append(result_str)
                            messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                            check = False                         
        #Kiểm tra từng câu
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_askbyroute") ,"Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)
                row_excel = self.get_row_index(df_excel,i)
                for question_condition, value_condition, relation in conditions:
                    if pd.isnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() == ""):
                        result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu tên câu điều kiện nhưng lại có giá trị điều kiện là {value_condition}"
                        print(result_str)
                        results.append(result_str)
                        messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                        check = False          
                    if pd.isnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() == ""):
                        result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu giá trị điều kiện nhưng lại có tên câu điều kiện là {question_condition} "
                        print(result_str)
                        results.append(result_str)
                        messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)    
                        check = False  

        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_equal") | (df_excel["Funtion"] == "Valcheck_Not_equal") | (df_excel["Funtion"] == "Create_Compare_num") ,"Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                row_excel = self.get_row_index(df_excel,i)

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)
                
                if len(conditions) <=0:
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu tên câu điều kiện/value"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False     

        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Create_diff") | (df_excel["Funtion"] == "Create_inter") | (df_excel["Funtion"] == "Create_union") ,"Question Check"]
        
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                row_excel = self.get_row_index(df_excel,i)
                if pd.isnull(df_excel.loc[index, 'Current_Value']) or (isinstance(df_excel.loc[index, 'Current_Value'], str) and df_excel.loc[index, 'Current_Value'].strip() == ""):
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu thông tin ở cột Current_Value"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False  

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)
                
                if len(conditions) <=0:
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu tên câu điều kiện/value"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False          
 
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Create_pushdata") | (df_excel["Funtion"] == "Valcheck_Autocode_Logic"),"Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                row_excel = self.get_row_index(df_excel,i)
                if pd.isnull(df_excel.loc[index, 'Current_Value']) or (isinstance(df_excel.loc[index, 'Current_Value'], str) and df_excel.loc[index, 'Current_Value'].strip() == ""):
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu thông tin ở cột Current_Value"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False  

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)
                
                if len(conditions) <=0:
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu tên câu điều kiện/value"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False      

                for question_condition, value_condition, relation in conditions:
                    if pd.isnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() == ""):
                        result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu tên câu điều kiện nhưng lại có giá trị điều kiện là {value_condition}"
                        print(result_str)
                        results.append(result_str)
                        messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                        check = False          
              
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_filterbycat") | (df_excel["Funtion"] == "Valcheck_initialize") | (df_excel["Funtion"] == "Create_AnswerCount"),"Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                row_excel = self.get_row_index(df_excel,i)
                if pd.isnull(df_excel.loc[index, 'QRE_1']) or (isinstance(df_excel.loc[index, 'QRE_1'], str) and df_excel.loc[index, 'QRE_1'].strip() == ""):
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu thông tin ở cột QRE_1"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False

        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_num") | (df_excel["Funtion"] == "Valcheck_Selected"),"Question Check"]
        if not Question_Check.empty:                 
            for index,i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                row_excel = self.get_row_index(df_excel,i)
                if pd.isnull(df_excel.loc[index, 'Current_Value']) or (isinstance(df_excel.loc[index, 'Current_Value'], str) and df_excel.loc[index, 'Current_Value'].strip() == ""):
                    result_str = f"[Tại dòng: {row_excel}]: Kiểm tra câu hỏi {i} trong Check_list thuộc funtion: {df_excel.loc[index, 'Funtion']} bị thiếu thông tin ở cột Current_Value"
                    print(result_str)
                    results.append(result_str)
                    messagebox.showinfo("Thông báo [valcheck_checklist_import]", result_str)   
                    check = False  

        return results, check  
    
    def get_conditions_FULL(self, df_excel, index):
        #Get các điều kiện và giá trị của điều kiện ở file excel checklist
        conditions = []
        condition_prefix = 'QRE_'
        value_prefix = 'VALUE_'
        related_prefix = 'RELATED_'
        
        #Xác định số lượng điều kiện
        num_conditions = []
        for col in df_excel.columns:
            if col.startswith(condition_prefix):
                qre_condition = df_excel.loc[index, col]
                stt_qre = int(col.split("_")[1])
                value_col = f'{value_prefix}{stt_qre}'
                val_condition = df_excel.loc[index, value_col]
                if pd.notnull(qre_condition) or (isinstance(qre_condition, str) and qre_condition.strip() != ""):
                    num_conditions.append(col)
                else:
                    if pd.notnull(val_condition) or (isinstance(val_condition, str) and val_condition.strip() != ""):
                        num_conditions.append(col)

        for qre in num_conditions:
            stt_qre = int(qre.split("_")[1])
            value_col = f'{value_prefix}{stt_qre}'
            related_col = f'{related_prefix}{stt_qre}_AND_{stt_qre+1}'
            qre_condition = df_excel.loc[index, qre]
            value_condition = df_excel.loc[index, value_col]
            try:
                related_condition = df_excel.loc[index, related_col]
            except:
                related_condition= None
            conditions.append((qre_condition, value_condition, related_condition))   

        return conditions

    def get_conditions(self, df_excel, index):
        #Get các điều kiện và giá trị của điều kiện ở file excel checklist
        conditions = []
        condition_prefix = 'QRE_'
        value_prefix = 'VALUE_'
        related_prefix = 'RELATED_'
        
        #Xác định số lượng điều kiện
        num_conditions = []
        for col in df_excel.columns:
            if col.startswith(condition_prefix):
                question_condition = df_excel.loc[index, col]
                if pd.notnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() != ""):
                    num_conditions.append(col)

        for qre in num_conditions:
            stt_qre = int(qre.split("_")[1])
            value_col = f'{value_prefix}{stt_qre}'
            related_col = f'{related_prefix}{stt_qre}_AND_{stt_qre+1}'
            qre_condition = df_excel.loc[index, qre]
            value_condition = df_excel.loc[index, value_col]
            try:
                related_condition = df_excel.loc[index, related_col]
            except:
                related_condition= None
            conditions.append((qre_condition, value_condition, related_condition))   

        return conditions
    
    def check_conditions(self, id, conditions):
        overall_result = True
        current_relation = None
        #Run từng câu và giá trị điều kiện trong list conditions đã khai báo
        for question_condition, value_condition, relation in conditions:
            try:
                if pd.notnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() != ""):
                    if id == "2032655":
                        a=0
                    #Get giá trị hiện tại của câu điều kiện trong dataframe, split để tách các câu có thể là MA ra từng code và đưa về dạng list để so sánh
                    current_answers = self.df.at[id, question_condition]
                    current_answers = self.convert_value(current_answers)                    

                    value_condition_check = self.convert_value(value_condition)
                    arr_intersection = set(current_answers).intersection(set(value_condition_check))
                    #Với giá trị chuỗi cần cách check riêng or có các ký tự ! phủ định 
                    condition_result = False
                    if isinstance(value_condition, str):
                        if len(current_answers) > 0 and len(value_condition_check) > 0 and current_answers[0] != 'NULL' and value_condition_check[0] != 'NULL':
                            if ">" in value_condition:
                                if int(current_answers[0]) > int(value_condition_check[0]):
                                    condition_result = True
                            if "<" in value_condition:
                                if int(current_answers[0]) < int(value_condition_check[0]):
                                    condition_result = True                          
                        if "!" in value_condition:
                            if len(arr_intersection) <= 0:
                                condition_result = True
                        else:
                            if not "<" in value_condition and not ">" in value_condition:
                                if len(arr_intersection) > 0:
                                    condition_result = True                           
                    else:
                        if len(arr_intersection) > 0:
                            condition_result = True
                    #Kiểm tra điều kiện check logic giữa các câu, với câu đầu tiên thì chưa check do đang gán NONE
                    if current_relation == "&":
                        overall_result = overall_result and condition_result
                    elif current_relation == "|":
                        overall_result = overall_result or condition_result
                    else:
                        overall_result = condition_result
                    #Sau câu đầu tiên, bắt đầu lấy điều kiện check để check với câu sau đó
                    current_relation = relation
            except Exception as e:
                print(f"[check_conditions]: Lỗi xử lý điều kiện tại câu {question_condition}: {e}")
        return overall_result

    def save_results_to_file(self, results, filename):
        with open(filename, 'a', encoding='utf-8') as f:
            if results == None:
                return
            else:
                for result in results:
                    f.write(result + '\n')

    def get_row_index(self, df_excel, question):
        # Truy vấn dòng của câu đang xử lý
        row_index = df_excel.index[df_excel["Question Check"] == question].tolist()
        if row_index:
            return row_index[0] + 3
        return row_index

    def diff_lists(self, list1, list2):
        list = []
        for item in list1:
            if item not in list2:
                list.append(item)
        return list

    def convert_value(self, value):
        if isinstance(value, str):
            pass
        else:
            value = int(value)
            
        if (pd.notnull(value) and (isinstance(value, str) and value.strip() != "" or isinstance(value, (int, float)) )):
            if isinstance(value, str):
                value = value.replace(" ","").replace("!","").replace(">","").replace("<","").split(',')
                value = [int(x) if x.isdigit() else x.upper() for x in value if pd.notnull(x) and x != '']
            else:
                value = [int(value)]                                
        else:
            value = ['NULL']
        return value
    
    def get_qre_loop(self, Qre):
        Qre = pd.Series(Qre)
        for i in Qre:
            try:
                if "[..]" in i:
                    LoopName = i.split("[..]")[0]
                    LoopChidren = i.split("[..].")[1]   
                    for qre_loop in self.df.columns:
                        try:
                            if LoopName in qre_loop:
                                if LoopChidren == qre_loop.split(".")[-1]:
                                    Qre = pd.concat([Qre, pd.Series([qre_loop])], ignore_index=True)
                        except:
                            break
                    Qre = Qre.unique()
                    Qre = pd.Series(Qre)
                    Qre = Qre[Qre != i]         
            except Exception as e:
                print(f"[get_qre_loop]: Lỗi xử lý điều kiện tại câu {i}: {e}")              
        return Qre
        
    def Create_pushdata(self, df_excel, Question_Check,index):
        results = []
        if len(Question_Check) > 0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist thông qua funtion get conditions
                conditions = []
                conditions = self.get_conditions(df_excel,index)

                for id in self.df.index.tolist():
                #Get current_value ở file excel checklist
                    if (id == "2000817"):
                        a=0
                    Current_Value = df_excel.loc[index, 'Current_Value']
                    if not isinstance(Current_Value, str):
                        Current_Value = int(df_excel.loc[index, 'Current_Value'])
                    else:
                        for qre in self.df.columns:
                            if qre in Current_Value:
                                if len(qre) == len(Current_Value):
                                    Current_Value = self.df.at[id, qre]
                                    break   
                    # Dùng funtion check_conditions được viết riêng để so sánh giá trị hiện tại và giá trị tại các câu điều kiện
                    if self.check_conditions(id, conditions):
                        # Nếu kết quả trả về true thì thực hiện push data
                        if pd.isnull(Current_Value) or (isinstance(Current_Value, str) and Current_Value.strip() == "")  or (isinstance(Current_Value, str) and Current_Value.upper() == "NULL"):
                            pass
                        else: 
                            if self.df.at[id, Question_Check]:
                                self.df.at[id, Question_Check] += f",{Current_Value}"
                            else:
                                self.df.at[id, Question_Check] = f"{Current_Value}"
                self.df[Question_Check] = self.df[Question_Check].apply(lambda x: x.rstrip(','))
                a=0
            except Exception as e:
                print(f"[Create_pushdata]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results

    def Create_diff(self, df_excel, Question_Check, index):
        results = []
        if len(Question_Check)>0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist thông qua funtion get conditions
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)

                for id in self.df.index.tolist():
                    if id == "2055607":
                        a=0
                    main_question_value = df_excel.loc[index, 'Current_Value']
                    if not isinstance(main_question_value, str):
                        main_question_value = int(df_excel.loc[index, 'Current_Value'])
                    else:
                        for qre in self.df.columns:
                            if qre in main_question_value:
                                if len(qre) == len(main_question_value):
                                    main_question_value = self.df.at[id, qre]
                                    break    
                    main_question_value = self.convert_value(main_question_value)   
                    for question_condition, value_condition, related in conditions:
                        try:
                            if (pd.notnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() != "")):
                                current_answers = self.df.at[id, question_condition]
                                current_answers = self.convert_value(current_answers)
                                if pd.notnull(current_answers).any() or (isinstance(current_answers, str) and current_answers.strip() != "") or (isinstance(current_answers, list) and len(current_answers) == 1 and current_answers != ['NULL']):     
                                    main_question_value = self.diff_lists(main_question_value,current_answers)

                            if (pd.notnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() != "")):
                                current_answers = value_condition
                                current_answers = self.convert_value(current_answers)
                                if pd.notnull(current_answers).any() or (isinstance(current_answers, str) and current_answers.strip() != "") or (isinstance(current_answers, list) and len(current_answers) == 1 and current_answers != ['NULL']):  
                                    main_question_value = self.diff_lists(main_question_value,current_answers)
                        except Exception as e:
                            print(f"[Create_diff]: Lỗi xử lý điều kiện tại câu {question_condition}: {e}")

                    # Dùng funtion check_conditions được viết riêng để so sánh giá trị hiện tại và giá trị tại các câu điều kiện
                    if len(main_question_value) > 0:
                        self.df.at[id, Question_Check] += ','.join(map(str, main_question_value))

                    a =0    
            except Exception as e:
                print(f"[Create_diff]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results
    
    def Create_inter(self, df_excel, Question_Check, index):
        results = []
        if len(Question_Check)>0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist thông qua funtion get conditions
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)

                for id in self.df.index.tolist():
                    if id == "2001304:":
                        a=0
                    main_question_value = df_excel.loc[index, 'Current_Value']
                    if not isinstance(main_question_value, str):
                        main_question_value = int(df_excel.loc[index, 'Current_Value'])
                    else:
                        for qre in self.df.columns:
                            if qre in main_question_value:
                                if len(qre) == len(main_question_value):
                                    main_question_value = self.df.at[id, qre]
                                    break    

                    main_question_value = self.convert_value(main_question_value)
                    list_union = []
                    for question_condition, value_condition, related in conditions:
                        try:
                            list_union = []
                            if (pd.notnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() != "")):
                                current_answers = self.df.at[id, question_condition]
                                current_answers = self.convert_value(current_answers)
                                if pd.notnull(current_answers).any() or (isinstance(current_answers, str) and current_answers.strip() != "") or (isinstance(current_answers, list) and len(current_answers) == 1 and current_answers != ['NULL']):     
                                    list_union = set(list_union).union(set(current_answers))

                            if (pd.notnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() != "")):
                                current_answers = value_condition
                                current_answers = self.convert_value(current_answers)
                                if pd.notnull(current_answers).any() or (isinstance(current_answers, str) and current_answers.strip() != "") or (isinstance(current_answers, list) and len(current_answers) == 1 and current_answers != ['NULL']): 
                                    list_union = set(list_union).union(set(current_answers))
                        except Exception as e:
                            print(f"[Create_inter]: Lỗi xử lý điều kiện tại câu {question_condition}: {e}")
                    
                    main_question_value = set(main_question_value).intersection(set(list_union))

                    # Dùng funtion check_conditions được viết riêng để so sánh giá trị hiện tại và giá trị tại các câu điều kiện
                    if len(main_question_value) > 0:
                        self.df.at[id, Question_Check] += ','.join(map(str, main_question_value))
            except Exception as e:
                print(f"[Create_inter]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results
    
    def Create_get_iteration(self, df_excel, Question_Check, index):
        results = []
        if len(Question_Check) > 0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                main_question_value = self.convert_value(df_excel.loc[index, 'VALUE_1'])
                main_question = df_excel.loc[index, 'QRE_1']
                list_Question_loop = self.get_qre_loop(main_question)

                # Lặp qua từng câu hỏi trong list_Question_loop để kiểm tra giá trị của câu hỏi có nằm trong danh sách giá trị của câu điều kiện không
                for x in list_Question_loop:
                    for j, row in self.df[x].items():
                        row = self.convert_value(row)

                        arr_intersection = set(row).intersection(set(main_question_value))
                        if isinstance(df_excel.loc[index, 'VALUE_1'], str):
                            if "!" in df_excel.loc[index, 'VALUE_1']:
                                # Nếu có ký tự "!" trong giá trị điều kiện, kiểm tra xem giá trị của câu hỏi có nằm ngoài danh sách giá trị của câu điều kiện không
                                if len(arr_intersection) <= 0:
                                    numbers = ','.join(re.findall(r'\{_([^\}]+)\}', x))
                                    self.df.at[j,Question_Check] += numbers +","
                            else:
                                # Nếu không có ký tự "!" trong giá trị điều kiện, kiểm tra xem giá trị của câu hỏi có nằm trong danh sách giá trị của câu điều kiện không
                                if len(arr_intersection) > 0:
                                    numbers = ','.join(re.findall(r'\{_([^\}]+)\}', x))
                                    self.df.at[j,Question_Check] += numbers +"," 
                        else:
                            # Nếu giá trị điều kiện không phải là chuỗi, kiểm tra xem giá trị của câu hỏi có nằm trong danh sách giá trị của câu điều kiện không
                            if len(arr_intersection) > 0:
                                numbers = ','.join(re.findall(r'\{_([^\}]+)\}', x))
                                self.df.at[j,Question_Check] += numbers +","            
                
                self.df[Question_Check] = self.df[Question_Check].apply(
                    lambda x: ','.join(sorted(set(x.split(',')), key=lambda v: (v.isdigit(), int(v) if v.isdigit() else v)))
                )
                self.df[Question_Check] = self.df[Question_Check].apply(lambda x: x.rstrip(',')) #Loại bỏ dấu phẩy bị dư ở cuối khi thực hiện get iteration trên  
                
            except Exception as e:
                print(f"[Create_get_iteration]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results
    
    def Create_union(self, df_excel, Question_Check, index):
        results = []
        if len(Question_Check) > 0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                #Get các điều kiện và giá trị của điều kiện ở file excel checklist thông qua funtion get conditions
                conditions = []
                conditions = self.get_conditions_FULL(df_excel,index)

                for id in self.df.index.tolist():

                    main_question_value = df_excel.loc[index, 'Current_Value']
                    if not isinstance(main_question_value, str):
                        main_question_value = int(df_excel.loc[index, 'Current_Value'])
                    else:
                        for qre in self.df.columns:
                            if qre in main_question_value:
                                if len(qre) == len(main_question_value):
                                    main_question_value = self.df.at[id, qre]
                                    break    
                                    
                    main_question_value = self.convert_value(main_question_value)

                    if id == "2055607":
                        a=0
                    for question_condition, value_condition, related in conditions:
                        try:
                            if (pd.notnull(question_condition) or (isinstance(question_condition, str) and question_condition.strip() != "")):
                                current_answers = self.df.at[id, question_condition]
                                current_answers = self.convert_value(current_answers)   
                                if pd.notnull(current_answers).any() or (isinstance(current_answers, str) and current_answers.strip() != "") or (isinstance(current_answers, list) and len(current_answers) == 1 and current_answers != ['NULL']):     
                                    a=0                                                           
                                    main_question_value = set(main_question_value).union(set(current_answers))
                                    # main_question_value = main_question_value.unique()
                            if (pd.notnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() != "")):
                                current_answers = value_condition
                                current_answers = self.convert_value(current_answers)   
                                if pd.notnull(current_answers).any() or (isinstance(current_answers, str) and current_answers.strip() != "") or (isinstance(current_answers, list) and len(current_answers) == 1 and current_answers != ['NULL']):                                                                
                                    main_question_value = set(main_question_value).union(set(current_answers))                                  
                                    # main_question_value = main_question_value.unique()
                                    
                        except Exception as e:
                            print(f"[Create_union]: Lỗi xử lý điều kiện tại câu {question_condition}: {e}")

                    if len(main_question_value) > 0:
                        self.df.at[id, Question_Check] += ','.join(map(str, main_question_value))
                
            except Exception as e:
                print(f"[Create_union]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results

    def Create_AnswerCount(self, df_excel, Question_Check, index):
        results = []
        if len(Question_Check) > 0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                for id in self.df.index.tolist():

                    main_question_value = self.df.at[id, df_excel.loc[index, 'QRE_1']]
                    main_question_value = self.convert_value(main_question_value)
                    if len(main_question_value) > 0:
                        self.df.at[id, Question_Check] = len(main_question_value)

                    a =0 
            except Exception as e:
                print(f"[Create_AnswerCount]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results

    def Create_Compare_num(self, df_excel, Question_Check, index):
        results = []
        if len(Question_Check) > 0:
            try:
                #Thêm câu tạo mới này vào dataframe của df, kiểm tra nếu chưa tạo sẵn trong dataframe mới thực hiện
                if Question_Check not in self.df:
                    self.df[Question_Check] = ""

                for id in self.df.index.tolist():

                    if id == "2000817":
                        a=0
                    current_answers = df_excel.loc[index, 'QRE_1']

                    if (pd.notnull(current_answers) or (isinstance(current_answers, str) and current_answers.strip() != "")):
                        current_answers = self.df.at[id, current_answers]
                    else:
                        current_answers = df_excel.loc[index, 'VALUE_1']

                    current_answers = int(current_answers)

                    value_condition = df_excel.loc[index, 'QRE_2']
                    if (pd.notnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() != "")):
                        value_condition = self.df.at[id, value_condition]
                    else:
                        value_condition = df_excel.loc[index, 'VALUE_2']

                    value_condition = int(value_condition)

                    if current_answers > value_condition:
                        self.df.at[id, Question_Check] = ">"
                    elif current_answers < value_condition:
                        self.df.at[id, Question_Check] = "<" 
                    else:
                        self.df.at[id, Question_Check] = "="
            except Exception as e:
                print(f"[Create_Compare_num]: Lỗi xử lý điều kiện tại câu {Question_Check}: {e}")
        return results
    
    def Valcheck_askall(self, df_excel):
        results = []       
        #Filter các câu là hỏi tất cả trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_askall"), "Question Check"]
        Question_Check = self.get_qre_loop(Question_Check)
        if not Question_Check.empty:
            for i in Question_Check:
                try:
                    missing_data = self.df.loc[
                            self.df[i].isnull() | 
                            (self.df[i].astype(str).str.strip() == "") | 
                            (self.df[i].astype(str).str.upper() == "NULL"),
                            i
                    ]
                    if not missing_data.empty:                 
                        for id in missing_data.index.tolist():
                            result_str = f"{id}: {i}: Missing DATA"
                            print(result_str)
                            results.append(result_str)
                except Exception as e:
                    print(f"[Valcheck_askall]: Lỗi xử lý điều kiện tại câu {i}: {e}")                                 
        return results    
    
    def Valcheck_Selected(self, df_excel):
        results = []       
        #Filter các câu là logic code trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_Selected"), "Question Check"]
        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    qre_name = self.get_qre_loop(pd.Series([qre_name]))
                    for i in qre_name.tolist():
                        data_check = self.df.loc[
                            self.df[i].notnull() & 
                            (self.df[i].astype(str).str.strip() != "") & 
                            (self.df[i].astype(str).str.upper() != "NULL"),
                            i
                        ]
                        if not data_check.empty:
                            for id in data_check.index.tolist():       

                                if id == "2029213":
                                    a=0                     
                                Current_Value = df_excel.loc[index, 'Current_Value']
                                Name_condition = None
                                if not isinstance(Current_Value, str):
                                    Current_Value = int(df_excel.loc[index, 'Current_Value'])
                                else:
                                    for qre in self.df.columns:
                                        if "[..]" in Current_Value:
                                            iteration = ','.join(re.findall(r'\{_([^\}]+)\}', i))  
                                            Name_condition = Current_Value.replace("[..]", "[{_" + iteration + "}]")
                                            Current_Value = self.df.at[id, Name_condition]
                                            break
                                        if "!" in Current_Value:
                                            if qre in Current_Value.replace("!",""):
                                                if len(qre) == len(Current_Value.replace("!","")):
                                                    Current_Value = self.df.at[id, qre]
                                                    break                                
                                        else:
                                            if qre in Current_Value:
                                                if len(qre) == len(Current_Value):
                                                    Current_Value = self.df.at[id, qre]
                                Current_Value = self.convert_value(Current_Value)
                                current_answers = []
                                current_answers = self.df.at[id, i]
                                current_answers = self.convert_value(current_answers)   

                                arr_intersection = set(current_answers).intersection(set(Current_Value))

                                condition_result = False
                                if isinstance(df_excel.loc[index, 'Current_Value'], str):
                                    if "!" in df_excel.loc[index, 'Current_Value']:
                                        if len(arr_intersection) <= 0:
                                            condition_result = True
                                    else:
                                        if len(arr_intersection) > 0:
                                            condition_result = True
                                else:
                                    if len(arr_intersection) > 0:
                                        condition_result = True

                                if condition_result == False:
                                    if Name_condition is None:
                                        result_str = f"{id}: {i}: {current_answers} không có trong list code {df_excel.loc[index, 'Current_Value']}"
                                    else:
                                        result_str = f"{id}: {i}: {current_answers} không có trong list code {Name_condition}: {df_excel.loc[index, 'Current_Value']}"
                                    print(result_str)
                                    results.append(result_str)  
                except Exception as e:
                    print(f"[Valcheck_Selected]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")                                         
        return results
    
    def Valcheck_num(self, df_excel):
        results = []       
        #Filter các câu là logic code trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_num"), "Question Check"]
        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    qre_name = self.get_qre_loop(pd.Series([qre_name]))
                    for i in qre_name.tolist():
                        data_check = self.df.loc[
                            self.df[i].notnull() & 
                            (self.df[i].astype(str).str.strip() != "") & 
                            (self.df[i].astype(str).str.upper() != "NULL"),
                            i
                        ]
                        if not data_check.empty:
                            for id in data_check.index.tolist():
                                current_answers = int(self.df.at[id, i])
                                Value = df_excel.loc[index, 'Current_Value'].split("..")

                                if len(Value) > 2:
                                    Min = Value[0]
                                    Max = Value[-1]
                                else:
                                    if Value[0]:
                                        Min = Value[0]
                                    else:
                                        Min = None
                                    if Value[-1]:
                                        Max = Value[-1]
                                    else:
                                        Max = None

                                if Min != None:        
                                    try: #Kiểm tra thử trong phần điều kiện có khai báo giá trị Min là số hay tên câu hỏi
                                        Min = int(Min)
                                    except ValueError:
                                        Qre = Min
                                        Min = int(self.df.at[id, Qre]) #Get giá trị Min tại ID câu cụ thể
                                    if int(current_answers) < Min:
                                        result_str = f"{id}: {i}: {current_answers} nhỏ hơn {Min}"
                                        print(result_str)
                                        results.append(result_str)      
                            
                                if Max != None:
                                    try:
                                        Max = int(Max)
                                    except ValueError:
                                        Qre = Max
                                        Max = int(self.df.at[id, Qre])
                                    if int(current_answers) > Max:
                                        result_str = f"{id}: {i}: {current_answers} lớn hơn {Max}"
                                        print(result_str)
                                        results.append(result_str)          
                except Exception as e:
                    print(f"[Valcheck_num]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")                                         
        return results
    
    def Valcheck_Autocode_Logic(self, df_excel):
        results = []         
        #Filter các câu là logic question trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_Autocode_Logic"), "Question Check"]   

        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    #Get current_value ở file excel checklist
                    Current_Value = df_excel.loc[index, 'Current_Value']
                    check_value = False
                    if isinstance(Current_Value, str):
                        if "!" in Current_Value:
                            check_value = True
                            break

                    Current_Value = self.convert_value(Current_Value)

                    qre_name = self.get_qre_loop(pd.Series([qre_name]))
                    for i in qre_name.tolist():
                        #Filter DATA tại câu cần check
                        data_check = self.df.loc[
                            self.df[i].notnull() & 
                            (self.df[i].astype(str).str.strip() != "") & 
                            (self.df[i].astype(str).str.upper() != "NULL"),
                            i
                        ]
                        if not data_check.empty:
                            for id in data_check.index.tolist():

                                if id == "2055281":
                                    a =0
                                #Get các điều kiện và giá trị của điều kiện ở file excel checklist
                                conditions = []
                                conditions = self.get_conditions(df_excel, index)                                    
                                iteration = ','.join(re.findall(r'\{_([^\}]+)\}', i))    
                                for j in range(len(conditions)):
                                    if "[..]" in conditions[j][0]:
                                        # Tạo tuple mới với phần tử đầu đã thay đổi
                                        new_condition = (
                                            conditions[j][0].replace("[..]", "[{_" + iteration + "}]"),
                                            conditions[j][1],
                                            conditions[j][2]
                                        )
                                        conditions[j] = new_condition
                                 # Dùng funtion check_conditions được viết riêng để so sánh giá trị hiện tại và giá trị tại các câu điều kiện
                                check_condition = self.check_conditions(id, conditions)

                                current_answers = self.df.at[id, i]

                                current_answers = self.convert_value(current_answers)

                                arr_intersection = set(current_answers).intersection(set(Current_Value))

                                current_result = False
                                if check_value == True:
                                    if len(arr_intersection) <= 0:
                                        current_result = True
                                else:
                                    if len(arr_intersection) > 0:
                                        current_result = True
                                        
                                #So sánh kết quả check_condition với current_result
                                if (check_condition == True and  current_result == False):
                                    result_str = f"{id}: {i}: {current_answers} không có trong list code {conditions}"
                                    print(result_str)
                                    results.append(result_str)                                                                        
                                  
                except Exception as e:
                    print(f"[Valcheck_Autocode_Logic]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")                        
        return results    

    def Valcheck_equal(self, df_excel):
        results_list = []        
        #Filter các câu là logic question trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_equal"), "Question Check"] 
        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    qre_name = self.get_qre_loop(pd.Series([qre_name]))
                    for i in qre_name.tolist():
                        #Filter DATA tại câu cần check
                        data_check = self.df[i]
                        if not data_check.empty:
                            for id in data_check.index.tolist():
                                if id == "R_937BqmzQ20PHpfw":
                                    a=0
                                current_answers = self.df.at[id, i]
                                current_answers = self.convert_value(current_answers)

                                value_condition = df_excel.loc[index, 'QRE_1']
                                Name_condition = None
                                if (pd.notnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() != "")):
                                    if "[..]" in value_condition:
                                        iteration = ','.join(re.findall(r'\{_([^\}]+)\}', i))  
                                        Name_condition = value_condition.replace("[..]", "[{_" + iteration + "}]") 
                                        value_condition = self.df.at[id, Name_condition]
                                    else:  
                                        value_condition = self.df.at[id, value_condition]
                                else:
                                    value_condition = df_excel.loc[index, 'VALUE_1']

                                value_condition = self.convert_value(value_condition)
                                if set(current_answers) != set(value_condition):
                                    if Name_condition == None:
                                        results = f"{id}: {i}: {current_answers}  Không giống {value_condition}"
                                    else:
                                        results = f"{id}: {i}: {current_answers}  Không giống {Name_condition}: {value_condition}"
                                    print(results)
                                    results_list.append(results)                                                          
                except Exception as e:
                    print(f"[Valcheck_equal]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")    
        return results_list  
    
    def Valcheck_Not_equal(self, df_excel):
        results = []              
        #Filter các câu là logic question trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_Not_equal"), "Question Check"] 

        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    qre_name = self.get_qre_loop(pd.Series([qre_name]))
                    for i in qre_name.tolist():
                        #Filter DATA tại câu cần check
                        data_check = self.df[i]
                        if not data_check.empty:
                            for id in data_check.index.tolist():
                                if id == "2035833":
                                    a=0                            
                                current_answers = self.df.at[id, i]
                                current_answers = self.convert_value(current_answers)

                                value_condition = df_excel.loc[index, 'QRE_1']
                                Name_condition = None
                                if (pd.notnull(value_condition) or (isinstance(value_condition, str) and value_condition.strip() != "")):
                                    if "[..]" in value_condition:
                                        iteration = ','.join(re.findall(r'\{_([^\}]+)\}', i))  
                                        Name_condition = value_condition.replace("[..]", "[{_" + iteration + "}]") 
                                        value_condition = self.df.at[id, Name_condition]   
                                    else:  
                                        value_condition = self.df.at[id, value_condition]                                            
                                else:
                                    value_condition = df_excel.loc[index, 'VALUE_1']
                                value_condition = self.convert_value(value_condition)
                                arr_intersection = set(current_answers).intersection(set(value_condition))
                                if len(arr_intersection)>0:
                                    result_str = f"{id}: {i}: {current_answers}  có code giống {Name_condition}: {value_condition}"
                                    print(result_str)
                                    results.append(result_str)                                                                                             
                except Exception as e:
                    print(f"[Valcheck_Not_equal]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")   
        return results  
    
    def Valcheck_sum(self, df_excel):
        results = []         
        #Filter các câu là logic question trong file check list
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_sum"), "Question Check"] 

        if not Question_Check.empty:
            for index, i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    list_Question_loop = self.get_qre_loop(pd.Series([i]))
                    if len(list_Question_loop) > 0:  
                        self.df[i] = self.df[list_Question_loop].sum(axis=1, skipna=True, min_count=1) 
                    
                    Data_check = self.df.loc[
                        self.df[i].notnull() & 
                        (self.df[i].astype(str).str.strip() != "") & 
                        (self.df[i].astype(str).str.upper() != "NULL"),
                        i
                    ]

                    for id,Answer in zip(Data_check.index.tolist(),Data_check.tolist()):  
                        if id == "2032843":
                            a=0
                        if not Data_check.empty:  
                            try:
                                sum_value = int(df_excel.loc[index, 'VALUE_1'])
                                if sum_value != int(Answer):
                                    result_str = f"{id}: {Answer}: không bằng {df_excel.loc[index, 'VALUE_1']}"
                                    print(result_str)
                                    results.append(result_str)  
                            except:
                                sum_value = int(self.df.at[id, df_excel.loc[index, 'QRE_1']])
                                
                                if sum_value != int(Answer):
                                    result_str = f"{id}: {i}: {Answer} không bằng {df_excel.loc[index, 'QRE_1']}: {sum_value}"
                                    print(result_str)
                                    results.append(result_str) 
                except Exception as e:
                    print(f"[Valcheck_sum]: Lỗi xử lý điều kiện tại câu {i}: {e}")                       
        return results    

    def Valcheck_initialize(self, df_excel):
        results = []      
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_initialize"), "Question Check"] 

        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    
                    DK_Check = False
                    DK = []
                    if pd.notnull(df_excel.loc[index, 'VALUE_1']) and df_excel.loc[index, 'VALUE_1'] != '':
                        for x in df_excel.loc[index, 'VALUE_1'].split(','):
                            if "DK" in x.upper(): #Kiểm tra trong câu được filter có code loại trừ 98, 99 không thông qua "DK"
                                DK.append(int(x.replace("DK_","")))
                                DK_Check = True

                    qre_name = self.get_qre_loop(pd.Series([qre_name]))
                    for i in qre_name.tolist():
                        Data_check = self.df.loc[
                            self.df[i].notnull() & 
                            (self.df[i].astype(str).str.strip() != "") & 
                            (self.df[i].astype(str).str.upper() != "NULL"),
                            i
                        ]
                        
                        if not Data_check.empty:      
                            for id,Answer in zip(Data_check.index.tolist(),Data_check.tolist()):  
                                #get giá trị của câu filter và đưa ra thành list
                                current_answers = Answer
                                current_answers = self.convert_value(current_answers)
                                if id == "2055281":
                                    a=0
                                question_condition = df_excel.loc[index, 'QRE_1']
                                Name_condition = None
                                for qre in self.df.columns:
                                    if "[..]" in question_condition:
                                        iteration = ','.join(re.findall(r'\{_([^\}]+)\}', i))  
                                        Name_condition = question_condition.replace("[..]", "[{_" + iteration + "}]")
                                        question_condition = self.df.at[id, Name_condition]
                                        break
                                    else:
                                        if qre in question_condition:
                                            if len(qre) == len(question_condition):
                                                question_condition = self.df.at[id, qre]
                                
                                question_condition = self.convert_value(question_condition)
                                arr_intersection = set(current_answers).intersection(set(question_condition))
                                if DK_Check == True: #Xử lý các câu có note DK trong điều kiện
                                    if not any(item in current_answers for item in DK) and not any(item in question_condition for item in DK):
                                        if len(arr_intersection) == 0:
                                            if Name_condition == None:
                                                result_str = f"{id}: {i}: {current_answers} KHÔNG filter {df_excel.loc[index, 'QRE_1']} : {question_condition}"
                                            else:
                                                result_str = f"{id}: {i}: {current_answers} KHÔNG filter {Name_condition}: {question_condition}"
                                            print(result_str)
                                            results.append(result_str)
                                        for item in current_answers:
                                            if item not in question_condition: #Kiểm tra xem có code nào không thuộc trong list code filter không
                                                if Name_condition == None:
                                                    result_str = f"{id}: {i}: {current_answers} KHÔNG filter {df_excel.loc[index, 'QRE_1']} : {question_condition}"
                                                else:
                                                    result_str = f"{id}: {i}: {current_answers} KHÔNG filter {Name_condition}: {question_condition}"
                                                print(result_str)
                                                results.append(result_str)                                        
                                else:
                                    if len(arr_intersection) == 0:
                                        if Name_condition == None:
                                            result_str = f"{id}: {i}: {current_answers} KHÔNG filter {df_excel.loc[index, 'QRE_1']} : {question_condition}"
                                        else:
                                            result_str = f"{id}: {i}: {current_answers} KHÔNG filter {Name_condition}: {question_condition}"
                                        print(result_str)
                                        results.append(result_str)
                                    for item in current_answers:
                                        if item not in question_condition: #Kiểm tra xem có code nào không thuộc trong list code filter không
                                            if Name_condition == None:
                                                result_str = f"{id}: {i}: {current_answers} KHÔNG filter {df_excel.loc[index, 'QRE_1']} : {question_condition}"
                                            else:
                                                result_str = f"{id}: {i}: {current_answers} KHÔNG filter {Name_condition}: {question_condition}"
                                            print(result_str)
                                            results.append(result_str)   
                except Exception as e:
                    print(f"[Valcheck_initialize]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")                                 
        return results  
    
    def Valcheck_filterbycount(self, df_excel):
        results = []      
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_filterbycount"), "Question Check"] 

        if not Question_Check.empty:
            for index, i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    list_Question_loop = self.get_qre_loop(pd.Series([i]))
                    
                    self.df[i] = ""
                    for x in list_Question_loop:
                        for j, row in self.df[x].items():
                            if row is not None:
                                numbers = ','.join(re.findall(r'\{_(\d+)\}', x))
                                self.df.at[j,i] += numbers +"," #tạo ra 1 cột mới lấy giá trị iteration thỏa điều kiện nếu có giá trị khác null                
                    
                    self.df[i] = self.df[i].apply(lambda x: x.rstrip(',')) #Loại bỏ dấu phẩy bị dư ở cuối khi thực hiện get iteration trên
                    #Lọc data cho câu filter có giá trị khác null
                    Data_check = self.df.loc[
                        self.df[df_excel.loc[index, 'QRE_1']].notnull() & 
                        (self.df[df_excel.loc[index, 'QRE_1']].astype(str).str.strip() != "") & 
                        (self.df[df_excel.loc[index, 'QRE_1']].astype(str).str.upper() != "NULL"),
                        df_excel.loc[index, 'QRE_1']
                    ]
                    if not Data_check.empty:      
                        for id,Answer in zip(Data_check.index.tolist(),Data_check.tolist()):  
                            if id == "2000817":
                                a=0
                            current_answers = self.df.at[id, i]
                            current_answers = self.convert_value(current_answers) 

                            Answer_list = Answer
                            MAX = len(list_Question_loop)

                            if len(current_answers) == MAX: #Kiểm tra số code được chọn ở filter lớn hơn số lượng iteration ở câu được filter
                                if int(Answer_list) < MAX: #Kiểm tra số lượng iteration khi kiểm tra null ở trên có bằng tương ứng bẳng số iteration trong data không
                                    result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                    print(result_str)
                                    results.append(result_str)  
                            else:
                                if len(current_answers) != int(Answer_list):     
                                    result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                    print(result_str)
                                    results.append(result_str)   

                    #Lọc data cho câu được filter có giá trị khác null
                    Data_check = self.df.loc[
                        self.df[i].isnull() | 
                        (self.df[i].astype(str).str.strip() == "") | 
                        (self.df[i].astype(str).str.upper() == "NULL"),
                        i
                    ]    
                    if not Data_check.empty:      
                        for id,Answer in zip(Data_check.index.tolist(),Data_check.tolist()):  

                            current_answers = Answer
                            current_answers = self.convert_value(current_answers)                 
                            Answer_list = self.df.at[id, df_excel.loc[index, 'QRE_1']]

                            MAX = len(list_Question_loop)

                            if len(current_answers) == MAX: #Kiểm tra số code được chọn ở filter lớn hơn số lượng iteration ở câu được filter
                                if int(Answer_list) < MAX: #Kiểm tra số lượng iteration khi kiểm tra null ở trên có bằng tương ứng bẳng số iteration trong data không
                                    result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                    print(result_str)
                                    results.append(result_str)  
                            else:
                                if len(current_answers) != int(Answer_list):     
                                    result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                    print(result_str)
                                    results.append(result_str)   
                except Exception as e:
                    print(f"[Valcheck_filterbycount]: Lỗi xử lý điều kiện tại câu {i}: {e} với ID = {id}")   
            return results 

    def Valcheck_filterbycat(self, df_excel):
        results = []   
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_filterbycat"), "Question Check"] 

        if not Question_Check.empty:
            for index, i in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    list_Question_loop = self.get_qre_loop(pd.Series([i]))
                    self.df[i] = ""
                    for x in list_Question_loop:
                        for j, row in self.df[x].items():
                            if j == "2049767":
                                a=0
                            if pd.notnull(row):
                                matches = re.findall(r'\{_([^\}]+)\}', x)
                                numbers = matches[-1] if matches else ""
                                self.df.at[j,i] += numbers +"," #tạo ra 1 cột mới lấy giá trị iteration thỏa điều kiện nếu có giá trị khác null                
                    self.df[i] = self.df[i].apply(
                        lambda x: ','.join(sorted(set(x.split(',')), key=lambda v: (v.isdigit(), int(v) if v.isdigit() else v)))
                    )                 
                    self.df[i] = self.df[i].apply(lambda x: x.rstrip(',')) #Loại bỏ dấu phẩy bị dư ở cuối khi thực hiện get iteration trên
                    List_iteration_fix_for_filter = []
                    DK_Check = False
                    DK = []
                    value_check = df_excel.loc[index, 'VALUE_1']
                    if pd.notnull(value_check) or (isinstance(value_check, str) and value_check.strip() != ""):
                        for item in df_excel.loc[index, 'VALUE_1'].split(","):
                            try:
                                num = int(item)
                                List_iteration_fix_for_filter.append(num)                          
                            except ValueError:
                                if "DK" not in item:
                                    List_iteration_fix_for_filter.append(item.upper())       
                                else:
                                    DK.append(int(item.replace("DK_","")))
                                    DK_Check = True                                    

                    #Lọc data cho câu filter có giá trị khác null
                    Data_check = self.df.loc[
                        self.df[df_excel.loc[index, 'QRE_1']].notnull() & 
                        (self.df[df_excel.loc[index, 'QRE_1']].astype(str).str.strip() != "") & 
                        (self.df[df_excel.loc[index, 'QRE_1']].astype(str).str.upper() != "NULL"),
                        df_excel.loc[index, 'QRE_1']
                    ]
                        
                    if not Data_check.empty:  
                            
                        for id,Answer in zip(Data_check.index.tolist(),Data_check.tolist()):  
                            current_answers = self.df.at[id, i]
                            current_answers = self.convert_value(current_answers)
                            if id == "2049767":
                                a=0
                            Answer_list = Answer
                            Answer_list = self.convert_value(Answer_list)
                            if DK_Check == True:
                                if not any(item in Answer_list for item in DK):
                                    if len(List_iteration_fix_for_filter) > 0:
                                        arr_intersection = set(List_iteration_fix_for_filter).intersection(set(Answer_list))
                                        if len(current_answers) != len(arr_intersection):
                                            result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {arr_intersection}"
                                            print(result_str)
                                            results.append(result_str)
                                        for item in current_answers:
                                            if item not in arr_intersection:
                                                result_str = f"{id}: {i}: have iteration {item} not in {df_excel.loc[index, 'QRE_1']}: {arr_intersection}"
                                                print(result_str)
                                                results.append(result_str)     
                                    else:
                                        if len(current_answers) != len(Answer_list): #Kiểm tra số lượng iteration khác null tại câu được filter và số lượng code tại câu filter
                                            result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {len(Answer_list)}"
                                            print(result_str)
                                            results.append(result_str)    
                                        for item in current_answers:
                                            if item not in Answer_list:
                                                result_str = f"{id}: {i}: have iteration {item} not in {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                                print(result_str)
                                                results.append(result_str)                                          
                                                            
                            else:
                                if len(List_iteration_fix_for_filter) > 0:
                                    arr_intersection = set(List_iteration_fix_for_filter).intersection(set(Answer_list))
                                    if len(arr_intersection) == 0 and current_answers == ['NULL']:
                                        pass
                                    else:
                                        if len(current_answers) != len(arr_intersection):
                                            result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {arr_intersection}"
                                            print(result_str)
                                            results.append(result_str)
                                        for item in current_answers:
                                            if item not in arr_intersection:
                                                result_str = f"{id}: {i}: have iteration {item} not in {df_excel.loc[index, 'QRE_1']}: {arr_intersection}"
                                                print(result_str)
                                                results.append(result_str)     
                                else:
                                    if len(current_answers) != len(Answer_list):
                                        result_str = f"{id}: {i}: {current_answers} not filter {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                        print(result_str)
                                        results.append(result_str)
                                    for item in current_answers:
                                        if item not in Answer_list:
                                            result_str = f"{id}: {i}: have iteration {item} not in {df_excel.loc[index, 'QRE_1']}: {Answer_list}"
                                            print(result_str)
                                            results.append(result_str)          

                    #### Viết thêm trường hợp check data bị dư với self.df[question] có dữ liệu như bên data_check lại có data
                    miss_data = self.df.loc[
                        self.df[i].isnull() | 
                        (self.df[i].astype(str).str.strip() == "") | 
                        (self.df[i].astype(str).str.upper() == "NULL"),
                        i
                    ]                    
                    if not miss_data.empty:
                        for id,Answer in zip(miss_data.index.tolist(),miss_data.tolist()): 
                            Answer_list = self.convert_value(Answer)
                            Answer_filter = self.convert_value(self.df.at[id, df_excel.loc[index, 'QRE_1']])
                            for item in Answer_list:
                                if item not in Answer_filter:
                                    result_str = f"{id}: {i}: iteration {item} have DATA, BUT {item} not in {df_excel.loc[index, 'QRE_1']}: {Answer_filter}"
                                    print(result_str)
                                    results.append(result_str)  
                except Exception as e:
                    print(f"[Valcheck_filterbycat]: Lỗi xử lý điều kiện tại câu {i}: {e}")   
        return results          
    
    def Valcheck_askbyroute(self, df_excel):

        results = []  
        Question_Check = df_excel.loc[(df_excel["Funtion"] == "Valcheck_askbyroute"), "Question Check"] 

        if not Question_Check.empty:
            for index, qre_name in zip(Question_Check.index.tolist(),Question_Check.tolist()):
                try:
                    qre_name = self.get_qre_loop(pd.Series([qre_name]))

                    for i in qre_name.tolist():
                        #Get các điều kiện và giá trị của điều kiện ở file excel checklist
                        conditions = []
                        conditions = self.get_conditions(df_excel, index)

                        iteration = ','.join(re.findall(r'\{_([^\}]+)\}', i))    
                        for j in range(len(conditions)):
                            if "[..]" in conditions[j][0]:
                                # Tạo tuple mới với phần tử đầu đã thay đổi
                                new_condition = (
                                    conditions[j][0].replace("[..]", "[{_" + iteration + "}]"),
                                    conditions[j][1],
                                    conditions[j][2]
                                )
                                conditions[j] = new_condition

                        # Lọc data missing: là NaN, chuỗi rỗng hoặc "NULL"
                        missing_data = self.df.loc[
                            self.df[i].isnull() | 
                            (self.df[i].astype(str).str.strip() == "") | 
                            (self.df[i].astype(str).str.upper() == "NULL"),
                            i
                        ]
                        # Lọc data có giá trị thực sự (không phải NaN, không rỗng, không "NULL")
                        has_data = self.df.loc[
                            self.df[i].notnull() & 
                            (self.df[i].astype(str).str.strip() != "") & 
                            (self.df[i].astype(str).str.upper() != "NULL"),
                            i
                        ]
                        if not missing_data.empty:                 
                            for id in missing_data.index.tolist():
                                if id == "2032655":
                                    a = 0                              
                                # Dùng funtion check_conditions được viết riêng để so sánh giá trị hiện tại và giá trị tại các câu điều kiện          
                                check_condition = self.check_conditions(id, conditions)
    
                                if check_condition:
                                    result_str = f"{id}: {i}: Missing DATA - Khi thỏa điều kiện: {conditions}"
                                    print(result_str)
                                    results.append(result_str)         

                        if not has_data.empty:                 
                            for id in has_data.index.tolist():
                                check_condition = self.check_conditions(id, conditions)
                                if check_condition == False:
                                    result_str = f"{id}: {i}: DƯ DATA - Khi không thỏa điều kiện: {conditions}"
                                    print(result_str)
                                    results.append(result_str)  
   
                except Exception as e:
                    print(f"[Valcheck_askbyroute]: Lỗi xử lý điều kiện tại câu {qre_name}: {e}")   
            return results
        
