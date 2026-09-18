# -*- coding: utf-8 -*-
"""
Created on Mon Jul 15 13:36:13 2024
 
@author: R.Gayathri
"""
 
"""
Modified on Friday, April 23 17:08
(For extending the existing logics to "SimilationValues" sheet)
@author: R.Gayathri
"""
 
#%% Import Necessary Libraries
import os
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential
import io
import re
import requests
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from pathlib import Path
from tqdm import tqdm
import logging
from datetime import datetime
from time import time
import sys
import warnings
import traceback
import xlwings as xw
 
warnings.filterwarnings("ignore")
 
SHEETS_TO_VALIDATE = ["SimulationQty", "SimulationValues"]  # @Gayathri, validation for both the sheets
 
try:
    def import_arg_variables():
        if len(sys.argv)>1:
            local_folder_path = sys.argv[1]
        else:
            local_folder_path = ""
        return(local_folder_path)  
       
    local_folder_path = import_arg_variables()
    print("Hi! You have selected the following folder : ",local_folder_path,"\n")
except Exception as e:
    print(e)
 
BASE_DIR = Path(__file__).resolve().parent.parent
os.chdir(BASE_DIR)
print(BASE_DIR)
 
today = datetime.now()
today = today.strftime("%Y-%m-%d %H:%M:%S")
today = today.replace(".","-")
today = today.replace(":","-")
 
start_time = time()
error_code=0
 
logging.shutdown()
 
validation_report_name_date = 'validation_report_log_'+str(today)+'.log'
validation_report_name = os.path.join(BASE_DIR, 'Log Files', validation_report_name_date)
 
if os.path.exists(validation_report_name):
    os.remove(validation_report_name)
 
logging.basicConfig(filename=validation_report_name, level=logging.INFO,filemode='w',format='%(process)d-%(levelname)s-%(asctime)s-%(message)s')
validation_reporting_log = logging.getLogger(validation_report_name)
 
try:
    data_health_ui_path = os.path.join(BASE_DIR,"data_health_validator.xlsm")
 
    wb = xw.Book(data_health_ui_path)
    sheet = wb.sheets['run_local_files']
    data = sheet.range('G6:H10').options(pd.DataFrame, header=1).value
   
    start_month = int(sheet['G7'].value)
    start_year = int(sheet['H7'].value)
    end_month = int(sheet['G10'].value)
    end_year = int(sheet['H10'].value)
   
    validation_reporting_log.info(f"User Inputs read successfully as : {start_month}-{start_year} and {end_month}-{end_year}\n")
   
    def generate_folder_paths(start_month, start_year, end_month, end_year, local_folder_path):
        start_date = datetime(start_year, start_month, 1)
        end_date = datetime(end_year, end_month, 1)
        filtered_files = []
       
        for file_name in os.listdir(local_folder_path):
            if (file_name.endswith('.xlsx') and not file_name.startswith('~$') and ('siopwhatIf' in file_name or 'siop_whatIf_' in file_name)):
                try:  
                    date_str = file_name.split('siopwhatIf')[1][:6]
                except:
                    date_str = file_name.split('siop_whatIf_')[1][:6]
                file_date = datetime.strptime(date_str, '%Y%m')
               
                if start_date <= file_date <= end_date:
                    filtered_files.append(os.path.join(local_folder_path, file_name))
       
        return filtered_files
 
    def validate_file(file_paths,country_df,local_folder_path):
       
        print('''==========================================\nStarting the process for File Validation\n==========================================''')
        validation_reporting_log.info('''\n==========================================\nStarting the process for File Validation\n==========================================''')
       
        try:
            country_names = country_df['country'].tolist()
            records=[]
            file_presence_list = []
           
            if(len(file_paths)<=0):
                file_presence_list.append({'Foldername':local_folder_path,'Folder Exists':'No',"No. of Files Found":0,"Files Found":"-"})
                print("Folder not Found or No Files in Folder\n")
            else:
                file_iterations = tqdm(file_paths, desc=f"Validating files in the given folder")
                file_presence_list.append({'Foldername':local_folder_path,'Folder Exists':'Yes',"No. of Files Found":len(file_paths),"Files Found":file_paths}) 
               
                for file in file_iterations:
                    filename = os.path.basename(file)
                    file_iterations.set_description("Processing %s" % filename)
                   
                    is_excel = filename.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb'))
                    is_excel = "Excel Format" if is_excel else "Excel Format not found"
                   
                    starts_with_country = any(filename.startswith(country) for country in country_names)
                    matches_pattern_1 = bool(re.search(r'siopwhatIf\d{8}', filename))
                    matches_pattern_2 = bool(re.search(r'siop_whatIf_\d{8}', filename))
                    matches_pattern = matches_pattern_1 or matches_pattern_2
                   
                    flag_naming_convention = "Match" if (starts_with_country and matches_pattern) else "Does not match"
                   
                    if is_excel=="Excel Format":
                        file_url = os.path.join(local_folder_path, filename)
                        try:
                             sheet_names_list = pd.ExcelFile(file_url, engine="openpyxl").sheet_names
                        except Exception as e:
                             print(f"Skipping invalid file: {filename}")
                             validation_reporting_log.warning(f"Skipped file {filename}: {e}")
                             continue
                        sheet_status = {
                            sheet: ("Found" if sheet in sheet_names_list else "Not Found")
                            for sheet in SHEETS_TO_VALIDATE
                        }
                    else:
                        sheet_status = {}
                        file_url = os.path.join(local_folder_path, filename)
                   
                    records.append({
                        'Filename': filename,
                        'Foldername': local_folder_path,
                        'Is_Excel': is_excel,
                        'File_URL': file_url,
                        'Naming_Convention_Flag': flag_naming_convention,
                        **{f"{sheet}_Sheet_Exists": status for sheet, status in sheet_status.items()}
                    })
           
            df = pd.DataFrame(records)
            file_presence_df = pd.DataFrame(file_presence_list)
 
            def split_after_siopwhatIf(filename):
                if 'siopwhatIf' in filename:
                    split_point = filename.find('siopwhatIf') + len('siopwhatIf') + 6
                else:
                    split_point = filename.find('siop_whatIf_') + len('siop_whatIf_') + 6
                return filename[:split_point]
           
            df['prefix'] = df['Filename'].apply(split_after_siopwhatIf)
            df['Similar_File'] = df.duplicated(subset=['prefix', 'Foldername'], keep=False).replace({True: 'Duplicate', False: '-'})
            df.drop(columns=['prefix'], inplace=True)
           
            print("File Validation Completed\n")
            validation_reporting_log.info("File Validation Completed\n")
           
            return df, file_presence_df
       
        except Exception as e:
            print("Error occured while validating file: ", e)
            traceback.print_exc()
            validation_reporting_log.error(f"Error occured while validating file: {e}")
            global error_code
            error_code=1
            return pd.DataFrame(), pd.DataFrame()
 
    def validate_sheet(file_df, expected_columns_df):
        results = []
       
        print('''==========================================\nStarting the process for Sheet Validation\n==========================================''')
        validation_reporting_log.info('''\n==========================================\nStarting the process for Sheet Validation\n==========================================''')
   
        try:
            file_iterations = tqdm(file_df.iterrows(), desc="Validating sheet structure", total=file_df.shape[0])
            for index, row in file_iterations:
                file_iterations.set_description("Processing %s" % row['Filename'])
               
                if row['Naming_Convention_Flag'] == "Match":
                    for sheet_name in SHEETS_TO_VALIDATE:           # @Gayathri, Will validate if both the sheets are present or not
                        print(f"{row['Filename']} | {sheet_name} → {row.get(f'{sheet_name}_Sheet_Exists')}")
                        if row.get(f"{sheet_name}_Sheet_Exists") != "Found":
                            continue
 
                        file_url = os.path.join(row['Foldername'],row['Filename'])
                        folder_name = row['Foldername']
 
                        sheet_df = pd.read_excel(
                            file_url,
                            sheet_name=sheet_name,
                            header=2,
                            engine='openpyxl' if row['Filename'].endswith('.xlsx') else 'xlrd'
                        )
 
                        total_columns = len(sheet_df.columns)
                        column_match_status = "Total Columns match" if total_columns == 83 else "Total Columns do not match"
 
                        results.append({
                            'Filename': row['Filename'],
                            'Foldername': row['Foldername'],
                            'Sheet': sheet_name,
                            'Parameter':'Total Number of Columns',
                            'Column Values': total_columns,
                            'Column_Match_Status': column_match_status
                        })
                       
 
                        expected_columns = expected_columns_df['column_name'].tolist()
 
                        unmatched_columns = [expected for expected, actual in zip(expected_columns, sheet_df.columns[:23]) if expected != actual]
                        column_name_match = len(unmatched_columns) == 0
                        column_name_status = "All columns match" if column_name_match else f"Columns not found: {', '.join(unmatched_columns)}"
 
                        results.append({
                            'Filename': row['Filename'],
                            'Foldername': row['Foldername'],
                            'Sheet': sheet_name,
                            'Parameter':'Fixed Column Names (Col:A-W)',
                            'Column Values': len(unmatched_columns),
                            'Column_Match_Status': column_name_status
                        })
 
                        month_columns_unmatched = [col for col in sheet_df.columns[23:] if not re.match(r'\d{4}/\d{2}', str(col))]
                        column_format_status = "All month columns found" if len(month_columns_unmatched)==0 else f"Columns not in yyyy/mm format: {', '.join(map(str, month_columns_unmatched))}"
 
                        results.append({
                            'Filename': row['Filename'],
                            'Foldername': row['Foldername'],
                            'Sheet': sheet_name,
                            'Parameter':'Monthly Column Names',
                            'Column Values': len(month_columns_unmatched),
                            'Column_Match_Status': column_format_status
                        })
 
                        extra_columns = [col for col in sheet_df.columns if col not in expected_columns and not re.match(r'\d{4}/\d{2}', str(col))]
                        extra_columns_status = "No extra columns" if len(extra_columns)==0 else f"Columns not required: {', '.join(map(str, extra_columns))}"
 
                        results.append({
                            'Filename': row['Filename'],
                            'Foldername': row['Foldername'],
                            'Sheet': sheet_name,
                            'Parameter':'Extra Columns',
                            'Column Values': len(extra_columns),
                            'Column_Match_Status': extra_columns_status
                        })
 
            print('''Sheet Validation Process Completed\n''')
            validation_reporting_log.info('''Sheet Validation Process Completed\n''')
           
            return pd.DataFrame(results)
       
        except Exception as e:
            print("Error occured while validating sheets: ", e)
            validation_reporting_log.error(f"Error occured while validating sheets: {e}")
            global error_code
            error_code=1
            return()
 
    def validate_columns(file_df, metadata_df):
        print('''==========================================\nStarting the process for Column Validation\n==========================================''')
        validation_reporting_log.info('''\n==========================================\nStarting the process for Column Validation\n==========================================''')
        report = []
       
        try:
            file_iterations = tqdm(file_df.iterrows(), desc="Validating columns", total=file_df.shape[0])
            for index, row in file_iterations:
                file_error = False
            
                file_iterations.set_description("Processing %s" % row['Filename'])
 
                if row['Naming_Convention_Flag'] == "Match":
                      
                
                    for sheet_name in SHEETS_TO_VALIDATE:    # @Gayathri, will validate columns for both the sheets
                        print(f"{row['Filename']} | {sheet_name} → {row.get(f'{sheet_name}_Sheet_Exists')}")
                        if row.get(f"{sheet_name}_Sheet_Exists") != "Found":
                            continue
 
                        file_url = os.path.join(row['Foldername'], row['Filename'])
                        foldername = row['Foldername']
 
                        workbook = load_workbook(filename=file_url, data_only=True)
                        sheet = workbook[sheet_name]
                        data = sheet.values
 
                        sheet_df = pd.DataFrame(data)
                        sheet_df.columns = sheet_df.iloc[2]
                        sheet_df = sheet_df[3:]
                        sheet_df.reset_index(drop=True, inplace=True)
                        sheet_df.dropna(how="all", inplace=True)
                        

                        current_column = None
 
                        for _, meta_row in metadata_df.iterrows():
                            column_name = meta_row['column_name']
                            current_column = column_name
                            expected_type = meta_row['data_type']
                            allows_blank = meta_row['contains_blanks'] == 'Yes'

                            if column_name not in sheet_df.columns:
                                report.append({
                                    'Filename': row['Filename'],
                                    'sheet_name': sheet_name,
                                    'Foldername': foldername,
                                    'column_name': column_name,
                                    'datatype_check': 'Column not found',
                                    'blank_check': 'Column not found',
                                    'na_check': 'Column not found',
                                    'excel_error_check': 'Column not found',
                                    'TS_value_check': '-'
                                })
                                continue

                            column_data = sheet_df[column_name]

                            if column_name == "Noted":
                                ts_value_check = 'Correct' if 'TS' in column_data.values else 'TS value not found'
                            else:
                                ts_value_check = '-'

                            try:
                                # FIRST: compute type_issues
                                if sheet_name == "SimulationValues":
                                     if expected_type == 'string':
                                         column_data = column_data.replace(0, pd.NA)
                                         type_issues = column_data.apply(lambda x: not isinstance(x, str) and pd.notna(x) and str(x).strip() != "")
                                     elif expected_type == 'int': 
                                         type_issues = column_data.apply(lambda x: not (isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                                     elif expected_type == 'float':
                                         type_issues = column_data.apply(lambda x: not isinstance(x, float) and pd.notna(x))
                                     elif expected_type == 'mixed':
                                         type_issues = column_data.apply(lambda x: not (isinstance(x, str) or isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                                     else:
                                         type_issues = pd.Series([False] * len(column_data))
                                else:
                                     if expected_type == 'string':
                                          type_issues = column_data.apply(lambda x: not isinstance(x, str) and pd.notna(x) and str(x).strip() != "")
                                     elif expected_type == 'int':
                                          type_issues = column_data.apply(lambda x: not (isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                                     elif expected_type == 'float':
                                          type_issues = column_data.apply(lambda x: not isinstance(x, float) and pd.notna(x))
                                     elif expected_type == 'mixed':
                                         type_issues = column_data.apply(lambda x: not (isinstance(x, str) or isinstance(x, int) or isinstance(x, float)) and pd.notna(x))                                     
                                     else:
                                          type_issues = pd.Series([False] * len(column_data))

                                # THEN: use type_issues
                                datatype_check = 'Correct' if not type_issues.any() else 'Invalid data type found'
                                blank_check = 'Correct' if allows_blank or not column_data.isnull().any() else 'Blank values found'
                                na_check = 'Correct' if not (column_data == 'NA').any() and not (column_data == 'N/A').any() else 'Manual NA found'

                                excel_errors = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']
                                excel_error_check = 'Correct' if not column_data.isin(excel_errors).any() else 'Excel error found'

                                report.append({
                                    'Filename': row['Filename'],
                                    'sheet_name': sheet_name,
                                    'Foldername': foldername,
                                    'column_name': column_name,
                                    'datatype_check': datatype_check,
                                    'blank_check': blank_check,
                                    'na_check': na_check,
                                    'excel_error_check': excel_error_check,
                                    'TS_value_check': ts_value_check
                                })
                                if file_error:
                                    validation_reporting_log.warning(f"FILE SKIPPED DUE TO ERROR ➜ {filename}")
                                    continue
                                

                            except Exception as col_err:
                                 file_error = True
                                 validation_reporting_log.error(f"COLUMN ERROR ➜ File={row['Filename']} | Sheet={sheet_name} | Column={column_name}", exc_info=True)
                                 print(f"⚠ Column failed ➜ {row['Filename']} | {sheet_name} | {column_name} → {col_err}")
                                 continue
                                 
                            
                            if file_error:
                                break 

 
 
                            # SPECIAL HANDLING FOR SimulationValues

                        if sheet_name == "SimulationValues":
                            if expected_type == 'string':
                                column_data = column_data.replace(0, pd.NA)
                                type_issues = column_data.apply(lambda x: not isinstance(x, str) and pd.notna(x) and str(x).strip() != "")
                            elif expected_type == 'int':
                                    type_issues = column_data.apply(lambda x: not (isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                            elif expected_type == 'float': type_issues = column_data.apply(lambda x: not isinstance(x, float) and pd.notna(x))
                            elif expected_type == 'mixed':
                                type_issues = column_data.apply(lambda x: not (isinstance(x, str) or isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                            else:
                                type_issues = pd.Series([False] * len(column_data))
 
 
 
                            # ORIGINAL logic for SimulationQty (unchanged)
                        else:
                            if expected_type == 'string':
                                type_issues = column_data.apply(lambda x: not isinstance(x, str) and pd.notna(x) and str(x).strip() != "")
                            elif expected_type == 'int':
                                type_issues = column_data.apply(lambda x: not (isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                            elif expected_type == 'float':
                                    type_issues = column_data.apply(lambda x: not isinstance(x, float) and pd.notna(x))
                            elif expected_type == 'mixed':
                                type_issues = column_data.apply(lambda x: not (isinstance(x, str) or isinstance(x, int) or isinstance(x, float)) and pd.notna(x))
                            else:
                                    type_issues = pd.Series([False] * len(column_data))
 
 
                            datatype_check = 'Correct' if not type_issues.any() else 'Invalid data type found'
                            blank_check = 'Correct' if allows_blank or not column_data.isnull().any() else 'Blank values found'
                            na_check = 'Correct' if not (column_data == 'NA').any() and not (column_data == 'N/A').any() else 'Manual NA found'
 
                            excel_errors = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']
                            excel_error_check = 'Correct' if not column_data.isin(excel_errors).any() else 'Excel error found'
 
                            report.append({
                                'Filename': row['Filename'],
                                'sheet_name': sheet_name,
                                'Foldername': foldername,
                                'column_name': column_name,
                                'datatype_check': datatype_check,
                                'blank_check': blank_check,
                                'na_check': na_check,
                                'excel_error_check': excel_error_check,
                                'TS_value_check':ts_value_check
                            })
                        if file_error:
                            print(
                                f"SKIPPING FILE ➜ {row['Filename']} | "
                                f"Sheet={sheet_name} | "
                                f"Column={current_column}"
                            )
                            continue

                        if file_error:
                            print(f"Skipping file due to column error → {row['Filename']}")
 
               
            report_df = pd.DataFrame(report)
            print("Column Validation Process Completed\n")
            validation_reporting_log.info("Column Validation Process Completed\n")
            return report_df
       
        except Exception as e:
            err_msg = (
                f"FILE-LEVEL COLUMN ERROR ➜ "
                f"File={row.get('Filename', 'UNKNOWN')} | "
                f"Sheet={sheet_name if 'sheet_name' in locals() else 'UNKNOWN'} | "
                f"Column={current_column} | "
                f"Error={e}")
            print(err_msg)
            validation_reporting_log.error(err_msg, exc_info=True)
            return pd.DataFrame(report)

            
 
    country_df = pd.read_excel(os.path.join('Supporting Files','config_file.xlsx'), sheet_name="countries")
    expected_columns_df = pd.read_excel(os.path.join('Supporting Files','config_file.xlsx'), sheet_name= "sheet_checkpoints")
   
    folders_to_traverse = generate_folder_paths(start_month, start_year, end_month, end_year, local_folder_path)
 
    file_check_df, file_presence_df = validate_file(folders_to_traverse, country_df, local_folder_path)
    sheet_check_df = validate_sheet(file_check_df, expected_columns_df)
    column_check_df = validate_columns(file_check_df, expected_columns_df)
 
    def write_multiple_dfs_to_excel(dataframes, sheet_names, filename, check_values):
       
        print('''==========================================\nStarting the process for Saving Validation Report\n==========================================''')
        validation_reporting_log.info('''\nStarting the process for Saving Validation Report\n==========================================''')
       
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                for df, sheet_name,check_value_list in tqdm(zip(dataframes, sheet_names,check_values), desc="Saving Validation Report", total=len(dataframes)):
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    workbook = writer.book
                    sheet = writer.sheets[sheet_name]
               
                    sheet.sheet_view.showGridLines = False
                   
                    header_fill = PatternFill(start_color='000080', end_color='000080', fill_type='solid')
                    header_font = Font(color='FFFFFF', bold=True)
                    border_side = Side(style='thin', color='D3D3D3')
                    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
                   
                    for cell in sheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.border = border
                   
                    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
                        for cell in row:
                            cell.border = border
                   
                    for check_value in check_value_list:
                        dxf = DifferentialStyle(fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
                                                font=Font(color='9C0006'))
                        rule = Rule(type='containsText', operator='containsText', text=check_value, dxf=dxf)
                        rule.formula = ['NOT(ISERROR(SEARCH("'+check_value+'",A1)))']
                       
                        max_column_letter = get_column_letter(sheet.max_column)
                        sheet.conditional_formatting.add(f'A1:{max_column_letter}{sheet.max_row}', rule)
                   
                    for column in sheet.columns:
                        max_length = 0
                        column = [cell for cell in column]
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        sheet.column_dimensions[get_column_letter(column[0].column)].width = (max_length + 2)
           
            validation_reporting_log.info(f"Workbook '{filename}' saved with sheets: {', '.join(sheet_names)}")
            return()
       
        except Exception as e:
            print("Error occured while saving validation file: ", e)
            validation_reporting_log.error(f"Error occured while saving validation file: {e}\n")
            global error_code
            error_code=1
            return()
 
    try:
        file_check_df.drop(columns=['File_URL'], inplace=True)
    except:
        pass
   
    dataframes = [file_presence_df,file_check_df, sheet_check_df, column_check_df]
    sheet_names = ['Folder Check Status','File Check Status', 'Sheet Check Status', 'Column Check Status']
   
    report_name = 'Validation_Report_Local_File'+str(today)+'.xlsx'
    filename = os.path.join(BASE_DIR,'Validation Reports',report_name)
    check_values = [['No'],['not','Not found','duplicate'], ['not'],['found']]
   
    write_multiple_dfs_to_excel(dataframes, sheet_names, filename, check_values)
   
    end_time = time()
    process_time = (end_time-start_time)/60
    validation_reporting_log.info(f"Process completed in {process_time} minutes")
   
    logging.shutdown()
    input("Enter")
    sys.exit(error_code)
 
except Exception as e:
    print("Error occured while running data_validation tool: ", e)
    traceback.print_exc()
    validation_reporting_log.error(f"Error occured while running data_validation tool: {e}\n")
    logging.shutdown()
    error_code=1
    input("Enter")
    sys.exit(error_code)
# %%