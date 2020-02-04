import pandas as pd
import numpy as np
import os
import win32com.client
from pywintypes import com_error
import xlsxwriter


def read_file(io, index):
    f = open(io, "r")
    contents = f.readlines()[index]
    f.close()
    return contents

def process_data(data, q_num_list):
    data = data.replace(" ","") # strip all whitespaces
    list_mixed = data[:-1].split(",") #from 0 to -1 takes out the \n
    group_num = list_mixed[0]
    grades_array = np.array(list_mixed[1:])
    grades_array_2d = np.reshape(grades_array, (-1, 2))
    grades_df = pd.DataFrame(grades_array_2d, columns=['Question Number', 'Deduction in points'])

    #append correct questions
    for each in q_num_list:
        if each not in np.array(grades_df["Question Number"]):
             grades_df = grades_df.append({'Question Number' : each , 'Deduction in points' : '0'} , ignore_index=True)
    #sort according to question number
    grades_df = grades_df.sort_values(by = 'Question Number')
    # Delete the default index
    grades_df.set_index('Question Number', inplace=True)
    return group_num, grades_df


def make_dir(hw_num):
    #make a folder on Desktop
    dir = "hw{}".format(hw_num)
    parent_dir = os.path.join(os.environ["HOMEPATH"], "desktop")
    path = os.path.join(parent_dir, dir)
    os.mkdir(path)
    path_excel_folder = os.path.join(path,"Excel")
    os.mkdir(path_excel_folder)
    print("Directory '% s' created" % dir)
    return path, path_excel_folder


def export_to_excel(df, path, hw_num, group_number):
    # export to excel
    df['Deduction in points'] = pd.to_numeric(df['Deduction in points'],errors='coerce')
    print("Exporting to Excel...")
    excel_name = "HW{}_{}_report.xlsx".format(hw_num, group_number)
    excel_path = os.path.join(path, excel_name)
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Detailed Report')
    workbook  = writer.book
    worksheet1 = writer.sheets['Detailed Report']
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    format_red = workbook.add_format({'bg_color': '#FFC7CE',
                                      'font_color': '#9C0006'})
    format_green = workbook.add_format({'bg_color': '#C6EFCE',
                                        'font_color': '#006100'})
    worksheet1.set_column('A:A', 18)
    worksheet1.set_column('B:B', 20)
    worksheet1.conditional_format('B2:B200', {'type': 'cell',
                                        'criteria': '<',
                                        'value': 0,
                                        'format': format_red})
    worksheet1.conditional_format('B2:B200', {'type': 'cell',
                                        'criteria': '=',
                                        'value': 0,
                                        'format': format_green})
    writer.save()
    return excel_path


def to_pdf(path, excel_path, hw_num, group_number):
    # Path to original excel file
    WB_PATH = excel_path

    pdf_name = "HW{}_{}_report".format(hw_num, group_number)
    # PDF path when saving
    PATH_TO_PDF = os.path.join(path, pdf_name)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        print('Start conversion to PDF')
        # Open
        wb = excel.Workbooks.Open(WB_PATH)
        # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
        ws_index_list = [1]
        wb.WorkSheets(ws_index_list).Select()
        # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
    except com_error as e:
        print('failed.')
    else:
        print('Succeeded.')
    finally:
        wb.Close()
        excel.Quit()

def main():
    """run this"""
    HW_number = 1
    all_questions_list = ['1a','1b','1c','1d','2','3','4','5']
    file_io = 'test_data.txt'


    path, path_excel_folder = make_dir(HW_number)
    count = len(open(file_io).readlines()) # get the number of groups
    for num in range(count):
        print("Reading grades of group No.{}...".format(num+1))
        contents = read_file(file_io,num)
        print("Preparing dataframe...")
        g_num, df = process_data(contents, all_questions_list)
        path_excel = export_to_excel(df, path_excel_folder, HW_number, g_num)
        to_pdf(path, path_excel, HW_number, g_num)
    return



main()
