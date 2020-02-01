import pandas as pd
import numpy as np
import os
import win32com.client
from pywintypes import com_error


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
    return group_num, grades_df


def make_dir(hw_num):
    #make a folder on Desktop
    dir = "hw{}".format(hw_num)
    parent_dir = os.path.join(os.environ["HOMEPATH"], "desktop")
    path = os.path.join(parent_dir, dir)
    os.mkdir(path)
    print("Directory '% s' created" % dir)
    return path


def to_excel(df, path, hw_num, group_number):
    # export to excel
    print("Exporting to Excel...")
    excel_name = "HW{}_{}_report.xlsx".format(hw_num, group_number)
    excel_path = os.path.join(path, excel_name)
    df.to_excel(excel_path, sheet_name='Details about deduction')
    return excel_path


def to_pdf(path, excel_path, hw_num, group_number):
    # Path to original excel file
    WB_PATH = excel_path

    pdf_name = "HW{}_{}_report.xlsx".format(hw_num, group_number)
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


    path = make_dir(HW_number)
    count = len(open(file_io).readlines()) # get the number of groups
    for num in range(count):
        print("Reading grades of group No.{}...".format(num+1))
        contents = read_file(file_io,num)
        print("Preparing dataframe...")
        g_num, df = process_data(contents, all_questions_list)
        path_excel = to_excel(df, path, HW_number, g_num)
        to_pdf(path, path_excel, HW_number, g_num)
    return



main()
