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
    list_mixed = data[:-1].split(",") #from 0 to -1 takes out the \n
    group_num = list_mixed[0].strip()
    grades_array = np.array(list_mixed[1:])
    grades_array_2d = np.reshape(grades_array, (-1, 3))
    grades_df = pd.DataFrame(grades_array_2d, columns=['Question Number',
                                                      'Deduction in points',
                                                      'Comments'])
    # Strip all the spaces in first two attributes
    for index in range(len(grades_df)):
        grades_df['Question Number'][index] = grades_df['Question Number'][index].replace(" ","")
        grades_df['Deduction in points'][index] = grades_df['Deduction in points'][index].replace(" ","")
    # Append correct questions
    for each in q_num_list:
        if each not in np.array(grades_df["Question Number"]):
             grades_df = grades_df.append({'Question Number' : each ,
                                           'Deduction in points' : '0',
                                           'Comments': ''} , ignore_index=True)
    # Sort according to question number
    grades_df = grades_df.sort_values(by = 'Question Number')
    # Delete the default index
    grades_df.set_index('Question Number', inplace=True)
    # change deduction of points from str to float
    grades_df['Deduction in points'] = pd.to_numeric(grades_df['Deduction in points'],errors='coerce')
    return group_num, grades_df


def add_title_and_sum(grades_df, group_num, hw_num, full_score):
    title = "HW{} {}".format(hw_num, group_num.title())
    total_deduction = grades_df['Deduction in points'].sum()
    final_score = full_score + total_deduction
    return grades_df


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
    format_q_num = workbook.add_format({'valign': 'vcenter',
                                        'bold': False})
    format_point = workbook.add_format({'valign': 'vcenter',
                                        'align' : 'center',
                                        'bottom': True})
    format_comment = workbook.add_format({'text_wrap': True,
                                          'bottom': True,
                                          'right': True,
                                          'left': True})
    worksheet1.set_column('A:A', 18, format_q_num)
    worksheet1.set_column('B:B', 19, format_point)
    worksheet1.set_column('C:C', 50, format_comment)
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
    full_score = 37
    all_questions_list = ['1a','1b','1c','1d','2','3','4','5']
    file_io = 'test_data.txt'


    path, path_excel_folder = make_dir(HW_number)
    count = len(open(file_io).readlines()) # get the number of groups
    for num in range(count):
        print("Reading grades of group No.{}...".format(num+1))
        contents = read_file(file_io,num)
        print("Preparing dataframe...")
        g_num, df = process_data(contents, all_questions_list)
        df = add_title_and_sum(df, g_num, HW_number, full_score)
        path_excel = export_to_excel(df, path_excel_folder, HW_number, g_num)
        to_pdf(path, path_excel, HW_number, g_num)
    return


main()
