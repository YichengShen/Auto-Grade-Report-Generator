import pandas as pd
import numpy as np


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
             a = [[each, '0']]
             grades_df = grades_df.append({'Question Number' : each , 'Deduction in points' : 0} , ignore_index=True)
    #sort according to question number
    grades_df = grades_df.sort_values(by = 'Question Number')
    return grades_df


def main():
    """run this"""
    HW_number = 1
    all_questions_list = ['1a','1b','1c','1d','2','3','4','5']
    file_io = 'test_data.txt'


    count = len(open(file_io).readlines()) # get the number of groups
    for num in range(count):
        print("Reading grades of group No.{}...".format(num+1))
        contents = read_file(file_io,num)
        print("Preparing dataframe...")
        df = process_data(contents, all_questions_list)
        print(df)
    return



main()
