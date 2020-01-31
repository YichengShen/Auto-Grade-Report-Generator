import pandas as pd


def read_file(io):
    f = open(io, "r")
    contents = f.read()
    f.close()
    return contents

def process_data(all_groups):
    return


def main():
    """run this"""
    file_io = 'test_data.txt'
    contents = read_file(file_io)
    process_data(contents)
    return



main()
