import pandas as pd
import sys
import os
import glob
import logging


# This function reads the folder and extracts the excel files.

def get_files(path):
    path = os.getcwd()
    csv_files = glob.glob(os.path.join(path, "*.xlsx"))
    return csv_files


# now we want to extract excel files with all sheets from the folder
# and transfer it to dict of dataFrames in pandas format. 

def read_excel(excel):
    # read the csv file
    df = pd.read_excel(excel, sheet_name=None)
    return df


def turn_to_table(dataframe):
    hTable = dataframe.to_html(classes='table table-stripped')
    return hTable
    
def main():

    # this if else chain is used for testing comman line arguments.

    if len(sys.argv) > 2:
        sys.exit("this programs run with only one argument: please enter the directory of folder")
    elif type(sys.argv[1]) == int:
        sys.exit("please enter only the directory of folder")
    else:
        path = sys.argv[1]
    try:

        with open('index.html', 'w') as html:
            # path = "E:\Excel_analayzer\Orders-With Nulls.xlsx"
            excels = get_files(path)
            for excel in excels:
                df = read_excel(excel)
                for table in df.items():
                    table1 = turn_to_table(table[1])
                    # print(table1)
                    html.write(table1)
            print("operation completed successfully")        
            


    except Exception as Argument:
        logging.exception("Error occurred while doing operation")
    

if __name__ == "__main__":
    main()