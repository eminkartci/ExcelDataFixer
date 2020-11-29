#################################--DEFAULT CODES--#################################
# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


#def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    #print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
#if __name__ == '__main__':
    #print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

###############################-- END DEFAULT CODES--###############################

#################################-- IMPORT --#################################

# To read excel values import the function
from ReadExcelOnlyValues import read_excel_only_values

# To convert string numbers to integer
from ConvertStringNumbers import convert_string_numbers

# To delete repetitive data
from DeleteRepetitiveData import  delete_repetitive_data

# import input function
from GeneralLibrary import get_file_name

#################################-- END  IMPORT --#############################


#################################--DRIVER PROGRAM--#################################



# get file name from user
dataFileName = get_file_name()

# read values and bu calling the function
wb, dataSheet = read_excel_only_values(dataFileName)
    # Inputs:
        # file name or path
        # sheet name (Default: including "veri" names)
    # Outputs
        # workbook
        # worksheet


# convert if any string number exists
wbCorrectNumbers , dataSheetCorrectNumbers = convert_string_numbers(wb,dataSheet)
# INPUTS:
    #   wb            -> Excel Workbook (Read by "openpyxl" library)
    #   dataSheet     -> There might be several sheets at workbook. The function needs only "Values"/"veri" sheet
    #   savedFileName -> if the user wants to save the processed workbook, s/he must give a file name
    #   willPrint     -> if the user wants to see the process or errors type True

# OUTPUTS
    #   wb          -> Processed data is stored in the same variable
    #   dataSheet   -> Data sheet is stored in the same variable


# delete if the same datum is given twice
delete_repetitive_data(wb,dataSheet)
# INPUTS:
    #   wb            -> Excel Workbook (Read by "openpyxl" library)
    #   dataSheet     -> There might be several sheets at workbook. The function needs only "Values"/"veri" sheet
    #   savedFileName -> if the user wants to save the processed workbook, s/he must give a file name
    #   willPrint     -> if the user wants to see the process or errors type True









#################################--DRIVER PROGRAM--#################################