

#################################-- Function Description --#################################

# Purpose:
#       This function gets an excel file and checks if there is a wrong `,` or `.` in the values.
#    If a problem exists the function can change the punctuation errors.

# INPUTS:
#   wb            -> Excel Workbook (Read by "openpyxl" library)
#   dataSheet     -> There might be several sheets at workbook. The function needs only "Values"/"veri" sheet
#   savedFileName -> if the user wants to save the processed workbook, s/he must give a file name
#   willPrint     -> if the user wants to see the process or errors type True

# OUTPUTS
#   wb          -> Processed data is stored in the same variable
#   dataSheet   -> Data sheet is stored in the same variable

#################################-- END Function Description --##############################

#################################-- IMPORT --#################################

# To control ".xlsx" extension
from GeneralLibrary import control_xlsx_extension
# Save xlsx file
from GeneralLibrary import save_workbook_as_xlsx

#################################-- END  IMPORT --#############################



def convert_string_numbers(wb,dataSheet,savedFileName = "No",willPrint = False,surpressErrors=False):

    # Static Variables (Some values are static THEY WILL BE ARRANGED ^^ TO DO ^^ )
    rowRange = 8000
    colRange = 40

    # comma was stand for decimal point at "medas" data
    commaPosition = 1000

    # iterate rows
    for rowIndex in range(1, rowRange):
        # iterate columns
        for colIndex in range(1, colRange):

            # get current cell's value
            val1 = dataSheet.cell(row=rowIndex, column=colIndex).value

            # if the user wants to see process print the value
            if(willPrint):
                print("val1", val1, type(val1))

            # If the value is string and has at least "." or ","
            if (type(val1) == str) and (("," in val1) or ("." in val1)):

                # All cells might not appropriate for the process
                try:
                    # Delete the comma
                    val1 = val1.replace(",", "")
                    # Delete the dot
                    val1 = val1.replace(".", "")

                    # try to cast it into integer
                    intValue = int(val1)# if any problem occurs here this cell is not my target go to except

                    # I assume that they separate decimals
                    intValue /= commaPosition

                    # Update the cell's new value
                    dataSheet.cell(row=rowIndex, column=colIndex).value = intValue

                    # Inform the user (OPTIONAL)
                    if(willPrint):
                        print("{} is written".format(intValue))

                # If any problem occurs
                except:
                    if (surpressErrors):
                        i = 0
                    else:
                        # Inform the user
                        print("There is an error on CONVERT STRING NUMBERS METHOD !! \nCOL INDEX: {}\nROW INDEX: {}".format(colIndex,rowIndex))

    # save the workbook
    save_workbook_as_xlsx(savedFileName,wb)

    # return the updated worksheet and data sheet
    return wb, dataSheet

