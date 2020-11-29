
#################################-- Function Description --#################################

# Purpose:
#       This function deletes repeated data


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


def delete_repetitive_data(wb,dataSheet,savedFileName = "No",willPrint = False):

    # Static Variables (Some values are static THEY WILL BE ARRANGED ^^ TO DO ^^ )
    rowRange = 8000
    targetColumn = 3

    # take first cell as initial
    previousCell = dataSheet.cell(row=1, column=3)

    # iterate all rows starting from 2
    for row in range(2,rowRange):# Because first cell is taken before



        # get current Cell from target column
        currentCell = dataSheet.cell(row=row, column=targetColumn)

        # If the user wants to see process
        if willPrint:
            # print the row
            print("\n{}.ROW: {}\n".format(row, dataSheet.row[row]))
            # print current value ofcell
            print("Current Value {}".format(currentCell.value))

        # if the values are the same
        if (previousCell.value == currentCell.value):
            # remove current row
            dataSheet.delete_rows(row, 1)

            # inform the user
            if (willPrint):
                print("THERE IS A REPETITIVE DATA AT: \nROW INDEX: {}\nCOL INDEX: {}".format(row,3))

        # update the previous cell for next cell
        previousCell = currentCell

    # save the workbook
    save_workbook_as_xlsx(savedFileName,wb)
