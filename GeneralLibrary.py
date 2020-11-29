

#################################-- Library Description --#################################

# Purpose:
#       This library contains functions that will be used frequently among the program


#################################-- END Function Description --##############################



def control_xlsx_extension(fileName):

    #################################-- Function Description --#################################

    # Purpose:
    #       This function controls the file name whether
    #    it contains ".xlsx" extension

    # INPUTS:
    #   fileName      -> a string file name & path

    # OUTPUTS:
    #   fileName      -> an output string file with extension

    #################################-- END Function Description --##############################

    # control if the name contains ".xlsx"
    if ".xlsx" in fileName:
        # return the fileName
        return fileName
    else:
        # update the filename
        fileName += ".xlsx"
        # return the fileName
        return fileName


def save_workbook_as_xlsx(savedFileName,wb):

    #################################-- Function Description --#################################

    # Purpose:
    #       This function controls the file name at first then saves the workbook

    # INPUTS:
    #   fileName      -> a string file name & path (IF IT IS "No" DON"T SAVE)
    #   wb            -> workbook that will be saved

    # OUTPUTS:
    #   Excel File (".xlsx")!!

    #################################-- END Function Description --##############################

    try:
        # if a filename is given except "No"
        if (savedFileName != "No"):

            # control if the name contains ".xlsx"
            savedFileName = control_xlsx_extension(savedFileName)

            # save the file
            wb.save(savedFileName)
    # In case any problem
    except:
        # inform the user
        print("There is a problem SAVE WORKBOOK AS XLSX !!")


def get_file_name():
    #################################-- Function Description --#################################

    # Purpose:
    #       This function gets a string input file name from user

    # OUTPUTS:
    #   fileName    -> a string file name with ".xlsx" extension

    #################################-- END Function Description --##############################

    fileName = input("Please type the file name: ")

    fileName = control_xlsx_extension(fileName)

    print(fileName)

    return fileName
