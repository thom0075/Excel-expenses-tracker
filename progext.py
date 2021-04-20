import time
import openpyxl as xl


def checkIfOpen(workbook):
    try:
        workbook.save("example.xlsx")
    except PermissionError:
        print("Permission Error, permission denied. Close the document in 5 seconds and wait for the\n \
program to try again")
        time.sleep(5)
        print("Trying again... wait")
        try:
            workbook.save("example.xlsx")
            print("Workbook saved successfully!")
        except PermissionError:
            print("Fatal error, exiting the program...")
        finally:
            print("Exiting the program...")
            exit()
            """WARNING!  This function can be used only at the end of the program"""


def nextRow(cell, sheet, value="", value1=0):
    return sheet[cell].row + 1


def sleep(seconds):
    time.sleep(seconds)
