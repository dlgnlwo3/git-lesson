import os
import win32com.client
import openpyxl
import xlrd
import time


def copy_sheet(old_file, new_file):

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True

        wb1 = excel.Workbooks.Open(old_file)
        ws1 = wb1.ActiveSheet
        print(f"{wb1}, {ws1}")

        time.sleep(1)

        wb1.SaveAs(new_file, FileFormat=51)

    except Exception as e:
        print(e)

    finally:
        excel.Quit()


if __name__ == "__main__":

    old_file = os.path.join(os.getcwd(), "report.xls")
    new_file = os.path.join(os.getcwd(), "copy.xlsx")

    try:
        copy_sheet(old_file, new_file)
    except Exception as e:
        print(e)
        print("실패")
