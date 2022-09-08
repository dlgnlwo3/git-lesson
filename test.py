import os
import win32com.client
import openpyxl
import xlrd
import time


def copy_sheet(old_file, new_file):

    try:
        print(f"메인 브랜치에서 수정함")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True

        wb1 = excel.Workbooks.Open(old_file)
        # wb2 = excel.Workbooks.Open(new_file)
        ws1 = wb1.ActiveSheet
        print(f"{wb1}, {ws1}")

        # ws1.Copy(After=wb2.Worksheets(f"Sheet5"))
        time.sleep(1)

        # wb1.Close(savechanges=0)
        # wb2.Close(savechanges=1)

        # wb1.Worksheets("Sheet1").Copy(After=wb2.Worksheets("Sheet1"))
        # wb2.Save()
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
