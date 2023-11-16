
from datetime import datetime
import time

import openpyxl



def main():

    TimeStamp = str(datetime.now())
    TimeStamp = TimeStamp.replace("-", ".")
    TimeStamp = TimeStamp.replace(":", ".")
    TimeStamp = TimeStamp[:len(TimeStamp)-7]
    TimeStamp = str(TimeStamp)
    FullPathForExcelReportFile = r'C:\Users\meirb\AppData\Local\Programs\Python\Python311\StructShare\Reports\Report_From_Date_And_Time_' + TimeStamp + r'.xlsx'
    print("Full Path For Excel Report File = ", FullPathForExcelReportFile)

    
    ExcelRowNumber = 1

    # Create a new workbook
    workbook = openpyxl.Workbook()
    # Select the active worksheet
    worksheet = workbook.active
    # Headers for the Excel Report
    worksheet['A1'] = 'Test Case Name'
    worksheet['B1'] = 'Check Name'
    worksheet['C1'] = 'Status'
    worksheet['D1'] = 'Comment (Exception)'
    worksheet['E1'] = 'The Screenshot File'
    

    # StructShare
    print("\n\StructShare = \n")
    import StructShare
    StructShare.setIndex(ExcelRowNumber)
    print ("Script INITIAL ExcelRowNumber = ", StructShare.getIndex())
    ExcelRowNumber = StructShare.StructShare(ExcelRowNumber, FullPathForExcelReportFile, worksheet, workbook)
    print ("Script LAST ExcelRowNumber = ", ExcelRowNumber)


main()