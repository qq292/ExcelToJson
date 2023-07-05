import openpyxl
import json
import os
import argparse


def GetRange(sheet):

    r, c, title = str, str, []
    if isinstance(sheet, str):
        sheet = wb.get_sheet_by_name(sheet)
    elif isinstance(sheet, int):
        sheet = wb.worksheets[sheet]

    for row in list(sheet)[0]:
        if (row.value is None):
            break
        title.append(row.value)

    for row in sheet.iter_rows():
        if row[0].value is None:
            break
        else:
            r = row[0].coordinate

    for col in sheet.iter_cols():
        if col[0].value is None:
            break
        else:
            c = col[0].coordinate

    return title, sheet.title, sheet["A2":c.split("1")[0] +
                                     "".join(list(filter(str.isdigit, r)))]


def SheetToJson(sheetRange):
    title, fileName, sheet = GetRange(sheetRange)
    return json.dumps([
        dict(zip(title, [cell.value for cell in cellTupe]))
        for cellTupe in sheet
    ],
                      ensure_ascii=False,
                      indent=4), fileName


def SaveJson(data):
    jsonData, filename = data
    filePath = savePath + "\\" + filename + ".json"
    with open(filePath, "w") as f:
        f.write(jsonData)
    print("已保存到:  "+filePath)

def parse_args():

    parser = argparse.ArgumentParser(
        description="...->--")
    parser.add_argument('ExcelName', help="ExcelName")
    parser.add_argument('-SavePath', default=os.getcwd(), help="SavePath")
    parser.add_argument('-Sheet', default=0, type=int, help="Sheet")
    args = parser.parse_args()
    return args


if __name__ == "__main__":

    args = parse_args()
    excelPath = args.ExcelName
    savePath = args.SavePath
    sheets = args.Sheet
    wb = openpyxl.load_workbook(excelPath, data_only=True)
    if(sheets==-1):
        for sheet in wb.worksheets:
            SaveJson(SheetToJson(sheet))
    else:
         SaveJson(SheetToJson(sheets))





