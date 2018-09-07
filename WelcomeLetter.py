from mailmerge import MailMerge
import xlrd
import os
import shutil


def get_cell_value(sheet, rownum, colnum):
    cell = sheet.cell(rownum, colnum)
    cell_value = cell.value
    if cell.ctype in (2, 3) and int(cell_value) == cell_value:
        cell_value = str(int(cell_value))
    else:
        cell_value = str(cell_value)
    return cell_value


directory_name = os.path.dirname(os.path.abspath(__file__))
folder_name = os.path.join(directory_name, 'WelcomeLetters')

if os.path.exists(folder_name):
    shutil.rmtree(folder_name)
os.makedirs(folder_name)  # Made WelcomeLetters folder.

template = "Parent-Welcome-Letter-Template.docx"
excel_path = "Parent-Welcome-Letter-Data.xlsx"

worksheet = xlrd.open_workbook(excel_path).sheet_by_index(0)

RowToMergeField = {}
for row_num in range(worksheet.nrows):
    if row_num is 0:
        for col_num in range(worksheet.ncols):
            RowToMergeField[get_cell_value(worksheet, row_num, col_num)] = col_num
    else:  # Merge field entries
        dict_merge = {}
        document = MailMerge(template)
        for merge_field in document.get_merge_fields():
            dict_merge[str(merge_field)] = get_cell_value(worksheet, row_num, RowToMergeField[str(merge_field)])
        document.merge(**dict_merge)
        document.write(os.path.join(folder_name,'WelcomeLetter' + str(row_num) + '.docx'))
        document.close()


