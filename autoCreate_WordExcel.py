#
#
#

import sys
import os
import shutil
import time
import openpyxl
import io
import win32com.client

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')


replaceSourceList = []
sourceColumn = 2
replaceDestList = []
DestColumn = 3
replaceRule = openpyxl.load_workbook(".\\replace_rule.xlsx").get_sheet_by_name('replace_rule')
cutTime = time.strftime('_%Y%m%d-%H_%M_%S', time.localtime(time.time()))
sourceDir = ".\\source_template"


# create replace rule
def create_replace_rule():
    for i in range(3, 102):
        if replaceRule.cell(row=i, column=sourceColumn).value is not None:
            replaceSourceList.append((str(replaceRule.cell(row=i, column=sourceColumn).value)).rstrip('').rstrip('\n').rstrip('\r'))
        if replaceRule.cell(row=i, column=DestColumn).value is not None:
            replaceDestList.append((str(replaceRule.cell(row=i, column=DestColumn).value)).rstrip('').rstrip('\n').rstrip('\r'))
    print(replaceSourceList)
    print(replaceDestList)
    global destDir
    destDir = ".\\dest_document\\[autoGen] " + replaceDestList[0] + cutTime
    if (len(replaceSourceList) == len(replaceDestList)) is not True:
        print(" !!!!!! [Error] Failed to generate file. Please check file replace_rule.xlsx !!!!!! \n")
        return False
    else:
        return True


# Gen a new copy
def gen_new_copy():
    shutil.copytree(sourceDir, destDir)


def list_all_files(rootdir):
    _files = []
    list = os.listdir(rootdir)
    for i in range(0, len(list)):
        path = os.path.join(rootdir, list[i])
        if os.path.isdir(path):
            _files.extend(list_all_files(path))
        if os.path.isfile(path):
            _files.append(path)
    return _files


def check_file_format(temp_file):
    if 'doc' in (temp_file.lstrip(destDir)).split('.')[1]:
        return 'doc'
    elif 'wps' in (temp_file.lstrip(destDir)).split('.')[1]:
        return 'doc'
    elif 'xls' in (temp_file.lstrip(destDir)).split('.')[1]:
        return 'xls'
    elif 'ignore' in (temp_file.lstrip(destDir)).split('.')[1]:
        return 'ignore'
    else:
        return 'no_found'


# List all documents and replace according to the rules
def precess_files():
    allFilesList = list_all_files(destDir)
    i = 0
    for each_file in allFilesList:
        i = i + 1
        if check_file_format(each_file) == 'doc':
            print("[ " + str(i) + " ] Start processing : " + each_file.lstrip(destDir))
            abs_path = os.path.abspath(each_file)
            print(abs_path)
            process_docx(abs_path)
            print("       Replace completed")
        elif check_file_format(each_file) == 'xslx':
            process_xlsx(each_file)
            print("[ " + str(i) + " ] Start processing : " + each_file.lstrip(destDir))
            print("       Replace completed")
        else:
            print("[ " + str(i) + " ] Start processing : " + each_file.lstrip(destDir))
            print("       Replace completed")
    return


# Docx process lib start
def process_docx(file_doc):
    # start word process
#    try:
#        w = win32com.client.gencache.EnsureDispatch('kwps.application')
#        time.sleep(3)
#    except:
#        w = win32com.client.gencache.EnsureDispatch('wps.application')
#        time.sleep(3)
#    else:
    time.sleep(3)
    w = win32com.client.gencache.EnsureDispatch('word.application')
    # running in the background, display program interface, no warning
    # w.Visible = 0
    # w.DisplayAlerts = 0
    # open doc
    # debug
    print(file_doc)
    doc = w.Documents.Open(file_doc)
    # doc = w.Documents.Open(file_doc, "utf8")
    # doc = w.Documents.Open2000(file_doc, "utf8")
    for i in range(len(replaceDestList)):
        # replace text body
        w.Selection.Find.ClearFormatting()
        w.Selection.Find.Replacement.ClearFormatting()
        w.Selection.Find.Execute(replaceSourceList[i], False, False, False, False, False, True, 1, True, replaceDestList[i], 2)
        # replace text table
#        for table_i in range(1, w.ActiveDocument.Tables.Count + 1):
#            table = w.ActiveDocument.Tables[table_i]
#            numRows = table.Rows.Count
#            numcolumns = table.Columns.Count
#            for row in range(1, numRows + 1):
#                for column in range(1, numcolumns + 1):
#                    print(table.cell(row, column).text)
        # replace text header
#        for header_i in range(len(w.ActiveDocument.Sections)):
#            w.ActiveDocument.Sections[header_i].Headers[0].Range.Find.ClearFormatting()
#            w.ActiveDocument.Sections[header_i].Headers[0].Range.Find.Replacement.ClearFormatting()
#            w.ActiveDocument.Sections[header_i].Headers[0].Range.Find.Execute(replaceSourceList[i], False, False, False, False, False, True, 1, False, replaceDestList[i], 2)
        # replace text footer
#        for header_i in range(len(w.ActiveDocument.Sections)):
#            w.ActiveDocument.Sections[header_i].Footers[0].Range.Find.ClearFormatting()
#            w.ActiveDocument.Sections[header_i].Footers[0].Range.Find.Replacement.ClearFormatting()
#            w.ActiveDocument.Sections[header_i].Footers[0].Range.Find.Execute(replaceSourceList[i], False, False, False, False, False, True, 1, False, replaceDestList[i], 2)
    doc.SaveAs(file_doc)
    # doc.Close()
    w.Documents.Close()
    w.Quit()
    return
# Docx process lib end


# Excel process lib start
def process_xlsx(file_xslx):
    # start xlsx process
    ex=win32com.client.Dispatch('Excel.Application')
    # running in the background, # running in the background, display program interface, no warning 
    ex.Visible = 0
    ex.DisplayAlerts = 0
    # open xlsx
    # debug
    print(file_xslx)
    wk=ex.Workbooks.Open(file_xslx, "utf8")
    for sheet_i in range(wk.sheetnames):
        ws=wk.Worksheets(sheet_i)
        ws.Activate
        for i in range(len(replaceDestList)):
            ex.Selection.Replace(replaceSourceList[i], False, False, False, False, False, True, 1, False, replaceDestList[i], 2)
    wk.Save()
    ex.quit()
    return
# Excel process lib end


if __name__ == "__main__":
    try:
        if create_replace_rule() is True:
            gen_new_copy()
            precess_files()
#    except ValueError as e:
#        os.system('pause')
#    except ZeroDivisionError as e:
#        os.system('pause')
#    else:
#        os.system('pause')
    finally:
        print("end")
#        os.system('pause')

