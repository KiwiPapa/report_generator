import os, xlrd, openpyxl
from changeOffice import Change

# 格式转换
c = Change(".\\极睿导出")
c.doc2docx()
c.xls2xlsx()

# 现在开始提取成果表中的内容
PATH = ".\\极睿导出"
for fileName in os.listdir(PATH):
    if 'Layer' in fileName and '$' not in fileName and 'New' not in fileName:#避免临时文件报错
        fileDirLayer = PATH + "\\" + fileName
        wb1 = xlrd.open_workbook(fileDirLayer)
        wb1_openpyxl = openpyxl.load_workbook(fileDirLayer)
    elif 'Result' in fileName and '$' not in fileName and 'New' not in fileName:#避免临时文件报错
        fileDirResult = PATH + "\\" + fileName
        wb2 = xlrd.open_workbook(fileDirResult)
        wb2_openpyxl = openpyxl.load_workbook(fileDirResult)

##############################################
# 处理Layer表
sheet1 = wb1.sheets()[0]
nrow1 = sheet1.nrows
ncol1 = sheet1.ncols

# 用openpyxl进行处理
sheet1_openpyxl = wb1_openpyxl[wb1_openpyxl.sheetnames[0]]
for row in range(nrow1):
    for col in range(ncol1):
        if sheet1.cell_value(row, col) == '垂直定位':
            delete_Row = row
sheet1_openpyxl.delete_rows(delete_Row + 1)
# deleterows(sheet1_openpyxl, delete_Row + 1)#openpyxl中数行数从1开始
sheet1_openpyxl['C2'] = None

for row in range(1, nrow1):
    for col in range(1, ncol1):
        if sheet1_openpyxl[row][col].value == '龙一2':
            sheet1_openpyxl[row][col].value = '龙一^2'
        elif sheet1_openpyxl[row][col].value == '龙一１1':
            sheet1_openpyxl[row][col].value = '龙一^１1'
        elif sheet1_openpyxl[row][col].value == '龙一１2':
            sheet1_openpyxl[row][col].value = '龙一^１2'
        elif sheet1_openpyxl[row][col].value == '龙一１3':
            sheet1_openpyxl[row][col].value = '龙一^１3'
        elif sheet1_openpyxl[row][col].value == '龙一１4':
            sheet1_openpyxl[row][col].value = '龙一^１4'
        else:
            pass
wb1_openpyxl.save('.\\极睿导出\\Layer_New.xlsx')
##############################################
# 处理Result表
sheet2 = wb2.sheets()[0]
nrow2 = sheet2.nrows
ncol2 = sheet2.ncols

# 用openpyxl进行处理
sheet2_openpyxl = wb2_openpyxl[wb2_openpyxl.sheetnames[0]]
for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '自然伽马':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '补偿声波':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '补偿密度':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '电阻率':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '孔隙度':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '含水饱和度':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '有机碳含量':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

for row in range(1, sheet2_openpyxl.max_row):
    for col in range(1, sheet2_openpyxl.max_column):
        if sheet2_openpyxl[row][col - 1].value == '总含气量':
            delete_Col = col
            sheet2_openpyxl.delete_cols(delete_Col)

# insert column
sheet2_openpyxl.insert_cols(4)
for row in range(3, sheet2_openpyxl.max_row + 1):
    sheet2_openpyxl[row][2].value = sheet2_openpyxl[row][2].value.split('--')[1]
    sheet2_openpyxl[row][3].value = sheet2_openpyxl[row][2].value.split('--')[0]
wb2_openpyxl.save('.\\极睿导出\\Result_New.xlsx')
