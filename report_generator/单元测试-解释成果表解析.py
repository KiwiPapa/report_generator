import os, xlrd

# 现在开始提取成果表中的内容
PATH = ".\\2统计表"
for fileName in os.listdir(PATH):
    if '1统' in fileName:
        fileDir = PATH + "\\" + fileName
        workbook1 = xlrd.open_workbook(fileDir)
    elif '2统' in fileName:
        fileDir = PATH + "\\" + fileName
        workbook2 = xlrd.open_workbook(fileDir)

##########################
# 解析解释成果表-1统
sheet1 = workbook1.sheets()[0]

nrow1 = sheet1.nrows
ncol1 = sheet1.ncols

# 统计结论
good_Length1 = str(sheet1.cell_value(3, 2))
good_Ratio1 = str(sheet1.cell_value(3, 3))

median_Length1 = str(sheet1.cell_value(4, 2))
median_Ratio1 = str(sheet1.cell_value(4, 3))

bad_Length1 = str(sheet1.cell_value(5, 2))
bad_Ratio1 = str(sheet1.cell_value(5, 3))

# 合格率
pass_Percent1 = str(round((sheet1.cell_value(3, 3) + sheet1.cell_value(4, 3)), 2))

##########################
# 解析解释成果表-2统
sheet2 = workbook2.sheets()[0]

nrow2 = sheet2.nrows
ncol2 = sheet2.ncols

# 统计结论
good_Length2 = str(sheet2.cell_value(3, 2))
good_Ratio2 = str(sheet2.cell_value(3, 3))

median_Length2 = str(sheet2.cell_value(4, 2))
median_Ratio2 = str(sheet2.cell_value(4, 3))

bad_Length2 = str(sheet2.cell_value(5, 2))
bad_Ratio2 = str(sheet2.cell_value(5, 3))

# 合格率
pass_Percent2 = str(round((sheet2.cell_value(3, 3) + sheet2.cell_value(4, 3)), 2))

input()