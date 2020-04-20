from changeOffice import Change

# 转换文件，有问题，目前采用wps手动转换方法
# import win32com.client as wc
# word = wc.Dispatch("Word.Application")
# doc = word.Documents.Open(r"D:\#Programming Lab\Python_TestLab\\高石001-X43_20200217_原始资料收集登记表.doc")
# 上面的地方只能使用完整绝对地址，相对地址找不到文件，且，只能用“\\”，不能用“/”，哪怕加了 r 也不行，涉及到将反斜杠看成转义字符。
# doc.SaveAs(r"D:\#Programming Lab\Python_TestLab\\test.docx", 16, False, "", True, "", False, \
# False, False, False)  # 转换后的文件
# doc.Close
# word.Quit

#转换文件，可能转出的文件读写空值，那么还得利用WPS或者LIBRE OFFICE
c = Change("./test")
c.doc2docx()
c.xls2xlsx()