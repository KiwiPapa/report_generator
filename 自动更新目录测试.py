import win32com.client
word = win32com.client.DispatchEx("Word.Application")
doc = word.Documents.Open("123456.docx")
doc.TablesOfContents(1).Update()
doc.Close(SaveChanges=True)
word.Quit()

