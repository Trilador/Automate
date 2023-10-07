Dim xlApp
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = True  ' Set to False if you don't want Excel to show
Dim xlBook
Set xlBook = xlApp.Workbooks.Open("C:\Users\timot\Desktop\Python\DailyEmails.xlsm", 0, False)
xlApp.Run "Chep"
xlBook.Close False  ' Updated line
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing
