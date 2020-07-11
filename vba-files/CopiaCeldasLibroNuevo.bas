Dim oExcel As Object    
Dim oBook As Object    
Dim oSheet As Object 
    
'Start a new workbook in Excel    
Set oExcel = CreateObject("Excel.Application")    
Set oBook = oExcel.Workbooks.Add
      
'Add data to cells of the first worksheet in the new workbook    
Set oSheet = oBook.Worksheets(1)   



oSheet.Range("A1").Value = "Last Name"    
oSheet.Range("B1").Value = "First Name"    
oSheet.Range("A1:B1").Font.Bold = True    
oSheet.Range("A2").Value = "Doe"    
oSheet.Range("B2").Value = "John"     

'Save the Workbook and Quit Excel    
oBook.SaveAs "C:\Book1.xls"    
oExcel.Quit