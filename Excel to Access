Sub ImportExcelWorksheetsIntoAccess()

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim xlApp As Excel.Application
Dim xlWB As Excel.Workbook
Dim xlWS As Excel.Worksheet

Set db = CurrentDb

' Change the file path to the location of your Excel workbook
Set xlApp = New Excel.Application
Set xlWB = xlApp.Workbooks.Open("C:\ExcelWorkbook.xlsx")

For Each xlWS In xlWB.Sheets
    ' Change the table name to the name of your Access table
    Set tdf = db.CreateTableDef("Sheet" & xlWS.Name)
    tdf.SourceTableName = "ExcelWorkbook.xlsx" & "!" & xlWS.Name & "$"
    tdf.Connect = "Excel 8.0;" & xlWB.FullName
    db.TableDefs.Append tdf
Next xlWS

xlWB.Close
xlApp.Quit

Set tdf = Nothing
Set xlWS = Nothing
Set xlWB = Nothing
Set xlApp = Nothing
Set db = Nothing

End Sub
