Sub ImportColumnsByColumnName()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    Dim lastRow As Long
    Dim colName As String
    Dim importCol As Range
    Dim destCol As Range
    Dim i As Integer
    
    ' Open the source workbook
    Set sourceWB = Workbooks.Open("C:\Users\user\archive\iris.xlsx")
    Set sourceWS = sourceWB.Sheets("iris") ' Change "Sheet1" to your sheet name
    
    ' Set your destination workbook and sheet
    Set wb = ThisWorkbook ' Destination workbook (the workbook with the VBA code)
    Set ws = wb.Sheets("Sheet1") ' Change "DestinationSheet" to your sheet name
    
    ' Clear existing data in the destination sheet (optional)
    ws.Cells.Clear
    
    ' Column names you want to import (change these to match your column names)
    Dim columnNames As Variant
    columnNames = Array("sepal_length", "sepal_width", "petal_length", "petal_width", "species") ' Change these to your column names
    
    ' Loop through each column name
    For i = LBound(columnNames) To UBound(columnNames)
        colName = columnNames(i)
        
        ' Find the column by its header name
        On Error Resume Next
        Set importCol = sourceWS.Rows(1).Find(colName, LookIn:=xlValues, lookat:=xlWhole)
        On Error GoTo 0
        
        ' If the column name is found, copy its data to the destination sheet
        If Not importCol Is Nothing Then
            Set destCol = ws.Cells(1, i + 1) ' Place data in the destination sheet starting from column 1
            
            ' Find the last row with data in the source column
            lastRow = sourceWS.Cells(sourceWS.Rows.Count, importCol.Column).End(xlUp).Row
            
            ' Copy data from source to destination
            sourceWS.Range(importCol.Offset(1), sourceWS.Cells(lastRow, importCol.Column)).Copy destCol
        Else
            MsgBox "Column '" & colName & "' not found.", vbExclamation
        End If
    Next i
    
    ' Close the source workbook without saving changes
    sourceWB.Close SaveChanges:=False
End Sub

