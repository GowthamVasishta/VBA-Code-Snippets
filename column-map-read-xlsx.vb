Sub ImportColumnsByColumnName()
    Dim sourceSheet As Worksheet
    Dim destSheet As Worksheet
    Dim sourceCol As Range
    Dim destCol As Range
    Dim sourceColName As String
    Dim destColName As String
    Dim lastRow As Long
    Dim i As Integer
    
    ' Set your source and destination sheets
    Set sourceSheet = ThisWorkbook.Sheets("SourceSheet")
    Set destSheet = ThisWorkbook.Sheets("DestinationSheet")
    
    ' Define column names to map
    Dim columnsMap As Variant
    columnsMap = Array( _
        Array("SourceColumn1", "DestinationColumn1"), _
        Array("SourceColumn2", "DestinationColumn2"), _
        Array("SourceColumn3", "DestinationColumn3") _
        ' Add more columns as needed
    )
    
    ' Find last row in source sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each column mapping and copy data
    For i = LBound(columnsMap) To UBound(columnsMap)
        sourceColName = columnsMap(i)(0)
        destColName = columnsMap(i)(1)
        
        ' Find source and destination columns by name
        On Error Resume Next
        Set sourceCol = sourceSheet.Rows(1).Find(sourceColName, LookIn:=xlValues, LookAt:=xlWhole)
        Set destCol = destSheet.Rows(1).Find(destColName, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        ' Check if both columns were found
        If Not sourceCol Is Nothing And Not destCol Is Nothing Then
            ' Copy data from source to destination column
            sourceSheet.Range(sourceCol.Offset(1), sourceSheet.Cells(lastRow, sourceCol.Column)).Copy destSheet.Cells(2, destCol.Column)
        Else
            MsgBox "Column " & sourceColName & " or " & destColName & " not found.", vbExclamation
        End If
    Next i
End Sub
