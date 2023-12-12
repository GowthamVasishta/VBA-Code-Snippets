Sub FilterDataBasedOnRanges()
    Dim wbFilter As Workbook
    Dim wsFilter As Worksheet
    Dim wsData As Worksheet
    Dim wsOutput As Worksheet
    Dim filterRange As Range
    Dim lastRow As Long
    Dim col As Long
    Dim filterColumn As Range
    Dim cell As Range
    Dim dataRange As Range
    Dim outputRow As Long
    
    ' Open the workbook containing filter ranges
    Set wbFilter = ThisWorkbook ' Update with the name of your workbook
    Set wsFilter = wbFilter.Sheets("Filters") ' Update with the name of your filter sheet
    
    ' Set the data sheet (sheet where data to be filtered is located)
    Set wsData = wbFilter.Sheets("DataSheet") ' Update with the name of your data sheet
    
    ' Set the output sheet
    Set wsOutput = wbFilter.Sheets("Output") ' Update with the name of your output sheet
    
    ' Find the last row with data in the data sheet
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Define the range containing filter criteria
    Set filterRange = wsFilter.Range("A1").CurrentRegion ' Assumes filter range starts from A1
    
    ' Loop through each column in the filter range
    For col = 1 To filterRange.Columns.Count
        Set filterColumn = filterRange.Columns(col)
        
        ' Apply filter for each column on the data
        For Each cell In filterColumn.Cells(2, 1).Resize(filterColumn.Rows.Count - 1, 1)
            Set dataRange = wsData.Range(wsData.Cells(1, col), wsData.Cells(lastRow, col))
            
            If WorksheetFunction.CountIf(dataRange, cell.Value) > 0 Then
                dataRange.AutoFilter Field:=col, Criteria1:=cell.Value
                
                ' Copy filtered data to output sheet
                outputRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1
                dataRange.SpecialCells(xlCellTypeVisible).EntireRow.Copy wsOutput.Cells(outputRow, 1)
                
                ' Turn off the filter
                wsData.AutoFilterMode = False
            End If
        Next cell
    Next col
End Sub
