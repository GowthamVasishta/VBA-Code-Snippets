Sub FilterColumnWithValue()
    Dim ws As Worksheet
    Dim filterRange As Range
    Dim columnToFilter As Range
    Dim filterValue As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Define the range to apply the filter (excluding headers)
    Set filterRange = ws.Range("A1").CurrentRegion ' Change this range to fit your data
    
    ' Define the column to filter
    Set columnToFilter = filterRange.Columns(1) ' Filter based on the first column (Column A)
    
    ' Set the value to filter
    filterValue = "YourValue" ' Replace 'YourValue' with the value you want to filter
    
    ' Apply the filter
    filterRange.AutoFilter Field:=columnToFilter.Column, Criteria1:=filterValue
    
    ' Example: To clear the filter
    ' filterRange.AutoFilter Field:=columnToFilter.Column
    
End Sub
