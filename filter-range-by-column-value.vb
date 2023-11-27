Sub FilterRangeByColumnValue()
    Dim ws As Worksheet
    Dim filterRange As Range
    Dim filterColumn As Range
    Dim filterValue As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Define the range to be filtered
    Set filterRange = ws.Range("A1:D100") ' Change this to your range
    
    ' Define the column to filter
    Set filterColumn = ws.Range("B1:B100") ' Change this to your column
    
    ' Set the value to filter by
    filterValue = "YourFilterValue" ' Change this to your desired value
    
    ' Turn off any existing filters
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' Apply the filter based on the value in the column
    filterRange.AutoFilter Field:=filterColumn.Column, Criteria1:=filterValue
    
    ' Your data is now filtered based on the specified value in the specified column
End Sub
