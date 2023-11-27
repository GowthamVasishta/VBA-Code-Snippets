Sub RangeToArrayExample()
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim i As Long, j As Long
    
    ' Define the range you want to convert to an array
    Set dataRange = ThisWorkbook.Sheets("Sheet1").Range("A1:C10")
    
    ' Read the range values into a variant array
    dataArray = dataRange.Value
    
    ' Loop through the array to access each element
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            ' Access each element of the array
            Debug.Print dataArray(i, j)
        Next j
    Next i
    
    ' You now have the data from the range in the 'dataArray' variable
End Sub
