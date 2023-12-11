Function selectModule()
    
    Dim startRow As Integer
    Dim lastColumn As String
    
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim checkRange As Range
    Dim cell As Range
    
    Dim moduleName As String
    Dim outputSheet As String
    
    ' control sheet configuration
    startRow = 4
    lastColumn = "k"
    
    ' Set the control worksheet
    Set ws = ThisWorkbook.Sheets("Control")
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Define the control range from A4 to the last row in column F
    Set ctrlRange = ws.Range("A" & startRow & ":" & lastColumn & lastRow)
    
    ' Loop through each row in the control range
    For Each ctrlRow In ctrlRange.Rows
        
        ' Check if row needs to be skipped
        If ctrlRow.Cells(1, 11).Value <> True Then
            
            ' Update status Messages
            Application.StatusBar = "Executing: " & ctrlRow.Cells(1, 3)
            ctrlRow.Cells(1, 12).Value = "Step executed at " & Now()
            
            ' Get inputs and execute modules
            moduleName = ctrlRow.Cells(1, 4).Value
            outputSheet = ctrlRow.Cells(1, 9).Value
            
            ' check the module name and call it
            If moduleName = "clearSheets" Then
                On Error Resume Next
                Application.Run moduleName, outputSheet
                If Err.Number <> 0 Then
                    MsgBox "Error calling the module: " & moduleName
                End If
                On Error GoTo 0
            End If
                
            
        Else
            ' Update status Messages
            Application.StatusBar = "Step Skipped"
            ctrlRow.Cells(1, 12).Value = "Step Skipped"
        End If
        
        
        Application.StatusBar = False
        
    Next ctrlRow
End Function

Sub ClearSheets(sheetNamesString As String)
    Dim sheetNames() As String
    ' Splitting the input string by semicolon
    sheetNames = Split(sheetNamesString, ";")

    Dim ws As Worksheet
    Dim i As Integer

    For i = LBound(sheetNames) To UBound(sheetNames)
        ' Trimming extra spaces around sheet names
        Set ws = ThisWorkbook.Sheets(Trim(sheetNames(i)))
        If Not ws Is Nothing Then
            ' Checking if sheet has more than one row
            If ws.Rows.Count > 1 Then
                ' Clearing data except for the first row
                ws.Range("A2:" & ws.Cells(ws.Rows.Count, ws.Columns.Count).Address).ClearContents
            End If
        End If
    Next i
End Sub

Sub ReadAndMapData()
    Dim inputFilePath As String
    Dim columnMapName As String
    Dim mappingTable As Range
    Dim inputWB As Workbook
    Dim outputWB As Workbook
    Dim inputWS As Worksheet
    Dim outputWS As Worksheet
    Dim lastRowInput As Long
    Dim lastRowOutput As Long
    Dim i As Long, j As Long
    Dim columnFound As Boolean
    
    ' Get input file path and column map name from Excel cells
    inputFilePath = ThisWorkbook.Sheets("Sheet1").Range("A1").Value ' Change the cell reference as needed
    columnMapName = ThisWorkbook.Sheets("Sheet1").Range("B1").Value ' Change the cell reference as needed
    
    ' Open input workbook
    Set inputWB = Workbooks.Open(inputFilePath)
    Set inputWS = inputWB.Sheets(1) ' Change the sheet index as needed
    
    ' Get mapping table
    Set outputWB = ThisWorkbook
    Set outputWS = outputWB.Sheets("Mapping") ' Change the sheet name as needed
    Set mappingTable = outputWS.Range("A1").CurrentRegion
    
    ' Find the column map name in the mapping table
    For i = 1 To mappingTable.Rows.Count
        If mappingTable.Cells(i, 1).Value = columnMapName Then
            ' Get last row of input and output sheets
            lastRowInput = inputWS.Cells(inputWS.Rows.Count, 1).End(xlUp).Row
            lastRowOutput = outputWS.Cells(outputWS.Rows.Count, 1).End(xlUp).Row
            
            ' Loop through input data and map columns to output sheet
            For j = 1 To lastRowInput
                For Each cell In mappingTable.Rows(i).Offset(0, 1).Cells
                    columnFound = False
                    For k = 1 To inputWS.Cells(1, inputWS.Columns.Count).End(xlToLeft).Column
                        If inputWS.Cells(1, k).Value = cell.Value Then
                            outputWS.Cells(lastRowOutput + 1, cell.Column).Value = inputWS.Cells(j, k).Value
                            columnFound = True
                            Exit For
                        End If
                    Next k
                    If Not columnFound Then
                        outputWS.Cells(lastRowOutput + 1, cell.Column).Value = inputWS.Cells(j, 1).Value ' Input column name populated for all rows
                    End If
                Next cell
                lastRowOutput = lastRowOutput + 1
            Next j
            Exit For
        End If
    Next i
    
    ' Close input workbook without saving changes
    inputWB.Close SaveChanges:=False
End Sub

