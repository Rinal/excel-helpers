' Select 2 ranges and copy them into a separate sheet. Sort data based on the ID column



Sub MergeRangesByID_OneIDColumn()
    Dim range1 As Range, range2 As Range
    Dim wsNew As Worksheet
    Dim chunkSize As Long
    Dim totalChunks1 As Long, totalChunks2 As Long
    Dim newRow As Long, i As Long, j As Long
    Dim id1 As String, id2 As String
    Dim matchFound As Boolean
    
    ' Prompt user to select RANGE_1
    On Error Resume Next
    Set range1 = Application.InputBox("Select RANGE_1 (multiple of 4 cells):", Type:=8)
    If range1 Is Nothing Or range1.Cells.Count Mod 4 <> 0 Then
        MsgBox "Please select a range with a multiple of 4 cells for RANGE_1.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt user to select RANGE_2
    Set range2 = Application.InputBox("Select RANGE_2 (multiple of 4 cells):", Type:=8)
    If range2 Is Nothing Or range2.Cells.Count Mod 4 <> 0 Then
        MsgBox "Please select a range with a multiple of 4 cells for RANGE_2.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Validate the ranges
    chunkSize = 4
    totalChunks1 = range1.Cells.Count / chunkSize
    totalChunks2 = range2.Cells.Count / chunkSize
    
    ' Create a new sheet
    Set wsNew = ThisWorkbook.Sheets.Add
    wsNew.Name = "MergedByIDSheet"
    
    ' Add headers to the new sheet
    wsNew.Cells(1, 1).Value = "ID"
    wsNew.Cells(1, 2).Value = "RANGE_1 Data"
    wsNew.Cells(1, 5).Value = "RANGE_2 Data"
    newRow = 2
    
    ' Iterate through chunks in RANGE_1
    For i = 0 To totalChunks1 - 1
        ' Get ID from RANGE_1
        id1 = range1.Cells(i * chunkSize + 1).Value
        
        ' Look for matching ID in RANGE_2
        matchFound = False
        For j = 0 To totalChunks2 - 1
            id2 = range2.Cells(j * chunkSize + 1).Value
            
            If id1 = id2 Then
                matchFound = True
                
                ' Write ID to the new sheet
                wsNew.Cells(newRow, 1).Value = id1
                
                ' Write RANGE_1 chunk to the new sheet (skip first cell - ID)
                For k = 2 To chunkSize
                    wsNew.Cells(newRow, k).Value = range1.Cells(i * chunkSize + k).Value
                Next k
                
                ' Write RANGE_2 chunk to the new sheet (skip first cell - ID)
                For k = 2 To chunkSize
                    wsNew.Cells(newRow, k + 3).Value = range2.Cells(j * chunkSize + k).Value
                Next k
                
                newRow = newRow + 1
                Exit For
            End If
        Next j
        
        ' If no match is found, write only RANGE_1 data
        If Not matchFound Then
            wsNew.Cells(newRow, 1).Value = id1
            For k = 2 To chunkSize
                wsNew.Cells(newRow, k).Value = range1.Cells(i * chunkSize + k).Value
            Next k
            newRow = newRow + 1
        End If
    Next i
    
    MsgBox "Ranges merged successfully with a single ID column!", vbInformation
End Sub