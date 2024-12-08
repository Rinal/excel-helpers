Sub CalculateTotalDaysBetweenDateRanges()
    Dim startRange As Range
    Dim endRange As Range
    Dim resultRange As Range
    Dim startDates() As String
    Dim endDates() As String
    Dim i As Long
    Dim daysDifference As Long
    Dim totalDays As Long
    
    ' Prompt the user to select the range of start dates
    On Error Resume Next
    Set startRange = Application.InputBox("Select the range of start dates:", Type:=8)
    If startRange Is Nothing Then Exit Sub ' Exit if user cancels
    On Error GoTo 0
    
    ' Prompt the user to select the range of end dates
    On Error Resume Next
    Set endRange = Application.InputBox("Select the range of end dates:", Type:=8)
    If endRange Is Nothing Then Exit Sub ' Exit if user cancels
    On Error GoTo 0
    
    ' Prompt the user to select the range for results
    On Error Resume Next
    Set resultRange = Application.InputBox("Select the range for results:", Type:=8)
    If resultRange Is Nothing Then Exit Sub ' Exit if user cancels
    On Error GoTo 0
    
    ' Check if both ranges have the same number of rows
    If startRange.Rows.Count <> endRange.Rows.Count Or startRange.Rows.Count <> resultRange.Rows.Count Then
        MsgBox "The number of rows in start dates, end dates, and results ranges must match.", vbExclamation
        Exit Sub
    End If
    
    ' Loop through each row
    For i = 1 To startRange.Rows.Count
        Dim startCell As Range
        Dim endCell As Range
        Dim resultCell As Range
        
        Set startCell = startRange.Cells(i, 1)
        Set endCell = endRange.Cells(i, 1)
        Set resultCell = resultRange.Cells(i, 1)
        
        ' Split the cell values by line breaks into arrays
        startDates = Split(startCell.Value, vbLf)
        endDates = Split(endCell.Value, vbLf)
        
        ' Check if the number of dates match
        If UBound(startDates) <> UBound(endDates) Then
            MsgBox "The number of dates in each cell does not match in row " & i & ". Please check and try again.", vbExclamation
            Exit Sub
        End If
        
        ' Initialize total days for this row
        totalDays = 0
        
        ' Loop through the date pairs
        For j = 0 To UBound(startDates)
            If IsDate(Trim(startDates(j))) And IsDate(Trim(endDates(j))) Then
                ' Calculate the difference in days, inclusive of the final day
                daysDifference = CDate(Trim(endDates(j))) - CDate(Trim(startDates(j))) + 1
                ' Add to total days for this row
                totalDays = totalDays + daysDifference
            Else
                MsgBox "Invalid date in row " & i & ", date pair " & (j + 1) & ". Calculation stopped.", vbExclamation
                Exit Sub
            End If
        Next j
        
        ' Output the total days to the corresponding result cell
        resultCell.Value = totalDays
    Next i
End Sub
