Option Explicit

Sub CheckMultiplePeriodsForOverlap_WithSickLeaveAndAssignment()
    Dim cRange As Range
    Dim totalCells As Long, scenarioCount As Long
    Dim scenarioIndex As Long
    
    Dim startWorkCell As Range, endWorkCell As Range
    Dim startVacationCell As Range, endVacationCell As Range
    Dim startSickleaveCell As Range, endSickleaveCell As Range
    Dim startAssignmentCell As Range, endAssignmentCell As Range
    Dim resultCell As Range
    
    Dim startWorkDates() As String, endWorkDates() As String
    Dim startVacationDates() As String, endVacationDates() As String
    Dim startSickleaveDates() As String, endSickleaveDates() As String
    Dim startAssignmentDates() As String, endAssignmentDates() As String
    
    Dim i As Integer, j As Integer
    
    Dim startWork As Date, endWork As Date
    Dim startVacation As Date, endVacation As Date
    Dim startSickleave As Date, endSickleave As Date
    Dim startAssignment As Date, endAssignment As Date
    
    Dim overlapFound As Boolean
    Dim message As String

    ' Prompt the user to select a range whose cell count is a multiple of 9
    On Error Resume Next
    Set cRange = Application.InputBox("Select a range whose cell count is a multiple of 9:" & vbCrLf & _
                                      "Index(mod 9):" & vbCrLf & _
                                      "0: START WORKING" & vbCrLf & _
                                      "1: END WORKING" & vbCrLf & _
                                      "2: START VACATION" & vbCrLf & _
                                      "3: END VACATION" & vbCrLf & _
                                      "4: START SICKLEAVE" & vbCrLf & _
                                      "5: END SICKLEAVE" & vbCrLf & _
                                      "6: START ASSIGNMENT" & vbCrLf & _
                                      "7: END ASSIGNMENT" & vbCrLf & _
                                      "8: RESULT", "Select Range", Type:=8)
    If cRange Is Nothing Then Exit Sub
    On Error GoTo 0
    
    totalCells = cRange.Cells.Count
    If totalCells = 0 Or (totalCells Mod 9) <> 0 Then
        MsgBox "The selected range must have a number of cells that is a multiple of 9.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If
    
    scenarioCount = totalCells \ 9
    
    ' Process each scenario
    For scenarioIndex = 0 To scenarioCount - 1
        ' Identify the cells for this scenario
        Set startWorkCell = cRange.Cells(scenarioIndex * 9 + 1)       ' mod 9 = 0
        Set endWorkCell = cRange.Cells(scenarioIndex * 9 + 2)         ' mod 9 = 1
        Set startVacationCell = cRange.Cells(scenarioIndex * 9 + 3)   ' mod 9 = 2
        Set endVacationCell = cRange.Cells(scenarioIndex * 9 + 4)     ' mod 9 = 3
        Set startSickleaveCell = cRange.Cells(scenarioIndex * 9 + 5)  ' mod 9 = 4
        Set endSickleaveCell = cRange.Cells(scenarioIndex * 9 + 6)    ' mod 9 = 5
        Set startAssignmentCell = cRange.Cells(scenarioIndex * 9 + 7) ' mod 9 = 6
        Set endAssignmentCell = cRange.Cells(scenarioIndex * 9 + 8)   ' mod 9 = 7
        Set resultCell = cRange.Cells(scenarioIndex * 9 + 9)          ' mod 9 = 8
        
        ' Clear the result cell and formatting first
        resultCell.ClearContents
        resultCell.Interior.ColorIndex = xlColorIndexNone

        ' Split dates in each cell into arrays (multiple lines per cell)
        startWorkDates = Split(startWorkCell.Value, vbLf)
        endWorkDates = Split(endWorkCell.Value, vbLf)
        startVacationDates = Split(startVacationCell.Value, vbLf)
        endVacationDates = Split(endVacationCell.Value, vbLf)
        startSickleaveDates = Split(startSickleaveCell.Value, vbLf)
        endSickleaveDates = Split(endSickleaveCell.Value, vbLf)
        startAssignmentDates = Split(startAssignmentCell.Value, vbLf)
        endAssignmentDates = Split(endAssignmentCell.Value, vbLf)

        ' Validate counts
        If UBound(startWorkDates) <> UBound(endWorkDates) Then
            resultCell.Value = "Error: Start/End Working dates do not match in count."
            resultCell.Interior.Color = vbRed
            GoTo NextScenario
        End If
        
        If UBound(startVacationDates) <> UBound(endVacationDates) Then
            resultCell.Value = "Error: Start/End Vacation dates do not match in count."
            resultCell.Interior.Color = vbRed
            GoTo NextScenario
        End If
        
        If UBound(startSickleaveDates) <> UBound(endSickleaveDates) Then
            resultCell.Value = "Error: Start/End Sickleave dates do not match in count."
            resultCell.Interior.Color = vbRed
            GoTo NextScenario
        End If
        
        If UBound(startAssignmentDates) <> UBound(endAssignmentDates) Then
            resultCell.Value = "Error: Start/End Assignment dates do not match in count."
            resultCell.Interior.Color = vbRed
            GoTo NextScenario
        End If

        overlapFound = False
        message = "Overlapping Periods Found:" & vbCrLf

        ' Check each working period against each of the other periods (Vacation, Sickleave, Assignment)
        ' A helper function approach would be cleaner, but for clarity, we’ll just inline the checks.
        
        ' Define a local routine for overlap checking to reduce repetition
        Dim CheckOverlap As Boolean
        CheckOverlap = CheckAndReportOverlap(startWorkDates, endWorkDates, startVacationDates, endVacationDates, "Vacation", resultCell, message, overlapFound)
        If CheckOverlap = False Then GoTo NextScenario
        
        CheckOverlap = CheckAndReportOverlap(startWorkDates, endWorkDates, startSickleaveDates, endSickleaveDates, "Sickleave", resultCell, message, overlapFound)
        If CheckOverlap = False Then GoTo NextScenario
        
        CheckOverlap = CheckAndReportOverlap(startWorkDates, endWorkDates, startAssignmentDates, endAssignmentDates, "Assignment", resultCell, message, overlapFound)
        If CheckOverlap = False Then GoTo NextScenario

        ' After checks
        If overlapFound Then
            resultCell.Value = message
            resultCell.Interior.Color = vbRed
        Else
            resultCell.Value = "OK"
            resultCell.Interior.Color = vbGreen
        End If

NextScenario:
        ' Move to the next scenario
    Next scenarioIndex
End Sub

Private Function CheckAndReportOverlap(startWorkDates() As String, endWorkDates() As String, _
                                       startOtherDates() As String, endOtherDates() As String, _
                                       periodName As String, resultCell As Range, _
                                       ByRef message As String, ByRef overlapFound As Boolean) As Boolean
    Dim i As Integer, j As Integer
    Dim startWork As Date, endWork As Date
    Dim startOther As Date, endOther As Date
    
    CheckAndReportOverlap = True ' Assume success unless an error occurs
    
    ' Check each working period against the other period
    For i = LBound(startWorkDates) To UBound(startWorkDates)
        If Not IsDate(startWorkDates(i)) Or Not IsDate(endWorkDates(i)) Then
            resultCell.Value = "Error: Invalid working period dates."
            resultCell.Interior.Color = vbRed
            CheckAndReportOverlap = False
            Exit Function
        End If
        
        startWork = CDate(startWorkDates(i))
        endWork = CDate(endWorkDates(i))
        
        If startWork > endWork Then
            resultCell.Value = "Error: A working period's start date is later than its end date."
            resultCell.Interior.Color = vbRed
            CheckAndReportOverlap = False
            Exit Function
        End If
        
        For j = LBound(startOtherDates) To UBound(startOtherDates)
            If UBound(startOtherDates) >= 0 Then
                ' Validate other period dates
                If Not IsDate(startOtherDates(j)) Or Not IsDate(endOtherDates(j)) Then
                    resultCell.Value = "Error: Invalid " & periodName & " period dates."
                    resultCell.Interior.Color = vbRed
                    CheckAndReportOverlap = False
                    Exit Function
                End If
                
                startOther = CDate(startOtherDates(j))
                endOther = CDate(endOtherDates(j))
                
                If startOther > endOther Then
                    resultCell.Value = "Error: A " & periodName & " period's start date is later than its end date."
                    resultCell.Interior.Color = vbRed
                    CheckAndReportOverlap = False
                    Exit Function
                End If
                
                ' Check for overlap
                If (startOther <= endWork) And (endOther >= startWork) Then
                    overlapFound = True
                    message = message & "Working Period: " & Format(startWork, "dd.mm.yyyy") & " to " & Format(endWork, "dd.mm.yyyy") & _
                              ", " & periodName & " Period: " & Format(startOther, "dd.mm.yyyy") & " to " & Format(endOther, "dd.mm.yyyy") & vbCrLf
                End If
            End If
        Next j
    Next i
End Function
