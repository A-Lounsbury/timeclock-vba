' timesheet.vbs
' Author: Andrew W. Lounsbury
' Date: 3/19/24
' Description: a simple time clock for keeping track of hours worked

Sub startShift():
    ' Inserts the current date and the current time in the appropriate cells in columns A and B
    '
    ' insert the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim emptyCellFound As Boolean
    Dim i As Integer
    Dim Index As Integer
    Index = -1
    i = 2
    emptyCellFound = False
    Do While Not emptyCellFound
        If IsEmpty(Range("A" & i).Value) Then
            Index = i
            emptyCellFound = True
        End If
        i = i + 1
    Loop
    
    ' inserting the date into that Cell
    Dim d As Date
    d = Date
    Range("A" & Index) = d
    
    ' inserting the time into the empty cell in column B
    Dim t As Date
    t = Time()
    Range("B" & Index) = t
End Sub

Sub startLunch():
    ' Inserts the current time in the appropriate cell in column C
    '
    ' inserting the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim emptyCellFound As Boolean
    Dim i As Integer
    Dim Index As Integer
    Index = -1
    i = 2
    emptyCellFound = False
    Do While Not emptyCellFound
        If IsEmpty(Range("C" & i).Value) Then
            Index = i
            emptyCellFound = True
        End If
        i = i + 1
    Loop
    
    If IsEmpty(Range("B" & Index).Value) Then
        MsgBox ("Error: cannot start lunch before starting shift")
    Else
        ' inserting the time into the empty cell in column B
        Dim t As Date
        t = Time()
        Range("C" & Index) = t
    End If
End Sub

Sub endLunch():
    ' Inserts the current time in the appropriate cell in column D
    '
    ' inserting the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim emptyCellFound As Boolean
    Dim i As Integer
    Dim Index As Integer
    Dim morningStretch As Integer
    Dim targetEndShift As Date
    Index = -1
    i = 2
    emptyCellFound = False
    Do While Not emptyCellFound
        If IsEmpty(Range("D" & i).Value) Then
            Index = i
            emptyCellFound = True
        End If
        i = i + 1
    Loop
    
    If IsEmpty(Range("B" & Index).Value) Then
        MsgBox ("Error: cannot end lunch before starting shift")
    ElseIf IsEmpty(Range("C" & Index).Value) Then
        MsgBox ("Error: cannot end lunch before starting lunch")
    Else
        ' inserting the time into the empty cell in column B
        Dim t As Date
        t = Time()
        Range("D" & Index) = t
        
        Dim shiftStart As Date
        shiftStart = Range("B" & i).Value
        Dim lunchStart As Date
        lunchStart = Range("C" & i).Value
        
        morningStretch = DateDiff("n", Range("B" & Index).Value, Range("C" & Index).Value)
        Debug.Print ("morningStretch: " & morningStretch)
        
        Dim timeLeft As Integer
        Dim hoursLeft As Double
        Dim minutesLeft As Integer
        timeLeft = Range("B" & Index).Value + ((8 * 60) - morningStretch)
        Debug.Print ("timeLeft: " & timeLeft)
        hoursLeft = CInt(timeLeft / 60)
        minutesLeft = timeLeft Mod 60
        
        Dim newHours As Integer
        Dim newMinutes As Integer
        Dim seconds As Integer
        Dim oldHours As Integer
        Dim oldMinutes As Integer
        
        oldHours = Hour(Range("D" & Index).Value)
        oldMinutes = Minute(Range("D" & Index).Value)
        seconds = Second(Range("D" & Index).Value)
        newHours = oldHours + hoursLeft
        newMinutes = oldMinutes + minutesLeft
        
        targetShiftEnd = newHours & ":" & newMinutes & ":" & seconds & " PM"
        
        MsgBox ("Target Shift End: " & targetShiftEnd)
    End If
    
End Sub

Sub endShift():
    ' Inserts the current time in the appropriate cell in column E
    '
    ' inserting the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim emptyCellFound As Boolean
    Dim i As Integer
    Dim Index As Integer
    Index = -1
    i = 2
    emptyCellFound = False
    Do While Not emptyCellFound
        If IsEmpty(Range("E" & i).Value) Then
            Index = i
            emptyCellFound = True
        End If
        i = i + 1
    Loop
    
    If IsEmpty(Range("B" & Index).Value) Then
        MsgBox ("Error: cannot end shift before starting shift")
    ElseIf IsEmpty(Range("C" & Index).Value) Then
        MsgBox ("Error: cannot end shift before starting lunch")
    ElseIf IsEmpty(Range("D" & Index).Value) Then
        MsgBox ("Error: cannot end shift before ending lunch")
    Else
        ' inserting the time into the empty cell in column B
        Dim t As Date
        t = Time()
        Range("E" & Index) = t
    End If
End Sub

Sub computeTotal():
    ' computes the total number of hours worked and the difference between the expected number of hours and displays them in the appropriate cells
    Dim differenceHours As Integer
    Dim differenceMinutes As Integer
    Dim i As Integer
    Dim numDays As Integer
    Dim totalHours As Integer
    Dim totalMinutes As Integer
    Dim remainingMinutes As Integer
    expectedHours = -1
    numDays = 0
    total = 0
    i = 2
    Do While Not IsEmpty(Range("B" & i).Value)
        ' Adding the length of time from the start of the shift to the end of the shift, subtracting the time spent at lunch
        totalMinutes = totalMinutes + DateDiff("n", Range("B" & i).Value, Range("E" & i).Value) - DateDiff("n", Range("C" & i).Value, Range("D" & i).Value)
        numDays = numDays + 1
        i = i + 1
    Loop
    ' converting the total number of minutes to hours and minutes
    Debug.Print (totalMinutes)
    totalHours = Int(totalMinutes / 60)
    remainingMinutes = totalMinutes Mod 60
    expectedHours = numDays * 8
    expectedMinutes = expectedHours * 60
    
    differenceMinutes = totalMinutes - expectedMinutes
    Debug.Print ("differenceMinutes: " & differenceMinutes)
    
    If Abs(differenceMinutes) >= 60 Then
        diffHours = CInt(differenceMinutes / 60)
    Else
        diffHours = 0
    End If
    diffMinutes = differenceMinutes Mod 60
    Debug.Print ("diffHours: " & diffHours)
    Debug.Print ("diffMinutes: " & diffMinutes)
    
    ' formatting the cells
    Range("G2").NumberFormat = "General"
    Range("H2").NumberFormat = "General"
    Range("I2").NumberFormat = "General"
    
    ' inserting output into cells
    If remainingMinutes <> 0 Then
        If totalHours <> 0 Then
            Range("G2") = totalHours & "hr " & remainingMinutes & "min"
        Else
            Range("G2") = remainingMinutes & "min"
        End If
    Else
        Range("G2") = totalHours & "hr"
    End If
    Range("H2") = expectedHours & "hr"
    
    differenceMinutes = totalMinutes - (expectedHours * 60)
    If Abs(differenceMinutes) > 60 Then
        differenceHours = Int((totalHours - expectedHours) / 60)
    Else
        differenceHours = 0
    End If
    
    If differenceMinutes >= 60 Then
        differenceHours = CInt(differenceMinutes / 60)
    ElseIf differenceMinutes < 60 Then
        differenceHours = -7
    End If
    
    ' Entering Difference column
    If diffHours <> 0 Then
        If diffMinutes < 0 Then
            Range("I2") = diffHours & "hr " & Abs(diffMinutes Mod 60) & "min"
        ElseIf diffMinutes > 0 Then
            Range("I2") = diffHours & "hr " & diffMinutes Mod 60 & "min"
        Else ' just hours, no extra minutes
            Range("I2") = diffHours & "hr"
        End If
    Else ' just minutes
        Range("I2") = diffMinutes & "min"
    End If
    
    ' Coloring cell I2 according to its value
    If differenceMinutes < 0 Then
        Range("I2").Interior.Color = RGB(255, 128, 128) ' red
    ElseIf differenceMinutes > 0 Then
        Range("I2").Interior.Color = RGB(51, 153, 102) 'green
    Else ' on track
        Range("I2").Interior.ColorIndex = 6 'yellow
    End If
End Sub