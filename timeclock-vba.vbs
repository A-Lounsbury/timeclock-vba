' timeclock.vbs
' Author: Andrew W. Lounsbury
' Date: 3/10/24
' Description: a simple time clock for keeping tract of hours worked

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
    
    If IsEmpty(Range("B" & i).Value) Then
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
    
    If IsEmpty(Range("B" & i).Value) Then
        MsgBox ("Error: cannot end lunch before starting shift")
    ElseIf IsEmpty(Range("C" & i).Value) Then
        MsgBox ("Error: cannot end lunch before starting lunch")
    Else
        ' inserting the time into the empty cell in column B
        Dim t As Date
        t = Time()
        Range("D" & Index) = t
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
    
    If IsEmpty(Range("B" & i).Value) Then
        MsgBox ("Error: cannot end shift before starting shift")
    ElseIf IsEmpty(Range("C" & i).Value) Then
        MsgBox ("Error: cannot end shift before starting lunch")
    ElseIf IsEmpty(Range("D" & i).Value) Then
        MsgBox ("Error: cannot end shift before ending lunch")
    Else
        ' inserting the time into the empty cell in column B
        Dim t As Date
        t = Time()
        Range("E" & Index) = t
    Else
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
    totalHours = Int(totalMinutes / 60)
    remainingMinutes = totalMinutes Mod 60
    expectedHours = numDays * 8
    
    ' formatting the cells
    Range("G2").NumberFormat = "General"
    Range("H2").NumberFormat = "General"
    Range("I2").NumberFormat = "General"
    
    ' inserting output into cells
    If remainingMinutes <> 0 Then
        Range("G2") = totalHours & "hr " & remainingMinutes & "min"
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
    
    If Int(differenceMinutes / 60) <> 0 Then
        If differenceMinutes < 0 Then
            Range("I2") = CInt(differenceMinutes / 60) & "hr " & Abs(differenceMinutes Mod 60) & "min"
        ElseIf differenceMinutes > 0 Then
            Range("I2") = CInt(differenceMinutes / 60) & "hr " & differenceMinutes Mod 60 & "min"
        Else ' just hours, no extra minutes
            Range("I2") = CInt(differenceMinutes / 60) & "hr"
        End If
    Else ' just minutes
        Range("I2") = differenceMinutes & "min"
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