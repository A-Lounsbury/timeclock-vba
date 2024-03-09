' timeclock.vbs
' Author: Andrew W. Lounsbury
' Date: 3/7/24
' Description: a simple time clock for keeping tract of hours worked

Sub startShift():
    ' Inserts the current date and the current time in the appropriate cells in columns A and B
    '
    ' insert the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim Index As Integer
    Dim i As Integer
    Dim emptyCellFound As Boolean
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
    ' insert the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim Index As Integer
    Dim i As Integer
    Dim emptyCellFound As Boolean
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
    
    ' inserting the time into the empty cell in column B
    Dim t As Date
    t = Time()
    Range("C" & Index) = t
End Sub

Sub endLunch():
    ' Inserts the current time in the appropriate cell in column D
    '
    ' inserting the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim Index As Integer
    Dim i As Integer
    Dim emptyCellFound As Boolean
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
    
    ' inserting the time into the empty cell in column B
    Dim t As Date
    t = Time()
    Range("D" & Index) = t
End Sub

Sub endShift():
    ' Inserts the current time in the appropriate cell in column E
    '
    ' inserting the date into the first empty cell in column A
    ' getting the first empty cell in column A
    Dim Index As Integer
    Dim i As Integer
    Dim emptyCellFound As Boolean
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
    
    ' inserting the time into the empty cell in column B
    Dim t As Date
    t = Time()
    Range("E" & Index) = t
End Sub

Sub computeTotal():
    ' computes the total number of hours worked
    Dim i As Integer
    Dim total As Double
    Dim numDays As Integer
    Dim expectedHours
    expectedHours = -1
    numDays = 0
    i = 2
    Do While Not IsEmpty(Range("B" & i).Value)
        ' Adding the length of time from the start of the shift to the end of the shift, subtracting the time spent at lunch
        total = total + DateDiff("h", Range("B" & i).Value, Range("E" & i).Value) - DateDiff("h", Range("C" & i).Value, Range("D" & i).Value)
        numDays = numDays + 1
        i = i + 1
    Loop
    expectedHours = numDays * 8
    
    Range("G2").NumberFormat = "General"
    Range("H2").NumberFormat = "General"
    Range("I2").NumberFormat = "General"
    Range("G2") = total
    Range("H2") = expectedHours
    Range("I2") = total - expectedHours
    If Range("I2").Value < 0 Then
        Range("I2").Interior.Color = RGB(255, 128, 128)
    ElseIf Range("I2").Value > 0 Then
        Range("I2").Interior.Color = RGB(51, 153, 102)
    Else ' on track
        Range("I2").Interior.ColorIndex = 6
    End If
End Sub
