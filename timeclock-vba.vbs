' timeclock-vba.vbs
' Author: Andrew Lounsbury
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
    ' UNFINISHED; computes the total number of hours worked
    Dim i As Integer
    Dim total As Double
    Debug.Print ("TYPE: " & VarType(total))
    Dim hr
    hr = Minute(total)
    Debug.Print ("mins: " & hr)
    Debug.Print ("total: " & total)
    Debug.Print ("dbl total: " & CDbl(total))
    i = 2
    Dim Add As Double
    Do While Not IsEmpty(Range("B" & i).Value)
        Debug.Print ("ADDING: " & CDbl(Minute(Range("E" & i).Value)) - CDbl(Minute(Range("B" & i).Value)))
        total = total + CDbl(DateDiff("n", Minute(Range("B" & i).Value), Minute(Range("E" & i).Value)))
        i = i + 1
    Loop
    Debug.Print ("TYPE AGAIN: " & VarType(total))
    Range("G" & 2).Value = CDbl(total)
    Range("H" & 2).Value = "PLACEHOLDER 2"
End Sub