' timeclock.vbs
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
    i = 2
    Do While Not IsEmpty(Range("B" & i).Value)
        total = total + DateDiff("h", Range("B" & i).Value, Range("E" & i).Value)
        i = i + 1
    Loop
    
    Range("G2").NumberFormat = "General"
    Range("H2").NumberFormat = "General"
    Range("G2") = total
    Range("H2") = 2080 - total
End Sub