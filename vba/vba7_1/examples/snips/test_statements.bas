Attribute VB_Name = "statements"

Public Static Function Module()
    ' Call
    Call bar
    bar.Type baz:="7", foo:=5
    Call baz(4)

    'While
    While True
    Wend

    ' For
    For i = 0 To 7 Step 2
    Next
    
    For i = 0 To 7 Step 2
    Next i

    For i = 0 To 7 Step 2
        For j = 0 To 7
            sum = sum + i * j
        Next j
        for each j in k
next j, i
End Function
