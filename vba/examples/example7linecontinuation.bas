Attribute VB_Name = "LineContinuation"

Sub TestLineContinuedMemberCalls(foo, _
        bar)
'   Valid line continuation syntax with
'   method / property chaining.
    With A.B _
        .B
    End With

    With A _
        .B
    End With

    With A. _
        B
    End With
    Err.Raise Number:=vbObjectError + 42, _
              Description:="Fnord"
End _
    Sub
