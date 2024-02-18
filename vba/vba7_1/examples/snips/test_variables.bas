Attribute VB_Name = "TestVariables"
Function Foo()
    Dim Bar()
    Dim vbaTests(1 To 99, 1 To 2) As Variant
    ' MS-VBAL 3.2.1 states that final line may be non-terminated
End Function