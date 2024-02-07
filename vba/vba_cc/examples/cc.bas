Attribute VB_Name = "CCTEST"
#Const conDebug = 1

#If conDebuge = 1 Then 'Run debug code
    Debug.Print "oops"
#Elseif conDebug > 1 Then
    Debug.Print "help"
#Else
    #If Cdbl(Abs(conDebug)) < (1 Mod 3) Then
    
        foo = 2
    #Endif
#End If

OtherStuff = 3
