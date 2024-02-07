Attribute VB_Name = "CCTEST"
#Const conDebug = 1

#If conDebuge = 1 Then 'Run debug code
    Debug.Print "oops"
#Elseif conDebug > 1 Then
    Debug.Print "help"
#Else
    #If conDebug < 1 Then
    
        foo = 2
    #Endif
#End If

OtherStuff = 3
