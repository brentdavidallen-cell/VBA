Attribute VB_Name = "modVba2mdl"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
' Visual Basic for Applications Routine to communicate with MicroStation MDL
' through shared(published) structure and mdl functions and commands. Load the
' vba2mdl MDL example prior to running this.
'
' MDL Routines:
'   mdl2VbaInfo_function (double, long, short)
'       - prints list of values and adds 3 to long parameter and returns result
'   mdl2VbaInfo_printStructure ()
'       - prints the shared structure values
' MDL Commands:
'   vba2mdlCommand ( char*)
'       - copies input string into shared structure string member
'
Sub Vba2MdlTest()
    Dim result As Long
' Variables corresponding to shared structure
'                                   struct Mdl2VbaInfo
    Dim asciiString As String     ' char    asciiString[200]
    Dim wideString As String      ' MSWChar wideString[200]
    Dim s1 As Integer             ' short   s1
    Dim d1 As Double              ' double  d1
    Dim i1 As Long                ' long    i1
    Debug.Print ""
    Debug.Print "MDL - VBA Communication Example"
    Debug.Print "-------------------------------"
    Debug.Print ""
    
' Access shared structure using C expression evaluation
    d1 = GetCExpressionValue("mdl2VbaInfo.d1", "vba2mdl")
    s1 = GetCExpressionValue("mdl2VbaInfo.s1", "vba2mdl")
    i1 = GetCExpressionValue("mdl2VbaInfo.i1", "vba2mdl")
    asciiString = GetCExpressionValue("mdl2VbaInfo.asciiString", "vba2mdl")
    wideString = GetCExpressionValue("mdl2VbaInfo.wideString", "vba2mdl")
'
' Call mdl function to print out initial values to text window
'
   GetCExpressionValue "mdl2VbaInfo_printStructure()", "vba2mdl"
   
' Call mdl function to set and print list of values from shared structure and
' modifying the long structure member ( adding 3 to i1 and returning result )
    result = GetCExpressionValue("mdl2VbaInfo_function (100.5, 10, 20)", "vba2mdl")
    
    If result = 13 Then
        Debug.Print "Test for executing mdl2VbaInfo_function passed"
    Else
        Debug.Print "mdl2VbaInfo_function failed"
    End If
        
    If GetCExpressionValue("mdl2VbaInfo.d1", "vba2mdl") = 100.5 Then
        Debug.Print ""
        Debug.Print "Test for double structure member update passed"
        Debug.Print "Initial Double value = " & d1
        Debug.Print "New Double value = " & GetCExpressionValue("mdl2VbaInfo.d1", "vba2mdl")
    Else
        Debug.Print "mdl2VbaInfo_function double failed"
    End If
    If GetCExpressionValue("mdl2VbaInfo.i1", "vba2mdl") = 10 Then
        Debug.Print ""
        Debug.Print "Test for long structure member  passed"
        Debug.Print "Initial Long value = " & i1
        Debug.Print "New Long value = " & GetCExpressionValue("mdl2VbaInfo.i1", "vba2mdl")
    Else
        Debug.Print "mdl2VbaInfo_function long failed"
    End If
    If GetCExpressionValue("mdl2VbaInfo.s1", "vba2mdl") = 20 Then
        Debug.Print ""
        Debug.Print "Test for short structure member passed"
        Debug.Print "Initial Short value = " & s1
        Debug.Print "New Short value = " & GetCExpressionValue("mdl2VbaInfo.s1", "vba2mdl")
    Else
        Debug.Print "mdl2VbaInfo_function short failed"
    End If

' Use SetCExpressionValue to directly modify shared structure members double, long, short

    SetCExpressionValue "mdl2VbaInfo.d1", 200.5, "vba2mdl"
    SetCExpressionValue "mdl2VbaInfo.i1", 2000, "vba2mdl"
    SetCExpressionValue "mdl2VbaInfo.s1", 2, "vba2mdl"
    
    If GetCExpressionValue("mdl2VbaInfo.d1", "vba2mdl") = 200.5 Then
        Debug.Print ""
        Debug.Print "Test to set double structure member passed"
        Debug.Print "Initial Double value = " & d1
        Debug.Print "New Double value = " & GetCExpressionValue("mdl2VbaInfo.d1", "vba2mdl")
    Else
        Debug.Print "SetCExpressionValue double failed"
    End If
    If GetCExpressionValue("mdl2VbaInfo.i1", "vba2mdl") = 2000 Then
        Debug.Print ""
        Debug.Print "Test to set long structure member passed"
        Debug.Print "Initial Long value = " & i1
        Debug.Print "New Long value = " & GetCExpressionValue("mdl2VbaInfo.i1", "vba2mdl")
      Else
        Debug.Print "SetCExpressionValue long failed"
    End If
    If GetCExpressionValue("mdl2VbaInfo.s1", "vba2mdl") = 2 Then
        Debug.Print ""
        Debug.Print "Test to set short structure member passed"
        Debug.Print "Initial Short value = " & s1
        Debug.Print "New Short value = " & GetCExpressionValue("mdl2VbaInfo.s1", "vba2mdl")
    Else
        Debug.Print "SetCExpressionValue short failed"
    End If
'
' Call mdl function to print out intermediate values to text window
'
   GetCExpressionValue "mdl2VbaInfo_printStructure()", "vba2mdl"
'
' Use SetCExpressionValue to directly modify shared structure members - string, Unicode string
'
    SetCExpressionValue "mdl2VbaInfo.asciiString", "FOR THE ASCII STRING", "vba2mdl"
    SetCExpressionValue "mdl2VbaInfo.wideString", "FOR THE WIDE STRING", "vba2mdl"
    
    If GetCExpressionValue("mdl2VbaInfo.asciiString", "vba2mdl") = "FOR THE ASCII STRING" Then
        Debug.Print ""
        Debug.Print "Test to set ASCII string structure member passed"
        Debug.Print "Initial Ascii String = " & asciiString
        Debug.Print "New Ascii String = " & GetCExpressionValue("mdl2VbaInfo.asciiString", "vba2mdl")
     Else
        Debug.Print "SetCExpressionValue ascii string failed"
    End If
    If GetCExpressionValue("mdl2VbaInfo.wideString", "vba2mdl") = "FOR THE WIDE STRING" Then
        Debug.Print ""
        Debug.Print "Test to set Unicode string structure member passed"
        Debug.Print "Initial Wide(Unicode) String = " & wideString
        Debug.Print "New Wide(Unicode) String = " & GetCExpressionValue("mdl2VbaInfo.wideString", "vba2mdl")
    Else
        Debug.Print "SetCExpressionValue Wide(Unicode) string failed"
    End If
'
' Call mdl command through CadInputQueue keyin to copy ascii string into shared structure ascii string
'
    CadInputQueue.SendKeyin "mdl co vba2mdlCommand COPY THIS STRING"
    
    If GetCExpressionValue("mdl2VbaInfo.asciiString", "vba2mdl") = "COPY THIS STRING" Then
        Debug.Print ""
        Debug.Print "Test to copy string into shared structure through MDL passed"
        Debug.Print "Initial Ascii String = " & asciiString
        Debug.Print "New Ascii String = " & GetCExpressionValue("mdl2VbaInfo.asciiString", "vba2mdl")
    Else
        Debug.Print "Test to copy string into shared structure through MDL failed"
    End If
'
' Set the shared structure to the local variable values (initial values)
'

'    SetCExpressionValue "mdl2VbaInfo.d1", d1, "vba2mdl"
'    SetCExpressionValue "mdl2VbaInfo.s1", s1, "vba2mdl"
'    SetCExpressionValue "mdl2VbaInfo.i1", i1, "vba2mdl"
'    SetCExpressionValue "mdl2VbaInfo.asciiString", asciiString, "vba2mdl"
'    SetCExpressionValue "mdl2VbaInfo.wideString", wideString, "vba2mdl"
'
' Print out shared structure values
'
    Debug.Print ""
    Debug.Print "Final Double value = " & GetCExpressionValue("mdl2VbaInfo.d1", "vba2mdl")
    Debug.Print "Final Long value = " & GetCExpressionValue("mdl2VbaInfo.i1", "vba2mdl")
    Debug.Print "Final Short value = " & GetCExpressionValue("mdl2VbaInfo.s1", "vba2mdl")
    Debug.Print "Final Ascii String = " & GetCExpressionValue("mdl2VbaInfo.asciiString", "vba2mdl")
    Debug.Print "Final Wide(Unicode) String = " & GetCExpressionValue("mdl2VbaInfo.wideString", "vba2mdl")
'
' Call mdl function to print out final values to text window
'
   GetCExpressionValue "mdl2VbaInfo_printStructure()", "vba2mdl"
End Sub


