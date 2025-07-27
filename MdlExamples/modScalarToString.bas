Attribute VB_Name = "modScalarToString"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Declare PtrSafe Sub mdlCurrTrans_scaleDoubleArray Lib "stdmdlbltin.dll" (ByRef out As Double, ByRef in_ As Double, ByVal numValues As Long)
Declare PtrSafe Sub mdlString_fromUors Lib "stdmdlbltin.dll" (ByVal uor_string As String, ByVal uors As Double)
Function ScalarToString(ByVal dbl As Double) As String
    Dim str As String
    
    str = Space(1024)
    mdlCurrTrans_scaleDoubleArray dbl, dbl, 1
    mdlString_fromUors str, dbl
    
    ScalarToString = TruncateAtEOS(str)
End Function
Sub TestValue()
    Debug.Print ScalarToString(10.5)
End Sub
