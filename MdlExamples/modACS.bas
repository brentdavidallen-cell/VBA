Attribute VB_Name = "modACS"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Declare PtrSafe Function mdlACS_getCurrent Lib "stdmdlbltin.dll" (ByRef originP As Point3d, ByRef rotMatrixP As Matrix3d, ByVal typeP As LongPtr, ByVal nameP As LongPtr, ByVal descriptionP As LongPtr) As Long

'  This example illustrates how to retrieve strings using a function that has MSWChar *
'  arguments. In this case, both the name and description arguments are MSWChar *.
Sub GetACS(pntACSOrigin As Point3d, mtrxRotation As Matrix3d, intType As Integer, strName As String, strDescription As String)
    Dim status As Long
    
    '  Create long strings
    strName = Space(1024)
    strDescription = Space(1024)
    
    ' Use StrPtr to pass the address of the String's data buffer
    status = mdlACS_getCurrent(pntACSOrigin, mtrxRotation, intType, StrPtr(strName), StrPtr(strDescription))
    If status <> 0 Then
        Err.Raise msdErrorNoAcsDefined, "No ACS defined"
        Exit Sub
    End If
    
    '  Now truncate the String at the vbNullChar
    strName = TruncateAtEOS(strName)
    strDescription = TruncateAtEOS(strDescription)
End Sub
Sub TestCurr()
    Dim org As Point3d
    Dim rot As Matrix3d
    Dim intType As Integer
    Dim strName As String
    Dim strDescr As String
    
    GetACS org, rot, intType, strName, strDescr
    Debug.Print "ACS " & strName & "[" & strDescr & "] is " & intType
End Sub
