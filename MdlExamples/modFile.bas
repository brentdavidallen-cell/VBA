Attribute VB_Name = "modFile"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Declare PtrSafe Function mdlFile_find Lib "stdmdlbltin.dll" (ByVal outName As String, ByVal inName As String, ByVal envVar As String, ByVal iext As String) As Long


'  strInName -  a filename that has a format a user would enter. It can contain a logical filename
'               and a path. It can contain an extension.  The information in strInName is
'               used first to find the file.
'  strEnvVar -  usually provides the name of an environment variable. The information in the environment variable
'               generally provides a list of paths to examine when searching for the file. It can be vbNothing.
'  strExt -     usually provides a file extension (example .dgn or dgn). ext can be vbNothing.
Function FindFile(strInName As String, strEnvVar As String, ByVal strExt As String) As String
    Dim status As Long
    
    If strExt <> "" Then
        If Mid(strExt, 1, 1) <> "." Then
            strExt = "." & strExt
        End If
    End If
    FindFile = Space(1024)
    status = mdlFile_find(FindFile, strInName, strEnvVar, strExt)
    If status <> 0 Then
        FindFile = ""
        Exit Function
    End If
    
    FindFile = TruncateAtEOS(FindFile)
End Function

Sub TestFindFile()
    Dim filename As String
    
    filename = FindFile("MdlExamples", "MS_VBASEARCHDIRECTORIES", "mvba")
    Debug.Print "FindFile returned " & filename
End Sub
