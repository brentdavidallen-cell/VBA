Attribute VB_Name = "modFileOpenDialog"
'
'  Copyright: (c) 2013 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

'  This module shows how to access MicroStation's standard file open dialogs.

Declare PtrSafe Function mdlDialog_defFileCreate Lib "stdmdlbltin.dll" (ByVal filename As String, ByVal rFileH As Long, ByVal dlogBoxId As Long, ByVal suggestedFileName As String, ByVal filterString As String, ByVal defaultDirectory As String, ByVal titleString As String, ByVal defaultFileId As Long, ByVal userPrefH As Long) As Long
Declare PtrSafe Function mdlDialog_defFileOpen Lib "stdmdlbltin.dll" (ByVal filename As String, ByVal rFileH As Long, ByVal dlogBoxId As Long, ByVal suggestedFileName As String, ByVal filterString As String, ByVal defaultDirectory As String, ByVal titleString As String, ByVal defaultFileId As Long, ByVal userPrefH As Long) As Long
Declare PtrSafe Function mdlDialog_fileOpen Lib "stdmdlbltin.dll" (ByVal filename As String, ByVal rFileH As Long, ByVal resourceId As Long, ByVal suggestedFileName As String, ByVal filterString As String, ByVal defaultDirectory As String, ByVal titleString As String) As Long
Declare PtrSafe Function mdlDialog_create Lib "stdmdlbltin.dll" (ByVal rFileH As Long, ByVal ownerMD As LongPtr, ByVal dialogType As Long, ByVal dialogId As Long, ByVal noWarnResourceError As Long) As LongPtr   '  Returns a pointer to a structure
Declare PtrSafe Function mdlDialog_fileCreateFromSeed Lib "stdmdlbltin.dll" (ByVal filename As String, ByVal rFileH As Long, ByVal resourceId As Long, ByVal suggestedFileName As String, ByVal filterString As String, ByVal defaultDirectory As String, ByVal titleString As String, ByVal seedFile As String, ByVal seedDirectory As String, ByVal seedFilter As String) As Long

Const DEFDGNFILE_ID As Long = -101

Sub TryOpen()
    Dim strNewFile As String
    
    ' Use Space to force the String to be long enough for whatever mdlDialog_defFileOpen puts in it
    strNewFile = Space(1024)
    If mdlDialog_defFileOpen(strNewFile, 0, 0, "MyFile", "*.dgn", "c:\ustation\dgn\", "Select file for test", DEFDGNFILE_ID, 0) = 0 Then
        '  Find the vbNullChar that mdlDialog_defFileOpen and truncate the String there
        Debug.Print TruncateAtEOS(strNewFile)
    End If
End Sub
Sub TrySimple()
    Dim strNewFile As String
    
    strNewFile = Space(1024)
    If mdlDialog_fileOpen(strNewFile, 0, 0, "MyFile.dgn", "*.dgn", "c:\ustation\dgn\", "Select file for test") = 0 Then
        Debug.Print TruncateAtEOS(strNewFile)
    End If
End Sub

Sub TryCreateFromSeed()
    Dim strNewFile As String
    Dim strSeedFile As String
    
    strSeedFile = ActiveWorkspace.ConfigurationVariableValue("MS_DESIGNSEED")
    
    strNewFile = Space(1024)
    If mdlDialog_fileCreateFromSeed(strNewFile, 0, 0, "test.dgn", "*.dgn", "d:\ustation\dgn\", "Specify name of file to create", strSeedFile, "d:\ustation\dgn", "*.dgn;*.dwg") = 0 Then
        Debug.Print TruncateAtEOS(strNewFile)
    End If
End Sub
