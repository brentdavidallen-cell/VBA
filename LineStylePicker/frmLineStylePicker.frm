VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLineStylePicker 
   Caption         =   "LineStyle Picker"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2070
   OleObjectBlob   =   "frmLineStylePicker.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLineStylePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'

Option Explicit

Private Function GetTempEnhMetafile() As String
    GetTempEnhMetafile = GetTemporaryFileName(GetTemporaryPath, "EMF")
    If GetTempEnhMetafile = "" Then
        Err.Raise 1, "Get Temporary path failed"
    End If
   
End Function

Sub DrawToFile(ls As LineStyle, strEnhMetafileName As String, Width As Long, Height As Long)
    On Error Resume Next
    ls.DrawToFile strEnhMetafileName, PointsToPixelsX(Width), PointsToPixelsY(Height), True
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub listboxLineStyles_Click()
    Dim lss As LineStyles
    Dim ls As LineStyle
    
    If listboxLineStyles.ListIndex = 0 Then
        Set ls = ByLevelLineStyle
        ActiveSettings.LineStyle = ls
    ElseIf listboxLineStyles.ListIndex = 1 Then
        Set ls = ByLevelLineStyle
        ActiveSettings.LineStyle = ls
    Else
        Set lss = ActiveDesignFile.LineStyles
        Set ls = lss.Item(listboxLineStyles.ListIndex - 1)
        ActiveSettings.LineStyle = ls
    End If
    
    With frmLineStylePicker.Image1
        On Error GoTo DrawToMetaFile
        .Picture = ls.GetPicture(PointsToPixelsX(.Width), PointsToPixelsY(.Height), True)
    End With
    Exit Sub
    
DrawToMetaFile:
    Dim strEnhMetafileName As String
    ' On Error Resume Next
    strEnhMetafileName = GetTempEnhMetafile
    With frmLineStylePicker.Image1
        DrawToFile ls, strEnhMetafileName, .Width, .Height
        .Picture = LoadPicture(strEnhMetafileName)
    End With
    Kill strEnhMetafileName
End Sub


Private Sub UserForm_Initialize()
    Dim lss As LineStyles
    Dim ls As LineStyle
    Dim i As Integer
    
    Set lss = ActiveDesignFile.LineStyles
       
    frmLineStylePicker.listboxLineStyles.AddItem ByLevelLineStyle.Name 'at index 0
    frmLineStylePicker.listboxLineStyles.AddItem ByCellLineStyle.Name  'at index 1
    
    For i = 1 To lss.Count
        Set ls = lss.Item(i)
        Debug.Print "name = " & ls.Name & ", number = " & ls.Number
        frmLineStylePicker.listboxLineStyles.AddItem ls.Name                       ' at index 2, 3, ...
    Next
    
    frmLineStylePicker.listboxLineStyles.ListIndex = 0
    Set ls = ByLevelLineStyle
'        Set ls = lss.Item(1)
    With frmLineStylePicker.Image1
        ActiveSettings.LineStyle = ls
        On Error GoTo DrawToMetaFile
        .Picture = ls.GetPicture(PointsToPixelsX(.Width), PointsToPixelsY(.Height), True)
    End With
   
    frmLineStylePicker.Show
    Exit Sub
    
DrawToMetaFile:
    ' GetPicture raises an error if the LineStyle object is not from the current
    ' process. For example, it raises an error if this code is used in a VB program
    ' that gets a LineStyle object from a MicroStation that is running as a separate
    ' process.  This error does not occur if this code is used within MicroStation's VBA,
    ' or in a VB DLL accessed from MicroStation's VBA.
    Dim strEnhMetafileName As String
    strEnhMetafileName = GetTempEnhMetafile
    With frmLineStylePicker.Image1
        DrawToFile ls, strEnhMetafileName, .Width, .Height
        .Picture = LoadPicture(strEnhMetafileName)
    End With
    frmLineStylePicker.Show
    Kill strEnhMetafileName
End Sub











