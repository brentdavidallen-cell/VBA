VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangeCase 
   Caption         =   "Change Case"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   OleObjectBlob   =   "frmChangeCase.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangeCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'---------------------------------------------------------------------
'
'   Implementation for CHANGECASE Example
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------

Private Sub cmdExit_Click()
Unload Me ' unload this form
End Sub

Private Sub cmdOK_Click()

Dim FontCase As String ' tells us if the user has selected Upper or Lower case
Dim index As Integer

CadInputQueue.SendCommand "NULL"

If optUpper.Value = True Then
    FontCase = "Upper"
ElseIf optLower.Value = True Then
    FontCase = "Lower"
Else 'user has not selected any option
  MsgBox "Please select one of the two options.", vbInformation
  Exit Sub
End If

'  Set up to read all of the text, text node, and tag elements
'  from the model
' now transform selected elements
Dim ElEnum As ElementEnumerator
' GetElementEnumerator is in the module modTextCommands
Set ElEnum = GetElementEnumerator
If ElEnum Is Nothing Then Exit Sub

While ElEnum.MoveNext
    With ElEnum.Current
        .Redraw msdDrawingModeErase
        index = index + 1
        If .Type = msdElementTypeText Then
            Dim oTextEl As TextElement
            Set oTextEl = ElEnum.Current
            
            If FontCase = "Lower" Then
                oTextEl.Text = LCase(oTextEl.Text)
            ElseIf FontCase = "Upper" Then
                oTextEl.Text = UCase(oTextEl.Text)
            End If
        ElseIf .Type = msdElementTypeTextNode Then
            Dim oTextNodeEl As TextNodeElement
            Set oTextNodeEl = ElEnum.Current
            
            oTextNodeEl.Redraw msdDrawingModeErase
            
            'change case of individual text lines
            Dim LineCount As Long
            LineCount = 1
            While (LineCount <= oTextNodeEl.TextLinesCount)
                If FontCase = "Lower" Then
                   oTextNodeEl.TextLine(LineCount) = LCase(oTextNodeEl.TextLine(LineCount))
                ElseIf FontCase = "Upper" Then
                   oTextNodeEl.TextLine(LineCount) = UCase(oTextNodeEl.TextLine(LineCount))
                End If
                LineCount = LineCount + 1
            Wend
        ElseIf .Type = msdElementTypeTag Then
            Dim oTagEl As TagElement
            Set oTagEl = ElEnum.Current
            
            If FontCase = "Lower" Then
                oTagEl.Value = LCase(oTagEl.Value)
            ElseIf FontCase = "Upper" Then
                oTagEl.Value = UCase(oTagEl.Value)
            End If
        End If
        .Redraw msdDrawingModeNormal
        .Rewrite
    End With
Wend

ShowStatus "Changed " & index & " elements."
CommandState.StartDefaultCommand

End Sub

Private Sub UserForm_Click()

End Sub
