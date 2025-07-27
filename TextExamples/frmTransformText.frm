VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTransformText 
   Caption         =   "Transform Text Element"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   OleObjectBlob   =   "frmTransformText.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTransformText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'
'  Implementation for TRANSFORMTEXT Example
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'

Private Sub chkRotate_Change()
If chkRotate.Value = False Then
    txtAngle.Value = ""
    txtAngle.Enabled = False
Else
    txtAngle.Enabled = True
End If
End Sub

Private Sub chkScale_change()
If chkScale.Value = False Then
    txtScale.Text = ""
    txtScale.Enabled = False
Else
    txtScale.Enabled = True
End If
End Sub

Private Sub cmdClear_Click()
' clear the text boxes
txtScale.Value = ""
txtAngle.Value = ""
End Sub

Private Sub cmdExit_Click()
Unload Me 'unload this form
End Sub

Private Sub cmdTransform_Click()
'used for transforming the element
Dim ScalingMatrix As Matrix3d
Dim RotationMatrix As Matrix3d
Dim TransMatrix As Matrix3d
Dim Eltrans As Transform3d

ScalingMatrix = Matrix3dIdentity
RotationMatrix = Matrix3dIdentity

If chkScale Then
    'check for non-numeric values
     If txtScale.Value = "" Or txtScale.Value = 0 Then
         MsgBox "Please enter valid numeric values only", vbInformation
         txtScale.SetFocus
         Exit Sub
     End If
     
     Dim ElScale As Double
     ElScale = txtScale.Value
    
    'set up the scaling matrix
     ScalingMatrix = Matrix3dFromScale(ElScale)
End If

If chkRotate Then
    'check for non-numeric values
    If txtAngle.Value = "" Then
        MsgBox "Please enter valid numeric values only", vbInformation
        txtAngle.SetFocus
        Exit Sub
    End If

    Dim Angle As Double
    Angle = Radians(txtAngle.Value)  'retrieve values from user and change to radians

    RotationMatrix = Matrix3dFromAxisAndRotationAngle(2, Angle)
End If

'  Now combined the transformations
TransMatrix = Matrix3dFromMatrix3dTimesMatrix3d(ScalingMatrix, RotationMatrix)

' now transform selected elements
Dim oElEnum As ElementEnumerator

' GetElementEnumerator is in the module modTextCommands
Set oElEnum = GetElementEnumerator
If oElEnum Is Nothing Then Exit Sub

While oElEnum.MoveNext
   'check to see if selected element is a Text element
    If (oElEnum.Current.Type <> msdElementTypeText) Then
        MsgBox "Please select TextElements only", vbInformation
        Exit Sub
    End If

    Dim oEl As TextElement ' create an element object
    Set oEl = oElEnum.Current 'copy the current element
    
    oEl.Redraw msdDrawingModeErase
    
    Eltrans = Transform3dFromMatrix3dAndFixedPoint3d(TransMatrix, oEl.Origin)
    oEl.Transform Eltrans

    oEl.Rewrite
    oEl.Redraw
Wend

ShowPrompt "Done"
End Sub

Private Sub UserForm_Click()

End Sub
