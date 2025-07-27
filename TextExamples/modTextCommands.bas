Attribute VB_Name = "modTextCommands"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit
'  The ChangeCase and TransformText commands both use this to get an
'  ElementEnumerator to retrieve the list of elements to process.
Function GetElementEnumerator() As ElementEnumerator
    If ActiveModelReference.AnyElementsSelected Then
        Set GetElementEnumerator = ActiveModelReference.GetSelectedElements ' get the selected elements
    ElseIf ActiveDesignFile.Fence.IsDefined Then
        Set GetElementEnumerator = ActiveDesignFile.Fence.GetContents
    Else
        If MsgBox("No fence or selection set defined. Process all elements?", vbOKCancel) = vbCancel Then
            Exit Function
        End If
        Dim sc As New ElementScanCriteria
        sc.ExcludeAllTypes
        sc.IncludeType msdElementTypeText
        sc.IncludeType msdElementTypeTextNode
        Set GetElementEnumerator = ActiveModelReference.Scan(sc)
    End If
End Function
Sub ChangeCase()
    frmChangeCase.Show
End Sub

Sub TransformText()
    frmTransformText.Show
End Sub

