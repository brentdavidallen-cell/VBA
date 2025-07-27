VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPictures 
   Caption         =   "UserForm1"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   OleObjectBlob   =   "frmPictures.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Public Sub ShowElement(ele As Element)
    Dim width As Long
    Dim height As Long
        
    width = PointsToPixelsX(imgElement.width)
    height = PointsToPixelsY(imgElement.height)
    
    Set imgElement.Picture = ele.GetPicture(width, height, cbkBackground.Value)
End Sub
'  This indents the image by creating a cell with a hidden line surrounding the image to be displayed.
Function GetCellForPicture(aEle() As Element, ByVal xIndent As Double, ByVal yIndent As Double) As CellElement
    Dim eleTemp As CellElement
    Dim outer(0 To 1) As Element
    Dim vertices(0 To 3) As Point3d
    
    Set eleTemp = CreateCellElement1(vbNull, aEle, aEle(0).Range.Low, False)
    
    Set GetCellForPicture = eleTemp
    ' Exit Function
    
    Dim xRange As Double
    Dim yRange As Double
    Dim cellRange As Range3d
    
    cellRange = eleTemp.Range
    
    xRange = cellRange.High.X - cellRange.Low.X + 1
    yRange = cellRange.High.Y - cellRange.Low.Y + 1

    xIndent = xRange / xIndent
    yIndent = yRange / yIndent
    
    With eleTemp.Range
        vertices(0) = .Low
        vertices(1).X = .Low.X
        vertices(1).Y = .High.Y
        vertices(2) = .High
        vertices(3).X = .High.X
        vertices(3).Y = .Low.Y
    End With
    
    vertices(0).X = vertices(0).X - xIndent / 2
    vertices(1).X = vertices(1).X - xIndent / 2
    vertices(2).X = vertices(2).X + xIndent / 2
    vertices(3).X = vertices(3).X + xIndent / 2
    
    vertices(0).Y = vertices(0).Y - yIndent / 2
    vertices(1).Y = vertices(1).Y + yIndent / 2
    vertices(2).Y = vertices(2).Y + yIndent / 2
    vertices(3).Y = vertices(3).Y - yIndent / 2
    
    Set outer(0) = eleTemp
    Set outer(1) = CreateShapeElement1(Nothing, vertices)
    outer(1).IsHidden = True
    Set GetCellForPicture = CreateCellElement1(vbNullString, outer, outer(1).Range.Low)
End Function
Sub ShowAllElements()
    Dim ee As ElementEnumerator
    Dim count As Long
    On Error GoTo NothingToShow
    
    Set ee = ActiveModelReference.GraphicalElementCache.Scan
    
    Do While ee.MoveNext
        count = count + 1
    Loop
    
    ReDim arr(0 To count - 1) As Element
    Dim index As Long
    
    Set ee = ActiveModelReference.GraphicalElementCache.Scan
    Do While ee.MoveNext
        Set arr(index) = ee.Current
        index = index + 1
    Loop

    Dim eleCell As CellElement
    
    Set eleCell = GetCellForPicture(arr, 5, 5)
    
    frmPictures.ShowElement eleCell
    frmPictures.Caption = "Displaying design file"
    Exit Sub
    
NothingToShow:
    ShowError "Did not find anything to show"
End Sub
Sub ShowElements()
    Dim ee As ElementEnumerator
    Dim nSelected As Long
    Dim index As Long
    
    '  If there are no elements selected, display everything in
    '  the design file
    If ActiveModelReference.AnyElementsSelected = False Then
        ShowAllElements
        Exit Sub
    End If
    
    '
    '  Create a cell from the selected elements and then display the cell
    '
    '  Count the selected elements and then create an array
    '  to use when creating the cell
    Set ee = ActiveModelReference.GetSelectedElements
    Do While ee.MoveNext
        nSelected = nSelected + 1
    Loop
    
    ee.Reset
    ReDim aEle(0 To nSelected) As Element
    Do While ee.MoveNext
        Set aEle(index) = ee.Current
        index = index + 1
    Loop
    
    
    ' Create the cell
    Dim eleCell As CellElement
    Set eleCell = GetCellForPicture(aEle, 5, 5)
    
    '  Display the cell
    frmPictures.Caption = "Displaying selected elements"
    frmPictures.ShowElement eleCell
    ActiveModelReference.UnselectAllElements
End Sub
Public Sub btnShowElements_Click()
    ShowElements
End Sub

Private Sub imgElement_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub imgElement_Click()

End Sub
