Attribute VB_Name = "modPlaceShapeCommands"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit
Sub InscribeEllipse()
    CommandState.StartPrimitive New clsInscribeEllipse
End Sub
Sub DefineBlock()
    CommandState.StartPrimitive New clsPlaceBlock
End Sub
Sub PlaceGrid()
    frmDefineGrid.Show
End Sub

Sub PlaceCircle()
    CommandState.StartPrimitive New clsPlaceCircleCommand
End Sub

Sub PlaceCurve()
      CommandState.StartPrimitive New clsPlaceCurve
End Sub
Sub PlaceLine()
    CommandState.StartPrimitive New clsPlaceLineCommand
End Sub

Sub PlaceArc()
    CommandState.StartPrimitive New clsPlaceArcCommand
End Sub
