VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDefineGrid 
   Caption         =   "Define a Grid"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   OleObjectBlob   =   "frmDefineGrid.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDefineGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------
'Implementation for PLACEGRID Example
'Description: This command draws a grid based on the following information
'             entered by the user
'Author:  Bentley Systems
'Adapted from "101 MDL Commands" by Bill Steinback
'
'Copyright (c) 1999-2001;  Bentley Systems, Inc., 685 Stockton Drive,
'                      Exton PA, 19341-0678, USA.  All Rights Reserved.
'
'This program is confidential, proprietary and unpublished property of Bentley Systems
'Inc. It may NOT be copied in part or in whole on any medium, either electronic or
'printed, without the express written consent of Bentley Systems, Inc.
'---------------------------------------------------------------------

Option Explicit

Private Sub cmdCreateGrid_Click()

' set all the component elements of the Grid to have the same graphic group number
Dim mtrxRotation As Matrix3d
Dim trans As Transform3d
Dim GraphicState As Boolean
GraphicState = ActiveSettings.GraphicGroupLockEnabled 'save original graphic lock settings
ActiveSettings.GraphicGroupLockEnabled = True
ActiveSettings.CurrentGraphicGroup = ActiveSettings.CurrentGraphicGroup + 1

'create local variables
Dim xscale As Double
Dim yscale As Double
Dim xmin As Double
Dim ymin As Double
Dim xmax As Double
Dim ymax As Double
Dim eleLine As LineElement

' check for non-numeric values
If (txtXScale.Value = "") Or (txtYScale.Value = "") Or (txtXMin.Value = "") Or (txtYMin.Value = "") Or (txtXMax.Value = "") Or (txtYMax.Value = "") Then
    MsgBox "Please enter numneric values only", vbInformation
    Exit Sub
End If

mtrxRotation = ActiveDesignFile.Views(1).Rotation
mtrxRotation = Matrix3dInverse(mtrxRotation)

'initialize variables
xscale = txtXScale.Value
yscale = txtYScale.Value
xmin = txtXMin.Value
ymin = txtYMin.Value
xmax = txtXMax.Value
ymax = txtYMax.Value

Dim startpt As Point3d 'point that will determine the start of the grid lines
Dim endpt As Point3d ' point that will determine the end of the grid lines
Dim pntCorner As Point3d
Dim Origin As Point3d ' Orign of the text element used for labelling the numbers
Dim coord As Double

'initialize the points for the vertical grid lines
startpt.Y = ymin
endpt.Y = ymax
Origin.Y = startpt.Y - 0.5
'  For the vertical lines, the x coordinate varies
coord = xmin

pntCorner = Point3dFromXYZ(xmin, ymin, 0)

Do While (coord <= xmax)
   'draw grid's vertical lines
   startpt.X = coord
   endpt.X = coord

   'draw the line
   Set eleLine = CreateLineElement2(Nothing, startpt, endpt)
   trans = Transform3dFromMatrix3dAndFixedPoint3d(mtrxRotation, pntCorner)
   eleLine.Transform trans
   ActiveModelReference.AddElement eleLine
   eleLine.Redraw

   ' move to the next line
   coord = coord + xscale
Loop

'intialize the points for the horizontal grid lines
startpt.X = xmin
endpt.X = xmax
Origin.X = startpt.X + 0.2
'  For the horizontal lines, the y coordinate varies
coord = ymin

Do While (coord <= ymax)
   'draw grid's horizontal lines
   startpt.Y = coord
   endpt.Y = coord

   'draw horizontal lines
   Set eleLine = CreateLineElement2(Nothing, startpt, endpt)
   trans = Transform3dFromMatrix3dAndFixedPoint3d(mtrxRotation, pntCorner)
   eleLine.Transform trans
   ActiveModelReference.AddElement eleLine
   eleLine.Redraw
   
   'move to the next line
   coord = coord + yscale
Loop

ActiveSettings.GraphicGroupLockEnabled = GraphicState 'reset original graphic lock settings
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
