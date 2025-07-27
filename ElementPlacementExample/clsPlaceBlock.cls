VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlaceBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------
'Class Implementation for PLACEBLOCK Example
'Description: This command allows the user to define the height, width
'             and angle of a simple quadrilateral.
'
'Adapted from "101 MDL Commands" by Bill Steinback
'---------------------------------------------------------------------

Option Explicit

Implements IPrimitiveCommandEvents

'----------------------------------------------------------------------
'
'           Private Data Members
'
'----------------------------------------------------------------------
Private m_Width As Double
Private m_Height As Double

Private Sub IPrimitiveCommandEvents_Cleanup()
Unload frmDefineBlock
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
Dim atPoints(0 To 3) As Point3d

' check for empty textboxes indicating missing information
If (frmDefineBlock.txtHeight.Value = "") Or (frmDefineBlock.txtWidth.Value = "") Then
    MsgBox "Please type in data in all fields", vbOKOnly + vbExclamation
    Exit Sub
End If

'intialize variables
m_Height = frmDefineBlock.txtHeight.Value
m_Width = frmDefineBlock.txtWidth.Value

ShowPrompt "Select center of block"
CommandState.StartDynamics

'given values for height and width and knowing the coords of the
'centroid of the block, calculate the coords of the 4 vertices
atPoints(0).X = Point.X - (m_Width / 2)
atPoints(0).Y = Point.Y - (m_Height / 2)

atPoints(1).X = Point.X + (m_Width / 2)
atPoints(1).Y = Point.Y - (m_Height / 2)

atPoints(2).X = Point.X + (m_Width / 2)
atPoints(2).Y = Point.Y + (m_Height / 2)

atPoints(3).X = Point.X - (m_Width / 2)
atPoints(3).Y = Point.Y + (m_Height / 2)

'draw the block
Dim oShapeEl As ShapeElement
Set oShapeEl = CreateShapeElement1(Nothing, atPoints)
ActiveModelReference.AddElement oShapeEl
oShapeEl.Redraw

ShowPrompt "Done"
End Sub


Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'method called to show dynamics

Dim atPoints(0 To 3) As Point3d

'  Refreshing m_Height and m_Width, makes the dynamics reflect
'  the most current values in the dialog box. DO not
'  refresh when the draw mode is msdDrawingModeTemporaryErase
'  because this must erase what was drawn in the last call to
'  IPrimitiveCommandEvents_Dynamics.
If DrawMode <> msdDrawingModeTemporaryErase Then
    m_Height = frmDefineBlock.txtHeight.Value
    m_Width = frmDefineBlock.txtWidth.Value
End If

'given values for height and width and knowing the coords of the possible
'centroid of the block, calculate the coords of the 4 vertices
atPoints(0).X = Point.X - (m_Width / 2)
atPoints(0).Y = Point.Y - (m_Height / 2)

atPoints(1).X = Point.X + (m_Width / 2)
atPoints(1).Y = Point.Y - (m_Height / 2)

atPoints(2).X = Point.X + (m_Width / 2)
atPoints(2).Y = Point.Y + (m_Height / 2)

atPoints(3).X = Point.X - (m_Width / 2)
atPoints(3).Y = Point.Y + (m_Height / 2)

' display a temporary outline of the block
Dim oShapeEl As ShapeElement
Set oShapeEl = CreateShapeElement1(Nothing, atPoints)
oShapeEl.Color = 5 ' color to indicate its a temporary display
oShapeEl.Redraw DrawMode
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal KeyIn As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
' method called when the reset event occurs

' On reset, start the command again.
CommandState.StartPrimitive Me
End Sub

Private Sub IPrimitiveCommandEvents_Start()
'method called at the start of the command

ShowCommand "VBA Define Block Example"

'show the form that enables the user to enter required information
frmDefineBlock.Show
ShowPrompt "Enter Block Information"
End Sub

