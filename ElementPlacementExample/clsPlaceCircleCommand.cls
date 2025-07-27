VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlaceCircleCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------
'Class Implementation for PLACECIRCLE Example
'Description: This command draws a circle after the user has entered its
'             desired radius
'---------------------------------------------------------------------

Option Explicit

Implements IPrimitiveCommandEvents

'----------------------------------------------------------------------
'
'           Private Data Members
'
'----------------------------------------------------------------------
Private m_Radius As Double
Private m_Center As Point3d

Private Sub IPrimitiveCommandEvents_Cleanup()
    Unload frmEnterRadius
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
' check for empty textbox indicating missing information and check for
' negative value of radius.
If (frmEnterRadius.txtRadius.Value = "") Or (frmEnterRadius.txtRadius.Value <= 0) Then
    MsgBox "Please enter a valid value for radius", vbOKOnly + vbExclamation
    Exit Sub
End If

m_Radius = frmEnterRadius.txtRadius.Value

ShowPrompt "Select Center Of Circle"
m_Center = Point
CommandState.StartDynamics

'draws the circle given a center and radius
Dim oEllipseEl As EllipseElement
Set oEllipseEl = CreateEllipseElement2(Nothing, m_Center, m_Radius, m_Radius, View.Rotation)
ActiveModelReference.AddElement oEllipseEl
oEllipseEl.Redraw

End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'method called to show dynamics

If DrawMode = msdDrawingModeTemporary Then
    '  Never refresh on the erase.  Erasing
    '  must use the same radius as the previous draw.
    m_Radius = frmEnterRadius.txtRadius.Value
End If

'display temoprary outline of the circle
Dim oEllipseEl As EllipseElement
Set oEllipseEl = CreateEllipseElement2(Nothing, Point, m_Radius, m_Radius, View.Rotation)
oEllipseEl.Color = 5 ' color to indicate its a temporary display
oEllipseEl.Redraw DrawMode
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

ShowCommand "VBA PlaceCircle Example"

'shows a pop-up window and asks user to enter radius
frmEnterRadius.Show
frmEnterRadius.txtRadius.Value = 1
CommandState.StartDynamics
ShowPrompt "Enter Radius Of Circle"
End Sub


