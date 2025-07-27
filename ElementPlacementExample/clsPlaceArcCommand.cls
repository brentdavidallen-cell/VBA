VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlaceArcCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------
'Class Implementation for PLACEARC Example
'Description: This command creates an arc after the user has entered
'             its Startpoint, Centerpoint and Endpoint.
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
Private m_Start As Point3d
Private m_Center As Point3d
Private m_End As Point3d
Private m_nPoints As Integer


Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
If m_nPoints = 0 Then
    'accepts starting point of the arc
    CommandState.StartDynamics
    m_Start = Point
    m_nPoints = 1
    ShowPrompt "Identify Center point"

ElseIf m_nPoints = 1 Then
    'accepts the center of the arc
    m_Center = Point
    m_nPoints = 2
    ShowPrompt "Select Ending Point of Arc"

ElseIf m_nPoints = 2 Then
    'accepts the end point of the arc
    m_End = Point

    'draws the arc element
    Dim oArcEl As ArcElement
    Set oArcEl = CreateArcElement1(Nothing, m_Start, m_Center, m_End)
    ActiveModelReference.AddElement oArcEl
    oArcEl.Redraw

    ShowPrompt "Done"
End If
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'method called to show dynamics

Dim radius As Double
If m_nPoints = 1 Then
    radius = Point3dDistance(Point, m_Start)
    
    'draw a dynamic circle to aid user in selecting arc center
    Dim oEllipseEl As EllipseElement
    Set oEllipseEl = CreateEllipseElement2(Nothing, Point, radius, radius, View.Rotation)
    oEllipseEl.Color = 5 ' color to indicate its a temporary display
    oEllipseEl.Redraw DrawMode

ElseIf m_nPoints = 2 Then
   
   'shows the user a dynamic display of the final ellipse.
    Dim oArcEl As ArcElement
    Set oArcEl = CreateArcElement1(Nothing, m_Start, m_Center, Point)
    oArcEl.Color = 5 ' color to indicate its a temporary display
    oArcEl.Redraw DrawMode
End If

End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal KeyIn As String)
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
' method called when the reset event occurs

' On reset, start the command again.
CommandState.StartPrimitive Me

m_nPoints = 0
End Sub

Private Sub IPrimitiveCommandEvents_Start()
'method called at the start of the command

ShowCommand "VBA Place Arc Example"
ShowPrompt "Select Starting Point of Arc"
End Sub

