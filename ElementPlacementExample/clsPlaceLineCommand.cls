VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlaceLineCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------
'Class Implementation for PLACELINE Example
'Description: This command draws simple line segments by connecting the
'             data points provided by the user.
'---------------------------------------------------------------------
Option Explicit

Implements IPrimitiveCommandEvents

'----------------------------------------------------------------------
'
'           Private Data Members
'
'----------------------------------------------------------------------
Private m_atPoints(0 To 1) As Point3d
Private m_nPoints As Integer

Private Sub IPrimitiveCommandEvents_Cleanup()

End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
If m_nPoints = 0 Then 'user has entered Startpoint of the line segment
    CommandState.StartDynamics
    m_atPoints(0) = Point
    m_nPoints = 1
    ShowPrompt "Place end point"
ElseIf m_nPoints = 1 Then 'user has entered Endpoint of the line segment
    m_atPoints(1) = Point
    
    'draws the line segment
    Dim oLineEl As LineElement
    Set oLineEl = CreateLineElement1(Nothing, m_atPoints)
    ActiveModelReference.AddElement oLineEl
    oLineEl.Redraw
    
    'set Startpoint for next line segment to the Endpoint of the previous one
    m_atPoints(0) = m_atPoints(1)
End If
End Sub


Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'method called to show dynamics

If m_nPoints = 1 Then
    m_atPoints(1) = Point
    
    'display the temporary line element
    Dim oLineEl As LineElement
    Set oLineEl = CreateLineElement1(Nothing, m_atPoints)
    oLineEl.Color = 5 ' color to indicate its a temporary display
    oLineEl.Redraw DrawMode
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

ShowCommand "VBA PlaceLine Example"
ShowPrompt "Select start of line"
End Sub
