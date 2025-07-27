VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlaceCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------
'Class Implementation for PLACECURVE Example
'Description: This command draws a curve given the key data points by the user.
'---------------------------------------------------------------------
Option Explicit

Implements IPrimitiveCommandEvents

'----------------------------------------------------------------------
'
'           Private Data Members
'
'----------------------------------------------------------------------
Private m_atPoints() As Point3d 'declare a dynamic array
Private m_nPoints As Integer
Private m_CurveEl As CurveElement
Private Sub IPrimitiveCommandEvents_Cleanup()

End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)

ReDim Preserve m_atPoints(m_nPoints)
m_atPoints(m_nPoints) = Point
m_nPoints = m_nPoints + 1

If (m_nPoints >= 2) Then
CommandState.StartDynamics
End If

ShowPrompt "Enter next point, reset to finish"
End Sub


Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'method called to show dynamics
If m_nPoints >= 2 Then
    'temporarily increase the size of the array so that we can store the current dynamic point
    ReDim Preserve m_atPoints(m_nPoints)
    m_atPoints(m_nPoints) = Point

   'display the temporary curve element
    Dim oCurveEl As CurveElement
    Set oCurveEl = CreateCurveElement1(Nothing, m_atPoints)
    oCurveEl.Color = 5 ' color to indicate its a temporary display
    oCurveEl.Redraw DrawMode
    
    ReDim Preserve m_atPoints(m_nPoints - 1)
End If

End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal KeyIn As String)

End Sub

Private Sub IPrimitiveCommandEvents_Reset()
' method called when the reset event occurs

'draw the final curve
Set m_CurveEl = CreateCurveElement1(Nothing, m_atPoints)
ActiveModelReference.AddElement m_CurveEl
m_CurveEl.Redraw

'clean up
ReDim m_atPoints(0) 'clear up the dynamic array
m_nPoints = 0

' On reset, start the command again.
CommandState.StartPrimitive Me
End Sub

Private Sub IPrimitiveCommandEvents_Start()
'method called at the start of the command

ShowCommand "VBA PlaceCurve Example"
ShowPrompt "Enter first point, reset to finish"
End Sub

