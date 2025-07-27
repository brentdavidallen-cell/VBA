VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInscribeEllipse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'---------------------------------------------------------------------
'Class Implementation for INSCRIBEELLIPSE Example
'Description: This command creates an ellipse by letting the user define
'             the bounding rectangle.
'Adapted from "101 MDL Commands" by Bill Steinback
'---------------------------------------------------------------------

Option Explicit

Implements IPrimitiveCommandEvents

'----------------------------------------------------------------------
'
'           Private Data Members
'
'----------------------------------------------------------------------
Private m_atPoints(0 To 3) As Point3d
Private m_nPoints As Integer
        
        
Private Sub IPrimitiveCommandEvents_Cleanup()
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
If m_nPoints = 0 Then 'user has entered first corner of bounding rectangle
    CommandState.StartDynamics 'start showing dynamics
    m_atPoints(0) = Point
    m_nPoints = 1
    ShowPrompt "Place end point"
    
ElseIf m_nPoints = 1 Then 'user has entered the diagonally opposite corner of bounding rect.
    m_atPoints(2) = Point
    m_nPoints = 2
      
    'determine primary and secondary radii
    Dim Prim_Radius As Double
    Dim Sec_Radius As Double
    Prim_Radius = Abs(m_atPoints(0).X - m_atPoints(2).X) / 2 'half of horizontal distance
    Sec_Radius = Abs(m_atPoints(0).Y - m_atPoints(2).Y) / 2 ' half of vertical distance
    
    'determine origin
    Dim Origin As Point3d ' midpoint
    Origin.X = (m_atPoints(0).X + m_atPoints(2).X) / 2
    Origin.Y = (m_atPoints(0).Y + m_atPoints(2).Y) / 2
    
    'draws the ellipse given primary and secondary radii
    Dim oEllipseEl As EllipseElement
    Set oEllipseEl = CreateEllipseElement2(Nothing, Origin, Prim_Radius, Sec_Radius, View.Rotation)
    ActiveModelReference.AddElement oEllipseEl
    oEllipseEl.Redraw
    
    '  Restart the command
    m_nPoints = 0
    CommandState.StartPrimitive Me
End If
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
'method called to show dynamics

If m_nPoints = 1 Then
    ' At this point, we have the coords for one vertex of the bounding rectangle.
    ' Hence, the current point is the point at the opposite corner of this rectangle i.e.
    
    m_atPoints(2) = Point
    
   ' given the coords of the 2 opposite vertices, determine the coords
   ' of the rest of the vertices.
    m_atPoints(1).X = m_atPoints(0).X
    m_atPoints(1).Y = Point.Y
    
    m_atPoints(3).X = Point.X
    m_atPoints(3).Y = m_atPoints(0).Y
    
    ' display the temporary boundary rectangle after knowing all of its four vertices.
    Dim oShapeEl As ShapeElement
    Set oShapeEl = CreateShapeElement1(Nothing, m_atPoints)
    oShapeEl.Color = 5 ' change color to indicate its a temporary shape
    oShapeEl.Redraw DrawMode
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

ShowCommand "VBA InscribeEllipse Example"
ShowPrompt "Enter first point"
End Sub



