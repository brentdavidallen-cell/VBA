Attribute VB_Name = "SmartSolidExamples"
'TestTrimSolidWithSurface is an example of SmartSolid.TrimSolidWithSurface
'TestCreateSectionFromSolid is an example of SmartSolid.CreateSectionFromSolid
'TestBlendEdges is an example of SmartSolid.BlendEdges
'TestBlendAllEdgesOfFace is an example of SmartSolid.BlendAllEdgesOfFace
'TestBlendEdgeWithVariableRadius is an example of SmartSolid.BlendEdgeWithVariableRadius
'TestChamferEdge is an example of SmartSolid.ChamferEdge
'TestOffsetFace is an example of SmartSolid.OffsetFace
'TestRemoveFace is an example of SmartSolid.RemoveFace
'TestProjectCurveOntoSolid is an example of SmartSolid.ProjectCurveOntoSolid
'TestRevolveProfileAxis is an example of SmartSolid.RevolveProfileAxis
'TestSpinFace is an example of SmartSolid.SpinFace
'TestSweepProfileAlongPathAndThinshell is an example of SmartSolid.SweepProfileAlongPath and SmartSolid.Thinshell
'TestExtrudeClosedPlanarCurve is an example of SmartSolid.ExtrudeClosedPlanarCurve

Public Sub TestTrimSolidWithSurface()
Dim result As SmartSolidElement
Dim tr As Transform3d

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim oBsplineSurfaceElement As BsplineSurfaceElement

Set oBsplineSurfaceElement = GetBsplineSurface()

Dim oEnumerator As ElementEnumerator
Dim pt As Point3d

Dim m As SmartSolidElement
Dim oe As ElementEnumerator
Set oe = SmartSolid.ConvertToSmartSolidElement(oBsplineSurfaceElement)
Do While oe.MoveNext
    Set m = oe.Current
Loop

Dim ray As Ray3d

ray.Origin.X = 10
ray.Origin.Y = 0
ray.Origin.Z = 9
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

pt = result.RaySolidIntersection(ray)

Set oEnumerator = SmartSolid.TrimSolidWithSurface(result, m, pt)

Do While oEnumerator.MoveNext
        Dim oElement As Element
        Set oElement = oEnumerator.Current
        ActiveModelReference.AddElement oElement
Loop

tr = Transform3dFromXYZ(20, 0, 0)
result.Transform tr
oBsplineSurfaceElement.Transform tr

ActiveModelReference.AddElement oBsplineSurfaceElement
oBsplineSurfaceElement.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub

Public Sub TestCreateSectionFromSolid()
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim oBsplineSurfaceElement As BsplineSurfaceElement

Set oBsplineSurfaceElement = GetBsplineSurface()

ActiveModelReference.AddElement oBsplineSurfaceElement
oBsplineSurfaceElement.Redraw

Dim oEnumerator As ElementEnumerator

Set oEnumerator = SmartSolid.CreateSectionFromSolid(result, oBsplineSurfaceElement)

Do While oEnumerator.MoveNext
    Dim oElement As Element
    Set oElement = oEnumerator.Current
    
    Dim tr As Transform3d
    tr = Transform3dFromXYZ(15, 0, 0)
    oElement.Transform tr

    ActiveModelReference.AddElement oElement
Loop

ActiveModelReference.AddElement result
result.Redraw

End Sub

Public Sub TestBlendEdges()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 0

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim pts(0 To 1) As Point3d

pts(0).X = 0
pts(0).Y = 2.5
pts(0).Z = 10


pts(1).X = 0
pts(1).Y = -2.5
pts(1).Z = 10

Set SSolid = SmartSolid.BlendEdges(result, pts, 0.5, True)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub
Public Sub TestBlendAllEdgesOfFace()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim pt As Point3d

Dim ray As Ray3d

ray.Origin.X = 10
ray.Origin.Y = 0
ray.Origin.Z = 8
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

pt = result.RaySolidIntersection(ray)

Set SSolid = SmartSolid.BlendAllEdgesOfFace(result, pt, 0.5, True)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub
Public Sub TestBlendEdgeWithVariableRadius()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim pt As Point3d

Dim ray As Ray3d

ray.Origin.X = 10
ray.Origin.Y = 0
ray.Origin.Z = Sqr(3) * 2.5
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

pt = result.RaySolidIntersection(ray)

Set SSolid = SmartSolid.BlendEdgeWithVariableRadius(result, pt, 0.5, False, 0.1, 0.8)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub

Public Sub TestChamferEdge()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim pt As Point3d

Dim ray As Ray3d

ray.Origin.X = 10
ray.Origin.Y = 0
ray.Origin.Z = Sqr(3) * 2.5
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

pt = result.RaySolidIntersection(ray)

Set SSolid = SmartSolid.ChamferEdge(result, pt, 0.5, 0.5, False)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw


End Sub

Public Sub TestOffsetFace()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim pt As Point3d

Dim ray As Ray3d

ray.Origin.X = 25
ray.Origin.Y = 0
ray.Origin.Z = 0
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

pt = result.RaySolidIntersection(ray)

Set SSolid = SmartSolid.OffsetFace(result, pt, 3)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub
Public Sub TestRemoveFace()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement
Dim pt As Point3d

Dim ray As Ray3d

ray.Origin.X = 25
ray.Origin.Y = 0
ray.Origin.Z = 0
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

pt = result.RaySolidIntersection(ray)

Set SSolid = SmartSolid.RemoveFace(result, pt, 3, True)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub

Public Sub TestProjectCurveOntoSolid()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)

Dim SSolid As SmartSolidElement
Dim c As EllipseElement


Set c = Application.CreateEllipseElement1(Nothing, Point3dFromXYZ(20, 0, 2.5), Point3dFromXYZ(20, 2, -2), Point3dFromXYZ(20, -2, -2))
Set SSolid = SmartSolid.ProjectCurveOntoSolid(result, c, Vector3dFromXYZ(-1, 0, 0), 0.00001)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub


Public Sub TestRevolveProfileAxis()
Dim c As LineElement
Set c = Application.CreateLineElement2(Nothing, Point3dFromXYZ(0, 0, 0), Point3dFromXYZ(0, 0, 5))
ActiveModelReference.AddElement c

Dim SSolid As SmartSolidElement
Set SSolid = SmartSolid.RevolveProfile(c, Point3dFromXYZ(10, 0, 0), Vector3dFromXYZ(0, 0, 1), 1.5, 1)

ActiveModelReference.AddElement SSolid
SSolid.Redraw

End Sub
Public Sub TestSpinFace()
Dim tr As Transform3d
Dim result As SmartSolidElement

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Set result = SmartSolid.SolidUnion(sp, slab1)
Dim SSolid As SmartSolidElement

Set SSolid = SmartSolid.SpinFace(result, Point3dFromXYZ(2.5, 0, 9), Point3dFromXYZ(10, 10, 10), Vector3dFromXYZ(0.5, 0.5, 1), 1.5)

tr = Transform3dFromXYZ(15, 0, 0)
SSolid.Transform tr

ActiveModelReference.AddElement SSolid
SSolid.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub
Public Sub TestSweepProfileAlongPathAndThinshell()

Dim path As BsplineCurveElement

Dim pathcurve As New BsplineCurve
Dim profilecurve As New BsplineCurve

Dim aPoints() As Point3d

'Dim profile As BsplineCurveElement
'ReDim aPoints(0 To 2)
'aPoints(0) = Point3dFromXYZ(0, 0, 0)
'aPoints(1) = Point3dFromXYZ(1, 3, 0)
'aPoints(2) = Point3dFromXYZ(2, 0, 0)
'profilecurve.SetPoles aPoints
'Set profile = Application.CreateBsplineCurveElement1(Nothing, profilecurve)

Dim profile As ShapeElement
ReDim aPoints(0 To 3)
aPoints(0) = Point3dFromXYZ(1.5, -1.5, 4)
aPoints(1) = Point3dFromXYZ(-1.5, -1.5, 4)
aPoints(2) = Point3dFromXYZ(-1.5, 1.5, 4)
aPoints(3) = Point3dFromXYZ(1.5, 1.5, 4)
Set profile = CreateShapeElement1(Nothing, aPoints, msdFillModeNotFilled)

ReDim aPoints(0 To 2)
aPoints(0) = Point3dFromXYZ(0, 0, 0)
aPoints(1) = Point3dFromXYZ(0, 0, 1)
aPoints(2) = Point3dFromXYZ(0, 0, 2)
pathcurve.SetPoles aPoints
Set path = Application.CreateBsplineCurveElement1(Nothing, pathcurve)

Dim result As SmartSolidElement
Set result = SmartSolid.SweepProfileAlongPath(profile, path)
Dim result1 As SmartSolidElement
Dim aa() As Double
ReDim aa(0 To 2)
aa(0) = 2
aa(1) = 0
aa(2) = 0
Set result1 = SmartSolid.ThinShell(result, 0.1)

ActiveModelReference.AddElement result1
result1.Redraw

End Sub

Sub TestExtrudeClosedPlanarCurve()

Dim base As EllipseElement
Set base = Application.CreateEllipseElement2(Nothing, Point3dFromXYZ(0, 0, 0), 5, 5, Matrix3dIdentity, msdFillModeNotFilled)

Dim ConeSurface As SmartSolidElement
Set ConeSurface = SmartSolid.ExtrudeClosedPlanarCurve(base, 10, 10, True)

ActiveModelReference.AddElement ConeSurface
ConeSurface.Redraw

End Sub

' fills an 8x8 pole array (through which we "hang" a torus-like biperiodic surface)
Function computePoles(aPoles() As Point3d, nUPoles As Long, nVPoles As Long, radiusOuter As Double, radiusInner As Double) As Point3d()
    Dim radiusMid As Double, radius As Double, height As Double
    Dim col As Long
    
    ' this algorithm only works for an 8x8 array
    If nUPoles <> 8 Or nVPoles <> 8 Then
        Exit Function
    End If
    ReDim aPoles(0 To nVPoles - 1, 0 To nUPoles - 1)
    
    radiusMid = (radiusOuter + radiusInner) / 2#
    
    ' array has dimensions nVPole x nUPole, with faster increase in u-direction
    For col = LBound(aPoles, 2) To UBound(aPoles, 2)
     
        Dim yy As Double
        yy = -8 + col * 2
        
        aPoles(0, col) = Point3dFromXYZ(3, yy, 12)
        aPoles(1, col) = Point3dFromXYZ(1.5, yy, 9)
        aPoles(2, col) = Point3dFromXYZ(0, yy, 6)
        aPoles(3, col) = Point3dFromXYZ(-1.5, yy, 3)
        aPoles(4, col) = Point3dFromXYZ(-3, yy, 0)
        aPoles(5, col) = Point3dFromXYZ(-0.5, yy, -3)
        aPoles(6, col) = Point3dFromXYZ(2, yy, -6)
        aPoles(7, col) = Point3dFromXYZ(4, yy, -11)
    Next col
    
    computePoles = aPoles
End Function

' sets normalized nonuniform interior knots
Function computeKnots(aKnots() As Double, nKnot As Long, intraKnotClusterGap As Double) As Double()
    Dim interKnotClusterGap As Double, knotVal As Double
    Dim i As Long, nCluster As Long
    
    If nKnot Mod 2 = 0 Then
        Exit Function ' algorithm only works for odd # interior knots
    End If
    ReDim aKnots(0 To nKnot - 1)
    
    nCluster = (nKnot + 1) \ 2
    If nCluster * intraKnotClusterGap >= 1# Then
        Exit Function
    End If
        
    interKnotClusterGap = (1# - (nCluster * intraKnotClusterGap)) / nCluster
    knotVal = 0
    For i = LBound(aKnots) To UBound(aKnots)
        If i Mod 2 = 0 Then
            knotVal = knotVal + interKnotClusterGap
        Else
            knotVal = knotVal + intraKnotClusterGap
        End If
        
        aKnots(i) = knotVal
    Next i
    
    computeKnots = aKnots
End Function
Function GetBsplineSurface() As BsplineSurfaceElement
    Dim oBsplineSurfaceElement As BsplineSurfaceElement
    Dim oBsplineSurface As New BsplineSurface
    Dim aPoles() As Point3d
    Dim aKnots() As Double
    Dim nUPole As Long, nVPole As Long, nUKnot As Long, nVKnot As Long
    
    ' Construct a biperiodic biquadratic punctured NURBS surface
    oBsplineSurface.VOrder = 3
    oBsplineSurface.UOrder = 3
    oBsplineSurface.VClosed = False
    oBsplineSurface.UClosed = False
    
    '...set 8x8 poles array (and uniform knots)...
    nUPole = 8
    nVPole = 8
    oBsplineSurface.SetPoles computePoles(aPoles, nUPole, nVPole, 3#, 1#)
    
    '...set nonuniform u- and v-knots...
    nUKnot = Bspline.ComputeKnotsCount(oBsplineSurface.UPolesCount, oBsplineSurface.UOrder, oBsplineSurface.UClosed)
    nVKnot = Bspline.ComputeKnotsCount(oBsplineSurface.VPolesCount, oBsplineSurface.VOrder, oBsplineSurface.VClosed)
    oBsplineSurface.SetUKnots computeKnots(aKnots, nUKnot, 0.1)
    oBsplineSurface.SetVKnots computeKnots(aKnots, nVKnot, 0.1)
    
    ' Create the element from our working definition, and add it to the active model
    Set oBsplineSurfaceElement = CreateBsplineSurfaceElement1(Nothing, oBsplineSurface)
    Set GetBsplineSurface = oBsplineSurfaceElement
End Function

