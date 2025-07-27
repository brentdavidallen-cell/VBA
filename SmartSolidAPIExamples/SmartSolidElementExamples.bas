Attribute VB_Name = "SmartSolidElementExamples"
'TestComputeVolume is an example of SmartSolidElement.ComputeVolume
'RayTest is an example of SmartSolidElement.RaySolidIntersection
'TestCapSurface is an example of SmartSolidElement.CapSurface
'TestGetVertices is an example of SmartSolidElement.GetVertices
'TestFacetSolidAsShapes is an example of SmartSolidElement.FacetSolidAsShapes
'TestFacetSolidAsMesh is an example of SmartSolidElement.FacetSolidAsMesh
'TestExtractSurfaceFromSolid is an example of SmartSolidElement.ExtractSurfaceFromSolid
'TestExtractAllSurfaceFromSolid is an example of SmartSolidElement.ExtractAllSurfaceFromSolid


Public Sub TestComputeVolume()
Dim vol As Double

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 10, 10, 10)
slab1.Color = 2

vol = slab1.ComputeVolume()

ActiveModelReference.AddElement slab1
slab1.Redraw

End Sub

Public Sub RayTest()
Dim sp As SmartSolidElement
Dim ray As Ray3d
Dim p As Point3d

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 10, 10, 10)
slab1.Color = 2

ray.Origin.X = 5
ray.Origin.Y = 5
ray.Origin.Z = 20
ray.Direction.X = 0
ray.Direction.Y = 0
ray.Direction.Z = -1

p = slab1.RaySolidIntersection(ray)

ActiveModelReference.AddElement slab1
slab1.Redraw

End Sub

Sub TestCapSurface()

Dim base As EllipseElement
Set base = Application.CreateEllipseElement2(Nothing, Point3dFromXYZ(0, 0, 0), 5, 5, Matrix3dIdentity, msdFillModeNotFilled)

Dim ConeSurface As SmartSolidElement
Set ConeSurface = SmartSolid.ExtrudeClosedPlanarCurve(base, 10, 10)

ActiveModelReference.AddElement ConeSurface
ConeSurface.Redraw

If (ConeSurface.IsSheetBody()) Then
    ConeSurface.Color = 2
    Dim tr As Transform3d
    tr = Transform3dFromXYZ(17.5, 17.5, 0)
    ConeSurface.Transform tr
    
    ConeSurface.CapSurface

    ConeSurface.Redraw msdDrawingModeNormal
    ConeSurface.Rewrite
End If

End Sub

Public Sub TestGetVertices()
Dim vol As Double

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 10, 10, 10)
slab1.Color = 2

Dim points() As Point3d
points = slab1.GetVertices()

ActiveModelReference.AddElement slab1
slab1.Redraw

End Sub

Public Sub TestFacetSolidAsShapes()
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

Dim oEnumerator As ElementEnumerator
Set oEnumerator = result.FacetSolidAsShapes(3, 1, 1, 90)

Do While oEnumerator.MoveNext
        Dim oElement As Element
        Set oElement = oEnumerator.Current
        ActiveModelReference.AddElement oElement
Loop

ActiveModelReference.AddElement result
result.Redraw

End Sub

Public Sub TestFacetSolidAsMesh()
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

Dim meshelm As Element
Set meshelm = result.FacetSolidAsMesh(3, 1, 1, 90)

ActiveModelReference.AddElement meshelm
meshelm.Redraw

End Sub

Public Sub TestExtractSurfaceFromSolid()
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

Dim ray As Ray3d
Dim p As Point3d

ray.Origin.X = 15
ray.Origin.Y = 0
ray.Origin.Z = 0
ray.Direction.X = -1
ray.Direction.Y = 0
ray.Direction.Z = 0

Dim bs As BsplineSurfaceElement
Dim pt As Point3d

pt = result.RaySolidIntersection(ray)
Set bs = result.ExtractSurfaceFromSolid(pt)
tr = Transform3dFromXYZ(20, 0, 0)
bs.Transform tr

ActiveModelReference.AddElement bs
result.Redraw

ActiveModelReference.AddElement result
result.Redraw

End Sub
Public Sub TestExtractAllSurfaceFromSolid()
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

Dim oEnumerator As ElementEnumerator
Set oEnumerator = result.ExtractAllSurfaceFromSolid()

Do While oEnumerator.MoveNext
        Dim oElement As Element
        Set oElement = oEnumerator.Current
        ActiveModelReference.AddElement oElement
Loop

tr = Transform3dFromXYZ(20, 0, 0)
result.Transform tr
ActiveModelReference.AddElement result
result.Redraw


End Sub

