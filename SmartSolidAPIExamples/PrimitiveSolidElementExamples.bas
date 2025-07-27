Attribute VB_Name = "PrimitiveSolidElementExamples"
'TestSlab is an example of SmartSolid.CreateSlab
'TestSphere is an example of SmartSolid.CreateSphere
'TestCylinder is an example of SmartSolid.CreateCylinder
'TestCone is an example of SmartSolid.CreateCone
'TestTous is an example of SmartSolid.CreateTorus
'TestSlab is an example of SmartSolid.CreateSlab

Public Sub TestSlab()
Dim vol As Double

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 10, 10, 10)
slab1.Color = 2

ActiveModelReference.AddElement slab1
slab1.Redraw

End Sub
Public Sub TestSphere()
Dim tr As Transform3d
Dim sp1 As SmartSolidElement

Set sp1 = SmartSolid.CreateSphere(Nothing, 10)
sp1.Color = 2

'tr = Transform3dFromXYZ(20, 0, 0)
'sp1.Transform tr

ActiveModelReference.AddElement sp1
sp1.Redraw

End Sub
Public Sub TestCylinder()
Dim tr As Transform3d
Dim cy As SmartSolidElement

Set cy = SmartSolid.CreateCylinder(Nothing, 10, 20)
cy.Color = 1

tr = Transform3dFromXYZ(40, 0, 0)
cy.Transform tr

ActiveModelReference.AddElement cy
cy.Redraw

End Sub

Public Sub TestCone()
Dim tr As Transform3d
Dim cone As SmartSolidElement

Set cone = SmartSolid.CreateCone(Nothing, 5, 10, 20)
cone.Color = 2

tr = Transform3dFromXYZ(60, 0, 0)
cone.Transform tr

ActiveModelReference.AddElement cone
cone.Redraw

End Sub

Public Sub TestTous()
Dim tr As Transform3d
Dim torus As SmartSolidElement

Set torus = SmartSolid.CreateTorus(Nothing, 10, 2, 300)
torus.Color = 2

tr = Transform3dFromXYZ(80, 0, 0)
torus.Transform tr

ActiveModelReference.AddElement torus
torus.Redraw

End Sub

Public Sub TestWedge()
Dim tr As Transform3d
Dim torus As SmartSolidElement

Set torus = SmartSolid.CreateWedge(Nothing, 10, 2, 270)
torus.Color = 2

tr = Transform3dFromXYZ(80, 0, 0)
torus.Transform tr

ActiveModelReference.AddElement torus
torus.Redraw

End Sub
