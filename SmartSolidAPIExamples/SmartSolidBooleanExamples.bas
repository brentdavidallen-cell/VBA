Attribute VB_Name = "SmartSolidBooleanExamples"
'TestSolidIntersect is an example of SmartSolid.SolidIntersect
'TestSolidSubtract is an example of SmartSolid.SolidSubtract
'TestSolidUnion is an example of SmartSolid.SolidUnion
'TestBooleanDisjoint is an example of SmartSolid.BooleanDisjoint
'TestDifference shows the difference bewteen SolidIntersect and BooleanDisjoint with
'  MODELER_BOOLEAN_difference = 2 as input operation.

Public Sub TestSolidIntersect()
Dim tr As Transform3d

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Dim resultbody As SmartSolidElement
Set resultbody = SmartSolid.SolidIntersect(sp, slab1)

ActiveModelReference.AddElement resultbody

End Sub
Public Sub TestSolidSubtract()
Dim tr As Transform3d

'group 1
Dim sp As SmartSolidElement
Set sp = SmartSolid.CreateSphere(Nothing, 5)
sp.Color = 1

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 5, 5, 20)
slab1.Color = 2

Dim resultbody As SmartSolidElement
Set resultbody = SmartSolid.SolidSubtract(sp, slab1)

ActiveModelReference.AddElement resultbody

End Sub
Public Sub TestSolidUnion()
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

ActiveModelReference.AddElement result
result.Redraw

End Sub

Public Sub TestBooleanDisjoint()
Dim tr As Transform3d
Dim result As SmartSolidElement
Dim sp(4) As SmartSolidElement

'group 1
'1
Set sp(1) = SmartSolid.CreateSphere(Nothing, 2)
sp(1).Color = 1

tr = Transform3dFromXYZ(10, 10, 0)
sp(1).Transform tr

'2
Set sp(0) = SmartSolid.CreateSphere(Nothing, 2)
sp(0).Color = 1

tr = Transform3dFromXYZ(-10, 10, 0)
sp(0).Transform tr

'3
Set sp(2) = SmartSolid.CreateSphere(Nothing, 2)
sp(2).Color = 1

tr = Transform3dFromXYZ(-10, -10, 0)
sp(2).Transform tr

'4
Set sp(3) = SmartSolid.CreateSphere(Nothing, 2)
sp(3).Color = 1

tr = Transform3dFromXYZ(10, -10, 0)
sp(3).Transform tr

Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 20, 20, 20)
slab1.Color = 2


Dim oEnumerator As ElementEnumerator

Set oEnumerator = SmartSolid.BooleanDisjoint(slab1, sp, 1)

Do While oEnumerator.MoveNext
        Dim oElement As Element
        Set oElement = oEnumerator.Current
        ActiveModelReference.AddElement oElement
Loop

End Sub

Public Sub TestDifference()
Dim tr As Transform3d
Dim result As SmartSolidElement
Dim sp(1) As SmartSolidElement

'group 1
'1
Set sp(1) = SmartSolid.CreateSlab(Nothing, 5, 30, 15)
sp(1).Color = 1

'main body
Dim slab1 As SmartSolidElement
Set slab1 = SmartSolid.CreateSlab(Nothing, 15, 15, 15)
slab1.Color = 2


Dim oEnumerator As ElementEnumerator

Set oEnumerator = SmartSolid.BooleanDisjoint(slab1, sp, 2)

Do While oEnumerator.MoveNext
        Dim oElement As Element
        Set oElement = oEnumerator.Current
        tr = Transform3dFromXYZ(10, -10, 0)
        oElement.Transform tr
        ActiveModelReference.AddElement oElement
Loop

'2
Dim second As SmartSolidElement
Set second = SmartSolid.CreateSlab(Nothing, 5, 30, 15)
second.Color = 3

'main body
Dim first As SmartSolidElement
Set first = SmartSolid.CreateSlab(Nothing, 15, 15, 15)
first.Color = 3

Dim onresult As SmartSolidElement
Set onresult = SmartSolid.SolidSubtract(first, second)
ActiveModelReference.AddElement onresult
End Sub



