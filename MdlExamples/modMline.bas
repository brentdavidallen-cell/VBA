Attribute VB_Name = "modMline"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Declare PtrSafe Function mdlMline_create Lib "stdmdlbltin.dll" (ByVal mline As LongPtr, ByVal seed As LongPtr, ByRef normal As Point3d, ByRef points As Point3d, ByVal nPoints As Long) As Long
Declare PtrSafe Function mdlElement_add Lib "stdmdlbltin.dll" (ByVal mline As LongPtr) As Long
'  This shows how to use an array to allocate space for an element buffer
Sub CreateMultiline()
    Dim index As Long
    Dim status As Integer
    Dim elementBuffer(0 To 65000) As Byte
    Dim pnts(0 To 2) As Point3d
    Dim normal As Point3d
    Dim ele As element
    
    normal.Z = 1
    
    pnts(0) = Point3dFromXY(-15, 10)
    pnts(1) = Point3dFromXY(13, -18)
    pnts(2) = Point3dFromXY(33, 3)
    
    status = mdlMline_create(VarPtr(elementBuffer(0)), 0, normal, pnts(0), 3)
    If status <> 0 Then Exit Sub
        
    ' Now add the element to the design file
    Dim filePos As Long
    filePos = mdlElement_add(VarPtr(elementBuffer(0)))
    
    If filePos = 0 Then Exit Sub
    
    '  Use the methods of GraphicalElementCache to get an
    '  Element object representing the new element
    With ActiveModelReference.GraphicalElementCache
        index = .IndexFromFilePosition(filePos)
        Set ele = .GetElement(index)
    End With
    
    ele.Redraw
End Sub
