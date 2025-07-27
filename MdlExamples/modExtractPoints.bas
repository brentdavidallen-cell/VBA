Attribute VB_Name = "modExtractPoints"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Declare PtrSafe Function mdlMline_extractPoints Lib "stdmdlbltin.dll" (ByRef outPoints As Point3d, ByVal mline As LongPtr, ByVal modelRef As LongPtr, ByVal pointNo As Long, ByVal nPoints As Long) As Long
'  Extracts the points from an Element that is a multiline
Sub ExtractPoints(eleMline As element, apnt() As Point3d)
    Dim ID As DLong
    Dim elDescrP As LongPtr
    Dim eleP As LongPtr
    Dim modelRef As LongPtr
    ReDim points(0 To 100) As Point3d
    Dim nPoints As Long
    
    If eleMline.Type <> msdElementTypeMultiLine Then
        Err.Raise msdErrorBadElement, "ExtractPoints", "The element is not a multiline"
        Exit Sub
    End If
    
    ' mdlMline_extractPoints requires a pointer to an MSElement.
    ' Use the hidden method MdlElementDescrP to get the
    ' address of the element descriptor corresponding to
    ' eleMline, and then use the accessor function
    ' to get the address of the element in the element
    ' descriptor
    elDescrP = eleMline.MdlElementDescrP
    eleP = ElmdscrAccessor_getMSElement(elDescrP)
    
    '  mdlMline_extractPoints requires a DgnModelRefP
    '  parameter.  Use the hidden method to get
    '  the DgnModelRefP corresponding to eleMline's
    '  ModelReference
    modelRef = eleMline.ModelReference.MdlModelRefP
    
    nPoints = mdlMline_extractPoints(points(0), eleP, modelRef, 0, 101)
    
    '  Now resize the array to be exactly the right size to
    '  hold the points.
    ReDim Preserve points(0 To nPoints - 1)
    apnt = points
End Sub

Sub ListPoints(ele As element)
    Dim pts() As Point3d
    Dim index As Long
    
    ExtractPoints ele, pts
    For index = LBound(pts) To UBound(pts)
        With pts(index)
            Debug.Print "(" & .X & ", " & .Y & ", " & .Z & ")"
        End With
    Next
End Sub

Sub TstListPoints()
    Dim ee As ElementEnumerator
    Set ee = ActiveModelReference.GraphicalElementCache.Scan
    Do While ee.MoveNext
        If ee.Current.Type = msdElementTypeMultiLine Then
            ListPoints ee.Current
        End If
    Loop
End Sub
