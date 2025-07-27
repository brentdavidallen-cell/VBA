Attribute VB_Name = "modLineStyleLinkage"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Private Const StyleLinkageID As Long = &H79F9

Private Type MdlStyleParam
    modifiers  As Long
    reserved   As Long
    mainScale  As Double
    dashScale  As Double
    gapScale   As Double
    startWidth As Double
    endwidth   As Double
    distPhase  As Double
    fractPhase As Double
    lineMask   As Long
    mlineFlags As Long
    normal     As Point3d
    rotation   As Matrix3d
    '  Following fields just allow room for growth of this structure
    reserved1  As Matrix3d
    reserved2  As Matrix3d
    reserved3  As Matrix3d
    reserved4  As Matrix3d
End Type

Declare PtrSafe Function mdlLineStyle_extractParams Lib "stdmdlbltin.dll" (ByRef param As MdlStyleParam, ByVal pElementIn As LongPtr) As Long

Sub ExtractLStyleInfo(ele As element, param As MdlStyleParam)
    Dim eleDescr As LongPtr
    Dim eleP As LongPtr
    Dim db() As DataBlock
    
    db = ele.GetUserAttributeData(StyleLinkageID)
    If UBound(db) < 0 Then
        Err.Raise -1, "ExtractLineStyle", "Element does not have a style linkage"
    End If
    
    '  First see if it has a line style linkage
    eleDescr = ele.MdlElementDescrP
    eleP = ElmdscrAccessor_getMSElement(eleDescr)
    mdlLineStyle_extractParams param, eleP
End Sub

Sub TestLSInfo()
    Dim params As MdlStyleParam
    Dim ele As element
    Dim ee As ElementEnumerator
    Set ee = ActiveModelReference.GetSelectedElements
    If Not ee.MoveNext Then Exit Sub
    
    ExtractLStyleInfo ee.Current, params
    
End Sub
