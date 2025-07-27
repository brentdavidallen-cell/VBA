Attribute VB_Name = "modElementAccess"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

'  This module has functions for converting between the various types representing
'  elements.  These types include the VBA Element and the MDL types MSElementDescr
'  and ElementRef.
'
#If VBA7 Then
Declare PtrSafe Function ElmdscrAccessor_getMSElement Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getNext Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getPrevious Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getHeader Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getFirst Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getDgnModelRef Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getElementRef Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getIsHeader Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As Long
Declare PtrSafe Function ElmdscrAccessor_getIsValid Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As Long
Declare PtrSafe Function ElmdscrAccessor_getUserData1 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Function ElmdscrAccessor_getUserData2 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr) As LongPtr
Declare PtrSafe Sub ElmdscrAccessor_setUserData1 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr, ByVal NewValue As LongPtr)
Declare PtrSafe Sub ElmdscrAccessor_setUserData2 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As LongPtr, ByVal NewValue As LongPtr)
Declare PtrSafe Function DialogItemAccessor_getRawItem Lib "stdmdlaccessor.dll" (ByVal DialogItem As LongPtr) As LongPtr
Declare PtrSafe Function elementRef_getElemID Lib "stdmdlbltin.dll" (ByVal elemRef As LongPtr) As DLong
#Else
Declare  Function ElmdscrAccessor_getMSElement Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getNext Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getPrevious Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getHeader Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getFirst Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getDgnModelRef Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getElementRef Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getIsHeader Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getIsValid Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getUserData1 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Function ElmdscrAccessor_getUserData2 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long) As Long
Declare Sub ElmdscrAccessor_setUserData1 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long, ByVal NewValue As Long)
Declare Sub ElmdscrAccessor_setUserData2 Lib "stdmdlaccessor.dll" (ByVal ElementDescr As Long, ByVal NewValue As Long)
Declare Function elementRef_getElemID Lib "stdmdlbltin.dll" (ByVal elemRef As Long) As DLong
#End If

'  Gets the element descriptor but does not detach it from TheElement
'  Since it is not detached, the lifetime of the element descriptor is
'  identical to the lifetime of  TheElement.  The VBA code should not free
'  the element descriptor, and should not use it unless it also has a reference to
'  TheElement
Function GetElmdscrP(TheElement As Element) As LongPtr
    If TheElement Is Nothing Then
        Err.Raise msdErrorBadElement, "MDL examples", "The Element is Nothing"
        Exit Function
    End If
    GetElmdscrP = TheElement.MdlElementDescrP(False)
End Function
'  Gets an ElementRef from an Element.  The ElementRef
'  is returned as a Long
Function GetElementRefFromElement(TheElement As Element) As LongPtr
    Dim elemDescr As LongPtr
    
    elemDescr = GetElmdscrP(TheElement)
    GetElementRefFromElement = ElmdscrAccessor_getElementRef(elemDescr)
End Function
'  Gets a pointer to the MSElement in the Element object.
'  The pointer is returned as a Long
Function GetMSElementFromElement(TheElement As Element) As LongPtr
    Dim elemDescr As LongPtr
    
    elemDescr = GetElmdscrP(TheElement)
    GetMSElementFromElement = ElmdscrAccessor_getMSElement(elemDescr)
End Function
'  Creates an Element object from the element descriptor.
'  The element descriptor becomes part of the Element, so the
'  VBA program must not free it.
Function GetElementObjectFromElmdscr(ElmdscrP As LongPtr) As Element
    If ElmdscrP = 0 Then
        Err.Raise msdErrorBadElement, "MDL examples", "The Element is Nothing"
        Exit Function
    End If
    
    Set GetElementObjectFromElmdscr = MdlCreateElementFromElementDescrP(ElmdscrP)
End Function
'  Create an Element object from an ElementRef and a DgnModelRefP
Function GetElementFromElementRef(ElementRef As LongPtr, DgnModelRefP As LongPtr) As Element
    Dim ID As DLong
    Dim oModelRef As ModelReference
    
    '  Use the accessor method to get the element ID
    '  from the ElementRef
    ID = elementRef_getElemID(ElementRef)
    
    '  Use Application's hidden method MdlGetModelReferenceFromModelRefP
    '  to get the ModelReference object corresponding to DgnModelRef
    If DgnModelRefP = 0 Then
        Set oModelRef = ActiveModelReference
    Else
        Set oModelRef = MdlGetModelReferenceFromModelRefP(DgnModelRefP)
    End If
    
    '  Get the element
    Set GetElementFromElementRef = oModelRef.GetElementByID(ID)
End Function

'  This is similar to GetElementObjectFromElmdscr, but it reads the element.
'  The existing element descriptor does not become associated with the element.
Function LoadElementObjectFromElmdscr(ElmdscrP As LongPtr) As Element
    If ElmdscrP = 0 Then
        Err.Raise msdErrorBadElement, "MDL examples", "The Element is Nothing"
        Exit Function
    End If
    
    Dim DgnModelRefP As LongPtr
    Dim elemRef As LongPtr
    
    '  Use the accessor function to get the MDL DgnModelRefP from
    '  the element descriptor
    DgnModelRefP = ElmdscrAccessor_getDgnModelRef(ElmdscrP)
    
    '  Use the accessor function to get the ElementRef from
    '  the element descriptor
    elemRef = ElmdscrAccessor_getElementRef(ElmdscrP)
    
    Set LoadElementObjectFromElmdscr = GetElementFromElementRef(elemRef, DgnModelRefP)
End Function
'
