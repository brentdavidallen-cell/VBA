Attribute VB_Name = "modViewConversions"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit
#If VBA7 Then
Declare PtrSafe Function mdlView_indexFromWindow Lib "stdmdlbltin.dll" (ByVal window As LongPtr) As Long
Declare PtrSafe Function mdlWindow_viewWindowGet Lib "stdmdlbltin.dll" (ByVal viewNum As Long) As LongPtr
Declare PtrSafe Function mdlWindow_isMaximized Lib "stdmdlbltin.dll" (ByVal windowP As LongPtr) As Long
#Else
Declare Function mdlView_indexFromWindow Lib "stdmdlbltin.dll" (ByVal window As Long) As Long
Declare Function mdlWindow_viewWindowGet Lib "stdmdlbltin.dll" (ByVal viewNum As Long) As Long
Declare Function mdlWindow_isMaximized Lib "stdmdlbltin.dll" (ByVal windowP As Long) As Long
#End If

Function MDLWindowToView(window As Long) As View
    Dim index As Long
    
    index = mdlView_indexFromWindow(window)
    Set MDLWindowToView = ActiveDesignFile.Views(index + 1)
End Function
Function ViewToMDLWindow(oView As View) As LongPtr
    ViewToMDLWindow = mdlWindow_viewWindowGet(oView.index - 1)
End Function
#If COMPILE_TEST_CODE Then
'  Set COMPILE_TEST_CODE to 1 in Tools->Properties->Conditional Compilation Arguments
'  to compile this test case
Sub Test()
    Dim ov As View
    Dim window As LongPtr
    
    Set ov = ActiveDesignFile.Views(2)
    window = ViewToMDLWindow(ov)
    Debug.Print "IsMaximized = " & CBool(mdlWindow_isMaximized(window))
End Sub
#End If
