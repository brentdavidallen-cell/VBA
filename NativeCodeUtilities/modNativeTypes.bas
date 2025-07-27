Attribute VB_Name = "modNativeTypes"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

#If VBA7 Then
'  Use this if both pointers are represented as longs.  Typically only need this if the neither the source nor the
'  destination is in memory allocated by VBA
Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal strDest As LongPtr, ByVal lpSource As LongPtr, ByVal length As Long)

'  Use this if source is outside of VBA and therefore represented as a Long, but the
'  destination is in an area allocated by VBA
Declare PtrSafe Sub CopyMemoryToVBA Lib "kernel32" Alias "RtlMoveMemory" _
  (ByRef VBALocation As Any, ByVal SourceLoc As LongPtr, ByVal length As Long)
  
'  Use this if destination is outside of VBA and therefore represented as a Long, but the
'  source is in an area allocated by VBA
Declare PtrSafe Sub CopyMemoryFromVBA Lib "kernel32" Alias "RtlMoveMemory" _
  (ByVal Destination As LongPtr, ByRef VBALocation As Any, ByVal length As Long)
  
Private Declare PtrSafe Function SysAllocString Lib "oleaut32" (ByVal CharPtr As LongPtr) As String
#Else
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal strDest As Long, ByVal lpSource As Long, ByVal length As Long)
Declare Sub CopyMemoryToVBA Lib "kernel32" Alias "RtlMoveMemory" _
  (ByRef VBALocation As Any, ByVal SourceLoc As Long, ByVal length As Long)
Declare Sub CopyMemoryFromVBA Lib "kernel32" Alias "RtlMoveMemory" _
  (ByVal Destination As Long, ByRef VBALocation As Any, ByVal length As Long)
  
Private Declare Function SysAllocString Lib "oleaut32" (ByVal CharPtr As Long) As String
#End If
'  Creates a String from a char *
Function CharPtrToString(CharPtr As Long) As String
    If CharPtr = 0 Then
        CharPtrToString = ""
    Else
        CharPtrToString = SysAllocString(CharPtr)
    End If
End Function

'  Truncates a buffer to just contain the string that
'  the C function returned.
Function TruncateAtEOS(str As String) As String
    If str <> "" Then  ' This test avoids exception on zero-length string
        TruncateAtEOS = Left$(str, InStr(1, str, vbNullChar) - 1)
    End If
End Function

Sub CopyPoint3dToVBA(ByRef pnt As Point3d, ByVal lPointer As Long)
    CopyMemoryToVBA pnt, lPointer, 24
End Sub
Sub CopyPoint3dFromVBA(ByVal lPointer As Long, ByRef pnt As Point3d)
    CopyMemoryFromVBA lPointer, pnt, 24
End Sub
