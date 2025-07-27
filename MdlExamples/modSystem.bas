Attribute VB_Name = "modSystem"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit
Declare PtrSafe Function mdlSystem_nextMdlApp Lib "stdmdlbltin.dll" (ByVal descP As LongPtr) As LongPtr   '  Returns a pointer to a structure
Declare PtrSafe Sub mdlSystem_getMdlTaskID Lib "stdmdlbltin.dll" (ByVal result As String, ByVal mdlDescP As LongPtr)

'  This example illustrates how to handle a create a String from a char * value
'  that an MDL built-in returns
Sub ListMdlApps()
    Dim mdlDesc As LongPtr
    Dim strTaskId As String
    Dim length As Long
    
    mdlDesc = mdlSystem_nextMdlApp(0)
    Do While mdlDesc <> 0
        strTaskId = Space(512)
        ' mdlSystem_getMdlTaskID returns a char * value
        ' that points to the name of the MDL task
        mdlSystem_getMdlTaskID strTaskId, mdlDesc
        strTaskId = TruncateAtEOS(strTaskId)
        
        '  SysAllocString generates a String from a char * pointer
        Debug.Print "HAVE TASK ID: " & strTaskId
        
        ' Now get the next MDL application
        mdlDesc = mdlSystem_nextMdlApp(mdlDesc)
    Loop
End Sub
