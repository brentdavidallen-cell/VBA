Attribute VB_Name = "modCommands"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

Sub ScanExample()
    Dim scanCriteria As New ElementScanCriteria
    Dim i As Long
    
    ' Terminate any active commands
    CadInputQueue.SendCommand "NULL"
    
    '  Initially, the scan criteria is set up to handle all types.
    '  First, exclude all and then just include the ones we want
    scanCriteria.ExcludeAllTypes
    scanCriteria.IncludeType msdElementTypeText
    scanCriteria.IncludeType msdElementTypeEllipse
    scanCriteria.IncludeType msdElementTypeLineString
    scanCriteria.IncludeType msdElementTypeLine
        
    Dim enumerator As ElementEnumerator
    Set enumerator = ActiveModelReference.Scan(scanCriteria)
    ShowStatus "Updating elements"
    
    '  Step through the scan results
    Do While enumerator.MoveNext
        With enumerator.Current
            .Color = (.Color + 1) Mod 10
            .Rewrite
            .Redraw
        End With
        
        i = i + 1
        If (i Mod 50) = 0 Then
            ShowStatus "Updated element number " & i
        End If
    Loop
    
    ' Set MicroStation back to normal state
    CommandState.StartDefaultCommand
    ShowStatus "Updated element number " & i
End Sub


