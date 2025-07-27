VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrintEventHandlerExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPrintEvents

'This example demonstrates a simple print accounting utility,
'plus print text substitution and resymbolization.

Private Sub AppendToAcctLog()

    Dim fileNumber As Integer
    fileNumber = FreeFile

    Dim logicalAcctFilePath As String
    logicalAcctFilePath = "$(MS_PLTFILES)VbaPrintAcctLog.txt"
    
    Dim acctFilePath As String
    acctFilePath = ActiveWorkspace.ExpandConfigurationVariable(logicalAcctFilePath)
    
    Open acctFilePath For Append Access Write Lock Write As #fileNumber

    If (LOF(fileNumber) = 0) Then
        Print #fileNumber, "Date;Time;PsetFile;PltcfgFile;OutFile"
    End If
    
    'See "Tracking Printing Usage" in the Help for the list of available accounting variables.
    Dim printDate As String
    printDate = ActiveWorkspace.ConfigurationVariableValue("MS_PLTDATE")

    Dim printTime As String
    printTime = ActiveWorkspace.ConfigurationVariableValue("MS_PLTTIME")
    
    Dim psetFileName As String
    psetFileName = ActiveWorkspace.ConfigurationVariableValue("MS_PLTPSETFILE_SHORT")
    
    Dim pltcfgFileName As String
    pltcfgFileName = ActiveWorkspace.ConfigurationVariableValue("MS_PLTDRVFILE_SHORT")

    Dim outFileName As String
    outFileName = ActiveWorkspace.ConfigurationVariableValue("MS_PLTOUT")

    Dim logMsg As String
    logMsg = printDate & ";" & printTime & ";" & psetFileName & ";" & pltcfgFileName & ";" & outFileName
    
    Print #fileNumber, logMsg
    Close #fileNumber
        
End Sub

Private Sub IPrintEvents_BeginPrintSet()

    'When printing from Print Organizer, the BeginPrintSet and EndPrintSet events are invoked
    'in a different process than the per-page print events.  Therefore, the BeginPrintSet and
    'EndPrintSet event handlers must not share any global data with other print event handlers.
    'BeginPrintSet and EndPrintSet are never invoked when printing from the single-Print dialog.

End Sub

Private Sub IPrintEvents_EndPrintSet(ByVal isSuccess As Boolean)

    'When printing from Print Organizer, the BeginPrintSet and EndPrintSet events are invoked
    'in a different process than the per-page print events.  Therefore, the BeginPrintSet and
    'EndPrintSet event handlers must not share any global data with other print event handlers.
    'BeginPrintSet and EndPrintSet are never invoked when printing from the single-Print dialog.
    If ((isSuccess) And (Not PrintManager.IsPreview)) Then
        AppendToAcctLog
    End If

End Sub

Private Sub IPrintEvents_BeginPrintPage()
End Sub

Private Sub IPrintEvents_EndPrintPage(ByVal isSuccess As Boolean)

    If ((isSuccess) And (Not PrintManager.IsPreview) And (Not PrintManager.IsPrintSet)) Then
        AppendToAcctLog
    End If

End Sub

Private Sub IPrintEvents_BeginPrintElement(ByVal param As PrintBeginElementParam)

    'The element reference is *READ-ONLY* and *NON-COPYABLE*!
    'It is not legal to add, remove, or modify elements within a model while printing.
    Dim ele As Element
    Set ele = param.Element

    Dim eleType As MsdElementType
    eleType = ele.Type

    'Note that some elements, such as cell headers, may not have a valid level.
    Dim level As level
    Dim levelName As String
    Set level = ele.level
    If (level Is Nothing) Then
        levelName = vbNullString
    Else
        levelName = level.Name
    End If

    Dim newText As String
    newText = vbNullString

    If (levelName = "Border") Then
        If (eleType = msdElementTypeText) Then
            If (ele.AsTextElement.Text = "$PrintNotes$") Then
                newText = "Replacement Text"
            End If
        End If
    End If

    If (newText <> vbNullString) Then
    
        'This is the preferred technique for print text substitution, as it does not
        'requiring cloning the original element. It is valid for text and tag elements.
        param.SubstituteText newText
    
    End If

End Sub

Private Sub IPrintEvents_EndPrintElement(ByVal param As PrintEndElementParam)
End Sub

Private Sub IPrintEvents_PrintResymbolize(ByVal param As PrintResymbolizeParam)

    'The element reference is *READ-ONLY* and *NON-COPYABLE*!
    'It is not legal to add, remove, or modify elements within a model while printing.
    'Note that in some cases (i.e. pattern components) this may be a different element
    'than the one received in the BeginElement and EndElement events.
    Dim ele As Element
    Set ele = param.Element
    
    'The print display parameters may be altered inside this event handler.
    'For single-symbology elements, this event is invoked only once per element.
    'For multi-symbology elements, this event may be invoked more than once per
    'element with different print display parameters.
    Dim dp As PrintDisplayParams
    Set dp = param.PrintDisplayParams

    Dim eleType As MsdElementType
    eleType = ele.Type

    'Print text in blue, everything else in black.
    If (eleType = msdElementTypeText) Then
        dp.ColorRGB = RGB(0, 0, 255)
    Else
        dp.ColorRGB = RGB(0, 0, 0)
    End If

End Sub

Private Sub IPrintEvents_PrintDefCollectState(ByVal param As PrintDefCollectStateParam)
End Sub

