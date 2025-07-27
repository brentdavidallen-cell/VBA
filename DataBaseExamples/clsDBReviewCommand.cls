VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBReviewCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'

Implements ILocateCommandEvents

Private Sub ILocateCommandEvents_Accept(ByVal Element As Element, Point As Point3d, ByVal View As View)
    Set oElement = Element
    dbLinks = oElement.GetDatabaseLinks
    LinkNum = LBound(dbLinks)
    
    '  Show a modal dialog box to step throough all of the linkage
    frmDBLinkInfo.PrintLink LinkNum
    
    '  Restart the command so the user can select another element
    CommandState.StartLocate New clsDBReviewCommand
End Sub

Private Sub ILocateCommandEvents_Cleanup()
End Sub

Private Sub ILocateCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
End Sub

Private Sub ILocateCommandEvents_LocateFailed()
    CommandState.StartLocate New clsDBReviewCommand
End Sub

Private Sub ILocateCommandEvents_LocateFilter(ByVal Element As Element, Point As Point3d, Accepted As Boolean)
End Sub

Private Sub ILocateCommandEvents_LocateReset()
    CommandState.StartLocate New clsDBReviewCommand
End Sub

Private Sub ILocateCommandEvents_Start()
    '  Since this command does not modify the original element,
    '  set the locate criteria to allow the read-only elements.
    Dim lc As LocateCriteria
    Set lc = CommandState.CreateLocateCriteria(False)
    CommandState.SetLocateCriteria lc
    
    ShowCommand "Review Element"
    ShowPrompt "Select element to Review"
End Sub


