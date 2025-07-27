Attribute VB_Name = "modDBCommands"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'Global variables for connecting to the database
Public DNSname As String
Public UserId As String
Public Password As String
Public UIDPWD As String
Public ADOconn As ADODB.Connection

'Global variables needed for printing dblink info
Public LinkNum As Integer
Public oElement As Element
Public dbLinks() As DatabaseLink

Sub DBReview()
    frmDBLinkInfo.CommandButton1.Caption = "Review Next Link"
    CommandState.StartLocate New clsDBReviewCommand
    
    ' Only connect the first time this runs
    If ADOconn Is Nothing Then frmDBConnect.Show
End Sub

