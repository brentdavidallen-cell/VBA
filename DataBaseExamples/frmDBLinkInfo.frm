VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBLinkInfo 
   Caption         =   "DB Link Info"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDBLinkInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDBLinkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Private Sub CommandButton1_Click()

If CommandButton1.Caption = "Review Next Link" Then
    'print next link
    LinkNum = LinkNum + 1
    SetListBox LinkNum
Else
    Me.hide
End If

End Sub
Public Sub SetListBox(LinkIndex As Integer)
 frmDBLinkInfo.ListBox1.Clear
 
 frmDBLinkInfo.CommandButton1.Caption = "Close"

 If oElement.HasAnyDatabaseLinks Then
 'check to see if element has db link
    
    If LinkIndex < UBound(dbLinks) Then
        frmDBLinkInfo.CommandButton1.Caption = "Review Next Link"
    End If
    
    Dim mslink As Integer
    mslink = dbLinks(LinkIndex).mslink

    Dim EntNum As Integer
    EntNum = dbLinks(LinkIndex).EntityNumber
    
    Dim MSCATstr As String
    MSCATstr = GetMSCATALOG
    
    Dim TableName As String
    TableName = GetTableName(MSCATstr, EntNum)
     
    If TableName <> "" Then
    
        Dim RS1 As ADODB.Recordset
        Set RS1 = GetRecordSet(ADOconn, "SELECT * FROM " + TableName + " WHERE MSLINK = " & CStr(mslink))
        
        Dim RecordCount As Integer
        RecordCount = RS1.RecordCount
       
        If RecordCount = 0 Then
        ' no records found with matching mslink number
           frmDBLinkInfo.ListBox1.AddItem "No Records Found!!"
        
        ElseIf RecordCount = 1 Then
        ' one entry found for given mslink number
            Dim i As Integer
            For i = 0 To RS1.Fields.Count - 1
                frmDBLinkInfo.ListBox1.AddItem RS1.Fields.Item(i).Name
                If RS1.Fields.Item(i).Value <> "" Then
                   frmDBLinkInfo.ListBox1.List(i, 1) = RS1.Fields.Item(i).Value
                Else
                   frmDBLinkInfo.ListBox1.List(i, 1) = "NULL"
                End If
            Next
        
        Else
        ' multiple entries found for mslink number
            frmDBLinkInfo.ListBox1.AddItem "MSLINK is not Unique!!"
        End If
        
        RS1.Close
        Set RS1 = Nothing
    End If
Else
'element has no dblinks
   frmDBLinkInfo.ListBox1.AddItem "Element has no DB Links"
End If

End Sub
Public Sub PrintLink(LinkIndex As Integer)
    SetListBox LinkIndex
    frmDBLinkInfo.Show
End Sub

