VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBConnect 
   Caption         =   "Database Login"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmDBConnect.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Private Sub cmdDBConnect_Click()
On Error GoTo RDOCC_EH

'accept values from the user
DNSname = txtDNSName.Value
UserId = txtUserName.Value
Password = txtPasswd.Value

If DNSname <> "" And UserId <> "" And Password <> "" Then
    UIDPWD = "Provider=OraOLEDB.Oracle;User ID="
    UIDPWD = UIDPWD & UserId
    UIDPWD = UIDPWD & ";Password="
    UIDPWD = UIDPWD & Password
    UIDPWD = UIDPWD & ";Data Source="
    UIDPWD = UIDPWD & DNSname
    
    Set ADOconn = New Connection
    ADOconn.ConnectionString = UIDPWD
    ADOconn.Open
    Unload Me
End If

Exit Sub
    
RDOCC_EH:
    MsgBox ("ERROR " & Err.Number & "  " & Err.Description)
End Sub

Private Sub cmdCancel_Click()
End
End Sub

