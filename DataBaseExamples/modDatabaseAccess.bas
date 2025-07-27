Attribute VB_Name = "modDatabaseAccess"
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
Option Explicit

'Get the MSCATALOG name if it is set as a Configuration variable.
Public Function GetMSCATALOG() As String

If ActiveWorkspace.IsConfigurationVariableDefined("MS_MSCATALOG") Then
   GetMSCATALOG = UCase$(ActiveWorkspace.ExpandConfigurationVariable("$(MS_MSCATALOG)"))
Else
   GetMSCATALOG = "MSCATALOG"
End If

End Function

'Process Query string and return result set
Public Function GetRecordSet(myConn As ADODB.Connection, sQry As String) As ADODB.Recordset

Dim RecSet As Recordset
Set RecSet = New Recordset

RecSet.CursorType = adOpenKeyset
RecSet.LockType = adLockOptimistic
RecSet.CursorLocation = adUseClient
RecSet.Source = sQry
Set RecSet.ActiveConnection = myConn
RecSet.Open

Set GetRecordSet = RecSet

End Function

'Get the Table name for a given entity.
 Public Function GetTableName(MSCATstr As String, EntNum As Integer) As String

    Dim RecordCount As Integer
    Dim RecSet1 As ADODB.Recordset
    Set RecSet1 = GetRecordSet(ADOconn, "SELECT TABLE_NAME FROM USER_TABLES WHERE TABLE_NAME = '" + MSCATstr + "'")
    RecordCount = RecSet1.RecordCount

    If RecordCount < 1 Then
    'no entry found for given tablename
        MsgBox "MSCATALOG Not Found!"
        RecSet1.Close
        Set RecSet1 = Nothing
        GetTableName = ""
        
    ElseIf RecordCount = 1 Then
        RecordCount = 0
        Dim RecSet2 As ADODB.Recordset
        Set RecSet2 = GetRecordSet(ADOconn, "SELECT TABLENAME FROM " + MSCATstr + " WHERE ENTITYNUM = " & CStr(EntNum))
        RecordCount = RecSet2.RecordCount
               
        If RecordCount = 0 Then
            MsgBox "No Table for Entity: " & CStr(EntNum), vbOKOnly
            RecSet2.Close
            Set RecSet2 = Nothing
            GetTableName = ""
            
        ElseIf RecordCount = 1 Then
            GetTableName = RecSet2!TableName
            RecSet2.Close
            Set RecSet2 = Nothing
        
        Else
            MsgBox "Multiple Tables for Entity: " & CStr(EntNum), vbOKOnly
            RecSet2.Close
            Set RecSet2 = Nothing
            GetTableName = ""
        End If
        
    Else
        MsgBox "MSCATALOG is not Unique!"
        RecSet1.Close
        Set RecSet1 = Nothing
        GetTableName = ""
    End If

End Function
