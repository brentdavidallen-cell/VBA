Attribute VB_Name = "Module1"
Option Explicit
'
'  Copyright: (c) 2004 Bentley Systems, Incorporated. All rights reserved.
'
'  dbcheck.mvba
'
'  This sample application is designed to show how to use the Database Verification tool,
'  DBCHECK, in batch mode.  The application should be used with design file CD9.dgn or
'  CD10.dgn.  It will change all the linkages from the Database Type ODBC to OLEDB.
'

Sub UpdateLinkages()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Send a keyin to load DBCHECK application
    CadInputQueue.SendKeyin "mdl load dbcheck"

'   Send a keyin to review all elements in the file
    CadInputQueue.SendKeyin "dbcheck button review file"

'   Send a keyin to Select all Database linkages
    CadInputQueue.SendKeyin "dbcheck select all"

'   Send a keyin to check the DBTYPE Toggle
    CadInputQueue.SendKeyin "dbcheck toggle dbtype on"

'   Set a variable associated with a dialog box
'   Set the type variable to change the linkage to OLEDB
'   ODBC    -  24162
'   OLEDB   -  22528
'   ORACLE  -  24721
    SetCExpressionValue "dbGlobs->dbType", 22528, "DBCHECK"
    
'   Send keyin to process linkages
    CadInputQueue.SendKeyin "dbcheck button process"

'   Send keyin to unload application
    CadInputQueue.SendKeyin "mdl unload dbcheck"

    CommandState.StartDefaultCommand
End Sub
