VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestMacro 
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmTestMacro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' -----------------------------------------------------------------------
' Author:  Bentley Systems
' Copyright (c) 1999-2001;  Bentley Systems, Inc., 685 Stockton Drive,
'                      Exton PA, 19341-0678, USA.  All Rights Reserved.
'
' This program is confidential, proprietary and unpublished property of Bentley Systems
' Inc. It may NOT be copied in part or in whole on any medium, either electronic or
' printed, without the express written consent of Bentley Systems, Inc.
' ------------------------------------------------------------------------
Option Explicit

Private Sub CommandButton1_Click()
    Dim startpt As Point3d
    Dim endpt As Point3d
    
    startpt.X = txtx1.Value
    startpt.Y = txty1.Value
    startpt.Z = 0
    
    endpt.X = txtx2.Value
    endpt.Y = txty2.Value
    endpt.Z = 0
    
    Dim lineel As LineElement
    Set lineel = CreateLineElement2(Nothing, startpt, endpt)
    ActiveModelReference.AddElement lineel
    lineel.Redraw
    ShowPrompt "done"
    
End Sub

Private Sub txtx1_Change()

End Sub
