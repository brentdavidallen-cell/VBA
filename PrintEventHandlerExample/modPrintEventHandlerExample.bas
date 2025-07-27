Attribute VB_Name = "modPrintEventHandlerExample"
Option Explicit

Dim printEventHandler As clsPrintEventHandlerExample

Sub OnProjectLoad()
    'Registers event handler when loaded in interactive session.
    Set printEventHandler = New clsPrintEventHandlerExample
    Application.PrintManager.AddPrintEventsHandler printEventHandler
End Sub

Sub OnProjectLoadNonGraphics()
    'Registers event handler when loaded in non-graphics session,
    'i.e. the Print Organizer worker process.
    Set printEventHandler = New clsPrintEventHandlerExample
    Application.PrintManager.AddPrintEventsHandler printEventHandler
End Sub
