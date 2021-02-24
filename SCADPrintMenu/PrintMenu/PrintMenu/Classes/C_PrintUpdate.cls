VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_PrintUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim SourceDT As Date, _
     DestinationDT As Date

Dim SeeMyValue As String
Dim Choice As String

Const MainProg = "\PrintMenu.exe"
Const ServerProg = "\\A-SV14\TOOLS\STEP\The Exe\PrintMenu.exe"

Public Sub psubUpdateExe()
On Error GoTo CancelMe
     If Not FileExists(App.Path & MainProg) Then MsgBox "Cannot find " & MainProg, _
          vbCritical, "Error": End: Exit Sub

     If Not FileExists(ServerProg) Then MsgBox "Cannot find " & ServerProg _
          , vbCritical, "Error": GoTo Continue

          SourceDT = FileDateTime(ServerProg)
          DestinationDT = FileDateTime(App.Path & MainProg)

          If SourceDT > DestinationDT Then
               FileCopy ServerProg, App.Path & MainProg
          End If

Continue:
               Shell App.Path & MainProg, vbNormalFocus

               End
               Exit Sub
CancelMe:

    MsgBox "Access to A-SV14 Denied! Unable to Update PrintMenu program", vbCritical
'    GoTo Continue
End Sub

Function FileExists(filename As String) As Boolean
     FileExists = (Dir$(filename) <> "")
End Function

