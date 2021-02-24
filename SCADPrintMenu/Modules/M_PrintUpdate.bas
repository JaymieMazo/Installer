Attribute VB_Name = "M_PrintUpdate"
Option Explicit

Dim clsForm         As Object
Dim clsPrintUpdate  As C_PrintUpdate

Public Sub Main()
     
On Error GoTo lnCancel
     Set clsForm = CreateObject("PG_DLL.FormTools")
     Call clsForm.ShowMessage("Downloading updated program")
     Set clsPrintUpdate = New C_PrintUpdate
     Call clsPrintUpdate.psubUpdateExe
     Call clsForm.HideMessage
     Exit Sub
lnCancel:
     MsgBox Err.Description, vbCritical, App.EXEName
End Sub
