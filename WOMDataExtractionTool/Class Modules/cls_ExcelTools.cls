VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ExcelTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===C_ExcelTools===============================================================
'   Developed by SCAD MIS
'==============================================================================

Option Explicit

Public xlApp As Object   ' Excel.Application
Public xlBook As Object  ' Excel.Workbooks
Public xlSheet As Object ' Excel.Worksheet

Public Function fblnOpenTemplate(ByVal strFileName As String, _
                        Optional ByVal strSheet As String, _
                        Optional ByVal blnBySheet As Boolean = False, _
                        Optional ByVal strPassword As String = "") As Boolean
On Error GoTo lnError

        Dim strNewFile As String
cont:
    strSheet = Replace(strSheet, "/", "_")
    strFileName = Replace(strFileName, "/", "_") '20110303

    If Len(Dir$(strFileName)) = 0 Then
        MsgBox "Template for " & strSheet & " was not found!", vbExclamation
        fblnOpenTemplate = False
    Else
    
    strNewFile = Environ$("WINDIR") & "\Temp\" & "ForkLift_BreakDown_Summary"
    Call FileCopy(strFileName, strNewFile)
    
        Set xlApp = New Excel.Application  'CreateObject("Excel.Application")
        If strPassword = "" Then
          Set xlBook = xlApp.Workbooks.Open(strNewFile)
        Else
          Set xlBook = xlApp.Workbooks.Open(strFileName, , True, , strPassword, , True)
        End If
        If blnBySheet Then
            Set xlSheet = xlBook.Worksheets(strSheet)
        Else
            Set xlSheet = xlBook.Worksheets(1)
        End If
        fblnOpenTemplate = True
      xlApp.Visible = True
    End If
    Exit Function
lnError:
        If Err.Description = "Permission denied" Then
                MsgBox "Kindly close first the extracted forklift summary excel file."
                GoTo cont
        End If
    fblnOpenTemplate = False
    MsgBox "Error in Printing " & strSheet & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "System Error"
End Function

Private Sub Class_Terminate()
On Error Resume Next
    Set xlSheet = Nothing

'    xlBook.Close False
'    xlApp.Quit

    Set xlBook = Nothing
    Set xlApp = Nothing
    
End Sub


