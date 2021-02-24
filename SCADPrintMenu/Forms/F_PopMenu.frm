VERSION 5.00
Begin VB.Form F_PopMenu 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyToExcel 
         Caption         =   "Copy to excel"
      End
   End
   Begin VB.Menu mnuCopyPaste 
      Caption         =   "CopyPaste"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
   End
   Begin VB.Menu mnuInvoiceSearch 
      Caption         =   "Invoice Search"
      Visible         =   0   'False
      Begin VB.Menu mnuExcel 
         Caption         =   "Copy to Excel"
      End
   End
End
Attribute VB_Name = "F_PopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuCopy_Click()
   Clipboard.Clear
   Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub mnuCut_Click()
   Clipboard.Clear
   Clipboard.SetText Screen.ActiveControl.SelText
   Screen.ActiveControl.SelText = ""
End Sub

Private Sub mnuExcel_Click()
    Call connecttoserver
   Dim intRow As Integer, intCol As Integer
   
   Screen.MousePointer = vbHourglass
   Call psubShowStatMsg("Please wait... copying to excel.")
   Call clsPrintMenu.Utility.OpenExcel
   With F_InvoiceSearch.hflxPOSearch
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                  clsPrintMenu.Utility.ExcelWkSheet.Cells(intRow + 1, intCol + 1) = "'" & .TextMatrix(intRow, intCol)
            Next intCol
        Next intRow
   End With
   Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, 18)
   clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
   clsPrintMenu.Utility.ExcelApp.Visible = True
   Call psubHideStatMsg
   Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub mnuPaste_Click()
   Screen.ActiveControl.SelText = Clipboard.GetText()
End Sub

Private Sub mnuCopyToExcel_Click()
    Call connecttoserver
   Dim strSelect As String
   Dim intLoop As Integer
   Dim intRow As Integer
   
   Screen.MousePointer = vbHourglass
   Call psubShowStatMsg("Please wait.... copying to excel.")
   Call clsPrintMenu.Utility.OpenExcel
   With F_InvRecord.hflxInvRecord
        For intLoop = 0 To .Cols - 1
               clsPrintMenu.Utility.ExcelWkSheet.Cells(intRow + 1, intLoop + 1) = .TextMatrix(0, intLoop)
        Next
        For intRow = 1 To .Rows - 1
               For intLoop = 0 To .Cols - 1
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(intRow + 1, intLoop + 1) = .TextMatrix(intRow, intLoop)
               Next
        Next
   End With
   Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, 8)
   clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
   clsPrintMenu.Utility.ExcelApp.Visible = True
   Call psubHideStatMsg
   Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub


