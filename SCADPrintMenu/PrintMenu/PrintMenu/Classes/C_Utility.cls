VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objXLApp             As Object  ' Excel.Application
Dim objXLWkBook          As Object  ' Excel.Workbook
Dim objXLWkSheet         As Object  ' Excel.Worksheet

'---Open Excel Application
Public Sub OpenExcel()
    Set objXLApp = CreateObject("Excel.Application")
    Set objXLWkBook = objXLApp.Workbooks.Add
    Set objXLWkSheet = objXLWkBook.Worksheets.Item(1)
End Sub

Property Get ExcelApp() As Excel.Application
     Set ExcelApp = objXLApp
End Property

Property Get ExcelWkSheet() As Excel.Worksheet
     Set ExcelWkSheet = objXLWkSheet
End Property

Property Get ExcelWkBook() As Excel.Workbook
    Set ExcelWkBook = objXLWkBook
End Property

'---Close Excel Application
Public Sub CloseExcel()
    Set objXLApp = Nothing
    Set objXLWkBook = Nothing
    Set objXLWkSheet = Nothing
End Sub
'--- put highlight to selected cell
Public Sub SetCellColor(intFRow As Integer, intFCol As Integer _
            , intLRow As Integer, intLCol As Integer, Optional intColor As Integer = 36)
    With objXLApp
        .Range(.Cells(intFRow, intFCol), _
        .Cells(intLRow, intLCol)).Interior.ColorIndex = intColor 'Default is Yellow
        .Range(.Cells(intFRow, intFCol), _
        .Cells(intLRow, intLCol)).Interior.Pattern = xlSolid
    End With
End Sub

