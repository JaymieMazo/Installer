VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C04_POData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetPOData(ByVal dtpFrom As Object, ByVal dtpTo As Object) As Object
     Dim strSQLPOData As String
     
     strSQLPOData = _
            " SELECT PoNo, ItemId, Description, Qty, Division, AcknowledgerId, RequisitionSlipNo, Remarks, " _
          & "     convert(datetime,OrderDate,101) OrderDate, EtdDate, EtaDate, FtryDate, ReceivedAllInvoices, " _
          & "     Canceled, FinalDestinationId, PoDetailSeq, UnitPrice, JapaneseDescription, QtyUnit, " _
          & "     EquivalentQty, EquivalentQtyUnit, TotalPrice, InventoryItemType, FinalDestination, IsFixed, Supplier " _
          & " FROM PODataView " _
          & " WHERE "
                 
    If IsNull(dtpFrom.Value) Then
        strSQLPOData = strSQLPOData & " IssuedDate <= " & pfstrQt(Format(dtpTo.Value, "yyyy/mm/dd"))
    ElseIf IsNull(dtpTo.Value) Then
          strSQLPOData = strSQLPOData & " IssuedDate >= " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd"))
    Else
          strSQLPOData = strSQLPOData & " IssuedDate >= " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd")) _
           & " And IssuedDate <= " & pfstrQt(DateAdd("d", 1, (Format(dtpTo.Value, "yyyy/mm/dd"))))
    End If
    Set GetPOData = clsDB.GetRecordSet(strSQLPOData)
End Function

Public Sub ExportToExcel(ByVal hflexGrid As Object)
    Dim bytCol      As Byte, _
        lngRow      As Long
    
On Error GoTo lnError
     '---open excel application
    Call clsPrintMenu.Utility.OpenExcel
    With hflexGrid
        For lngRow = 0 To .Rows - 1
            For bytCol = 0 To .Cols - 1
                '--- set the cell as text first before putting the records
                If bytCol = 8 Then
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1) = Format(.TextMatrix(lngRow, bytCol), "yyyy/mm/dd")
                Else
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1) = .TextMatrix(lngRow, bytCol)
                End If
                If lngRow = 0 Then clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1).Font.Bold = True
            Next
        Next
        clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
        Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, .Cols, 33)
        Call subFormatExcel(hflexGrid)
        clsPrintMenu.Utility.ExcelApp.Visible = True
        Call clsPrintMenu.Utility.CloseExcel
    End With
    
    GoTo lnCleanUp
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
lnCleanUp:
     Call clsPrintMenu.Utility.CloseExcel
End Sub

Private Sub subFormatExcel(ByVal hflexGrid As Object)
Dim strLenght   As String

    '--- This sets the range for puting the grid lines
        strLenght = "A1:Z"
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows))
        .VerticalAlignment = xlCenter
    End With
    
    clsPrintMenu.Utility.ExcelWkSheet.Cells.Borders(xlDiagonalDown).LineStyle = xlNone
    clsPrintMenu.Utility.ExcelWkSheet.Cells.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLenght & CStr(hflexGrid.Rows)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub
