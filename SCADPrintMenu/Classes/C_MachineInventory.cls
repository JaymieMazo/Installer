VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C10_MachineInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Dim strSQLMachineInvRecord As String
 
 Function GetMachineInventoryRecord(ByVal strItemId As String, ByVal intMachineId As Integer, _
                                   ByVal strMachineName As String, ByVal dteFrom As Date, _
                                   ByVal dteUntil As Date) As Object
                                   
      strSQLMachineInvRecord = " SELECT " _
                      & "       TransactedDate, TransId, ItemId, Qty, TransactionTypeId, Description, Remarks,MachineID,MachineName " _
                      & " FROM  MachineInventoryWHSRecord "
                      strSQLMachineInvRecord = strSQLMachineInvRecord & "WHERE "
                      If strItemId = "" Or strItemId = Null Then
                      Else
                        strSQLMachineInvRecord = strSQLMachineInvRecord & " ItemId = " & pfstrQt(strItemId)
                      End If
                      
                      If intMachineId = Null Then
                      'Add if no item id
                      ElseIf strItemId = "" Or strItemId = Null Then
                        strSQLMachineInvRecord = strSQLMachineInvRecord & " MachineID = " & intMachineId
                      Else
                        strSQLMachineInvRecord = strSQLMachineInvRecord & " AND MachineID = " & intMachineId
                      End If
                      
                      If strMachineName = Null Or strMachineName = "" Then
                      
                      ElseIf intMachineId = Null And (strItemId = "" Or strItemId = Null) Then
                        strSQLMachineInvRecord = strSQLMachineInvRecord & " MachineName like " & pfstrQt("%" & strMachineName & "%")
                      Else
                        strSQLMachineInvRecord = strSQLMachineInvRecord & " AND MachineName like " & pfstrQt("%" & strMachineName & "%")
                      End If
                        strSQLMachineInvRecord = strSQLMachineInvRecord & "       and TransactedDate >= " _
                        & pfstrQt(Format(dteFrom, "yyyy/mm/dd"))
                        strSQLMachineInvRecord = strSQLMachineInvRecord & "       and TransactedDate <= " _
                        & pfstrQt(Format(dteUntil, "yyyy/mm/dd"))
                        strSQLMachineInvRecord = strSQLMachineInvRecord & " ORDER BY TransactedDate asc, TransactionTypeId desc"
     Set GetMachineInventoryRecord = clsDB.GetRecordSet(strSQLMachineInvRecord, False)
     
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
                If bytCol = 1 Then
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
        strLenght = "A1:G"
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
