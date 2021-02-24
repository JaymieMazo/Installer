VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C02_WaitingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetWaitingList(ByVal strItemId As String, ByVal strDivision As String _
                              , ByVal strPONo As String, ByVal strSupplierId As String, ByVal strSupplier _
                              , ByVal dtpFrom As Object, ByVal dtpUntil As Object) As Object
    Dim strSQLWaiting   As String
     
    
    strSQLWaiting = " (SELECT " _
                  & "       PoNo, PoDetailSeq, ItemId, Description, OrderQty, " _
                  & "       QtyOK,QtyExpected, Waiting, IssuedDate, FtryDate, " _
                  & "       Division, SupplierName " _
                  & " FROM PoWaitingView " _
                  & " WHERE " _
                  & "       ItemID Like " & pfstrQt(strItemId & "%") _
                  & "       and Canceled = 0 And Waiting <> 0"
                  
    If strDivision <> "All" Then _
          strSQLWaiting = strSQLWaiting _
                  & "       and DivisionID = " & pfstrQt(clsPrintMenu.pfstrDivisionID(strDivision))
    If Trim$(strPONo) <> "" Then _
          strSQLWaiting = strSQLWaiting & " and PONo = " & pfstrQt(strPONo)
    If strSupplierId <> "" Then
          strSQLWaiting = strSQLWaiting & " and SupplierID = " & pfstrQt(strSupplierId)
    ElseIf strSupplier <> "" Then
          strSQLWaiting = strSQLWaiting & " and SupplierName Like " & pfstrQt(strSupplier & "%")
    End If
                        
    If IsNull(dtpFrom.Value) And Not IsNull(dtpUntil.Value) Then
          strSQLWaiting = strSQLWaiting & " and IssuedDate <= " & pfstrQt(Format(dtpUntil.Value, "yyyy/mm/dd"))
    ElseIf IsNull(dtpUntil.Value) And Not IsNull(dtpFrom.Value) Then
          strSQLWaiting = strSQLWaiting & " and IssuedDate >= " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd"))
    ElseIf Not IsNull(dtpFrom.Value) And Not IsNull(dtpUntil.Value) Then
          strSQLWaiting = strSQLWaiting & " and IssuedDate > = " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd")) _
               & " and IssuedDate < " & pfstrQt(DateAdd("d", 1, Format(dtpUntil.Value, "yyyy/mm/dd")))
    End If
    
    strSQLWaiting = strSQLWaiting & ")"
    
    strSQLWaiting = strSQLWaiting & " UNION " _
                  & " (SELECT PoNo,PoDetailSeq, ItemId,Description, OrderQty, " _
                  & "         QtyOk,QtyExpected,Waiting, IssuedDate, FtryDate, Division, SupplierName " _
                  & " FROM PoWaitingInvView " _
                  & " WHERE ItemId Like " & pfstrQt(strItemId & "%") _
                  & "       And Canceled=0 And Waiting > 0"

    If strDivision <> "All" Then _
          strSQLWaiting = strSQLWaiting _
                  & "       and DivisionID = " & pfstrQt(clsPrintMenu.pfstrDivisionID(strDivision))
    If Trim$(strPONo) <> "" Then _
          strSQLWaiting = strSQLWaiting & " and PONo = " & pfstrQt(strPONo)
    If strSupplierId <> "" Then
          strSQLWaiting = strSQLWaiting & " and SupplierID = " & pfstrQt(strSupplierId)
    ElseIf strSupplier <> "" Then
          strSQLWaiting = strSQLWaiting & " and SupplierName Like " & pfstrQt(strSupplier & "%")
    End If
                        
    If IsNull(dtpFrom.Value) And Not IsNull(dtpUntil.Value) Then
          strSQLWaiting = strSQLWaiting & " and IssuedDate <= " & pfstrQt(Format(dtpUntil.Value, "yyyy/mm/dd"))
    ElseIf IsNull(dtpUntil.Value) And Not IsNull(dtpFrom.Value) Then
          strSQLWaiting = strSQLWaiting & " and IssuedDate >= " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd"))
    ElseIf Not IsNull(dtpFrom.Value) And Not IsNull(dtpUntil.Value) Then
          strSQLWaiting = strSQLWaiting & " and IssuedDate > = " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd")) _
               & " and IssuedDate < " & pfstrQt(DateAdd("d", 1, Format(dtpUntil.Value, "yyyy/mm/dd")))
    End If
    
    strSQLWaiting = strSQLWaiting & ")"
      
    strSQLWaiting = strSQLWaiting & " ORDER BY PONo, PODetailSeq"
    Set GetWaitingList = clsDB.GetRecordSet(strSQLWaiting)
    
 End Function

Public Sub SavePrintWaitingList(ParamArray varValues())
     Dim strSQLQuery As String
     Call clsDB.DB_PrintMenu.Execute("BEGIN TRANSACTION")
     strSQLQuery = "INSERT INTO " _
               & "       PrintMenu..WaitingList " _
               & "            (SeqNo,PoNo, ItemID,Description,DivCode,OrderDate,ETA,OrderQty,QtyExpected,QtyReceived,QtyWaiting,SupplierName) " _
               & " VALUES (" & varValues(0) & "," & pfstrQt(pfvarIs_Null(varValues(1))) _
               & "       ," & pfstrQt(pfvarIs_Null(varValues(2))) & "," & pfstrQt(pfvarIs_Null(varValues(3))) _
               & "       ," & pfstrQt(pfvarIs_Null(varValues(4))) & "," & pfstrQt(pfvarIs_Null(Format(varValues(5), "yyyy/mm/dd"))) _
               & "       ," & pfstrQt(pfvarIs_Null(Format(varValues(6), "yyyy/mm/dd"))) & "," & pfvarIs_Null(varValues(7), False) _
               & "       ," & pfvarIs_Null(varValues(8), False) & "," & pfvarIs_Null(varValues(9), False) _
               & "       ," & pfstrQt(pfvarIs_Null(varValues(10))) & "," & pfstrQt(pfvarIs_Null(varValues(11))) & ")"

     Call clsDB.DB_PrintMenu.Execute(strSQLQuery)
     Call clsDB.DB_PrintMenu.Execute("COMMIT TRANSACTION")
End Sub

Public Sub DeleteWaiting()
     Call clsDB.DB_PrintMenu.Execute("BEGIN TRANSACTION")
     Call clsDB.DB_PrintMenu.Execute("DELETE from PrintMenu..WaitingList")
     Call clsDB.DB_PrintMenu.Execute("COMMIT TRANSACTION")
End Sub

Public Function SQLWaitingList() As String
     SQLWaitingList = " SELECT SeqNo, PONo, ItemID, Description, DivCode,OrderQty,OrderDate, ETA, QtyExpected, " _
                    & "        QtyReceived, QtyWaiting, SupplierName FROM WaitingList " _
                    & " ORDER BY PONo, SeqNo"
End Function

'--- Returns the Supplier Name given the SupplierID
Public Function GetSupplierName(ByVal strSupplierId As String) As String
    Dim adoRSSupplierName As Object
    Dim strSQLQuery       As String
    
    strSQLQuery = " SELECT SupplierName FROM Suppliers " _
                & " WHERE " _
                & "      SupplierID = " & pfstrQt(strSupplierId)
    Set adoRSSupplierName = clsDB.GetRecordSet(strSQLQuery)
    
    If adoRSSupplierName.EOF Then
        MsgBox "Invalid Supplier ID!", vbExclamation, pstrMessage
        Exit Function
    End If
    GetSupplierName = pfvarIs_Null(adoRSSupplierName.Fields("SupplierName").Value)
    Set adoRSSupplierName = Nothing
End Function

'--- Loads All SupplierID in combobox
Public Sub LoadSupplierID(ByVal cboSupplier As Object)
    Dim adoRSSupplierID  As Object
    Dim strSQLQuery      As String
    
On Error GoTo lnError
    strSQLQuery = "SELECT SupplierID FROM Suppliers ORDER BY SupplierID"
    Set adoRSSupplierID = clsDB.GetRecordSet(strSQLQuery)
    
    With adoRSSupplierID
        cboSupplier.Clear
        cboSupplier.AddItem ""
        Do Until .EOF
            cboSupplier.AddItem pfvarIs_Null(.Fields(0).Value)
            .MoveNext
        Loop
    End With
    Set adoRSSupplierID = Nothing
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub


