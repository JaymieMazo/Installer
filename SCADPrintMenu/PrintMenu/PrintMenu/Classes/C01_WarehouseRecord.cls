VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C01_WarehouseRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQLWHRecord As String

Public Function GetWHRecord(ByVal dteFrom As Date, ByVal dteUntil As Date, _
                             ByVal strDivision As String, ParamArray strItemId()) As Object
    Dim varItemId As Variant
    
    varItemId = Split(strItemId(0), "-")
    
    strSQLWHRecord = _
                 " SELECT " _
                 & "     ItemId, TransactedDate, Description, Sum(Qty) as Qty, TransactionTypeId, DivisionId " _
               & " FROM  InventoryRecordView " _
               & " WHERE TransactedDate >= " & pfstrQt(Format(dteFrom, "yyyy/mm/dd")) _
               & "       and TransactedDate <= " & pfstrQt(Format(dteUntil, "yyyy/mm/dd")) _
               & "       and DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision)
     '---load by item id
     If strItemId(0) <> "" Then _
                strSQLWHRecord = strSQLWHRecord _
                    & " and ItemId between " & pfstrQt(varItemId(0) & "%") _
                                            & " And " & pfstrQt(strItemId(1) & "%")
                    
     strSQLWHRecord = strSQLWHRecord _
               & " GROUP BY TransactedDate, ItemId, Description, TransactionTypeId, DivisionId " _
               & " ORDER BY ItemId, TransactedDate"
    
    Set GetWHRecord = clsDB.GetRecordSet(strSQLWHRecord)
End Function

Public Sub SaveToWHRecordWT(ByVal strFields As String, ByVal lngSeqNo As Long, ParamArray arrValues())
    
     clsDB.DB_PrintMenu.BeginTrans
     strSQLWHRecord = "INSERT INTO " _
                   & " PrintMenu..WH_Record " & strFields _
                   & " VALUES (" & lngSeqNo & "," & pfstrQt(arrValues(0)) _
                         & "," & pfstrQt(arrValues(1)) _
                         & "," & pfstrQt(arrValues(2)) _
                         & "," & pfvarIs_Null(arrValues(3), False) _
                         & "," & pfstrQt(fstrGetDivName(arrValues(4))) & ")"
     Call clsDB.DB_PrintMenu.Execute(strSQLWHRecord)
     clsDB.DB_PrintMenu.CommitTrans
End Sub

'--- Retuns the Division Name given the DivisionID
Private Function fstrGetDivName(ByVal strDivID As String) As String
    Dim adoGetDivName As Object, strSQLQuery As String
    
    strSQLQuery = "SELECT " _
               & "       Description " _
               & " FROM Divisions " _
               & " WHERE " _
               & "       DivisionID = " & pfstrQt(strDivID)
    Set adoGetDivName = clsDB.GetRecordSet(strSQLQuery)
    
    If adoGetDivName.EOF Then
        Exit Function
    End If
    fstrGetDivName = adoGetDivName.Fields(0).Value
End Function

Public Sub DeleteWHRecordWT()
     clsDB.DB_PrintMenu.BeginTrans
     Call clsDB.DB_PrintMenu.Execute("DELETE FROM PrintMenu..WH_Record")
     clsDB.DB_PrintMenu.CommitTrans
End Sub

Public Function GetWHRecordsInWT() As Object
     
     strSQLWHRecord = "SELECT " _
                  & "       ItemID, TransDate, Description, Division, " _
                  & "       SUM(InQty) as InQty, SUM(OutQty) as OutQty, SUM(AC) as ACQty " _
                  & " FROM  PrintMenu..WH_RECORD " _
                  & " GROUP BY " _
                  & "       TransDate, ItemID, Description, Division Order by ItemId, TransDate"
    
    Set GetWHRecordsInWT = clsDB.GetRecordSet(strSQLWHRecord, False)
End Function

Public Function GetACWHRecordsInWT() As Object

      strSQLWHRecord = "SELECT " _
                    & "       TransDate, ItemID, Description, Division, " _
                    & "       SUM(InQty) as InQty, SUM(OutQty) as OutQty, sum(Balance) as Balance " _
                    & " FROM  WH_RECORD " _
                    & " GROUP BY " _
                    & "       TransDate, ItemID, Description, Division Order by ItemId, TransDate"
    
    Set GetACWHRecordsInWT = clsDB.GetRecordSet(strSQLWHRecord, False)
End Function

Public Function GetTotalWHRecordsInWT() As Object

      strSQLWHRecord = "SELECT " _
                    & "       ItemID, Description, Division, " _
                    & "       SUM(InQty) as InQty, SUM(OutQty) as OutQty, BALANCE " _
                    & " FROM  WH_RECORD " _
                    & " GROUP BY " _
                    & "       ItemID, Description, Division, Balance"
     Set GetTotalWHRecordsInWT = clsDB.GetRecordSet(strSQLWHRecord, False)
End Function

Public Sub GetConsumptionPerDivision(ByRef hflxGrid As Object, ByVal dteFrom As Date, ByVal dteUntil As Date, _
                                     ByVal strDivision As String, ParamArray strItemId())
    Dim objDivisionConsumption As Object
    Dim bytCol                 As Byte
    
    strSQLWHRecord = _
                  " SELECT upper(t1.ItemId), t1.Description, t1.DivisionName, Sum(t1.Out) as Out " _
                & " FROM RequestConsumptionView t1 Inner Join DivisionItems t2 On " _
                        & " t1.ItemId = t2.ItemId " _
                & " WHERE   t1.ShippedDate >= " & pfstrQt(Format(dteFrom, "yyyy/mm/dd")) _
                & "      and t1.ShippedDate <= " & pfstrQt(Format(dteUntil, "yyyy/mm/dd")) _
                & "      and t2.DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision)
              
    If Not strItemId(0) = "" Then _
        strSQLWHRecord = strSQLWHRecord _
                & "      and t1.ItemId Between " & pfstrQt(strItemId(0)) _
                & "      and " & pfstrQt(strItemId(1))
                 
    strSQLWHRecord = strSQLWHRecord _
                & " GROUP BY t1.ItemId, t1.Description, t1.DivisionName " _
                & " ORDER BY t1.ItemId"

    Set objDivisionConsumption = clsDB.GetRecordSet(strSQLWHRecord)
    
    Set hflxGrid.DataSource = objDivisionConsumption
    
    With hflxGrid
        For bytCol = 0 To 3
            .TextMatrix(0, bytCol) = Choose(bytCol + 1, "ITEM ID", "DESCRIPTION", "REQUESTER", "CONSUMPTION")
            .ColWidth(bytCol) = Choose(bytCol + 1, 1400, 5500, 2000, 2000)
            .Row = 0: .Col = bytCol
            .CellFontBold = True
            .CellAlignment = 4
        Next bytCol
    End With
End Sub

Public Sub GetConsumptionDetails(ByRef hflxGrid As Object, ByVal dteFrom As Date, ByVal dteUntil As Date, _
                                 ByVal strDivision As String, ParamArray strItemId())
     Dim objConsumption As Object
     Dim bytCol As Byte
     
     strSQLWHRecord = _
                    " Select RequestConsumptionDetailsView.* " _
                  & " From RequestConsumptionDetailsView Inner Join" _
                  & "       DivisionItems On RequestConsumptionDetailsView.ItemId = DivisionItems.ItemId" _
                  & " Where TransDate >= " & pfstrQt(Format(dteFrom, "yyyy/mm/dd")) _
                  & "              And TransDate <= " & pfstrQt(Format(dteUntil, "yyyy/mm/dd")) _
                  & "       And DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision)
                  
     If Not strItemId(0) = "" Then
          strSQLWHRecord = strSQLWHRecord _
                  & "      and RequestConsumptionDetailsView.ItemId Between " & pfstrQt(strItemId(0)) _
                  & "      and " & pfstrQt(strItemId(1))
     End If
    strSQLWHRecord = strSQLWHRecord _
                  & " Order By RequestConsumptionDetailsView.TransDate, RequestConsumptionDetailsView.TransId, " _
                  & "              RequestConsumptionDetailsView.ItemId"

    Set objConsumption = clsDB.GetRecordSet(strSQLWHRecord)
    Set hflxGrid.DataSource = objConsumption
    
    With hflxGrid
          For bytCol = 0 To .Cols - 1
               .Row = 0: .Col = bytCol
               .ColWidth(bytCol) = Choose(bytCol + 1, 1300, 1000, 900, 2000, 1000, 2000, 2000, 1500, 2000, 1000)
               .CellFontBold = True
               .CellAlignment = 4
          Next
    End With
End Sub

