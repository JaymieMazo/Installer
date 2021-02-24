VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C06_InventoryCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim intOptionAC       As Integer
Dim strSQLInvCheck    As String

Public Sub LoadConsumptionData(ByRef hflxGrid As Object, ByVal strItemId As String _
                             , ByVal dteFrom As Date, ByVal dteUntil As Date _
                             , ByVal strDivision As String, ByVal blnAvg As Boolean)
    Dim adoConsOrdData  As Object
    
    If blnAvg Then
          strSQLInvCheck = " SELECT TransactedDate, Avg(Qty) AS Consumption " _
                         & " FROM TransactionView "
    Else
          strSQLInvCheck = " SELECT TransactedDate, Sum(Qty) AS Consumption " _
                         & " FROM TransactionView "
    End If
    strSQLInvCheck = strSQLInvCheck _
                         & " WHERE TransactedDate >= " & pfstrQt(Format(dteFrom, "yyyy/mm/dd")) _
                         & "       and TransactedDate <= " & pfstrQt(Format(dteUntil, "yyyy/mm/dd")) _
                         & "       and ItemId = " & pfstrQt(strItemId)

    Select Case intOptionAC
        Case 0
            strSQLInvCheck = strSQLInvCheck & " and TransactionTypeId = 3 And Qty > 0 "
        Case 1
            strSQLInvCheck = strSQLInvCheck & " and Qty > 0 And (TransactionTypeId=4 or TransactionTypeId=3) "
        Case 2
            strSQLInvCheck = strSQLInvCheck & " and (TransactionTypeId=4 or TransactionTypeId=3) "
    End Select
    
    If strDivision <> "" And strDivision <> "All" Then _
          strSQLInvCheck = strSQLInvCheck & " and DivisionID = " & clsPrintMenu.pfstrDivisionID(strDivision)
    
    strSQLInvCheck = strSQLInvCheck & " GROUP BY TransactedDate " _
                                    & " ORDER BY TransactedDate"
       
    Set adoConsOrdData = clsDB.GetRecordSet(strSQLInvCheck)
    
    If adoConsOrdData.EOF Then
        hflxGrid.Rows = 2
        Exit Sub
    End If
    Set hflxGrid.DataSource = adoConsOrdData
    Set adoConsOrdData = Nothing
End Sub

'--- Returns the Average Consumption on item from specified date range
Public Function AVGCons(ByVal strdteFrom As String, ByVal strdteUntil As String _
                      , ByVal strItemId As String, ByVal strDivision As String) As String
    Dim adoRSConsUntil  As Object
    
    strSQLInvCheck = " SELECT SUM(Qty) AS Consumption " _
                   & " FROM TransactionView " _
                   & " WHERE TransactedDate >= " & pfstrQt(Format(strdteFrom, "yyyy/mm/dd")) _
                   & "       and TransactedDate <= " & pfstrQt(Format(strdteUntil, "yyyy/mm/dd")) _
                   & "       and ItemId = " & pfstrQt(strItemId)

    Select Case intOptionAC
        Case 0
            strSQLInvCheck = strSQLInvCheck & " and TransactionTypeId = 3 "
        Case 1
            strSQLInvCheck = strSQLInvCheck & " and Qty > 0 And (TransactionTypeId=4 or TransactionTypeId=3) "
        Case 2
            strSQLInvCheck = strSQLInvCheck & " and (TransactionTypeId=4 or TransactionTypeId=3) "
    End Select
    If strDivision <> "" And strDivision <> "All" Then _
          strSQLInvCheck = strSQLInvCheck & " and DivisionID = " & clsPrintMenu.pfstrDivisionID(strDivision)
    
    Set adoRSConsUntil = clsDB.GetRecordSet(strSQLInvCheck)
    
    If adoRSConsUntil.EOF Then
        AVGCons = 0
        Exit Function
    End If
    
    Dim intDays As Integer
    
    intDays = CDate(strdteUntil) - CDate(strdteFrom)
    
    AVGCons = pfvarIs_Null(adoRSConsUntil.Fields("Consumption").Value, False) / intDays
    Set adoRSConsUntil = Nothing
End Function

Public Function GetWaitingDetails(ByVal strItemId As String, strDivision As String) As Object
     strSQLInvCheck = " (SELECT " _
                    & "       ItemId, Waiting, FtryDate " _
                    & " FROM " _
                    & "       PoWaitingView " _
                    & " WHERE " _
                    & "       Waiting <>0 " _
                    & "       and ItemID = " & pfstrQt(strItemId) _
                    & "       and DivisionID = " & clsPrintMenu.pfstrDivisionID(strDivision) _
                    & "       and Canceled = 0) " _
                    & " UNION " _
                    & " (SELECT " _
                    & "        ItemId, Waiting, FtryDate " _
                    & " FROM PoWaitingInvView " _
                    & " WHERE Waiting>0 " _
                    & "       and ItemId=" & pfstrQt(strItemId) _
                    & "       and DivisionID=" & clsPrintMenu.pfstrDivisionID(strDivision) _
                    & "       and Canceled=0)"
        
    Set GetWaitingDetails = clsDB.GetRecordSet(strSQLInvCheck)
End Function

Property Let OptionAC(ByVal intValue As Integer)
     intOptionAC = intValue
End Property

'--- Returns the Standard Deviation of an Item
Public Function GetStDev(ByVal strItemId As String, ByVal strdteFrom As String, _
                         ByVal strDteTo As String, ByVal strDivision As String) As Double
        Dim adoRSSafetyStock As Object
        
On Error GoTo lnError
        strSQLInvCheck = " SELECT " _
                       & "         StDev(ItemShipmentDetails.Qty) as StandardDeviation " _
                       & " FROM " _
                       & "         Items INNER JOIN " _
                       & "         ItemShipmentDetails ON " _
                       & "         Items.ItemId = ItemShipmentDetails.ItemId " _
                       & " INNER JOIN ItemShipmentHeaders ON " _
                       & "         ItemShipmentDetails.ItemShipmentId = ItemShipmentHeaders.ItemShipmentId " _
                       & " INNER JOIN Divisions ON " _
                       & "         ItemShipmentHeaders.DivisionId = Divisions.DivisionId " _
                       & " WHERE " _
                       & "         ShippedDate Between " & pfstrQt(Format(strdteFrom, "yyyy/mm/dd")) _
                       & "         and " & pfstrQt(Format(strDteTo, "yyyy/mm/dd")) _
                       & "         and Items.ItemID = " & pfstrQt(strItemId) _
                       & "         and ItemShipmentDetails.Qty > 0 " _

        If strDivision <> "" Or strDivision <> "All" Then
            strSQLInvCheck = strSQLInvCheck _
                        & "        and Divisions.DivisionID = " & clsPrintMenu.pfstrDivisionID(strDivision)
        End If

        Set adoRSSafetyStock = clsDB.GetRecordSet(strSQLInvCheck)
        
        If adoRSSafetyStock.EOF Then
            MsgBox "No Record Found", vbExclamation, pstrMessage
            Exit Function
        End If
        GetStDev = pfvarIs_Null(adoRSSafetyStock.Fields(0).Value, False)
        Set adoRSSafetyStock = Nothing
        Exit Function
lnError:
        MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Function
'--- Returns the Total Waiting Qty of an item
Public Function GetWaitingQty(ByVal strItemId As String, ByVal strDivision As String) As Double
    Dim adoRSWaitingQty As Object
    
    strSQLInvCheck = " (SELECT " _
                   & "        SUM(Waiting) AS Waiting " _
                   & " FROM " _
                   & "        PoWaitingView" _
                   & " WHERE " _
                   & "        ItemID = " & pfstrQt(strItemId) _
                   & "        and Canceled = 0 " _
                   & "        and DivisionID = " & clsPrintMenu.pfstrDivisionID(strDivision) _
                   & "         and Waiting > 0)"
                   
    Set adoRSWaitingQty = clsDB.GetRecordSet(strSQLInvCheck)
    
    If adoRSWaitingQty.EOF Then Exit Function
    GetWaitingQty = pfvarIs_Null(adoRSWaitingQty.Fields("Waiting").Value, False)
End Function
'--- Returns the LeadTime of an Item from Suppliers
Public Function GetLeadTime(ByVal strItemId As String, Optional ByVal strSupplierId As String) As String
    Dim adoRSLeadTime   As Object
    
    strSQLInvCheck = " SELECT " _
                   & "       LeadTime " _
                   & " FROM " _
                   & "       SupplierItemPrices " _
                   & " WHERE " _
                   & "       ItemID = " & pfstrQt(strItemId)
                   
    If strSupplierId <> "" Then _
               strSQLInvCheck = strSQLInvCheck & " and SupplierID = " & pfstrQt(strSupplierId)
               
    strSQLInvCheck = strSQLInvCheck & " Order by AppliedDate desc"
    
    Set adoRSLeadTime = clsDB.GetRecordSet(strSQLInvCheck)
    
    If adoRSLeadTime.EOF Then
        GetLeadTime = "0"
        Exit Function
    End If
    GetLeadTime = pfvarIs_Null(adoRSLeadTime.Fields("LeadTime").Value, False)
    Set adoRSLeadTime = Nothing
End Function

Public Function GetIQCQty(ByVal strItemId As String, ByVal strDivision As String)
    Dim objIQCQty As Object
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " SUM(POInvoiceSearchViewNull.QtyReceived) as TQty "
    strSQL = strSQL & " FROM POInvoiceSearchViewNull "
    strSQL = strSQL & " INNER JOIN Divisions "
    strSQL = strSQL & " ON POInvoiceSearchViewNull.DivisionId=Divisions.DivisionId "
    strSQL = strSQL & " WHERE POInvoiceSearchViewNull.ItemId='" & strItemId & "' "
    strSQL = strSQL & " AND Divisions.DivisionId='" & clsPrintMenu.pfstrDivisionID(strDivision) & "'"
    
    Set objIQCQty = clsDB.GetRecordSet(strSQL)
    
        If Not IsNull(objIQCQty.Fields("TQty").Value) Then
            GetIQCQty = objIQCQty.Fields("TQty").Value
         Else
              GetIQCQty = 0
         End If
         Set objIQCQty = Nothing

End Function

Public Function GetQtyExpected(ByVal strItemId As String, ByVal strDivision As String)
    Dim objQtyExpected As Object
    Dim strSQLQty As String
    
    strSQLQty = ""
    strSQLQty = strSQLQty & " SELECT "
    strSQLQty = strSQLQty & " SUM(Qty) as QtyExpected"
    strSQLQty = strSQLQty & " FROM InvoiceDetailsView"
    strSQLQty = strSQLQty & " WHERE ItemId='" & strItemId & "' AND"
    strSQLQty = strSQLQty & " DivisionId='" & clsPrintMenu.pfstrDivisionID(strDivision) & "' AND"
    strSQLQty = strSQLQty & " QtyReceived is null AND QtyOk is Null"
    
    Set objQtyExpected = clsDB.GetRecordSet(strSQLQty)
    
    If Not IsNull(objQtyExpected.Fields("QtyExpected").Value) Then
        GetQtyExpected = objQtyExpected.Fields("QtyExpected").Value
    Else
        GetQtyExpected = 0
    End If
    
    Set objQtyExpected = Nothing
    
End Function

Public Function GetExcessWaitingQty(ByVal strItemId As String, ByVal strDivision As String)
    Dim objExcessWaiting As Recordset
    Dim strSQLExcess As String

    strSQLExcess = ""
    strSQLExcess = strSQLExcess & " SELECT "
    strSQLExcess = strSQLExcess & "          SUM(Waiting) as ExcessWaiting "
    strSQLExcess = strSQLExcess & " FROM PoWaitingInvView "
    strSQLExcess = strSQLExcess & " WHERE ItemId='" & strItemId & "' "
    strSQLExcess = strSQLExcess & " AND DivisionId='" & clsPrintMenu.pfstrDivisionID(strDivision) & "'"
    strSQLExcess = strSQLExcess & " AND Waiting > 0"

    Set objExcessWaiting = clsDB.GetRecordSet(strSQLExcess)

    If Not IsNull(objExcessWaiting.Fields("ExcessWaiting").Value) Then
        GetExcessWaitingQty = objExcessWaiting.Fields("ExcessWaiting").Value
    Else
        GetExcessWaitingQty = 0
    End If

    Set objExcessWaiting = Nothing

End Function
