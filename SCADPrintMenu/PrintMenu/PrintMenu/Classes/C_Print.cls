VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_PrintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objWHRecord          As C01_WarehouseRecord
Dim objWaitingList       As C02_WaitingList
Dim objInvRecord         As C03_InventoryRecord
Dim objPOData            As C04_POData
Dim objMaterialList      As C05_MaterialList
Dim objInventoryCheck    As C06_InventoryCheck
Dim objInvoiceSearch     As C07_InvoiceSearch
Dim objMatConsumption    As C09_MaterialConsumption
Dim objUtility           As C_Utility
Dim strSQLQuery          As String
Dim strSupCategory As String

Property Get WHRecord() As C01_WarehouseRecord
     If objWHRecord Is Nothing Then Set objWHRecord = New C01_WarehouseRecord
     Set WHRecord = objWHRecord
End Property

Property Get WaitingList() As C02_WaitingList
     If objWaitingList Is Nothing Then Set objWaitingList = New C02_WaitingList
     Set WaitingList = objWaitingList
End Property

Property Get InvRecord() As C03_InventoryRecord
     If objInvRecord Is Nothing Then Set objInvRecord = New C03_InventoryRecord
     Set InvRecord = objInvRecord
End Property

Property Get POData() As C04_POData
     If objPOData Is Nothing Then Set objPOData = New C04_POData
     Set POData = objPOData
End Property

Property Get MaterialList() As C05_MaterialList
     If objMaterialList Is Nothing Then Set objMaterialList = New C05_MaterialList
     Set MaterialList = objMaterialList
End Property

Property Get InventoryCheck() As C06_InventoryCheck
     If objInventoryCheck Is Nothing Then Set objInventoryCheck = New C06_InventoryCheck
     Set InventoryCheck = objInventoryCheck
End Property

Property Get InvoiceSearch() As C07_InvoiceSearch
     If objInvoiceSearch Is Nothing Then Set objInvoiceSearch = New C07_InvoiceSearch
     Set InvoiceSearch = objInvoiceSearch
End Property

Property Get Utility() As C_Utility
     If objUtility Is Nothing Then Set objUtility = New C_Utility
     Set Utility = objUtility
End Property

Property Get MatConsumption() As C09_MaterialConsumption
    If objMatConsumption Is Nothing Then Set objMatConsumption = New C09_MaterialConsumption
    Set MatConsumption = objMatConsumption
End Property

'--- Loads all the Division Name in a ComboBox
Public Sub psubLoadDivision(ByVal cboCombo As MSForms.ComboBox, Optional ByVal blnDefaultAll As Boolean = False)
    Dim objLoadDivision As Object
    
On Error GoTo lnError
    strSQLQuery = "SELECT DivisionId, Description FROM Divisions"
    Set objLoadDivision = clsDB.GetRecordSet(strSQLQuery)
    
    If objLoadDivision.EOF Then Exit Sub
    
    With objLoadDivision
        cboCombo.Clear
        If blnDefaultAll Then cboCombo.AddItem "All"
        Do Until .EOF
            cboCombo.AddItem .Fields("Description").Value
            .MoveNext
        Loop
    End With
    cboCombo.ListIndex = 0
    Set objLoadDivision = Nothing
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub

Public Function pfstrDivisionID(ByVal strDivName As String) As String
    Dim objDivision As Object

    strSQLQuery = " SELECT DivisionID FROM DIVISIONS " _
                & " WHERE " _
                & "      Description = " & pfstrQt(strDivName)
    Set objDivision = clsDB.GetRecordSet(strSQLQuery)

    If objDivision.EOF Then Exit Function

    pfstrDivisionID = objDivision.Fields(0).Value
    Set objDivision = Nothing
End Function

Public Function pfstrGetItemDescription(ByVal strItemId As String) As String
    Dim objDescription As Object
    
    strSQLQuery = " SELECT " _
                & "      Description " _
                & " FROM Items " _
                & " WHERE ItemID = " & pfstrQt(strItemId)
    Set objDescription = clsDB.GetRecordSet(strSQLQuery, True)
    
    If objDescription.EOF Then Exit Function
    
    pfstrGetItemDescription = objDescription.Fields(0).Value
    Set objDescription = Nothing
End Function

Public Function pfvarStockBalance(ByVal strItemId As String, ByVal strDivision As String, ByVal strDate As String) As Variant
     Dim objStockBalance As Object
     
     strSQLQuery = " SELECT Qty FROM GetStockBalance('" & strDate & "') " _
                 & " WHERE  ItemId = " & pfstrQt(strItemId) _
                 & "          and DivisionId= " & clsPrintMenu.pfstrDivisionID(strDivision)
     Set objStockBalance = clsDB.GetRecordSet(strSQLQuery)
     If Not objStockBalance.EOF Then
          pfvarStockBalance = pfvarIs_Null(objStockBalance.Fields("Qty").Value, False)
     Else
          pfvarStockBalance = 0
     End If
     Set objStockBalance = Nothing
End Function

Public Sub psubLoadSupplier(ByRef cboSource As Object)
     Dim objSupplier As Object
     Dim strSQLQuerySup As String
     
     strSQLQuerySup = strSQLQuerySup & " SELECT " _
                 & "     distinct SupplierName " _
                 & " FROM Suppliers "
                 
     If strSupCategory <> "ALL" Then
        If strSupCategory = "LOCAL" Then
            strSQLQuerySup = strSQLQuerySup & " WHERE ImportSupplier=0 "
        ElseIf strSupCategory = "IMPORTED" Then
            strSQLQuerySup = strSQLQuerySup & " WHERE ImportSupplier=1"
        End If
    End If
         
    strSQLQuerySup = strSQLQuerySup & " ORDER BY SupplierName"
    
     Set objSupplier = clsDB.GetRecordSet(strSQLQuerySup)
     cboSource.Clear
     With objSupplier
          cboSource.AddItem ""
          Do While Not .EOF
               cboSource.AddItem .Fields("SupplierName").Value
               .MoveNext
          Loop
          cboSource.ListIndex = 0
     End With
     Set objSupplier = Nothing
End Sub

Public Sub psubLoadSupplierId(ByRef objList As Object, ByVal strSupplier As String)
     Dim objSupplierId As Object
     
     strSQLQuery = " SELECT " _
                 & "     SupplierId " _
                 & " FROM Suppliers " _
                 & " WHERE SupplierName = " & pfstrQt(strSupplier)
     Set objSupplierId = clsDB.GetRecordSet(strSQLQuery)
     objList.Clear
     With objSupplierId
          Do While Not .EOF
               objList.AddItem .Fields("SupplierId").Value
               .MoveNext
          Loop
     End With
     Set objSupplierId = Nothing
End Sub

Public Sub psubLoadMaker(ByRef cboSource As Object)
     Dim objMaker As Object
     
     strSQLQuery = " SELECT " _
                 & "        distinct MakerName " _
                 & " FROM   Makers " _
                 & " ORDER BY 1"
                '& " WHERE  MakerId IN"
                '& "        (SELECT Distinct MakerId from PoHeaders)"
     Set objMaker = clsDB.GetRecordSet(strSQLQuery)
     cboSource.Clear
     With objMaker
          cboSource.AddItem ""
          Do While Not .EOF
               cboSource.AddItem .Fields("MakerName").Value
               .MoveNext
          Loop
          cboSource.ListIndex = 0
     End With
End Sub

Public Sub psubLoadCurrency(ByRef cboSource As Object)
Dim objCurrency As Object

strSQLQuery = " SELECT CurrencyUnit FROM CurrencyUnits"

Set objCurrency = clsDB.GetRecordSet(strSQLQuery)

cboSource.Clear
cboSource.AddItem "ALL"

With objCurrency
    Do While Not .EOF
        cboSource.AddItem .Fields("CurrencyUnit").Value
        .MoveNext
    Loop
    cboSource.ListIndex = 0
End With
    
End Sub

Property Let SupplierCategory(ByVal strValue As String)
    strSupCategory = strValue
End Property

