VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C08_SupplierItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strSupplierId As String
Dim strItemId     As String
Dim strPodateFrom As String
Dim strPoDateTo As String

Public Sub GetSupplierItemList(ByRef hflxGrid As Object, Optional ByVal blnBySupplier As Boolean = False _
                                   , Optional ByVal blnByItem As Boolean = False)
     Dim strSQLSupplierItem   As String
     Dim objSupplierItem      As Object
     
      strSQLSupplierItem = " SELECT " _
                        & "        Suppliers.SupplierName, SupplierItemPrices.SupplierId, SupplierItemPrices.ItemId, " _
                        & "        Items.Description, SupplierItemPrices.UnitPrice, CurrencyUnits.CurrencyUnit, " _
                        & "        SupplierItemPrices.LeadTime"
    
    If strPodateFrom <> "" And strPoDateTo <> "" Then _
    strSQLSupplierItem = strSQLSupplierItem & ",QtyUnits.QtyUnit, (Select sum(Qty) as Qty from PODetails Inner Join POHeaders On " _
                        & " PoDetails.PoNo = PoHeaders.PoNo Where PoDetails.UnitPrice = SupplierItemPrices.UnitPrice " _
                        & " And PoDetails.ItemId = SupplierItemPrices.ItemId And PoHeaders.SupplierId=SupplierItemprices.SupplierId " _
                        & " And PoHeaders.IssuedDate>=" & pfstrQt(strPodateFrom) & " And PoHeaders.IssuedDate<=" & pfstrQt(strPoDateTo) & " " _
                        & " And PoHeaders.Canceled<>1) As Qty"
                                                                
    If strPodateFrom <> "" And strPoDateTo = "" Then _
    strSQLSupplierItem = strSQLSupplierItem & ",QtyUnits.QtyUnit, (Select sum(Qty) as Qty from PODetails Inner Join POHeaders On " _
                        & " PoDetails.PoNo = PoHeaders.PoNo Where PoDetails.UnitPrice = SupplierItemPrices.UnitPrice " _
                        & " And PoDetails.ItemId = SupplierItemPrices.ItemId And PoHeaders.SupplierId=SupplierItemprices.SupplierId " _
                        & " And PoHeaders.IssuedDate>=" & pfstrQt(strPodateFrom) & " And PoHeaders.Canceled<>1) As Qty"
    
    If strPodateFrom = "" And strPoDateTo <> "" Then _
    strSQLSupplierItem = strSQLSupplierItem & ",QtyUnits.QtyUnit, (Select sum(Qty) as Qty from PODetails Inner Join POHeaders On " _
                        & " PoDetails.PoNo = PoHeaders.PoNo Where PoDetails.UnitPrice = SupplierItemPrices.UnitPrice " _
                        & " And PoDetails.ItemId = SupplierItemPrices.ItemId And PoHeaders.SupplierId=SupplierItemprices.SupplierId " _
                        & " And PoHeaders.IssuedDate<=" & pfstrQt(strPoDateTo) & " And PoHeaders.Canceled<>1) As Qty"
    
    strSQLSupplierItem = strSQLSupplierItem & " FROM " _
                        & "   SupplierItemPrices INNER JOIN " _
                        & "   Suppliers ON SupplierItemPrices.SupplierId = Suppliers.SupplierId INNER JOIN " _
                        & "   Items ON SupplierItemPrices.ItemId = Items.ItemId INNER JOIN" _
                        & "   CurrencyUnits ON Suppliers.CurrencyUnitId = CurrencyUnits.CurrencyUnitId" _
                        & "   INNER JOIN QtyUnits On Items.SupplierQtyUnitId=QtyUnits.QtyUnitId"
                       
     If blnBySupplier And Not strSupplierId = "" Then _
          strSQLSupplierItem = strSQLSupplierItem _
                        & " WHERE (SupplierItemPrices.SupplierId = " & pfstrQt(strSupplierId) & ")"
     If blnByItem And Not strItemId = "" Then _
          strSQLSupplierItem = strSQLSupplierItem _
                        & " WHERE (SupplierItemPrices.ItemId = " & pfstrQt(strItemId) & ")"
                        
     If strPodateFrom <> "" Or strPoDateTo <> "" Then _
     strSQLSupplierItem = strSQLSupplierItem & " GROUP BY Suppliers.SupplierName, SupplierItemPrices.SupplierId, SupplierItemPrices.ItemId, " _
                        & " Items.Description, SupplierItemPrices.UnitPrice, CurrencyUnits.CurrencyUnit, SupplierItemPrices.LeadTime, " _
                        & " QtyUnits.QtyUnit"
                             
     strSQLSupplierItem = strSQLSupplierItem & " ORDER BY SupplierName,SupplierItemprices.ItemId"
                        
     Set objSupplierItem = clsDB.GetRecordSet(strSQLSupplierItem)
     Set hflxGrid.DataSource = objSupplierItem
     Set objSupplierItem = Nothing
End Sub

Public Property Let SupplierId(ByVal strValue As String)
     strSupplierId = strValue
End Property

Public Property Let ItemId(ByVal strValue As String)
     strItemId = strValue
End Property
Public Property Let PoDateFrom(ByVal strValue As String)
strPodateFrom = strValue
End Property
Public Property Let PoDateTo(ByVal strValue As String)
strPoDateTo = strValue
End Property


