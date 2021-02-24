VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function pfLoadAccomplishments(Optional where As Variant) As Object
        strSQL = ""
        strSQL = strSQL & " SELECT DISTINCT WO.CompanyID,WO.DEPARTMENTID ,D.DepartmentName"
        strSQL = strSQL & " FROM WorkOrderDetails WO"
        strSQL = strSQL & " INNER JOIN Departments D"
        strSQL = strSQL & " ON D.DepartmentID = WO.DepartmentId"
        strSQL = strSQL & " Where"
        strSQL = strSQL & " WO.TYPEID = " & where(0)
        strSQL = strSQL & " AND WO.CompanyID = " & where(1)
        strSQL = strSQL & " AND"
        strSQL = strSQL & " ("
        strSQL = strSQL & "         ("
        strSQL = strSQL & "                 CONVERT(VARCHAR(10),ReceivedDate,111) >= '" & where(2) & "'"
        strSQL = strSQL & "                 AND CONVERT(VARCHAR(10),ReceivedDate,111) <= '" & where(3) & "'"
        strSQL = strSQL & "         )"
        strSQL = strSQL & "         OR"
        strSQL = strSQL & "         ("
        strSQL = strSQL & "                 CONVERT(VARCHAR(10),FinishedDate,111) >= '" & where(2) & "'"
        strSQL = strSQL & "                 AND CONVERT(VARCHAR(10),FinishedDate,111) <= '" & where(3) & "'"
        strSQL = strSQL & "         )"
        strSQL = strSQL & " )"
        
       Set pfLoadAccomplishments = GetRecordSet(strSQL)
End Function
Public Function pfLoadAccomplishmentsDetails(ByVal loopD As Date, _
                                                                                ByVal intType As Integer, _
                                                                                ByVal strCompanyID As String, _
                                                                                ByVal intDeptID As Integer, _
                                                                                ByVal intRecFin As Integer, _
                                                                                ByVal strAbbvrName As String) As Object
        
        
        strSQL = "SELECT COUNT(*) AS " & IIf(intRecFin = 0, "SCHEDULED", "FINISHED")
        strSQL = strSQL & " FROM WorkOrderDetails WHERE CompanyID = '" & strCompanyID & "'"
        strSQL = strSQL & " AND DepartmentID = " & intDeptID
        strSQL = strSQL & " AND TypeID = " & intType
        
        If strAbbvrName <> "" Then strSQL = strSQL & " AND AbbreviatedName = '" & strAbbvrName & "'"
        
        strSQL = strSQL & " AND CONVERT(VARCHAR(10),WorkOrderDetails." & IIf(intRecFin = 0, "ReceivedDate", "FinishedDate") & ", 111) = '" & loopD & "'"
        
       Set pfLoadAccomplishmentsDetails = GetRecordSet(strSQL)
End Function
Public Function pfLoadEmployeeMasterlist(Optional strSQLwhere) As Object
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " Companies.CompanyName,"
    strSQL = strSQL & " Users.EmployeeNo,"
    strSQL = strSQL & " Users.EmployeeName,"
    strSQL = strSQL & " Users.RegisteredDate,"
    strSQL = strSQL & " Users.UpdatedDate,"
    strSQL = strSQL & " Users.DeletedDate"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " Users"
    strSQL = strSQL & " INNER JOIN Companies"
    strSQL = strSQL & " ON Users.CompanyID = Companies.CompanyID"
    strSQL = strSQL & strSQLwhere
    Set pfLoadEmployeeMasterlist = GetRecordSet(strSQL)
End Function

Public Function pfLoadMachineMasterlist(Optional strSQLwhere) As Object
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "WorkOrderItems.ItemCode, "
    strSQL = strSQL & "WorkOrderItems.TypeID, "
    strSQL = strSQL & "WorkOrderItems.ItemName, "
    strSQL = strSQL & "WorkOrderItems.CompanyID, "
    strSQL = strSQL & "ItemsDepartmentSectionLines.DepartmentID, "
    strSQL = strSQL & "ItemsDepartmentSectionLines.SectionID, "
    strSQL = strSQL & "ItemsDepartmentLocation.LocationID, "
    strSQL = strSQL & "ItemsDepartmentSectionLines.LineID, "
    strSQL = strSQL & "WorkOrderItems.PriorityLevel, "
    strSQL = strSQL & "WorkOrderItems.MakerName, "
    strSQL = strSQL & "WorkOrderItems.Model, "
    strSQL = strSQL & "WorkOrderItems.SerialNo, "
    strSQL = strSQL & "WorkOrderItems.Capacity, "
    strSQL = strSQL & "WorkOrderItems.FixedAssetNo, "
    strSQL = strSQL & "WorkOrderItems.PreventiveMaintenance, "
    strSQL = strSQL & "WorkOrderItems.EngineModel, "
    strSQL = strSQL & "WorkOrderItems.EngineSerialNo, "
    strSQL = strSQL & "WorkOrderItems.Transmission, "
    strSQL = strSQL & "WorkOrderItems.MastType, "
    strSQL = strSQL & "WorkOrderItems.MastHeight, "
    strSQL = strSQL & "WorkOrderItems.AttachmentType, "
    strSQL = strSQL & "WorkOrderItems.FrontTire, "
    strSQL = strSQL & "WorkOrderItems.FrontTireHoles, "
    strSQL = strSQL & "WorkOrderItems.RearTire, "
    strSQL = strSQL & "WorkOrderItems.RearTireHoles, "
    strSQL = strSQL & "WorkOrderItems.WorkStatus, "
    strSQL = strSQL & "WorkOrderItems.AcquisitionAmount, "
    strSQL = strSQL & "WorkOrderItems.AcquisitionDate, "
    strSQL = strSQL & "WorkOrderItems.DisposalDate, "
    strSQL = strSQL & "WorkOrderItems.Remarks "
     
    strSQL = strSQL & "FROM WorkOrderItems "
    
    strSQL = strSQL & "INNER JOIN Types ON WorkOrderItems.CompanyID = Types.CompanyID "
    strSQL = strSQL & "AND WorkOrderItems.TypeID = Types.TypeID "
    
    strSQL = strSQL & "LEFT JOIN ItemsDepartmentSectionLines ON WorkOrderItems.ItemCode = ItemsDepartmentSectionLines.ItemCode "
    strSQL = strSQL & "AND WorkOrderItems.CompanyID = ItemsDepartmentSectionLines.CompanyID "
    
    strSQL = strSQL & "LEFT JOIN ItemsDepartmentLocation ON WorkOrderItems.ItemCode = ItemsDepartmentLocation.ItemCode "
    strSQL = strSQL & "AND WorkOrderItems.CompanyID = ItemsDepartmentLocation.CompanyID "
    
    strSQL = strSQL & "LEFT JOIN Departments ON ItemsDepartmentSectionLines.CompanyID = Departments.CompanyID "
    strSQL = strSQL & "AND ItemsDepartmentSectionLines.DepartmentID = Departments.DepartmentID "

    strSQL = strSQL & "LEFT JOIN Sections ON ItemsDepartmentSectionLines.CompanyID = Sections.CompanyID "
    strSQL = strSQL & "AND ItemsDepartmentSectionLines.DepartmentID = Sections.DepartmentID "
    strSQL = strSQL & "AND ItemsDepartmentSectionLines.SectionID = Sections.SectionID "
        
    strSQL = strSQL & "WHERE " & strSQLwhere
    Set pfLoadMachineMasterlist = GetRecordSet(strSQL)
End Function

Public Function pfLoadMachineControlStatus(Optional strSQLwhere) As Object
    Dim strSQL As String
    strSQL = " "
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " CASE"
    strSQL = strSQL & " WHEN WorkStatus = 0 THEN 'No Status'"
    strSQL = strSQL & " WHEN WorkStatus = 1 THEN 'Dispose'"
    strSQL = strSQL & " WHEN WorkStatus = 2 THEN 'Transfer'"
    strSQL = strSQL & " WHEN WorkStatus = 3 THEN 'Inactive'"
    strSQL = strSQL & " WHEN WorkStatus = 4 THEN 'Active'"
    strSQL = strSQL & " WHEN WorkStatus = 5 THEN 'Not Existing'"
    strSQL = strSQL & " END AS Status,"
    strSQL = strSQL & " ItemCode,"
    strSQL = strSQL & " ItemName,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " MakerName,"
    strSQL = strSQL & " typeName,"
    strSQL = strSQL & " Location,"
    strSQL = strSQL & " LineName,"
    strSQL = strSQL & " Capacity,"
    strSQL = strSQL & " FixedAssetNo,"
    strSQL = strSQL & " PreventiveMaintenance,"
    strSQL = strSQL & " EngineModel,"
    strSQL = strSQL & " EngineSerialNo,"
    strSQL = strSQL & " Transmission,"
    strSQL = strSQL & " MastType,"
    strSQL = strSQL & " AttachmentType,"
    strSQL = strSQL & " FrontTire,"
    strSQL = strSQL & " FrontTireHoles,"
    strSQL = strSQL & " AcquisitionAmount,"
    strSQL = strSQL & " AcquisitionDate,"
    strSQL = strSQL & " DisposalDate,"
    strSQL = strSQL & " Remarks"
    strSQL = strSQL & " From WorkOrderItemStatusView"
    strSQL = strSQL & strSQLwhere
    Set pfLoadMachineControlStatus = GetRecordSet(strSQL)
    
End Function

Public Function pfLoadHistory(dtFrom As Date, dtTo As Date, Optional strSQLwhere, _
                                                Optional strStatus, Optional strServer As String) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT"
    strSQL = strSQL & " WOControlNo,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " LineName,"
    strSQL = strSQL & " ControlNo,"
    strSQL = strSQL & " ItemName,"
    strSQL = strSQL & " MainCategoryName + ' - ' + MainSubCategoryName AS WorkCategory,"
    strSQL = strSQL & " MachineClassification,"
    strSQL = strSQL & " PartOfMachine,"
    strSQL = strSQL & " MachineProblemFound,"
    strSQL = strSQL & " ConditionProblem,"
    strSQL = strSQL & " ReceivedDate,"
    strSQL = strSQL & " RespondDate,"
    strSQL = strSQL & " StartedDate,"
    strSQL = strSQL & " FinishedDate,"
    strSQL = strSQL & " ActionTaken,"
    
    strSQL = strSQL & " c.ItemID,"
    strSQL = strSQL & " c.Description,"
    strSQL = strSQL & " c.Qty,"
    strSQL = strSQL & " c.CurrencyUnit,"
    strSQL = strSQL & " c.UnitPrice,"
    strSQL = strSQL & " c.Qty * c.UnitPrice AS TotalCost,"
    
    strSQL = strSQL & " EmployeeName = "
    strSQL = strSQL & " STUFF((SELECT DISTINCT  ' -- '  + EmployeeName  "
    strSQL = strSQL & "        FROM CostingAndHistoryView b"
    strSQL = strSQL & "        where b.AbbreviatedName = a.AbbreviatedName AND b.WorkOrderControlNo = a.WorkOrderControlNo"
    strSQL = strSQL & "       FOR XML PATH('')), 1, 2, ''), "
    strSQL = strSQL & " Status,"
    strSQL = strSQL & " Remarks,"
    strSQL = strSQL & " ManPowerAffected,"
    strSQL = strSQL & " TotalMinutesBreakDown,"
    strSQL = strSQL & " TotalManHourLoss,"
    strSQL = strSQL & " TotalMinutesRepairTime,"
    strSQL = strSQL & " TargetDate, AbbreviatedName, No"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " CostingAndHistoryView a"
    
    
    strSQL = strSQL & "  LEFT JOIN [" & strServer & "].WarehouseManagement.dbo.WOM_ItemShipmentDetailsView c ON"
    strSQL = strSQL & "  a.WOControlNo = c.ItemShipmentID"
    
    
    strSQL = strSQL & " WHERE CONVERT(VARCHAR(20)," & strStatus & "Date,111) >= '" & dtFrom & "'"
    strSQL = strSQL & " AND CONVERT(VARCHAR(20)," & strStatus & "Date,111) <= '" & dtTo & "'"
    
    strSQL = strSQL & strSQLwhere
    
    strSQL = strSQL & " ORDER BY AbbreviatedName, No ASC"
    DoEvents
    Set pfLoadHistory = GetRecordSet(strSQL)
    
End Function
Public Function pfLoadMaintenance(dtFrom As Date, dtTo As Date, _
                                        Optional strSQLwhere, Optional strServer As String) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT"
    strSQL = strSQL & " AbbreviatedName + ' ' + WorkOrderControlNo As WOControlNo,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " LineName,"
    strSQL = strSQL & " ControlNo,"
    strSQL = strSQL & " ItemName,"
    strSQL = strSQL & " ReceivedDate,"
    strSQL = strSQL & " RequestDate,"
    strSQL = strSQL & " FinishedDate,"
    strSQL = strSQL & " TotalMinutesBreakDown,"
    
    strSQL = strSQL & " c.ItemID,"
    strSQL = strSQL & " c.Description,"
    strSQL = strSQL & " c.Qty,"
    strSQL = strSQL & " c.UnitPrice,"
    strSQL = strSQL & " c.Qty * c.UnitPrice AS TotalCost,"
    
    strSQL = strSQL & " EmployeeName = "
    strSQL = strSQL & " STUFF((SELECT DISTINCT  ' -- '  + EmployeeName  "
    strSQL = strSQL & "        FROM CostingAndHistoryView b"
    strSQL = strSQL & "        where b.AbbreviatedName = a.AbbreviatedName AND b.WorkOrderControlNo = a.WorkOrderControlNo"
    strSQL = strSQL & "       FOR XML PATH('')), 1, 2, ''), "
    strSQL = strSQL & " Status,"
    strSQL = strSQL & " Remarks, AbbreviatedName, No"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " CostingAndHistoryView a"
    
    
    strSQL = strSQL & "  LEFT JOIN [" & strServer & "].WarehouseManagement.dbo.WOM_ItemShipmentDetailsView c ON"
    strSQL = strSQL & "  a.WOControlNo = c.ItemShipmentID"
    
    
    strSQL = strSQL & " WHERE CONVERT(VARCHAR(20)," & IIf(frmMaintenance.optReceived.Value = True, "ReceivedDate", "RequestDate") & ",111) >= '" & dtFrom & "'"
    strSQL = strSQL & " AND CONVERT(VARCHAR(20)," & IIf(frmMaintenance.optReceived.Value = True, "ReceivedDate", "RequestDate") & ",111)  <= '" & dtTo & "'"
 
    strSQL = strSQL & strSQLwhere
'
'    strSQL = strSQL & " GROUP BY"
'    strSQL = strSQL & " AbbreviatedName, No, WorkOrderControlNo,"
'    strSQL = strSQL & " CompanyName,"
'    strSQL = strSQL & " DepartmentName,"
'    strSQL = strSQL & " SectionName,"
'    strSQL = strSQL & " LineName,"
'    strSQL = strSQL & " ControlNo,"
'    strSQL = strSQL & " ItemName,"
'    strSQL = strSQL & " ReceivedDate,"
'    strSQL = strSQL & " FinishedDate,"
'    strSQL = strSQL & " TotalMinutesBreakDown,"
'    strSQL = strSQL & " ItemID,"
'    strSQL = strSQL & " Description,"
'    strSQL = strSQL & " Qty,"
'    strSQL = strSQL & " UnitPrice,"
'    strSQL = strSQL & " TotalCost,"
'    strSQL = strSQL & " EmployeeName,"
'    strSQL = strSQL & " Status,"
'    strSQL = strSQL & " Remarks"

    strSQL = strSQL & " ORDER BY No,AbbreviatedName ASC"

    Set pfLoadMaintenance = GetRecordSet(strSQL)
    
End Function
Public Function pfLoadCosting(dtFrom As Date, dtTo As Date, Optional strSQLwhere) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " Select"
    strSQL = strSQL & " *"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " CostingAndHistoryView "
    strSQL = strSQL & " WHERE CONVERT(VARCHAR(20),ReceivedDate,111) >= '" & dtFrom & "'"
    strSQL = strSQL & " AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '" & dtTo & "' "
    strSQL = strSQL & strSQLwhere
    Set pfLoadCosting = GetRecordSet(strSQL)
    
End Function

Public Function pfLoadStatus(ByVal intType As Integer, strWhere As String, _
                                Optional strServer As String) As Object
    Dim strSQL As String
    strSQL = ""
'    strSQL = strSQL & " SELECT *"
'    strSQL = strSQL & " FROM"
'    strSQL = strSQL & " DailyReportView"

    strSQL = strSQL & "SELECT DISTINCT"
    strSQL = strSQL & " b.CompanyName,"
    strSQL = strSQL & " b.DepartmentName,"
    strSQL = strSQL & " b.DateOfWorkOrder,"
    strSQL = strSQL & " b.WorkCategory,"
    strSQL = strSQL & " b.Section,"
    strSQL = strSQL & " b.Line,"
    strSQL = strSQL & " b.LeaderIncharge,"
    strSQL = strSQL & " b.AbbreviatedName + ' ' + b.WorkOrderControlNo AS WorkOrderControlNo , "
    strSQL = strSQL & " b.MachineItemNo,"
    strSQL = strSQL & " b.MakerName,"
    strSQL = strSQL & " b.Model,"
    strSQL = strSQL & " b.MachineName,"
    strSQL = strSQL & " b.TypeOfRequest,"
    strSQL = strSQL & " b.MachineProblem,"
    strSQL = strSQL & " b.Status,"
    strSQL = strSQL & " b.Description,"
    strSQL = strSQL & " sum(b.Qty) AS Qty,"
    strSQL = strSQL & " b.QtyDescription,"
    strSQL = strSQL & " b.PrsDate,"
    strSQL = strSQL & " b.PrsNo,"
    strSQL = strSQL & " a.PoNo,"
    strSQL = strSQL & " b.PrsExpectedDelivery,"
    strSQL = strSQL & " a.EtaDate AS PoExpectedDelivery,"
    strSQL = strSQL & " a.ReceivedDate AS ActualReceived,"
    strSQL = strSQL & " b.FinishedDate,"
    strSQL = strSQL & " b.Remarks,"
    strSQL = strSQL & " b.No,"
    strSQL = strSQL & " b.AbbreviatedName"
    strSQL = strSQL & " FROM DailyReportView b"
    
    strSQL = strSQL & " LEFT JOIN [" & strServer & "].WarehouseManagement.dbo.LatestPoNoView a ON"
    strSQL = strSQL & " b.PrsNo = a.RequisitionSlipNo"
    
    strSQL = strSQL & " WHERE TypeID = " & intType
    strSQL = strSQL & strWhere
    
    strSQL = strSQL & " GROUP BY "
    strSQL = strSQL & "b.CompanyName,"
    strSQL = strSQL & "b.DepartmentName,"
    strSQL = strSQL & "b.DateOfWorkOrder,"
    strSQL = strSQL & "b.WorkCategory,"
    strSQL = strSQL & "b.Section,"
    strSQL = strSQL & "b.Line,"
    strSQL = strSQL & "b.LeaderIncharge,"
    strSQL = strSQL & "b.AbbreviatedName,"
    strSQL = strSQL & "b.WorkOrderControlNo ,"
    strSQL = strSQL & "b.MachineItemNo,"
    strSQL = strSQL & "b.MakerName,"
    strSQL = strSQL & "b.Model,"
    strSQL = strSQL & "b.MachineName,"
    strSQL = strSQL & "b.TypeOfRequest,"
    strSQL = strSQL & "b.MachineProblem,"
    strSQL = strSQL & "b.Status,"
    strSQL = strSQL & "b.Description,"
    strSQL = strSQL & "b.QtyDescription,"
    strSQL = strSQL & "b.PrsDate,"
    strSQL = strSQL & "b.PrsNo,"
    strSQL = strSQL & "a.PoNo,"
    strSQL = strSQL & "b.PrsExpectedDelivery,"
    strSQL = strSQL & "a.EtaDate,"
    strSQL = strSQL & "a.ReceivedDate,"
    strSQL = strSQL & "b.FinishedDate,"
    strSQL = strSQL & "b.Remarks,"
    strSQL = strSQL & "b.No,"
    strSQL = strSQL & "b.AbbreviatedName"
    
    strSQL = strSQL & " ORDER BY AbbreviatedName, No ASC"
    Set pfLoadStatus = GetRecordSet(strSQL)
End Function

'Public Function pfLoadBreakdown1(ByVal d1 As Date, ByVal d2 As Date, intType As Integer) As Object
'    Dim strSQL As String
'    strSQL = ""
'    strSQL = strSQL & "SELECT * FROM GetDetails('" & d1 & "','" & d2 & "'," & intType & ")"
'    Set pfLoadBreakdown1 = GetRecordSet(strSQL)
'End Function

Public Function pfLoadBreakdown(ByVal d1 As Date, ByVal d2 As Date, strCompany As String, intType As Integer) As Object
    Dim strSQL As String
    strSQL = ""
       
        strSQL = strSQL & " SELECT * FROM "
        strSQL = strSQL & " GetBreakDownDetails('" & d1 & "','" & d2 & "','" & strCompany & "'," & intType & ")"
    DoEvents
    Set pfLoadBreakdown = GetRecordSet(strSQL)
End Function
Public Function pfLoadBreakdown_forklift(ByVal d1 As Date, ByVal d2 As Date, _
                            strType As String, Optional ByVal d3 As Date) As Object
    Dim strSQL As String
        strSQL = ""
        strSQL = strSQL & " SELECT  '" & strType & "' AS TYPE ,CompanyID,CompanyName,"
        strSQL = strSQL & " sum(forschedule) AS [FOR SCHEDULE],"
        strSQL = strSQL & " sum(Waitingparts) AS [WAITING PARTS],"
        strSQL = strSQL & " sum(Ongoing) As [ON GOING],"
        strSQL = strSQL & " sum(FinishedRepair) As [FINISHED REPAIR],"
        strSQL = strSQL & " sum(NEWBREAKDOWN) As [NEW BREAKDOWN],"
        strSQL = strSQL & " sum(PREVIOUSPENDING) As [PREVIOUS PENDING]"


        
        strSQL = strSQL & " FROM  GetBreakDownDetails_ForkLift('" & d1 & "','" & d2 & "','" & d3 & "','" & d3 - 1 & "')"
        strSQL = strSQL & " WHERE ItemName LIKE '%" & strType & "%'"
        strSQL = strSQL & " GROUP BY CompanyName,CompanyID"

   
    Set pfLoadBreakdown_forklift = GetRecordSet(strSQL)
End Function

Public Function pfGetsubdetail(SearchingTable As String, strSelectingField As String, _
                                    ByVal strWhereField As String, ByVal strFieldValue As String, _
                                    Optional strWhereField2 As String, Optional strFieldValue2 As String, _
                                    Optional strWhereField3 As String, Optional strFieldValue3 As String, _
                                    Optional strWhereField4 As String, Optional strFieldValue4 As String) As String
                                    
    Dim adoRecordset As ADODB.Recordset
    strSQL = "SELECT "
    strSQL = strSQL & strSelectingField
    strSQL = strSQL & " FROM "
    strSQL = strSQL & SearchingTable
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & strWhereField & " = " & strFieldValue
    
    If strWhereField2 <> "" And strFieldValue2 <> "" Then
        strSQL = strSQL & " AND " & strWhereField2 & " = " & strFieldValue2
    End If
    
    If strWhereField3 <> "" And strFieldValue3 <> "" Then
        strSQL = strSQL & " AND " & strWhereField3 & " = " & strFieldValue3
    End If
    
    If strWhereField4 <> "" And strFieldValue4 <> "" Then
        strSQL = strSQL & " AND " & strWhereField4 & " = " & strFieldValue4
    End If
    
    Set adoRecordset = New ADODB.Recordset
    If ((strFieldValue = "" And strFieldValue2 <> "") Or (strFieldValue = "" And strFieldValue2 = "")) Or _
            ((strFieldValue = "" And strFieldValue3 <> "") Or (strFieldValue = "" And strFieldValue3 = "")) Or _
                ((strFieldValue = "" And strFieldValue4 <> "") Or (strFieldValue = "" And strFieldValue4 = "")) Then
        pfGetsubdetail = "": Exit Function
    End If
    
    adoRecordset.Open strSQL, cn, adOpenDynamic, adLockReadOnly
    If Not adoRecordset.EOF Then pfGetsubdetail = adoRecordset.Fields(strSelectingField).Value: Exit Function
    pfGetsubdetail = ""
End Function

Public Function pfLoadPRSMonitoring(CompanyID As String, Server As String, Optional strSQLwhere) As Object

'    strSQL = ""
'    strSQL = strSQL & " SELECT "
'    strSQL = strSQL & " PurchaseRequest.IsImport,"
'    strSQL = strSQL & " PurchaseRequest.PurchaseRequestNo,"
'    strSQL = strSQL & " WorkOrder.AbbreviatedName + ' ' + WorkOrder.WorkOrderControlNo AS WorkOrderNo,"
'    strSQL = strSQL & " PurchaseRequest.ItemId,"
'    strSQL = strSQL & " PurchaseRequest.Description,"
'    strSQL = strSQL & " WorkOrder.DepartmentName,"
'    strSQL = strSQL & " WorkOrder.Location_Line,"
'    strSQL = strSQL & " WorkOrder.MachineItemNo,"
'    strSQL = strSQL & " WorkOrder.MachineName,"
'    strSQL = strSQL & " WorkOrder.ProblemFound,"
'    strSQL = strSQL & " PurchaseRequest.IssuedDate,"
'    strSQL = strSQL & " PurchaseRequest.DeliveryDate,"
'    strSQL = strSQL & " PurchaseRequest.DeliveryCompletedDate,"
'    strSQL = strSQL & " PurchaseRequest.Qty,"
'    strSQL = strSQL & " PurchaseRequest.QtyReceived,"
'    strSQL = strSQL & " PurchaseRequest.QtyUnit,"
'    strSQL = strSQL & " PurchaseRequest.PoNo,"
'    strSQL = strSQL & " PurchaseRequest.EtdDate,"
'    strSQL = strSQL & " PurchaseRequest.InvoiceNo,"
'    strSQL = strSQL & " WorkOrder.LeaderIncharge,"

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " PurchaseRequest.IsImport,"
    strSQL = strSQL & " PurchaseRequest.PurchaseRequestNo,"
    strSQL = strSQL & " AdditionalData.TagNo,"
    strSQL = strSQL & " WorkOrderPRS.WorkOrderNo,"
    strSQL = strSQL & " PurchaseRequest.ItemId,"
    strSQL = strSQL & " NewItem.NewItemId,"
    strSQL = strSQL & " PurchaseRequest.Description,"
    strSQL = strSQL & " WorkOrderPRS.DepartmentName,"
    strSQL = strSQL & " WorkOrderPRS.Location + ' / ' + WorkOrderPRS.LineName AS Location_Line,"
    strSQL = strSQL & " WorkOrderPRS.ItemCode,"
    strSQL = strSQL & " WorkOrderPRS.ItemName,"
    strSQL = strSQL & " WOMDE_ItemStatus.EquipmentStatus,"
    strSQL = strSQL & " WorkOrderPRS.ProblemFound,"
    strSQL = strSQL & " PurchaseRequest.IssuedDate,"
    strSQL = strSQL & " AdditionalData.ReceivedByPurchasing,"
    strSQL = strSQL & " PurchaseRequest.DeliveryDate,"
    strSQL = strSQL & " PurchaseRequest.DeliveryCompletedDate,"
    strSQL = strSQL & " PurchaseRequest.Qty,"
    strSQL = strSQL & " PurchaseRequest.QtyReceived,"
    strSQL = strSQL & " PurchaseRequest.QtyUnit,"
    strSQL = strSQL & " PurchaseRequest.PoNo,"
    strSQL = strSQL & " PurchaseRequest.EtdDate,"
    strSQL = strSQL & " PurchaseRequest.InvoiceNo,"
    strSQL = strSQL & " AdditionalData.Incharge,"
    strSQL = strSQL & " AdditionalData.ReceivedBy,"
    strSQL = strSQL & " AdditionalData.Remarks,"
    
    If Server = "wkn-appserver" Or Server = "a-sv17" Then
        strSQL = strSQL & " PurchaseRequest.CancelledDate,"
        strSQL = strSQL & " PurchaseRequest.Division,"
        strSQL = strSQL & " PurchaseRequest.FinalDestination,"
    ElseIf Server = "jd-004" Then
        strSQL = strSQL & " PurchaseRequest.Cancelled,"
    End If
    
    strSQL = strSQL & " PurchaseRequest.PurchaseRequestDetailSeq"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " [" & Server & "].WarehouseManagement.dbo.PurchaseRequestMaintenanceALLView PurchaseRequest"
    
    strSQL = strSQL & " LEFT JOIN WorkOrderPRS"
    strSQL = strSQL & " ON   WorkOrderPRS.PrsNo = PurchaseRequest.PurchaseRequestNo "
    
    strSQL = strSQL & " LEFT JOIN WOMDE_ItemStatus"
    strSQL = strSQL & " ON WOMDE_ItemStatus.ItemCode = WorkOrderPRS.ItemCode"

    strSQL = strSQL & " LEFT JOIN"
    strSQL = strSQL & " (SELECT * FROM WOMDE_PRSAdditionalData WHERE WOMDE_PRSAdditionalData.CompanyId = '" & CompanyID & "')AdditionalData"
    strSQL = strSQL & " ON AdditionalData.PurchaseRequestNo = PurchaseRequest.PurchaseRequestNo"

    strSQL = strSQL & " LEFT JOIN"
    strSQL = strSQL & " (SELECT * FROM WOMDE_PRSNewItem WHERE WOMDE_PRSNewItem.CompanyId = '" & CompanyID & "')NewItem"
    strSQL = strSQL & " ON NewItem.PurchaseRequestNo = PurchaseRequest.PurchaseRequestNo"
    strSQL = strSQL & " AND NewItem.PurchaseRequestDetailSeq=PurchaseRequest.PurchaseRequestDetailSeq "
    
    strSQL = strSQL & strSQLwhere
    
    Set pfLoadPRSMonitoring = GetRecordSet(strSQL)
    
End Function

Public Function pfLoadPending(strAbbvrName As String, dtFrom As String, dtTo As String, _
                                Optional strCompany As String, Optional intType As Integer, _
                                Optional intStatus As Integer, Optional blnForkLift As Boolean, _
                                Optional strWrkCategory As String) As Object
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "RequestDate, "
    strSQL = strSQL & "ReceivedDate, "
    strSQL = strSQL & "ControlNo, "
    strSQL = strSQL & "DepartmentName, "
    strSQL = strSQL & "SectionName, "
    strSQL = strSQL & "Location, "
    strSQL = strSQL & "ItemCode, "
    strSQL = strSQL & "ItemName, "
    strSQL = strSQL & "Problem + ' - ' + Reason AS 'Problem', "
    strSQL = strSQL & "RequesterName1, "
    strSQL = strSQL & "RequesterName2, "
    strSQL = strSQL & "LeaderIncharge, "
    strSQL = strSQL & "StartedDate, "
    strSQL = strSQL & "FinishedDate, "
    strSQL = strSQL & "Status "
    strSQL = strSQL & "FROM WorkOrderSummaryView "

    strSQL = strSQL & " WHERE CONVERT(VARCHAR(20),ReceivedDate,111) >= '" & dtFrom & "' "
    strSQL = strSQL & "AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '" & dtTo & "' "
    
    If strCompany <> "" Then strSQL = strSQL & "AND CompanyID = '" & strCompany & "'"
    
    If intType <> 0 Then
        strSQL = strSQL & "AND TypeID = " & intType
    ElseIf blnForkLift = True Then
        strSQL = strSQL & "AND TypeName LIKE '%FORKLIFT%'"
    End If
    
    If strWrkCategory <> "" Then strSQL = strSQL & " AND MainCategoryName Like '%" & strWrkCategory & "%' "
    
    If intStatus <> 0 Then strSQL = strSQL & "AND StatusID = " & intStatus
    
    If strAbbvrName <> "" Then strSQL = strSQL & " AND AbbreviatedName LIKE '%" & strAbbvrName & "%' "
    
    strSQL = strSQL & " ORDER BY AbbreviatedName, No ASC"
    
    Set pfLoadPending = GetRecordSet(strSQL)
    
End Function

Public Function pfLoadPRSAdditionalData(Optional strSQLwhere) As Object

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " * FROM WOMDE_PRSAdditionalData"
    strSQL = strSQL & strSQLwhere
    
    Set pfLoadPRSAdditionalData = GetRecordSet(strSQL)
    
End Function

Public Function pfLoadPRSNewItem(Optional strSQLwhere) As Object

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " * FROM WOMDE_PRSNewItem"
    strSQL = strSQL & strSQLwhere
    
    Set pfLoadPRSNewItem = GetRecordSet(strSQL)
    
End Function

Public Function pfLoadMachineStatus(Optional strSQLwhere) As Object

    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " * FROM WOMDE_MachineStatus"
    strSQL = strSQL & strSQLwhere
    
    Set pfLoadMachineStatus = GetRecordSet(strSQL)
End Function










