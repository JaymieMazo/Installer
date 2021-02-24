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
        strSQL = strSQL & " FROM WorkOrdersView WO"
        strSQL = strSQL & " INNER JOIN Departments D"
        strSQL = strSQL & " ON D.DepartmentID = WO.DepartmentId"
        strSQL = strSQL & " Where"
        strSQL = strSQL & " WO.TYPEID = " & where(0)
        strSQL = strSQL & " AND WO.CompanyID = " & where(1)
        strSQL = strSQL & " AND"
        strSQL = strSQL & " ("
        strSQL = strSQL & "         ("
        strSQL = strSQL & "                 CONVERT(VARCHAR(20),ReceivedDate,111) >= '" & where(2) & "'"
        strSQL = strSQL & "                 AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '" & where(3) & "'"
        strSQL = strSQL & "         )"
        strSQL = strSQL & "         OR"
        strSQL = strSQL & "         ("
        strSQL = strSQL & "                 CONVERT(VARCHAR(20),FinishedDate,111) >= '" & where(2) & "'"
        strSQL = strSQL & "                 AND CONVERT(VARCHAR(20),FinishedDate,111) <= '" & where(3) & "'"
        strSQL = strSQL & "         )"
        strSQL = strSQL & " )"
        
       Set pfLoadAccomplishments = GetRecordSet(strSQL)
End Function
Public Function pfLoadAccomplishmentsDetails(ByVal loopD As Date, _
                                                                                ByVal intType As Integer, _
                                                                                ByVal intCompanyID As Integer, _
                                                                                ByVal intDeptID As Integer, _
                                                                                ByVal intRecFin As Integer) As Object
        
        
        strSQL = "SELECT dbo.GetAccomplishment ('" & loopD & "'," & intType & "," & intCompanyID & "," & intDeptID & "," & intRecFin & ") as RESULT"
  
        
       Set pfLoadAccomplishmentsDetails = GetRecordSet(strSQL)
End Function
Public Function pfLoadEmployeeMasterlist(Optional strSQLwhere) As Object
    strSQL = ""
    strSQL = strSQL & " SELECT "
    strSQL = strSQL & " Companies.CompanyName,"
    strSQL = strSQL & " Employees.EmployeeNo,"
    strSQL = strSQL & " Employees.EmployeeName,"
    strSQL = strSQL & " Employees.RegisteredDate,"
    strSQL = strSQL & " Employees.UpdatedDate,"
    strSQL = strSQL & " Employees.DeletedDate"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " Employees"
    strSQL = strSQL & " INNER JOIN Companies"
    strSQL = strSQL & " ON Employees.CompanyID = Companies.CompanyID"
    strSQL = strSQL & strSQLwhere
    Set pfLoadEmployeeMasterlist = GetRecordSet(strSQL)
End Function

Public Function pfLoadMachineMasterlist(Optional strSQLwhere) As Object
    strSQL = ""
    strSQL = strSQL & " SELECT * FROM MachineItems Where" & strSQLwhere
    Set pfLoadMachineMasterlist = GetRecordSet(strSQL)
End Function

Public Function pfLoadMachineControlStatus(Optional strSQLwhere) As Object
    Dim strSQL As String
    strSQL = " "
    strSQL = strSQL & " SELECT"
    strSQL = strSQL & " CASE"
    strSQL = strSQL & " WHEN Status = 1 THEN 'Dispose'"
    strSQL = strSQL & " WHEN Status = 2 THEN 'Transfer'"
    strSQL = strSQL & " WHEN Status = 3 THEN 'Active'"
    strSQL = strSQL & " WHEN Status = 4 THEN 'Inactive'"
    strSQL = strSQL & " END AS Status,"
    strSQL = strSQL & " MachineItemNo,"
    strSQL = strSQL & " MachineName,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " MakerName,"
    strSQL = strSQL & " typeName,"
    strSQL = strSQL & " Location,"
    strSQL = strSQL & " Line,"
    strSQL = strSQL & " MotorCapacity,"
    strSQL = strSQL & " UnitName,"
    strSQL = strSQL & " InvoiceNo,"
    strSQL = strSQL & " FixedAssetNo,"
    strSQL = strSQL & " PreventiveMaintenance,"
    strSQL = strSQL & " AccreditationID,"
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
    strSQL = strSQL & " From MachineItemsView"
    strSQL = strSQL & strSQLwhere
    Set pfLoadMachineControlStatus = GetRecordSet(strSQL)
    
End Function
'Public Function pfActiontaken(strWO As String) As String
'    Dim rs As New ADODB.Recordset
'    pfActiontaken = ""
'    Set rs = pfGetActionTaken(strWO)
'
'
'    With rs
'        .MoveFirst
'        Do While Not .EOF
'            pfActiontaken = pfActiontaken & vbCrLf & .Fields("ActionTaken").Value
'            .MoveNext
'        Loop
'    End With
'
'End Function
'Public Function pfGetActionTaken(strWO As String) As Object
'     Dim strSQL As String
'     strSQL = ""
'     strSQL = strSQL & " SELECT ActionTaken"
'     strSQL = strSQL & " FROM CostingAndHistoryView "
'     strSQL = strSQL & " WHERE WOControlNo = '" & strWO & "'"
'
'     Set pfGetActionTaken = GetRecordSet(strSQL)
'
'End Function

Public Function pfLoadHistory(dtFrom As Date, dtTo As Date, Optional strSQLwhere, Optional strStatus) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " Select DISTINCT"
    strSQL = strSQL & " WOControlNo,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " Line,"
    strSQL = strSQL & " ControlNo,"
    strSQL = strSQL & " MachineName,"
    strSQL = strSQL & " WorkCategoryName,"
    strSQL = strSQL & " MachineClassification,"
    strSQL = strSQL & " PartOfMachine,"
    strSQL = strSQL & " MachineProblemFound,"
    strSQL = strSQL & " ConditionProblem,"
    strSQL = strSQL & " ReceivedDate,"
    strSQL = strSQL & " RespondDate,"
    strSQL = strSQL & " StartedDate,"
    strSQL = strSQL & " FinishedDate,"
    strSQL = strSQL & " TotalMinutesBreakDown,"
    strSQL = strSQL & " ActionTaken = "
    strSQL = strSQL & " STUFF((SELECT DISTINCT ', ' + ActionTaken"
    strSQL = strSQL & "        FROM CostingAndHistoryView b"
    strSQL = strSQL & "        where b.wocontrolno = a.wocontrolno"
    strSQL = strSQL & "       FOR XML PATH('')), 1, 2, ''), "
    strSQL = strSQL & " ItemID,"
    strSQL = strSQL & " Description,"
    strSQL = strSQL & " Qty,"
    strSQL = strSQL & " CurrencyUnit,"
    strSQL = strSQL & " UnitPrice,"
    strSQL = strSQL & " TotalCost,"
    strSQL = strSQL & " EmployeeName = "
    strSQL = strSQL & " STUFF((SELECT DISTINCT  ' -- '  + EmployeeName  "
    strSQL = strSQL & "        FROM CostingAndHistoryView b"
    strSQL = strSQL & "        where b.wocontrolno = a.wocontrolno"
    strSQL = strSQL & "       FOR XML PATH('')), 1, 2, ''), "
    strSQL = strSQL & " Status,"
    strSQL = strSQL & " Remarks,"
    strSQL = strSQL & " ManPowerAffected,"
    strSQL = strSQL & " RespondTimeInMinute,"
    strSQL = strSQL & " TotalManHourLoss,"
    strSQL = strSQL & " TargetDate"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " CostingAndHistoryView a"
    strSQL = strSQL & " WHERE CONVERT(VARCHAR(20)," & strStatus & "Date,111) >= '" & dtFrom & "'"
    strSQL = strSQL & " AND CONVERT(VARCHAR(20)," & strStatus & "Date,111) <= '" & dtTo & "'"
    strSQL = strSQL & strSQLwhere
    strSQL = strSQL & " ORDER BY WOControlNo ASC"
    Set pfLoadHistory = GetRecordSet(strSQL)
    
End Function
Public Function pfLoadMaintenance(dtFrom As Date, dtTo As Date, Optional strSQLwhere) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " Select"
    strSQL = strSQL & " WOControlNo,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " Line,"
    strSQL = strSQL & " ControlNo,"
    strSQL = strSQL & " MachineName,"
    strSQL = strSQL & " ReceivedDate,"
    strSQL = strSQL & " FinishedDate,"
    strSQL = strSQL & " TotalMinutesBreakDown,"
    strSQL = strSQL & " ItemID,"
    strSQL = strSQL & " Description,"
    strSQL = strSQL & " Qty,"
    strSQL = strSQL & " UnitPrice,"
    strSQL = strSQL & " TotalCost,"
    'strSQL = strSQL & " TotalExpenses,"
    strSQL = strSQL & " EmployeeName,"
    strSQL = strSQL & " Status,"
    strSQL = strSQL & " Remarks"
    
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " CostingAndHistoryView "
    strSQL = strSQL & " WHERE CONVERT(VARCHAR(20),ReceivedDate,111) >= '" & dtFrom & "'"
    strSQL = strSQL & " AND CONVERT(VARCHAR(20),ReceivedDate,111) <= '" & dtTo & "'"
    strSQL = strSQL & strSQLwhere
    strSQL = strSQL & " GROUP BY"
    strSQL = strSQL & " WOControlNo,"
    strSQL = strSQL & " CompanyName,"
    strSQL = strSQL & " DepartmentName,"
    strSQL = strSQL & " SectionName,"
    strSQL = strSQL & " Line,"
    strSQL = strSQL & " ControlNo,"
    strSQL = strSQL & " MachineName,"
    strSQL = strSQL & " ReceivedDate,"
    strSQL = strSQL & " FinishedDate,"
    strSQL = strSQL & " TotalMinutesBreakDown,"
    strSQL = strSQL & " ItemID,"
    strSQL = strSQL & " Description,"
    strSQL = strSQL & " Qty,"
    strSQL = strSQL & " UnitPrice,"
    strSQL = strSQL & " TotalCost,"
    'strSQL = strSQL & " TotalExpenses,"
    strSQL = strSQL & " EmployeeName,"
    strSQL = strSQL & " Status,"
    strSQL = strSQL & " Remarks"
    strSQL = strSQL & " ORDER BY WOControlNo ASC"
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

Public Function pfLoadStatus(ByVal intType As Integer, strWhere As String) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " Select"
    strSQL = strSQL & " *"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " DailyReportView"
    strSQL = strSQL & " WHERE"
   
    strSQL = strSQL & " TypeID = " & intType
    
    
   
    'strSQL = strSQL & " AND isfinished = 0"
    strSQL = strSQL & strWhere
    
    strSQL = strSQL & " ORDER BY WorkOrderControlNo ASC"
    Set pfLoadStatus = GetRecordSet(strSQL)
End Function



Public Function pfLoadBreakdown1(ByVal d1 As Date, ByVal d2 As Date, intType As Integer) As Object
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM GetDetails('" & d1 & "','" & d2 & "'," & intType & ")"
    Set pfLoadBreakdown1 = GetRecordSet(strSQL)
End Function

Public Function pfLoadBreakdown2(ByVal d1 As Date, ByVal d2 As Date, intType As Integer) As Object
    Dim strSQL As String
    strSQL = ""
       
        strSQL = strSQL & " SELECT * FROM "
        strSQL = strSQL & " GetBreakDownDetails_try('" & d1 & "','" & d2 & "'," & intType & ")"
   
    Set pfLoadBreakdown2 = GetRecordSet(strSQL)
End Function
Public Function pfLoadBreakdown_forklift(ByVal d1 As Date, ByVal d2 As Date, strType As String, Optional ByVal d3 As Date) As Object
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
        strSQL = strSQL & " WHERE machineName LIKE '%" & strType & "%'"
        strSQL = strSQL & " GROUP BY CompanyName,CompanyID"

   
    Set pfLoadBreakdown_forklift = GetRecordSet(strSQL)
End Function

Public Function pfGetsubdetail(SearchingTable As String, strSelectingField As String, _
                                    ByVal strWhereField As String, ByVal strFieldValue As String, _
                                    Optional strWhereField2 As String, Optional strFieldValue2 As String) As String
                                    
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
    Set adoRecordset = New ADODB.Recordset
    If (strFieldValue = "" And strFieldValue2 <> "") Or (strFieldValue = "" And strFieldValue2 = "") Then pfGetsubdetail = "": Exit Function
    
    adoRecordset.Open strSQL, cn, adOpenDynamic, adLockReadOnly
    If Not adoRecordset.EOF Then pfGetsubdetail = adoRecordset.Fields(strSelectingField).Value: Exit Function
    pfGetsubdetail = ""
End Function

Public Function pfLoadPRSMonitoring(CompanyID As Integer, Server As String, Optional strSQLwhere) As Object

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
    strSQL = strSQL & " WorkOrder.AbbreviatedName + ' ' + WorkOrder.WorkOrderControlNo AS WorkOrderNo,"
    strSQL = strSQL & " PurchaseRequest.ItemId,"
    strSQL = strSQL & " NewItem.NewItemId,"
    strSQL = strSQL & " PurchaseRequest.Description,"
    strSQL = strSQL & " WorkOrder.DepartmentName,"
    strSQL = strSQL & " WorkOrder.Location_Line,"
    strSQL = strSQL & " WorkOrder.MachineItemNo,"
    strSQL = strSQL & " WorkOrder.MachineName,"
    strSQL = strSQL & " WOMDE_MachineStatus.EquipmentStatus,"
    strSQL = strSQL & " WorkOrder.ProblemFound,"
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
    
    strSQL = strSQL & " LEFT JOIN [wkn-appserver].WarehouseManagement.dbo.WorkOrderDetails WorkOrder"
    strSQL = strSQL & " ON   WorkOrder.PrsNo = PurchaseRequest.PurchaseRequestNo "
    
    strSQL = strSQL & " LEFT JOIN WOMDE_MachineStatus"
    strSQL = strSQL & " ON WOMDE_MachineStatus.MachineItemNo = WorkOrder.MachineItemNo"

    strSQL = strSQL & " LEFT JOIN"
    strSQL = strSQL & " (SELECT * FROM WOMDE_PRSAdditionalData WHERE WOMDE_PRSAdditionalData.CompanyId = " & CompanyID & ")AdditionalData"
    strSQL = strSQL & " ON AdditionalData.PurchaseRequestNo = PurchaseRequest.PurchaseRequestNo"

    strSQL = strSQL & " LEFT JOIN"
    strSQL = strSQL & " (SELECT * FROM WOMDE_PRSNewItem WHERE WOMDE_PRSNewItem.CompanyId = " & CompanyID & ")NewItem"
    strSQL = strSQL & " ON NewItem.PurchaseRequestNo = PurchaseRequest.PurchaseRequestNo"
    strSQL = strSQL & " AND NewItem.PurchaseRequestDetailSeq=PurchaseRequest.PurchaseRequestDetailSeq "
    
    strSQL = strSQL & strSQLwhere
    
    Set pfLoadPRSMonitoring = GetRecordSet(strSQL)
    
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










