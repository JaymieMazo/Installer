Attribute VB_Name = "M_Tool"
Option Explicit

Public cn As New ADODB.Connection
Public rs As New Recordset

Public strSQL As String

Public blnCancel As Boolean
Public RowCtr As Long


Public cls_GetDetails As New cls_details

Dim oSM As Object


Public Const strConnectionString As String = "Provider=SQLOLEDB.1;Password=h56r13d;Persist Security Info=True;User ID=rhrdap;Initial Catalog=WorkOrderMaintenance;Data Source=hrdsql6"
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Sub Main()
    
    On Error GoTo lnError
    '--- check if application is already open
    If App.PrevInstance = True Then MsgBox "System is already open!", vbInformation, "System": Exit Sub
    frmSplash.Show
    
    DoEvents
    
    'Call Connect
    
    Sleep 2000
    
    
    FM_Main.Show
    
    Unload frmSplash
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical
End Sub


Public Sub Connect()
    If cn.State = 0 Then
        cn.ConnectionString = strConnectionString
        cn.CursorLocation = adUseClient
        cn.Open
        cn.CommandTimeout = 0
    End If
 End Sub
 
 Public Sub Disconnect()
    cn.Close
    Set cn = Nothing
 End Sub

Public Sub CloseConnection()
   If cn.State = 1 Then cn.Close
End Sub

Public Function CommonDialogSave(cmdialog As CommonDialog, strFileName As String) As String
     On Error GoTo lnCancel
    With cmdialog
        .FileName = strFileName & Format(Now, "YYYYMMDD") & " " & Format(Time, "HH-MM-SS")
        .Filter = "Microsoft Excel (*xls)|*.xls"
        .Flags = cdlOFNHideReadOnly
        .CancelError = True
        .ShowSave
        DoEvents
    
        CommonDialogSave = .FileName
        blnCancel = False
    End With
    Exit Function
    
lnCancel:
    If Err.Number = 32755 Then
        CommonDialogSave = Left(CurDir, 3) & "DailyReport" & ".xls"
        blnCancel = True
    Else
        MsgBox "Error Number: " & Err.Number & vbCrLf & Error
    End If
End Function
 
Public Function setUpAccomplishmentHeader(ByVal d As Date, ByRef flx As MSHFlexGrid)
        Dim nLoop As Long
        Dim strDate As String
        Dim i, a As Long
        Dim counter As Integer
        
        On Error GoTo ErrorHandler
        
        strDate = d
        strDate = Mid(strDate, 1, 8)
    
        nLoop = Right(d, 2) * 2
        
        With flx
                .Visible = 1
                .Redraw = 0
                .Clear
                .Cols = nLoop + 3
                For i = 0 To (nLoop - 1) / 2
                        a = i + 1
                        .TextMatrix(0, a * 2) = strDate & i + 1
                        .TextMatrix(1, a * 2) = "ACCOMPLISHED"
                        
                        .TextMatrix(0, ((a - 1) * 2) + 1) = strDate & i + 1
                        .TextMatrix(1, ((a - 1) * 2) + 1) = "SCHEDULE"
                         
                        .MergeCells = flexMergeFree
                        .MergeRow(0) = True
                       
                       .Col = a * 2
                         .Row = 0
                      If fDayInWeek(strDate & i + 1) = strDate & i + 1 Then
                                'Debug.Print "SUNDAY ! - " & strDate & i + 1
                                .CellBackColor = &HFFFFFF
                      End If
                Next i
                For i = .Cols - 2 To .Cols - 1
                        .TextMatrix(0, i) = "TOTAL"
                        .TextMatrix(1, i) = IIf(i Mod 2, "SCHEDULE", "ACCOMPLISHED")
                Next i
                .Row = 1
                For i = 1 To .Cols - 1
                        .Col = i
                        .ColWidth(.Col) = 500
                        .CellAlignment = flexAlignCenterCenter
                        .CellBackColor = IIf(i Mod 2, &H8080&, &H8000&)
                        .CellForeColor = vbBlack
                Next i
                
                .Redraw = 1
                .Visible = 1
        End With
        
Exit Function

ErrorHandler:
        MsgBox Err.Number & " - " & Err.Description & vbNewLine & "- Call SMD-SD now.", vbCritical, "SYSTEM ERROR"
        
End Function

Public Function fDayInWeek(Optional dtmDate As Date) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    fDayInWeek = dtmDate - Weekday(dtmDate, vbUseSystemDayOfWeek) + 1
    
End Function

Public Sub subFormatGrid(flx As MSHFlexGrid, strVal As String)
    Dim intcol As Integer
    Dim lngrow As Long
    With flx
        .Clear
        .Rows = 2
        For intcol = 0 To .Cols - 1
        
         Select Case strVal
            '***************** STATUS ****************************
            Case "status"
                .TextMatrix(0, intcol) = Choose(intcol + 1, "NO.", "COMPANY NAME", "DEPARTMENT NAME", "DATE OF WORK ORDER", "WORK CATEGORY", "SECTION", _
                                                        "LINE", "PERSON IN-CHARGED / TL", "W.O. #", "EQPT. CONTROL NO.", "MACHINE NAME", _
                                                        "TYPE OF REQUEST", "SPECIFIC TROUBLE", "STATUS", "PARTS NEEDED", _
                                                        "(B) DATE OF MAKING MRS / MACHINE PARTS FOR REQUEST", "(C) DATE OF MAKING PRS", _
                                                        "PRS #", "PO #", "(D1) EXPECTED DATE DELIVERY (FROM PRS)", _
                                                        "(D2) EXPECTED DATE DELIVERY (FROM PURCHASING)", "(E) DATE OF ACTUAL RECEIVING OF ITEM", _
                                                        "(F) DATE FINISHED", "REMARKS", "", "", "")
                .ColWidth(intcol) = Choose(intcol + 1, 950, 1500, 1500, 2500, 2100, 1300, 1100, 2700, 1800, 2300, 2000, 2200, 2200, 1600, 2000, 5500, 2700, 2000, 2000, _
                                                        4100, 4800, 3800, 2400, 2000, 0, 0, 0)
        Case "status_forklift"
                .TextMatrix(0, intcol) = Choose(intcol + 1, "NO.", "COMPANY NAME", "DEPARTMENT NAME", "SECTION", "DATE OF WORK ORDER", _
                                                        "PERSON IN-CHARGED", "W.O. #", "EQPT. CONTROL NO.", "BRAND", "MODEL", "MACHINE NAME", _
                                                        "TYPE OF REQUEST", "SPECIFIC TROUBLE", "STATUS", "PARTS NEEDED", "QTY", _
                                                        "DATE OF MAKING MRS / MACHINE PARTS FOR REQUEST", "DATE OF MAKING PRS", _
                                                        "PRS #", "PO #", "EXPECTED DATE DELIVERY (FROM PRS)", _
                                                        "EXPECTED DATE DELIVERY (FROM PURCHASING)", "DATE OF ACTUAL RECEIVING OF ITEM", _
                                                        "SCHEDULE OF REPAIR", "ACTUAL REPAIR DATE", "DATE FINISHED", "REMARKS")
                .ColWidth(intcol) = Choose(intcol + 1, 500, 3000, 1500, 1300, 2500, 2100, 1100, 2700, 1800, 2300, 2100, 2100, 2000, 2200, 3900, 700, 2800, 1600, 1000, 1000, 2000, 2400, _
                                                        2100, 2800, 2800, 2000, 2400)
                   
            '***************** BREAKDOWN ************************
            Case "breakdown"
                .TextMatrix(0, intcol) = Choose(intcol + 1, "RECEIVED DATE", "DEPARTMENT NAME", "SECTION NAME", "RECEIVED", _
                                                        "FINISHED")
                .ColWidth(intcol) = Choose(intcol + 1, 1600, 2100, 2100, 1100, 1100)
                
            '***************** SUMMARY **************************
            Case "summary"
                .TextMatrix(0, intcol) = Choose(intcol + 1, "COMPANY", "DEPARTMENT", "BACKLOG", "RECEIVED", "FINISHED", _
                                                        "FINISHED WO FROM PENDING WO", "FINISHED ON THE SUCCEEDING MONTH", "CANCELED", _
                                                        "TURNOVER", "WAITING PARTS", "FOR SCHEDULE", "FOR CONFIRMATION/ONGOING", _
                                                        "TOTAL UNFINISHED " & vbCrLf & "(" & Format(frmBreakDown.dtFrom, "mmmm dd,yyyy") & " - " & Format(frmBreakDown.dtTo, "mmmm dd,yyyy") & ")", _
                                                        "TOTAL UNFINISHED" & vbCrLf & "BACKLOGS FROM PREVIOUS", "TOTAL UNFINISHED")
                .ColWidth(intcol) = Choose(intcol + 1, 2100, 2100, 1100, 1100, 1100, 3500, 3800, 1100, 1100, 1600, 1600, 2400, 2400, 2400, 2000)
                
          Case "summary_forklift"
                .TextMatrix(0, intcol) = Choose(intcol + 1, "TYPE", "COMPANYID", "COMPANY", "FOR SCHEDULE", "WAITING PARTS", "ON GOING", "FINISHED REPAIR", "NEW BREAKDOWN UNIT", "", "", "", "", "", "", "", "")
                .ColWidth(intcol) = Choose(intcol + 1, 3000, 0, 3000, 1100, 1100, 1100, 1100, 2100, 0, 0, 0, 0, 0, 0, 0, 0)
                
            '***************** HISTORY ****************************
            Case "history"
                .TextMatrix(0, intcol) = Choose(intcol + 1, "", "RECEIVED DATE", "COMPANY NAME", "DEPARTMENT NAME", "CATEGORY", _
                                                        "WORK CATEGORY NAME", "WORKORDER CONTROL NO.", "MACHINE ITEM NO.", "STATUS")
                .ColWidth(intcol) = Choose(intcol + 1, 350, 1600, 2450, 2000, 2600, 2000, 2600, 2200, 1600)
                
            '***************** COSTING AND HISTORY **************
            Case "costing"
                .Cols = 31
                .TextMatrix(0, intcol) = Choose(intcol + 1, "WO CONTROL NO.", "COMPANY", "DEPARTMENT", "SECTION", "LINE", "CONTROL NO.", _
                                                        "EQUIPMENT NAME", "WORK CATEGORY", "MACHINE CLASSIFICATION", "PART OF MACHINE", _
                                                        "MACHINE PROBLEM FOUND", "CONDITION/PROBLEM", "RECEIVED", "RESPOND", "STARTED", "FINISHED", _
                                                        "ACTION TAKEN", "ITEMCODE", "MATERIAL NAME", "QTY", "CURRENCY UNIT", "UNIT COST", _
                                                        "TOTAL COST", "PREPARED BY", "STATUS", "REMARKS", _
                                                        "# OF MANPOWER AFFECTED OF BREAKDOWN", "TOTAL MINUTES OF BREAKDOWN (DOWNTIME)", _
                                                        "TOTAL MANHOUR LOSS (BREAKDOWN)", "TOTAL MINUTES OF REPAIR", "TARGET DATE/TIME")
                .ColWidth(intcol) = Choose(intcol + 1, 2000, 2450, 2300, 2400, 1700, 1700, _
                                                        3600, 2400, 3000, 2500, _
                                                        2600, 2600, 2000, 2000, 2000, 2000, _
                                                        2300, 2700, 2000, 3100, 2000, 2000, _
                                                        2000, 0, 3000, 2500, 2500, _
                                                        3000, 3000, _
                                                        3000, 3000, 3100)
                                                        
                Case "machinecontrol"
                    .TextMatrix(0, intcol) = Choose(intcol + 1, "STATUS", "MACHINE ITEM NO", "MACHINE NAME", "COMPANY", "DEPARTMENT", _
                                                        "SECTION", "MAKER", "TYPENAME", "LOCATION", "LINE", "MOTOR CAPACITY", _
                                                        "FIXED ASSET NO", "PREVENTIVE MAINTENANCE", "ENGINE MODEL", "ENGINE SERIAL NO", _
                                                        "TRANSMISSION", "MAST TYPE", "ATTACHMENT TYPE", "FRONT TIRE", "FRONT TIRE HOLES", _
                                                        "ACQUISITION AMOUNT", "ACQUISITION DATE", "DISPOSAL DATE", "REMARKS")
                    
                    .ColWidth(intcol) = Choose(intcol + 1, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, 1800, _
                                                         1800, 1800, 1800, 1800, 1800, 1800, 1800)
                                                         
                ' ********************* EMPLOYEE MASTERLIST ***************************
                Case "employee"
                    .TextMatrix(0, intcol) = Choose(intcol + 1, "COMPANY", "ID", "NAME", "REGISTERED DATE", "UPDATED DATE", _
                                                                        "DELETED DATE")
                                                   
                    .ColWidth(intcol) = Choose(intcol + 1, 2700, 900, 3200, 1300, 1300, 1300)
                    
                    
                ' ********************* WORKORDERITEMS MASTERLIST ***************************
                Case "WorkOrderItems"
                    .TextMatrix(0, intcol) = Choose(intcol + 1, _
                                                                    "ItemCode", "TypeID", _
                                                                    "ItemName", "Company", _
                                                                    "Department", "Section", _
                                                                    "Location", "Line", "PriorityLevel", "MakerName", _
                                                                    "Model", "SerialNo", "Capacity", _
                                                                    "FixedAssetNo", "PreventiveMaintenance", _
                                                                    "EngineModel", "EngineSerialNo", _
                                                                    "Transmission", "MastType", _
                                                                    "MastHeight", "AttachmentType", _
                                                                    "FrontTire", "FrontTireHoles", _
                                                                    "RearTire", "RearTireHoles", "Status", _
                                                                    "AcquisitionAmount", "AcquisitionDate", "DisposalDate", "Remarks")
                                                   
                    .ColWidth(intcol) = Choose(intcol + 1, 1300, 1300, 2000, 3500, 1300, 1300, _
                                                                        1300, 1300, 1300, 1300, 1300, 1300, _
                                                                        1300, 1300, 1300, 1300, 1300, 1300, _
                                                                        1300, 1300, 1300, 1300, 1300, 1300, _
                                                                        1300, 1300, 1300, 1300, 1300, 1300)
                                                    
                     '***************** MAINTENANCE ****************************
                        Case "maintenance"
                            .TextMatrix(0, intcol) = Choose(intcol + 1, "WO CONTROL NO", "COMPANY", "DEPARTMENT", "SECTION", "LINE", _
                                                                    "MACHINE CONTROL NO", "EQUIPMENT NAME", "RECEIVE DATE", "REQUEST DATE", _
                                                                    "FINISHED DATE", "RESPOND TIME IN MINUTES", _
                                                                    "ITEMCODE", "MATERIAL NAME", "QTY", "UNIT COST", "TOTAL COST", "PREPARED BY", "STATUS", "REMARKS")
                            .ColWidth(intcol) = Choose(intcol + 1, 2000, 3500, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 2000, 3000)
                    
                    '***************** MONITORING ****************************
                     Case "prsmonitoring"
                            .TextMatrix(0, intcol) = Choose(intcol + 1, "TYPE OF PRS", "RS NO.", "TAG NO.", "W.O #", "ITEM CODE", "NEW ITEM CODE", _
                                                                    "ITEM NAME AND DESCRIPTION", "DEPARTMENT", "LOCATION / LINE", _
                                                                    "MACHINE CONTROL NO.", "MACHINE NAME", _
                                                                    "EQUIPMENT STATUS", "PURPOSE", "DATE REQUESTED", "RECEIVED BY PURCHASING", "DATE EXPECTED", _
                                                                    "DATE RECEIVED", "QTY REQUESTED", "QTY RECEIVED", "UNIT", _
                                                                    "PO NO.", "ETD ON PO", "INVOICE NO.", "INCHARGE", "RECEIVED BY", "STATUS", "REMARKS", "DIVISION", "FINAL", "SEQ")
                            .ColWidth(intcol) = Choose(intcol + 1, 1300, 1300, 1300, 1500, 1300, 1300, 3000, 1300, 3000, 1300, 3000, 1300, 3000, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 1300, 2000, 1300, 0, 0, 0)
                                
                    '***************** PENDING WORK ORDER ****************************
                     Case "pending"
                            .TextMatrix(0, intcol) = Choose(intcol + 1, "REQUEST DATE", "RECEIVED DATE", "CONTROL NO.", "DEPARTMENT", "SECTION", "LOCATION", _
                                                                        "MACHINE NO.", "MACHINE NAME", "PROBLEM", "REQUESTOR", "ALTERNATE CONTACT PERSON", "TEAM LEADER", _
                                                                        "STARTED DATE", "FINISHED DATE", "STATUS")
                            .ColWidth(intcol) = Choose(intcol + 1, 1300, 1300, 1300, 3000, 3000, 3000, 3000, 3000, 8000, 3000, 3000, 3000, 1300, 1300, 1300)

            End Select
            .RowHeight(0) = 500
            .Col = intcol
            .Row = 0
            .WordWrap = True
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
        Next intcol
    End With
        
End Sub

Function pfvarNoValue(ByVal varNovalue As Variant, _
              Optional ByVal blnStringNovalue As Boolean = True) _
                            As Variant
    If IsNull(varNovalue) Then
        If blnStringNovalue Then
            pfvarNoValue = ""
        Else
            pfvarNoValue = "Null"
        End If
    Else
        pfvarNoValue = varNovalue
    End If
End Function

Function LoadDataToCombo(combos As Object, strTable As String, Optional nVal As String, Optional nVal2 As Integer, _
                                    Optional strFrmName As String, Optional blnStatus As Boolean, _
                                    Optional blnIsForkLift As Boolean)
    Dim rs As New ADODB.Recordset
    Dim lngRecCnt As Long
    Dim i, ii As Long
    Dim lngIndex As Long
  
     Set rs = New ADODB.Recordset
     
     strSQL = "SELECT * FROM " & strTable
     
     If strTable = "Departments" Then
        strSQL = strSQL & " WHERE CompanyID = '" & nVal & "'"
        strSQL = strSQL & " AND DeletedDate IS NULL"
        strSQL = strSQL & " ORDER BY DepartmentName ASC"
     End If
     
     If strTable = "Sections" Then
        strSQL = strSQL & " WHERE DepartmentID = '" & nVal & "'"
        strSQL = strSQL & " AND DeletedDate IS NULL"
        strSQL = strSQL & " ORDER BY SectionName ASC"
     End If
     
     If strTable = "MainCategories" Then
        strSQL = strSQL & " WHERE CompanyID = '" & nVal & "'"
        strSQL = strSQL & " AND DeletedDate IS NULL"
        strSQL = strSQL & " ORDER BY MainCategoryName ASC"
     ElseIf blnIsForkLift = True Then
        strSQL = "SELECT DISTINCT MainCategoryID,MainCategoryName FROM " & strTable
     End If
     
     If strTable = "MainSubCategories" Then
        strSQL = strSQL & " WHERE CompanyID = '" & nVal & "'"
        strSQL = strSQL & " AND MainCategoryID = " & nVal2
        strSQL = strSQL & " AND DeletedDate IS NULL"
     End If
     
     If strTable = "AbbreviatedTypes" Then
        strSQL = strSQL & " WHERE CompanyID = '" & nVal & "'"
        strSQL = strSQL & " AND TypeID = " & nVal2
        strSQL = strSQL & " AND DeletedDate IS NULL"
     End If
     
     If strTable = "AbbreviatedMachines" Then
        strSQL = "SELECT DISTINCT AbbreviatedName FROM " & strTable
     End If
     
     If strTable = "Status" And blnStatus = True Then
        strSQL = strSQL & " Where StatusID <> 3 "
     End If
     
     If strTable = "Types" Then
        strSQL = strSQL & " WHERE CompanyID = '" & nVal & "'"
        strSQL = strSQL & " AND DeletedDate IS NULL"
        
        If strFrmName = "frmBreakDown" Or blnIsForkLift = True Then
            strSQL = strSQL & " AND TypeName NOT IN (SELECT TypeName FROM Types WHERE TypeName LIKE '%FORKLIFT%')"
        End If
        
        strSQL = strSQL & " ORDER BY TypeName ASC"
     End If
     
     rs.Open strSQL, cn, adOpenDynamic, adLockReadOnly
     If rs.EOF Then Exit Function
     With rs
        .MoveLast
        lngRecCnt = .RecordCount
        .MoveFirst
        combos.Visible = False
        combos.Clear
        For i = 1 To lngRecCnt
            Select Case strTable
                Case "Companies"
                    combos.AddItem .Fields("CompanyID").Value
                    combos.Column(1, lngIndex) = .Fields("CompanyName").Value
                Case "Departments"
                    combos.AddItem .Fields("DepartmentID").Value
                    combos.Column(1, lngIndex) = .Fields("DepartmentName").Value
                Case "Types"
                    combos.AddItem .Fields("TypeID").Value
                    combos.Column(1, lngIndex) = .Fields("TypeName").Value
                Case "Status"
                    combos.AddItem .Fields("StatusID").Value
                    combos.Column(1, lngIndex) = .Fields("Status").Value
                Case "Sections"
                    combos.AddItem .Fields("SectionID").Value
                    combos.Column(1, lngIndex) = .Fields("SectionName").Value
                Case "MainCategories"
                    combos.AddItem .Fields("MainCategoryID").Value
                    combos.Column(1, lngIndex) = .Fields("MainCategoryName").Value
                Case "MainSubCategories"
                    combos.AddItem .Fields("MainSubCategoryID").Value
                    combos.Column(1, lngIndex) = .Fields("MainSubCategoryName").Value
                Case "AbbreviatedTypes"
                    combos.AddItem .Fields("TypeID").Value
                    combos.Column(1, lngIndex) = .Fields("AbbreviatedName").Value
                Case "AbbreviatedMachines"
                    combos.AddItem .Fields("AbbreviatedName").Value
                    combos.Column(1, lngIndex) = .Fields("AbbreviatedName").Value
            End Select
            lngIndex = lngIndex + 1
            .MoveNext
        Next i
        If blnStatus = True Then GoTo ext
       
ext:
    combos.Visible = True
     End With
     
End Function

Public Function GetRecordSet(ByVal strQuery As String) As Object
    Dim adoRecordset As ADODB.Recordset
    
    Set adoRecordset = New ADODB.Recordset
    On Error GoTo lnError
    If adoRecordset.State = 1 Then adoRecordset.Close
    
    DoEvents
    adoRecordset.Open strQuery, cn, adOpenDynamic, adLockReadOnly
    
    Set GetRecordSet = adoRecordset
    
    Exit Function
lnError:
     MsgBox strQuery & vbCrLf & vbCrLf & Err.Number & "-" & Err.Description, vbCritical, "System Error"
End Function

Public Function GetPreventiveMaintenance(strNumber As String) As String
If strNumber = "1" Then
    GetPreventiveMaintenance = "With"
ElseIf strNumber = "2" Then
    GetPreventiveMaintenance = "Without"
Else
    GetPreventiveMaintenance = "-"
End If
End Function

Public Function GetServerName(strCompanyID As String) As String
Dim strCompanyName As String
    Select Case strCompanyID
        Case "001"
            strCompanyName = "a-sv17"
        Case "002"
            strCompanyName = "jd-004"
        Case "003"
            strCompanyName = "impex-sv4"
        Case Else
            strCompanyName = "wkn-appserver"
    End Select
    GetServerName = strCompanyName
End Function

Public Function exportLibre(flxDetail As MSHFlexGrid, strTitle As String, lblMessage As Label) As Boolean
     Dim oDoc As Object, _
        oDesk As Object, _
        oSheet As Object, _
        oPar(1) As Object, _
        CellProp As Object, _
        CellStyle As Object, _
        NewStyle As Object, _
        oRange As Object, _
        oColumns As Object, _
        PageStyles As Object, _
        NewPageStyle As Object, _
        StdPage As Object, _
        basicBorder As Object, _
        oBorder As Object
    Dim oCol As Long, oRow As Long
    Dim Charts As Object
    Dim Chart As Object
    Dim Rect As Object
    Dim oChartRange As Object
    Dim RangeAddress(0) As Object
    
    exportLibre = False
    oRow = 0
    
    Set oSM = CreateObject("com.sun.star.ServiceManager")
    Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
    Set oPar(0) = MakePropertyValue("Hidden", True)
    Set oPar(1) = MakePropertyValue("Overwrite", True)
    Set oDoc = oDesk.LoadComponentFromURL("private:factory/scalc", "_blank", 0, oPar)
    Set oSheet = oDoc.Sheets.getByIndex("0")
    Set CellProp = oDoc.StyleFamilies.getByName("CellStyles")
    Set NewStyle = oDoc.createInstance("com.sun.star.style.CellStyle")
    Call CellProp.InsertbyName("MyStyle", NewStyle)
    NewStyle.ParentStyle = "Default"
    Set CellStyle = CellProp.getByName("MyStyle")
    Set PageStyles = oDoc.StyleFamilies.getByName("PageStyles")
    Set StdPage = PageStyles.getByName("Default")
    Set basicBorder = oDoc.Bridge_GetStruct("com.sun.star.table.BorderLine")
    
    basicBorder.Color = RGB(0, 0, 0)
    basicBorder.InnerLineWidth = 0
    basicBorder.OuterLineWidth = 11
    basicBorder.LineDistance = 0
    
'        With StdPage
'            .FooterIsOn = False
'            .HeaderIsOn = False
'            .IsLandscape = False
'            .Width = 29700
'            .Height = 21000
'            .LeftMargin = 1000
'            .RightMargin = 1000
'            .TopMargin = 1000
'            .BottomMargin = 1000
'        End With
        '-Header--
        
            FM_Main.Enabled = False
            flxDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        oSheet.getCellByPosition(0, oRow).String = strTitle
         
                
        For oCol = 0 To flxDetail.Cols - 1
           With oSheet
                With CellStyle
                    .CharWeight = 300
                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .CellBackColor = RGB(255, 184, 0)
                End With
                .getCellByPosition(oCol, 1).CellStyle = "MyStyle"
                .getCellByPosition(oCol, 1).String = flxDetail.TextMatrix(0, oCol)
                .getCellByPosition(oCol, 1).HoriJustify = 2
           End With
            Set oBorder = oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder
            oBorder.LeftLine = basicBorder
            oBorder.Topline = basicBorder
            oBorder.RightLine = basicBorder
            oBorder.BottomLine = basicBorder
            oBorder.VerticalLine = basicBorder
            oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder = oBorder
        Next oCol
    
        '-Content--
        For oRow = 0 To flxDetail.Rows - 1
            DoEvents
            
            lblMessage.Caption = "Exporting Data : " & oRow & " rows out of " & flxDetail.Rows - 1 & " rows"
            For oCol = 0 To flxDetail.Cols - 1
               With oSheet
                    .getCellByPosition(oCol, oRow + 1).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .getCellByPosition(oCol, oRow + 1).HoriJustify = 2
                    flxDetail.Row = oRow
                    flxDetail.Col = oCol
                    .getCellByPosition(oCol, oRow + 1).String = flxDetail.TextMatrix(oRow, oCol)
                    If flxDetail.CellBackColor = &H8080FF Then
                        .getCellByPosition(oCol, oRow + 1).CellBackColor = RGB(128, 128, 255)
                    End If
               End With
            Next oCol
                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder
                oBorder.LeftLine = basicBorder
                oBorder.Topline = basicBorder
                oBorder.RightLine = basicBorder
                oBorder.BottomLine = basicBorder
                oBorder.VerticalLine = basicBorder
                oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder = oBorder
        Next oRow
     
        
        
        Call oDoc.storeToURL("file:///C:/Exported.xls", oPar)
        Set oPar(0) = MakePropertyValue("Hidden", False)
        Set oDoc = oDesk.LoadComponentFromURL("file:///C:/Exported.xls", "_blank", 0, oPar)
        exportLibre = True
    
    
        Set oSM = Nothing
        Set oDesk = Nothing
        Set oDoc = Nothing
        Set oSheet = Nothing
        Set oPar(1) = Nothing
        Set CellProp = Nothing
        Set CellStyle = Nothing
        Set NewStyle = Nothing
        Set oRange = Nothing
        Set oColumns = Nothing
        Set PageStyles = Nothing
        Set NewPageStyle = Nothing
        Set StdPage = Nothing
        exportLibre = True
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        Exit Function
ErrLibre:
        exportLibre = False
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
End Function



Private Function MakePropertyValue(propName, propVal) As Object
    Dim oPropValue As Object
    Set oPropValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    oPropValue.Name = propName
    oPropValue.Value = propVal
    Set MakePropertyValue = oPropValue
End Function
