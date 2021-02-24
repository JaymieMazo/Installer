VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_OnhandSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On Hand Summary"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "F_OnhandSummary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   6315
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Top             =   210
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   23789569
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdOnHand 
      Height          =   375
      Index           =   0
      Left            =   3660
      TabIndex        =   4
      Top             =   1050
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&EXCEL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "F_OnhandSummary.frx":0CCA
      PICN            =   "F_OnhandSummary.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOnHand 
      Height          =   375
      Index           =   1
      Left            =   4860
      TabIndex        =   5
      Top             =   1050
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&CLOSE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "F_OnhandSummary.frx":1280
      PICN            =   "F_OnhandSummary.frx":129C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   3510
      TabIndex        =   11
      Top             =   690
      Width           =   195
   End
   Begin MSForms.TextBox txtItemdIdUntil 
      Height          =   315
      Left            =   3780
      TabIndex        =   2
      Top             =   630
      Width           =   2175
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3836;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   210
      Left            =   420
      TabIndex        =   10
      Top             =   660
      Width           =   600
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ITEM ID"
      Size            =   "1058;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemId 
      Height          =   315
      Left            =   1245
      TabIndex        =   1
      Top             =   630
      Width           =   2175
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3836;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   1245
      TabIndex        =   3
      Top             =   1050
      Width           =   2175
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "3836;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "MS UI Gothic"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   128
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label13 
      Height          =   210
      Left            =   390
      TabIndex        =   9
      Top             =   1080
      Width           =   705
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "DIVISION"
      Size            =   "1244;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Left            =   390
      TabIndex        =   6
      Top             =   270
      Width           =   780
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "END DATE"
      Size            =   "1376;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   1485
      Left            =   120
      TabIndex        =   7
      Top             =   90
      Width           =   6030
      BackColor       =   8421504
      Size            =   "10636;2619"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1485
      Left            =   180
      TabIndex        =   8
      Top             =   180
      Width           =   6060
      BackColor       =   0
      Size            =   "10689;2619"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_OnhandSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboDivision_KeyPress(KeyAscii As MSForms.ReturnInteger)
     If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmdOnHand_Click(Index As Integer)
     Select Case Index
          Case 0
               Screen.MousePointer = vbHourglass
               Call psubShowStatMsg("Please wait... creating onhand excel list.")
               Call GetOnHandSummary
               Call psubHideStatMsg
               Screen.MousePointer = vbDefault
          Case 1: Unload Me
     End Select
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
     dtpEnd.Value = clsDB.ServerDate
     Call clsPrintMenu.psubLoadDivision(cboDivision)
End Sub

Private Sub txtItemdIdUntil_GotFocus()
     With txtItemdIdUntil
          .SelStart = 0
          .SelLength = Len(.Text)
     End With
End Sub

Private Sub txtItemdIdUntil_LostFocus()
     If txtItemId <> "" And txtItemdIdUntil.Text = "" Then txtItemdIdUntil.Text = txtItemId.Text
End Sub

Private Sub txtItemID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then txtItemdIdUntil.SetFocus
End Sub

Private Sub txtItemId_LostFocus()
     If txtItemdIdUntil.Text = "" Then txtItemdIdUntil.Text = txtItemId.Text
End Sub

'---Prepare excel summary of On Hand Qty
Private Sub GetOnHandSummary()
    Dim objOnHand As Object
    Dim objOnHandNull As Object
    Dim intRow As Integer, intCol As Integer, lngRecord As Long
    Dim blnWithWaiting As Boolean
    
On Error GoTo lnErrHandler
    
    Set objOnHand = clsDB.GetRecordSet(fstrSQLOnHand)
    
    lngRecord = objOnHand.RecordCount
     
    If lngRecord = 0 Then MsgBox "No record found.", vbInformation, "System Message": Exit Sub
    
    With clsPrintMenu.Utility
          Call .OpenExcel
          objOnHand.MoveFirst
          .ExcelWkSheet.Cells(1, 1) = _
                    "On Hand Qty as of  " & Format(dtpEnd.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:mm:ss am/pm")
    
          For intCol = 1 To 10
                .ExcelWkSheet.Cells(2, intCol) = _
                                        Choose(intCol, "ItemId", "Description", "Division", "Disuse", _
                                        "On Hand Qty", "Qty Unit", "Location", "IQCQty", "QtyExpected", "Waiting Qty")
          Next
          blnWithWaiting = False
          Call .SetCellColor(2, 1, 2, 10)
          '--- waiting qty is optional
          If MsgBox("Do you want to include waiting qty?", vbQuestion + vbYesNo) = vbYes Then
                If MsgBox("This will take a few minutes to load the records." & Chr(13) _
                     & "Do you want to continue?", vbQuestion + vbOKCancel) = vbOK Then
                     blnWithWaiting = True
                End If
          End If
          For intRow = 3 To lngRecord + 2
               For intCol = 1 To 10
                    .ExcelWkSheet.Cells(intRow, intCol) = Choose(intCol, "'" & objOnHand.Fields("ItemId").Value, _
                                             objOnHand.Fields("Description").Value, objOnHand.Fields("Division").Value, _
                                             objOnHand.Fields("Disuse").Value, objOnHand.Fields("Qty").Value, _
                                             objOnHand.Fields("QtyUnit").Value, objOnHand.Fields("Location").Value, IIf(IsNull(objOnHand.Fields("TQty").Value), 0, objOnHand.Fields("TQty").Value), _
                                             IIf(IsNull(objOnHand.Fields("QtyExpected").Value), 0, objOnHand.Fields("QtyExpected").Value))
                    If intCol = 5 Then
                         Select Case objOnHand.Fields("Qty").Value
                              Case Is < 0: Call .SetCellColor(intRow, intCol, intRow, intCol, 35)
                              Case Is = 0: Call .SetCellColor(intRow, intCol, intRow, intCol, 28)
                         End Select
                    End If
                    If blnWithWaiting Then
                         '--- write waiting qty
                         If intCol = 10 Then _
                                   .ExcelWkSheet.Cells(intRow, 10) = clsPrintMenu.InventoryCheck.GetWaitingQty( _
                                                                    objOnHand.Fields("ItemId").Value, cboDivision) + clsPrintMenu.InventoryCheck.GetExcessWaitingQty(objOnHand.Fields("ItemId").Value, cboDivision)
                    End If
               Next
               Call psubShowStatMsg(intRow - 2 & "/" & lngRecord, 4)
               objOnHand.MoveNext
          Next
          Call psubHideStatMsg(4)
          .ExcelWkSheet.Columns.AutoFit
          
          .ExcelApp.Visible = True
        
          .CloseExcel
    
        Set objOnHand = Nothing
    
        Exit Sub
    
lnErrHandler:
        .CloseExcel
End With
     MsgBox Err.Number & Err.Description, vbInformation, "System Error"
End Sub

Private Function fstrSQLOnHand() As String



'  fstrSQLOnHand = " SELECT " _
'                   & "       Q.ItemId, Q.Description, Q.Qty, " _
'                   & "       (Select QtyUnit from QtyUnits Where QtyUnitId = Q.QtyUnit) as QtyUnit, " _
'                   & "       Q.Division , Q.Location, Q.Disuse," _
'                   & "       (Select SUM(POInvoiceSearchViewNull.QtyReceived) from POInvoiceSearchViewNull" _
'                   & "       WHERE POInvoiceSearchViewNull.ItemId=Q.ItemId) as TQty," _'
'                   & "        (SELECT SUM(InvoiceDetailsView.Qty) FROM InvoiceDetailsView" _
'                   & "       WHERE InvoiceDetailsView.ItemId=Q.ItemId AND InvoiceDetailsView.QtyReceived is Null AND InvoiceDetailsView.QtyOk is null) as QtyExpected" _
'                   & " FROM " _
'                   & "    (" _
'                   & "        SELECT " _
'                   & "              Items.ItemId, Items.Description, " _
'                   & "              CASE WHEN GetStockBalance.Qty Is Null THEN 0 ELSE GetStockBalance.Qty END Qty, " _
'                   & "              CASE WHEN Items.McQtyUnitId IS NULL THEN SupplierQtyUnitId ELSE Items.McQtyUnitId END QtyUnit, " _
'                   & "              (Select Description from Divisions Where DivisionId = DivisionItems.DivisionId) as Division, " _
'                   & "              (Select Location from Locations Where LocationId = Items.DefaultLocationId) as Location, " _
'                   & "              Items.Disuse " _
'                   & "        FROM  DivisionItems LEFT OUTER JOIN " _
'                   & "              GetStockBalance('" & dtpEnd.Value & "') GetStockBalance ON " _
'                   & "              DivisionItems.DivisionId = GetStockBalance.DivisionId AND " _
'                   & "              DivisionItems.ItemId = GetStockBalance.ItemId RIGHT OUTER JOIN " _
'                   & "              Items ON DivisionItems.ItemId = Items.ItemId " _
'                   & "    ) Q LEFT JOIN  QtyUnits On Q.QtyUnit = QtyUnits.QtyUnitId " _
'                   & " WHERE Q.Division = '" & cboDivision.Text & "'"
'



  fstrSQLOnHand = " SELECT "
fstrSQLOnHand = fstrSQLOnHand & "        Q.ItemId, Q.Description, Q.Qty, "
fstrSQLOnHand = fstrSQLOnHand & "        (Select QtyUnit from QtyUnits Where QtyUnitId = Q.QtyUnit) as QtyUnit, "
fstrSQLOnHand = fstrSQLOnHand & "        Q.Division , Q.Location, Q.Disuse,"
fstrSQLOnHand = fstrSQLOnHand & "       (Select SUM(DeliveryDetails.QtyReceived) from DeliveryDetails"
fstrSQLOnHand = fstrSQLOnHand & "      INNER JOIN InvoiceDetails ON InvoiceDetails.Invoiceno = DeliveryDetails.InvoiceNo"
fstrSQLOnHand = fstrSQLOnHand & "      AND  InvoiceDetails.SupplierId = DeliveryDetails.SupplierId"
fstrSQLOnHand = fstrSQLOnHand & "     AND  InvoiceDetails.InvoiceDetailSeq = DeliveryDetails.InvoiceDetailSeq"
fstrSQLOnHand = fstrSQLOnHand & "     Right Outer Join PoDetails INNER JOIN"
fstrSQLOnHand = fstrSQLOnHand & "      PoHeaders ON dbo.PoDetails.PoNo = PoHeaders.PoNo INNER JOIN"
fstrSQLOnHand = fstrSQLOnHand & "  Suppliers ON dbo.PoHeaders.SupplierId = Suppliers.SupplierId ON"
fstrSQLOnHand = fstrSQLOnHand & "  InvoiceDetails.PoNo = dbo.PoDetails.PoNo AND"
fstrSQLOnHand = fstrSQLOnHand & " InvoiceDetails.PoDetailSeq = PoDetails.PoDetailSeq AND"
fstrSQLOnHand = fstrSQLOnHand & " InvoiceDetails.SupplierId = PoHeaders.SupplierId"
fstrSQLOnHand = fstrSQLOnHand & "  Where (dbo.PoDetails.Canceled = 0) And (PoHeaders.Canceled = 0) And DeliveryDetails.QtyReceived Is Not Null And DeliveryDetails.QtyOk Is Null"
fstrSQLOnHand = fstrSQLOnHand & " AND  PoDetails.ItemId=Q.ItemId) as TQty,"
fstrSQLOnHand = fstrSQLOnHand & "        (SELECT SUM(InvoiceDetailsView.Qty) FROM InvoiceDetailsView"
fstrSQLOnHand = fstrSQLOnHand & "       WHERE InvoiceDetailsView.ItemId=Q.ItemId AND InvoiceDetailsView.QtyReceived is Null AND InvoiceDetailsView.QtyOk is null) as QtyExpected"
fstrSQLOnHand = fstrSQLOnHand & " FROM "
fstrSQLOnHand = fstrSQLOnHand & "    ("
fstrSQLOnHand = fstrSQLOnHand & "        SELECT "
fstrSQLOnHand = fstrSQLOnHand & "              Items.ItemId, Items.Description, "
fstrSQLOnHand = fstrSQLOnHand & "              CASE WHEN GetStockBalance.Qty Is Null THEN 0 ELSE GetStockBalance.Qty END Qty, "
fstrSQLOnHand = fstrSQLOnHand & "              CASE WHEN Items.McQtyUnitId IS NULL THEN SupplierQtyUnitId ELSE Items.McQtyUnitId END QtyUnit, "
fstrSQLOnHand = fstrSQLOnHand & "              (Select Description from Divisions Where DivisionId = DivisionItems.DivisionId) as Division, "
fstrSQLOnHand = fstrSQLOnHand & "              (Select Location from Locations Where LocationId = Items.DefaultLocationId) as Location, "
fstrSQLOnHand = fstrSQLOnHand & "              Items.Disuse "
fstrSQLOnHand = fstrSQLOnHand & "        FROM  DivisionItems LEFT OUTER JOIN "
fstrSQLOnHand = fstrSQLOnHand & "              GetStockBalance('" & dtpEnd.Value & "') GetStockBalance ON "
fstrSQLOnHand = fstrSQLOnHand & "              DivisionItems.DivisionId = GetStockBalance.DivisionId AND "
fstrSQLOnHand = fstrSQLOnHand & "              DivisionItems.ItemId = GetStockBalance.ItemId RIGHT OUTER JOIN "
fstrSQLOnHand = fstrSQLOnHand & "              Items ON DivisionItems.ItemId = Items.ItemId "
fstrSQLOnHand = fstrSQLOnHand & "    ) Q LEFT JOIN  QtyUnits On Q.QtyUnit = QtyUnits.QtyUnitId "
fstrSQLOnHand = fstrSQLOnHand & " WHERE Q.Division = '" & cboDivision.Text & "'"


  
          
    If txtItemId <> "" And txtItemdIdUntil <> "" Then _
          fstrSQLOnHand = fstrSQLOnHand _
                   & " And Q.ItemId Between " & pfstrQt(txtItemId.Text) _
                                        & " And " & pfstrQt(txtItemdIdUntil.Text)
    fstrSQLOnHand = fstrSQLOnHand & " ORDER BY Q.ItemId"
End Function


