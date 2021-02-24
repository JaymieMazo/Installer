VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_InvRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   Icon            =   "F_InvRecord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10545
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   6435
      TabIndex        =   4
      Top             =   1245
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&VIEW"
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
      MICON           =   "F_InvRecord.frx":0CCA
      PICN            =   "F_InvRecord.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   6945
      TabIndex        =   2
      Top             =   765
      Width           =   1320
      _ExtentX        =   2328
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
      Format          =   112721921
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   8520
      TabIndex        =   3
      Top             =   765
      Width           =   1320
      _ExtentX        =   2328
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
      Format          =   112721921
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdPrint 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   1245
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&PRINT"
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
      MICON           =   "F_InvRecord.frx":1282
      PICN            =   "F_InvRecord.frx":129E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   375
      Left            =   9015
      TabIndex        =   6
      Top             =   1245
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
      MICON           =   "F_InvRecord.frx":183A
      PICN            =   "F_InvRecord.frx":1856
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxInvRecord 
      Height          =   5520
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9737
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   16744576
      BackColorBkg    =   11049333
      GridColor       =   -2147483633
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.Label Label3 
      Height          =   5535
      Left            =   225
      TabIndex        =   16
      Top             =   2145
      Width           =   10215
      BackColor       =   0
      Size            =   "18018;9763"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   3885
      TabIndex        =   1
      Top             =   780
      Width           =   2325
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "4101;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   270
      Left            =   3090
      TabIndex        =   14
      Top             =   810
      Width           =   960
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DIVISION"
      Size            =   "1693;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblDescription 
      Height          =   525
      Left            =   1380
      TabIndex        =   13
      Top             =   1170
      Width           =   4830
      ForeColor       =   16777215
      BackColor       =   -2147483638
      VariousPropertyBits=   8388627
      Size            =   "8520;926"
      BorderColor     =   -2147483641
      SpecialEffect   =   6
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   -15
      TabIndex        =   12
      Top             =   105
      Width           =   10575
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Inventory Record"
      Size            =   "18653;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label6 
      Height          =   270
      Left            =   390
      TabIndex        =   11
      Top             =   810
      Width           =   960
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "ITEM CODE"
      Size            =   "1693;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   780
      Width           =   1455
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2566;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   8325
      TabIndex        =   10
      Top             =   750
      Width           =   165
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "291;635"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   6435
      TabIndex        =   9
      Top             =   795
      Width           =   540
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DATE"
      Size            =   "952;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   1245
      Left            =   135
      TabIndex        =   7
      Top             =   615
      Width           =   10200
      BackColor       =   8421504
      Size            =   "17992;2196"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1245
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   10200
      BackColor       =   0
      Size            =   "17992;2196"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_InvRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
     Call connecttoserver
     Call subPrintInvRecord
     Call disconnecttoserver
End Sub

Private Sub cmdView_Click()
    Call connecttoserver
    If txtItemCode = "" Then
        MsgBox "Please Input Item ID", vbInformation, pstrMessage
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait.... loading records")
    hflxInvRecord.Clear
    hflxInvRecord.Rows = 2
    hflxInvRecord.Refresh
    Call subFormatGrid
    Call subViewRecord
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then dtpTo.SetFocus
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then cmdView.SetFocus
End Sub

Private Sub Form_Load()
    Call connecttoserver
    Call clsDB.SQLServer(PrintMenuDb, App.Path & "\Print.ini")
    dtpFrom.Value = DateAdd("m", -1, Format(Date, "yyyy/mm/dd"))
    dtpTo.Value = Format(Date, "yyyy/mm/dd")
    
    Call clsPrintMenu.psubLoadDivision(cboDivision)
    Call subFormatGrid
    Call disconnecttoserver
End Sub
Private Sub hflxInvRecord_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 And hflxInvRecord.TextMatrix(1, 0) <> "" Then
          PopupMenu F_PopMenu.mnuFile
     End If
End Sub

Private Sub txtItemCode_Change()
     hflxInvRecord.Clear
     hflxInvRecord.Rows = 2
     Call subFormatGrid
End Sub

Private Sub txtItemCode_LostFocus()
    Call connecttoserver
    lblDescription.Caption = clsPrintMenu.pfstrGetItemDescription(txtItemCode)
    If lblDescription.Caption = "" And txtItemCode.Text <> "" Then
          MsgBox "Invalid Item Code!", vbExclamation, pstrMessage
          txtItemCode.SetFocus
          Exit Sub
    End If
    txtItemCode.Text = UCase(txtItemCode.Text)
    Call disconnecttoserver
End Sub

Private Sub subViewRecord()
    Dim objInvRecord          As Object
    Dim lngSeqNo              As Long
    Dim strTransactionDate    As Date
    Dim dblLastBalance        As Double, dblRunBalance As Double
    Dim bytTransTypeIsIn      As Byte
    
On Error GoTo lnError
     Set objInvRecord = clsPrintMenu.InvRecord.GetInventoryRecord(txtItemCode, dtpFrom.Value, dtpTo.Value, cboDivision)
     dblLastBalance = clsPrintMenu.pfvarStockBalance(txtItemCode, cboDivision, DateAdd("d", -1, dtpFrom.Value))
     dblRunBalance = dblLastBalance
     With objInvRecord
          Do While Not .EOF
               lngSeqNo = lngSeqNo + 1
               hflxInvRecord.Rows = lngSeqNo + 1
               Select Case .Fields("TransactionTypeId").Value
                    Case 1
                    '---Received/In
                         bytTransTypeIsIn = 1
                         If strTransactionDate <> .Fields("TransactedDate").Value Then
                              dblRunBalance = clsPrintMenu.pfvarStockBalance( _
                                        txtItemCode, cboDivision, DateAdd("d", -1, .Fields("TransactedDate").Value)) + .Fields("Qty").Value
                         Else
                              dblRunBalance = dblRunBalance + .Fields("Qty").Value
                         End If
                         hflxInvRecord.TextMatrix(lngSeqNo, 2) = pfvarIs_Null(.Fields("Qty").Value, False)
                    Case 3
                    '---Shipped/Out
                         bytTransTypeIsIn = 3
                         If strTransactionDate <> .Fields("TransactedDate").Value Then
                              dblRunBalance = clsPrintMenu.pfvarStockBalance( _
                                        txtItemCode, cboDivision, DateAdd("d", -1, .Fields("TransactedDate").Value)) - .Fields("Qty").Value
                         Else
                              dblRunBalance = dblRunBalance - .Fields("Qty").Value
                         End If
                         hflxInvRecord.TextMatrix(lngSeqNo, 3) = pfvarIs_Null(.Fields("Qty").Value, False)
                    Case 4
                    '---Stock Taking/AC
                         bytTransTypeIsIn = 4
                         If strTransactionDate <> .Fields("TransactedDate").Value Then
                              dblRunBalance = .Fields("Qty").Value
                         Else
                              dblRunBalance = .Fields("Qty").Value - _
                                            clsPrintMenu.InvRecord.TotalOut(txtItemCode, cboDivision, .Fields("TransactedDate").Value) + _
                                            clsPrintMenu.InvRecord.TotalIn(txtItemCode, .Fields("TransactedDate").Value)
                         End If
                         hflxInvRecord.TextMatrix(lngSeqNo, 4) = pfvarIs_Null(.Fields("Qty").Value, False)
               End Select
               hflxInvRecord.TextMatrix(lngSeqNo, 0) = .Fields("TransactedDate").Value
               hflxInvRecord.TextMatrix(lngSeqNo, 1) = .Fields("TransId").Value
               hflxInvRecord.TextMatrix(lngSeqNo, 5) = dblRunBalance
               hflxInvRecord.TextMatrix(lngSeqNo, 6) = pfvarIs_Null(.Fields("Remarks").Value)
               
               strTransactionDate = .Fields("TransactedDate").Value
               
               .MoveNext
          Loop
     End With
     Set objInvRecord = Nothing
     Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description & _
    Chr(13) & "Please try to view again.", vbExclamation, pstrMessage
End Sub

Private Sub subPrintInvRecord()
    Dim adoInvRecord As New ADODB.Recordset
    Dim dblStDev     As Double
    Dim strLeadTime  As String
    Dim lngStyStock  As Long
    Dim lngRow       As Long
    
On Error GoTo lnErrMsg
    '--- Get the Standard Deviation
    dblStDev = clsPrintMenu.InventoryCheck.GetStDev(txtItemCode, dtpFrom.Value, dtpTo.Value, cboDivision.Text)
    '--- get the lead time
    strLeadTime = clsPrintMenu.InventoryCheck.GetLeadTime(txtItemCode)
    '--- compute the Safety Stock
    lngStyStock = CLng(2 * Math.Sqr(CDbl(strLeadTime) * dblStDev))
    
'    Unload DataEnv

    '--- Delete all the records first
    Call clsPrintMenu.InvRecord.DeleteInvRecordsWT
    With hflxInvRecord
          For lngRow = 1 To hflxInvRecord.Rows - 1
                     Call clsPrintMenu.InvRecord.SaveToInvRecordWT(lngRow, pfvarIs_Null(.TextMatrix(lngRow, 0)), _
                               pfvarIs_Null(.TextMatrix(lngRow, 1)), pfvarIs_Null(.TextMatrix(lngRow, 2), False), _
                               pfvarIs_Null(.TextMatrix(lngRow, 3), False), pfvarIs_Null(.TextMatrix(lngRow, 4), False), _
                               pfvarIs_Null(.TextMatrix(lngRow, 5), False), _
                               pfvarIs_Null(.TextMatrix(lngRow, 6)))
          Next lngRow
    End With
    
    Set adoInvRecord = clsDB.GetRecordSet(clsPrintMenu.InvRecord.SQLInventoryRecord, False)
    Set DR_INV_Record.DataSource = adoInvRecord
    
    With DR_INV_Record.Sections("Section2")
          .Controls("lblItemId").Caption = txtItemCode.Text
          .Controls("lblDescription").Caption = lblDescription
          .Controls("lblDivision").Caption = cboDivision.Text
          .Controls("lblSafetyStock").Caption = CStr(lngStyStock)
    End With
    With DR_INV_Record.Sections("Section1")
          .Controls("txtTransDate").DataField = "TransDate"
          .Controls("txtDivision").DataField = "TransId"
          .Controls("txtInQty").DataField = "InQty"
          .Controls("txtOutQty").DataField = "OutQty"
          .Controls("txtAcQty").DataField = "AcQty"
          .Controls("txtBalance").DataField = "Balance"
          .Controls("txtRemarks").DataField = "Remarks"
    End With
    
    DR_INV_Record.Show vbModal
    Set adoInvRecord = Nothing
    'DR_INV_Record.PrintReport
    'Unload DataEnv
    Exit Sub
lnErrMsg:
     MsgBox Err.Number & "-" & Err.Description & Chr(13) & "Please try to print again.", vbExclamation, pstrMessage
End Sub

Private Sub subFormatGrid()
    Dim intLoop As Integer
    
    With hflxInvRecord
        .Cols = 7
        .RowHeight(0) = 450
        For intLoop = 0 To 6
               .TextMatrix(0, intLoop) = Choose(intLoop + 1, "TRANSACTION DATE", "TRANSID", "IN QTY", "OUT QTY", "AC QTY", "BALANCE", "REMARKS")
               .ColWidth(intLoop) = Choose(intLoop + 1, 1500, 1500, 1200, 1200, 1200, 1300, 1800)
               .Col = intLoop: .Row = 0
               .CellFontBold = True
               .CellAlignment = flexAlignCenterCenter
               .ColAlignment(intLoop) = flexAlignCenterCenter
        Next
    End With
End Sub

