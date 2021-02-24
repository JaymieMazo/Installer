VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_POSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PO Search"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13035
   Icon            =   "F_POSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSupCategory 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "F_POSearch.frx":0CCA
      Left            =   4635
      List            =   "F_POSearch.frx":0CD7
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   585
      Width           =   1350
   End
   Begin VB.ComboBox cboCurrency 
      Height          =   315
      Left            =   11760
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   600
      Width           =   915
   End
   Begin OsenXPCntrl.OsenXPButton cmdReport 
      Height          =   375
      Left            =   9240
      TabIndex        =   21
      Top             =   990
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "MONTHLY IMPEX REPORT"
      ENAB            =   0   'False
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
      MICON           =   "F_POSearch.frx":0CF1
      PICN            =   "F_POSearch.frx":0D0D
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
      Left            =   7170
      TabIndex        =   1
      Top             =   585
      Width           =   1470
      _ExtentX        =   2593
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
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   23724033
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   8910
      TabIndex        =   2
      Top             =   585
      Width           =   1470
      _ExtentX        =   2593
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
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   23724033
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   375
      Left            =   11580
      TabIndex        =   20
      Top             =   1410
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
      MICON           =   "F_POSearch.frx":12A9
      PICN            =   "F_POSearch.frx":12C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Top             =   1410
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
      MICON           =   "F_POSearch.frx":1861
      PICN            =   "F_POSearch.frx":187D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxPOSearch 
      Height          =   4650
      Left            =   150
      TabIndex        =   8
      Top             =   2145
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   8202
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   13536000
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
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin OsenXPCntrl.OsenXPButton cmdExcel 
      Height          =   375
      Left            =   10410
      TabIndex        =   18
      Top             =   1410
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&EXCEL"
      ENAB            =   0   'False
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
      MICON           =   "F_POSearch.frx":1E19
      PICN            =   "F_POSearch.frx":1E35
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000010&
      Caption         =   "SUPPLIER CATEG."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   285
      Left            =   3210
      TabIndex        =   25
      Top             =   660
      Width           =   1425
   End
   Begin VB.Label lblCurrency 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "CURRENCY UNIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   10440
      TabIndex        =   23
      Top             =   675
      Width           =   1260
   End
   Begin MSForms.Label Label10 
      Height          =   210
      Left            =   6090
      TabIndex        =   19
      Top             =   1080
      Width           =   585
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "MAKER"
      Size            =   "1032;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboMaker 
      Height          =   315
      Left            =   6705
      TabIndex        =   4
      Top             =   1020
      Width           =   2475
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "4366;556"
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
   Begin MSForms.TextBox txtDescription 
      Height          =   315
      Left            =   3135
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6045
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      Size            =   "10663;556"
      SpecialEffect   =   3
      FontName        =   "MS UI Gothic"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboSupplier 
      Height          =   315
      Left            =   1110
      TabIndex        =   3
      Top             =   1020
      Width           =   4860
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "8572;556"
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
   Begin MSForms.Label Label9 
      Height          =   210
      Left            =   285
      TabIndex        =   17
      Top             =   1080
      Width           =   780
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "SUPPLIER"
      Size            =   "1376;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemId 
      Height          =   315
      Left            =   1110
      TabIndex        =   5
      Top             =   1440
      Width           =   1965
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3466;556"
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
      Left            =   405
      TabIndex        =   16
      Top             =   1530
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
   Begin MSForms.TextBox txtPoNo 
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Top             =   600
      Width           =   1995
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3519;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label6 
      Height          =   210
      Left            =   570
      TabIndex        =   15
      Top             =   690
      Width           =   510
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "PO NO"
      Size            =   "900;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Left            =   6105
      TabIndex        =   11
      Top             =   630
      Width           =   1005
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ORDER DATE"
      Size            =   "1773;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   8670
      TabIndex        =   12
      Top             =   600
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
   Begin MSForms.Label Label2 
      Height          =   1425
      Left            =   150
      TabIndex        =   13
      Top             =   480
      Width           =   12660
      BackColor       =   8421504
      Size            =   "22331;2514"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   4665
      Left            =   255
      TabIndex        =   10
      Top             =   2250
      Width           =   12660
      BackColor       =   0
      Size            =   "22331;8229"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   -30
      TabIndex        =   9
      Top             =   30
      Width           =   13080
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Purchase Order Search"
      Size            =   "23072;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   1425
      Left            =   270
      TabIndex        =   14
      Top             =   600
      Width           =   12645
      BackColor       =   0
      Size            =   "22304;2514"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_POSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSupCategory_LostFocus()
    clsPrintMenu.SupplierCategory = cboSupCategory.Text
    Call clsPrintMenu.psubLoadSupplier(cboSupplier)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdExcel_Click()
     Dim lngLoop As Long
     Dim bytCol  As Byte
     Me.MousePointer = vbHourglass
     Call clsPrintMenu.Utility.OpenExcel
     With hflxPOSearch
          For lngLoop = 0 To .Rows - 1
               For bytCol = 0 To .Cols - 1
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngLoop + 1, bytCol + 1) = .TextMatrix(lngLoop, bytCol)
               Next
          Next
          clsPrintMenu.Utility.ExcelWkSheet.Range(clsPrintMenu.Utility.ExcelWkSheet.Cells(1, 1), _
                                                  clsPrintMenu.Utility.ExcelWkSheet.Cells(lngLoop, .Cols + 1)).ClearFormats
          Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, .Cols)
          clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
          clsPrintMenu.Utility.ExcelApp.Visible = True
          Call clsPrintMenu.Utility.CloseExcel
     End With
     Call clsPrintMenu.Utility.CloseExcel
     Me.MousePointer = vbDefault
End Sub

Private Sub cmdReport_Click()
Dim lngLoop As Long
Dim bytCol  As Byte
Dim lngRow As Long

     Me.MousePointer = vbHourglass
     Call clsPrintMenu.Utility.OpenExcel
     Call subFormatWidth
     Call subFormatXlWkSheet
     Call subBorder(hflxPOSearch)
     
    With hflxPOSearch
        For bytCol = 1 To 9
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, bytCol) = Choose(bytCol, "ORDER DATE", "ITEM NAME", "QTY", "UNIT PRICE", "CURRENCY", _
                                                                             "EXCHANGE RATE", "VALUE IN OTHER CURRENCY", "TOTAL VALUE IN PHP", "SUPPLIER")
        Next
          For lngLoop = 1 To .Rows - 1
             For bytCol = 1 To 9
                  clsPrintMenu.Utility.ExcelWkSheet.Cells(lngLoop + 2, bytCol).ShrinkToFit = True
                  clsPrintMenu.Utility.ExcelWkSheet.Cells(lngLoop + 2, bytCol) = Choose(bytCol, "'" & Format(.TextMatrix(lngLoop, 10), "yyyy/mm/dd"), .TextMatrix(lngLoop, 3), .TextMatrix(lngLoop, 4) * .TextMatrix(lngLoop, 5), _
                                                                               .TextMatrix(lngLoop, 8), .TextMatrix(lngLoop, 7), "", "", "", .TextMatrix(lngLoop, 18))
             Next
        Next
        Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, 9)
        clsPrintMenu.Utility.ExcelApp.Visible = True
        Call clsPrintMenu.Utility.CloseExcel
    End With
     Call clsPrintMenu.Utility.CloseExcel
     Me.MousePointer = vbDefault
End Sub

Private Sub cmdView_Click()
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Please wait.... loading records.")
     Call subGetPODetails
     Call subInitGrid
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
     
     If hflxPOSearch.Rows = 1 Then
        cmdReport.Enabled = False
     ElseIf hflxPOSearch.Rows <> 1 Then
        cmdReport.Enabled = True
    End If
End Sub

Private Sub dtpFrom_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub dtpTo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
     cboSupCategory.ListIndex = 0
     clsPrintMenu.SupplierCategory = cboSupCategory.Text
     Call clsPrintMenu.psubLoadSupplier(cboSupplier)
     Call clsPrintMenu.psubLoadMaker(cboMaker)
     Call clsPrintMenu.psubLoadCurrency(cboCurrency)
End Sub

Private Sub txtItemId_Change()
     txtDescription.Text = ""
End Sub

Private Sub txtItemID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then
          txtDescription.Text = clsPrintMenu.pfstrGetItemDescription(txtItemId)
'          SendKeys "{tab}"
     End If
End Sub

Private Sub txtPoNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub subInitGrid()
     Dim intCol As Integer

     With hflxPOSearch
          .RowHeight(0) = 300
          .Row = 0
          For intCol = 0 To .Cols - 1
               .ColWidth(intCol) = Choose(intCol + 1, 1200, 700, 1000, 4000, 1800, 1800, 1800, 1000, 2500, 2000, _
                         1000, 1500, 1500, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000, 1000)
               .ColAlignment(intCol) = Choose(intCol + 1, 4, 4, 4, 1, 7, 7, 7, 4, 1, 1, 1, 1, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4)
               .Col = intCol
               .CellAlignment = 4
               .CellFontBold = True
          Next
     End With
End Sub

Private Sub subGetPODetails()
Dim objPOSearch     As Object
Dim strSQLQuery     As String
          
On Error GoTo lnErrHandler
     strSQLQuery = "SELECT POSearchView.PoNo,POSearchView.PoSeq,POSearchView.ItemId," _
                    & " POSearchView.Description,POSearchView.Qty,POSearchView.ConvertingCoefficient,POSearchView.QtyUnit, " _
                    & " POSearchView.CurrencyUnit,POSearchView.UnitPrice,POSearchView.Division," _
                    & " POSearchView.Issueddate,POSearchView.Acknowledger,POSearchView.EtdDate," _
                    & " POSearchView.EtaDate,POSearchView.FtryDate,POSearchView.TermOfPayment," _
                    & " POSearchView.Remarks,POSearchView.SupplierId,POSearchView.SupplierName," _
                    & " POSearchView.MakerName,POSearchView.ReceivedAllInvoices,POSearchView.Canceled" _
                    & " FROM POSearchView INNER JOIN Suppliers ON Suppliers.SupplierId=POSearchView.SupplierId "
     
     '---Search by ItemId, Supplier and PONo
     If txtItemId.Text <> "" And cboSupplier.Text <> "" And txtPoNo.Text <> "" Then _
          strSQLQuery = strSQLQuery & " WHERE  POSearchView.ItemId like '" & txtItemId.Text & "%' " _
                    & " And POSearchView.SupplierName = " & pfstrQt(cboSupplier.Text) _
                    & " And POSearchView.PONo = " & pfstrQt(txtPoNo.Text)
     '---Search by ItemID only
     If txtItemId.Text <> "" And cboSupplier = "" And txtPoNo.Text = "" And cboMaker.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.ItemId = '" & txtItemId.Text & "'"
     '---Search by PONo only
     If txtItemId.Text = "" And cboSupplier.Text = "" And txtPoNo.Text <> "" And cboMaker.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.PoNo = " & pfstrQt(txtPoNo.Text)
     '---Search by Supplier only
     If txtItemId.Text = "" And cboSupplier.Text <> "" And txtPoNo.Text = "" And cboMaker.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.SupplierName = " & pfstrQt(cboSupplier.Text)
     '---Search by Maker Only
     If cboMaker.Text <> "" And txtItemId.Text = "" And cboSupplier.Text = "" And txtPoNo.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.MakerName = " & pfstrQt(cboMaker.Text)
     '---Search by ItemId and Supplier
     If txtItemId.Text <> "" And cboSupplier.Text <> "" And txtPoNo.Text = "" And cboMaker.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.ItemId like '" & txtItemId & "%' And " _
                    & " POSearchView.SupplierName = " & pfstrQt(cboSupplier.Text)
     '---Search by PONo and Supplier
     If txtItemId.Text = "" And cboSupplier.Text <> "" And txtPoNo.Text <> "" And cboMaker.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.SupplierName = " & pfstrQt(cboSupplier.Text) _
                    & " And POSearchView.PoNo = " & pfstrQt(txtPoNo.Text)
     '---Search by ItemId and PoNo and Currency
     If txtItemId.Text <> "" And cboSupplier.Text = "" And txtPoNo.Text <> "" Then
          strSQLQuery = strSQLQuery & " WHERE POSearchView.ItemId like '" & txtItemId.Text & "%' " _
                                    & " And POSearchView.PoNo = " & pfstrQt(txtPoNo.Text)
     End If
     '---Search by ItemId, Supplier, Maker
     If txtItemId.Text <> "" And cboSupplier.Text <> "" And cboMaker.Text <> "" And txtPoNo.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.ItemId like '" & txtItemId & "%' And " _
                    & " POSearchView.SupplierName = " & pfstrQt(cboSupplier.Text) _
                    & " And POSearchView.MakerName = " & pfstrQt(cboMaker)
     '---Search by Supplier, Maker
     If cboSupplier.Text <> "" And cboMaker.Text <> "" And txtItemId.Text = "" And txtPoNo.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.SupplierName = " & pfstrQt(cboSupplier.Text) _
                    & " And POSearchView.MakerName = " & pfstrQt(cboMaker)
     '---Search by ItemId, Maker
     If txtItemId.Text <> "" And cboSupplier.Text = "" And cboMaker.Text <> "" And txtPoNo.Text = "" Then _
          strSQLQuery = strSQLQuery & " WHERE POSearchView.ItemId like '" & txtItemId & "%' And " _
                    & " POSearchView.MakerName = " & pfstrQt(cboMaker)
                    
     '---Search by
     If IsNull(dtpFrom.Value) And IsNull(dtpTo.Value) Then
          '------no value
     ElseIf txtItemId.Text <> "" Or cboSupplier.Text <> "" Or txtPoNo.Text <> "" Or cboMaker.Text <> "" Then
          strSQLQuery = strSQLQuery & " And "
     ElseIf Not IsNull(dtpFrom.Value) Or Not IsNull(dtpTo.Value) Then
          strSQLQuery = strSQLQuery & " Where "
     End If
     
     If Not IsNull(dtpFrom.Value) And Not IsNull(dtpTo.Value) Then
          strSQLQuery = strSQLQuery & " IssuedDate >= " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd")) _
               & " And IssuedDate <= " & pfstrQt(DateAdd("d", 1, Format(dtpTo.Value, "yyyy/mm/dd")))
     ElseIf Not IsNull(dtpFrom.Value) And IsNull(dtpTo.Value) Then
          strSQLQuery = strSQLQuery & " IssuedDate >= " & pfstrQt(Format(dtpFrom.Value, "yyyy/mm/dd"))
     ElseIf IsNull(dtpFrom.Value) And Not IsNull(dtpTo.Value) Then
          strSQLQuery = strSQLQuery & " IssuedDate <= " & pfstrQt(Format(dtpTo.Value, "yyyy/mm/dd"))
     End If
     
     If cboCurrency.Text <> "ALL" Then _
            strSQLQuery = strSQLQuery & " And CurrencyUnit='" & cboCurrency.Text & "'"
     If cboSupCategory <> "ALL" Then
        If cboSupCategory = "LOCAL" Then
            strSQLQuery = strSQLQuery & " And ImportSupplier=0"
        ElseIf cboSupCategory = "IMPORTED" Then
            strSQLQuery = strSQLQuery & " And ImportSupplier=1"
        End If
     End If
     
     strSQLQuery = strSQLQuery & " ORDER BY PONo,PoSeq,CurrencyUnit"
     
     Set objPOSearch = clsDB.GetRecordSet(strSQLQuery)
     
     Set hflxPOSearch.DataSource = objPOSearch
     cmdExcel.Enabled = objPOSearch.RecordCount > 0
     If objPOSearch.EOF Then MsgBox "No record found.", vbExclamation, "PO Search"
     Set objPOSearch = Nothing
     Exit Sub
lnErrHandler:
     MsgBox Err.Number & Err.Description, vbCritical
End Sub

Private Sub subBorder(ByVal hflxGrid As Object)
Dim strLen As String
                   
    strLen = "A1:I"
    'format for the worksheet border style
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLen & CStr(hflxGrid.Rows + 1)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLen & CStr(hflxGrid.Rows + 1)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLen & CStr(hflxGrid.Rows + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLen & CStr(hflxGrid.Rows + 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLen & CStr(hflxGrid.Rows + 1)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With clsPrintMenu.Utility.ExcelWkSheet.Range(strLen & CStr(hflxGrid.Rows + 1)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Private Sub subFormatXlWkSheet()
Dim lngLoop As Long
Dim bytCol As Byte
    'format for the column header
    For bytCol = 1 To 9
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, bytCol).HorizontalAlignment = xlCenter
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, bytCol).VerticalAlignment = xlCenter
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, bytCol).Font.FontStyle = "Arial"
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, bytCol).Font.Bold = True
    Next
        clsPrintMenu.Utility.ExcelWkSheet.Range("A1:A2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("B1:B2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("C1:C2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("D1:D2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("E1:E2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("F1:F2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("G1:G2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("H1:H2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Range("I1:I2").MergeCells = True
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, 7).WrapText = True
        clsPrintMenu.Utility.ExcelWkSheet.Cells(1, 8).WrapText = True
    
        'format for the page margins
        clsPrintMenu.Utility.ExcelWkSheet.PageSetup.Orientation = xlLandscape
        clsPrintMenu.Utility.ExcelWkSheet.PageSetup.TopMargin = Application.InchesToPoints(0.5)
        clsPrintMenu.Utility.ExcelWkSheet.PageSetup.BottomMargin = Application.InchesToPoints(0.5)
        clsPrintMenu.Utility.ExcelWkSheet.PageSetup.LeftMargin = Application.InchesToPoints(0.5)
        clsPrintMenu.Utility.ExcelWkSheet.PageSetup.RightMargin = Application.InchesToPoints(0.5)
End Sub

Private Sub subFormatWidth()

    'format for the column width
    clsPrintMenu.Utility.ExcelWkSheet.Columns("A:A").ColumnWidth = 9
    clsPrintMenu.Utility.ExcelWkSheet.Columns("B:B").ColumnWidth = 32
    clsPrintMenu.Utility.ExcelWkSheet.Columns("C:C").ColumnWidth = 5
    clsPrintMenu.Utility.ExcelWkSheet.Columns("D:D").ColumnWidth = 9
    clsPrintMenu.Utility.ExcelWkSheet.Columns("E:E").ColumnWidth = 9
    clsPrintMenu.Utility.ExcelWkSheet.Columns("F:F").ColumnWidth = 13
    clsPrintMenu.Utility.ExcelWkSheet.Columns("G:G").ColumnWidth = 13
    clsPrintMenu.Utility.ExcelWkSheet.Columns("H:H").ColumnWidth = 15
    clsPrintMenu.Utility.ExcelWkSheet.Columns("I:I").ColumnWidth = 26
End Sub

    
    
