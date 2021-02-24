VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_SupplierItemList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Item List"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13650
   Icon            =   "F_SupplierItemList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   13650
   Begin MSComCtl2.DTPicker dtPoDateTo 
      Height          =   315
      Left            =   4875
      TabIndex        =   18
      Top             =   555
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   -2147483624
      CalendarTrailingForeColor=   -2147483632
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   120979457
      CurrentDate     =   38740
   End
   Begin MSComCtl2.DTPicker dtPoDateFrom 
      Height          =   315
      Left            =   2535
      TabIndex        =   17
      Top             =   555
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   -2147483624
      CalendarTitleBackColor=   -2147483627
      CalendarTitleForeColor=   -2147483634
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   120979457
      CurrentDate     =   38740
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   9540
      TabIndex        =   0
      Top             =   1290
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
      MICON           =   "F_SupplierItemList.frx":0CCA
      PICN            =   "F_SupplierItemList.frx":0CE6
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
      Left            =   11940
      TabIndex        =   1
      Top             =   1290
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
      MICON           =   "F_SupplierItemList.frx":1282
      PICN            =   "F_SupplierItemList.frx":129E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxList 
      Height          =   5520
      Left            =   135
      TabIndex        =   2
      Top             =   1935
      Width           =   13335
      _ExtentX        =   23521
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
   Begin OsenXPCntrl.OsenXPButton cmdExcel 
      Height          =   375
      Left            =   10740
      TabIndex        =   16
      Top             =   1290
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
      MICON           =   "F_SupplierItemList.frx":183A
      PICN            =   "F_SupplierItemList.frx":1856
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.Label Label9 
      Height          =   240
      Left            =   4485
      TabIndex        =   19
      Top             =   615
      Width           =   285
      ForeColor       =   -2147483628
      BackColor       =   -2147483632
      Caption         =   "TO"
      Size            =   "503;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Left            =   5850
      TabIndex        =   15
      Top             =   990
      Width           =   975
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "SUPPLIER ID"
      Size            =   "1720;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
      VariousPropertyBits=   746604569
      BackColor       =   -2147483624
      Size            =   "2566;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   210
      Left            =   390
      TabIndex        =   13
      Top             =   1350
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
   Begin MSForms.Label lblItemDescription 
      Height          =   315
      Left            =   2790
      TabIndex        =   12
      Top             =   1350
      Width           =   5565
      ForeColor       =   -2147483628
      VariousPropertyBits=   8388627
      Size            =   "9816;556"
      SpecialEffect   =   6
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optByItem 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   1590
      TabIndex        =   11
      Top             =   510
      Width           =   735
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   -2147483628
      DisplayStyle    =   5
      Size            =   "1296;635"
      Value           =   "0"
      Caption         =   "ITEM"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton optBySupplier 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   360
      TabIndex        =   10
      Top             =   510
      Width           =   1110
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   -2147483628
      DisplayStyle    =   5
      Size            =   "1958;635"
      Value           =   "1"
      Caption         =   "SUPPLIER"
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboSupplierId 
      Height          =   315
      Left            =   6870
      TabIndex        =   9
      Top             =   915
      Width           =   1455
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "2566;556"
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
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13695
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Supplier Item List"
      Size            =   "24156;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   210
      Left            =   390
      TabIndex        =   5
      Top             =   945
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
   Begin MSForms.ComboBox cboSupplier 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   915
      Width           =   4365
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "7699;556"
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
   Begin MSForms.Label Label3 
      Height          =   5535
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   13335
      BackColor       =   0
      Size            =   "23521;9763"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   1245
      Left            =   150
      TabIndex        =   7
      Top             =   510
      Width           =   13320
      BackColor       =   8421504
      Size            =   "23495;2196"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1245
      Left            =   255
      TabIndex        =   8
      Top             =   615
      Width           =   13320
      BackColor       =   0
      Size            =   "23495;2196"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_SupplierItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSupplier_Change()
     Call cboSupplier_Click
End Sub

Private Sub cboSupplier_Click()
     If optBySupplier.Value Then
            Call connecttoserver
          Call clsPrintMenu.psubLoadSupplierId(cboSupplierId, cboSupplier.Text)
          If cboSupplierId.ListCount > 0 Then cboSupplierId.ListIndex = 0
     End If
End Sub





Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExcel_Click()
    Call connecttoserver
     Dim lngLoop As Long
     Dim bytCol  As Byte
     Call psubShowStatMsg("Writing to excel....")
     Call clsPrintMenu.Utility.OpenExcel
     With clsPrintMenu.Utility.ExcelWkSheet
          For lngLoop = 0 To hflxList.Rows - 1
               For bytCol = 0 To hflxList.Cols - 1
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngLoop + 1, bytCol + 1).NumberFormat = "@"
                    .Cells(lngLoop + 1, bytCol + 1) = hflxList.TextMatrix(lngLoop, bytCol)
               Next
          Next
          Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, hflxList.Cols)
          clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
          clsPrintMenu.Utility.ExcelApp.Visible = True
          Call clsPrintMenu.Utility.CloseExcel
     End With
     Call clsPrintMenu.Utility.CloseExcel
     Call psubHideStatMsg
    Call disconnecttoserver
End Sub

Private Sub cmdView_Click()
    Call connecttoserver
     Dim clsSupplierItemList As New C08_SupplierItemList
          
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Loading Supplier Item List...")
     
     If optBySupplier.Value = True And IsNull(dtPoDateFrom.Value) And IsNull(dtPoDateTo.Value) Then
          clsSupplierItemList.SupplierId = cboSupplierId.Text
          Call clsSupplierItemList.GetSupplierItemList(hflxList, optBySupplier.Value, optByItem.Value)
     ElseIf optBySupplier.Value = True And (Not IsNull(dtPoDateFrom.Value) Or Not IsNull(dtPoDateTo.Value)) Then
          clsSupplierItemList.SupplierId = cboSupplierId.Text
          If IsNull(dtPoDateFrom.Value) Then
                clsSupplierItemList.PoDateFrom = ""
          ElseIf Not IsNull(dtPoDateFrom.Value) Then
                clsSupplierItemList.PoDateFrom = dtPoDateFrom.Value
          End If
          If IsNull(dtPoDateTo.Value) Then
                clsSupplierItemList.PoDateTo = ""
          ElseIf Not IsNull(dtPoDateTo.Value) Then
                clsSupplierItemList.PoDateTo = dtPoDateTo.Value
          End If
          Call clsSupplierItemList.GetSupplierItemList(hflxList, optBySupplier.Value, optByItem.Value)
    ElseIf optByItem.Value = True And IsNull(dtPoDateFrom.Value) And IsNull(dtPoDateTo.Value) Then
          clsSupplierItemList.ItemId = txtItemCode.Text
          Call clsSupplierItemList.GetSupplierItemList(hflxList, optBySupplier.Value, optByItem.Value)
    ElseIf optByItem.Value = True And (Not IsNull(dtPoDateFrom.Value) Or Not IsNull(dtPoDateTo.Value)) Then
          clsSupplierItemList.ItemId = txtItemCode
          If IsNull(dtPoDateFrom.Value) Then
                clsSupplierItemList.PoDateFrom = ""
          ElseIf Not IsNull(dtPoDateFrom.Value) Then
                clsSupplierItemList.PoDateFrom = dtPoDateFrom.Value
          End If
          If IsNull(dtPoDateTo.Value) Then
                clsSupplierItemList.PoDateTo = ""
          ElseIf Not IsNull(dtPoDateTo.Value) Then
                clsSupplierItemList.PoDateTo = dtPoDateTo.Value
          End If
          Call clsSupplierItemList.GetSupplierItemList(hflxList, optBySupplier.Value, optByItem.Value)
    End If
     cmdExcel.Enabled = True
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
     Call subGridInitialize
     Call disconnecttoserver
End Sub

Private Sub Form_Load()
     Call connecttoserver
     Call clsPrintMenu.psubLoadSupplier(cboSupplier)
     If cboSupplier.ListCount > 0 Then cboSupplier.ListIndex = 0
     Call subGridInitialize
     Call disconnecttoserver
End Sub

Private Sub subGridInitialize()
     Dim intCol As Integer
     Dim lngRow As Long
     
     With hflxList
          
          If (optByItem.Value = True Or optBySupplier.Value = True) And IsNull(dtPoDateFrom.Value) And IsNull(dtPoDateTo.Value) Then
          .Cols = 7
          
          For intCol = 0 To .Cols - 1
                 .TextMatrix(0, intCol) = Choose(intCol + 1, "Supplier Name", "SupplierID", "ItemID", "Description", "Unit Price" _
                                                         , "Currency", "Leadtime")
               
                    .ColWidth(intCol) = Choose(intCol + 1, 3500, 1000, 1500, 3500, 1500, 1000, 1000)
                    .Row = 0: .Col = intCol
                    .RowHeight(0) = 300
                    .CellAlignment = 4
                    .CellFontBold = True
                    
          Next
          
          ElseIf (optBySupplier.Value = True Or optByItem.Value = True) And Not IsNull(dtPoDateFrom.Value) Or Not IsNull(dtPoDateTo.Value) Then
          .Cols = 9
          
          For intCol = 0 To .Cols - 1
                 .TextMatrix(0, intCol) = Choose(intCol + 1, "Supplier Name", "SupplierID", "ItemID", "Description", "Unit Price" _
                                                         , "Currency", "Leadtime", "QtyUnit", "Qty")
          
                    .ColWidth(intCol) = Choose(intCol + 1, 3500, 1000, 1500, 3500, 1500, 1000, 1000, 1000, 2000)
               
               .Row = 0: .Col = intCol
               .RowHeight(0) = 300
               .CellAlignment = 4
               .CellFontBold = True
               .Col = 8: .Row = lngRow
          Next
          For lngRow = 0 To .Rows - 1
             If .TextMatrix(lngRow, 8) = "" Then
                .TextMatrix(lngRow, 8) = 0
              End If
          Next
          
    
          End If
          
          .MergeCells = flexMergeRestrictColumns
          .MergeCol(0) = True
          .MergeCol(1) = True
          
          
     End With
End Sub
Private Sub optByItem_Click()
     txtItemCode.Enabled = True
     cboSupplier.Locked = True
     cboSupplier.ListIndex = 0
     cboSupplierId.Clear
End Sub

Private Sub optBySupplier_Click()
     txtItemCode.Text = ""
     lblItemDescription.Caption = ""
     txtItemCode.Enabled = False
     cboSupplier.Locked = False
End Sub

Private Sub txtItemCode_Change()
    Call connecttoserver
     lblItemDescription.Caption = clsPrintMenu.pfstrGetItemDescription(txtItemCode.Text)
    Call disconnecttoserver
End Sub
