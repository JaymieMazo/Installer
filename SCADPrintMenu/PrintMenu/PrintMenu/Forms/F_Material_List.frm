VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_Material_List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material List"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13650
   Icon            =   "F_Material_List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   13650
   Begin VB.CommandButton cmdTest 
      Caption         =   "Command1"
      Height          =   315
      Left            =   13230
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   8850
      TabIndex        =   7
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
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
      MICON           =   "F_Material_List.frx":0CCA
      PICN            =   "F_Material_List.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxViewMaterial 
      Height          =   5520
      Left            =   105
      TabIndex        =   12
      Top             =   2190
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   9737
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483647
      BackColorBkg    =   11049333
      GridColor       =   -2147483633
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin OsenXPCntrl.OsenXPButton cmdExcel 
      Height          =   375
      Left            =   10275
      TabIndex        =   8
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
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
      MICON           =   "F_Material_List.frx":1282
      PICN            =   "F_Material_List.frx":129E
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
      Left            =   11955
      TabIndex        =   9
      Top             =   1230
      Width           =   1200
      _ExtentX        =   2117
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
      MICON           =   "F_Material_List.frx":1838
      PICN            =   "F_Material_List.frx":1854
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   1215
      TabIndex        =   0
      Top             =   810
      Width           =   2925
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "5159;556"
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
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   300
      TabIndex        =   19
      Top             =   885
      Width           =   960
      ForeColor       =   -2147483634
      BackColor       =   -2147483634
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
   Begin MSForms.Label Label6 
      Height          =   270
      Left            =   2970
      TabIndex        =   18
      Top             =   1320
      Width           =   1110
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "DESCRIPTION"
      Size            =   "1958;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   270
      Left            =   4410
      TabIndex        =   17
      Top             =   885
      Width           =   960
      ForeColor       =   -2147483634
      BackColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "LOCATION"
      Size            =   "1693;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboLocation 
      Height          =   315
      Left            =   5325
      TabIndex        =   1
      Top             =   810
      Width           =   3345
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "5900;556"
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
   Begin MSForms.OptionButton optAll 
      Height          =   360
      Left            =   8820
      TabIndex        =   4
      Top             =   810
      Width           =   615
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483632
      ForeColor       =   -2147483634
      DisplayStyle    =   5
      Size            =   "1085;635"
      Value           =   "1"
      Caption         =   "All"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton optActive 
      Height          =   360
      Left            =   9600
      TabIndex        =   5
      Top             =   810
      Width           =   915
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483632
      ForeColor       =   -2147483634
      DisplayStyle    =   5
      Size            =   "1614;635"
      Value           =   "0"
      Caption         =   "Active"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton optDisuse 
      Height          =   360
      Left            =   10590
      TabIndex        =   6
      Top             =   810
      Width           =   975
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483632
      ForeColor       =   -2147483634
      DisplayStyle    =   5
      Size            =   "1720;635"
      Value           =   "0"
      Caption         =   "Disuse"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   315
      Left            =   4065
      TabIndex        =   3
      Top             =   1275
      Width           =   4575
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "8070;556"
      SpecialEffect   =   3
      FontName        =   "MS Gothic"
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   390
      Left            =   -15
      TabIndex        =   16
      Top             =   90
      Width           =   13695
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Material Master Record"
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
   Begin MSForms.Label lblStatus 
      Height          =   270
      Left            =   105
      TabIndex        =   14
      Top             =   1920
      Width           =   13335
      ForeColor       =   -2147483634
      BackColor       =   8807750
      Size            =   "23521;476"
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1245
      TabIndex        =   2
      Top             =   1275
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2831;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   270
      Left            =   300
      TabIndex        =   13
      Top             =   1320
      Width           =   960
      ForeColor       =   -2147483634
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
   Begin MSForms.Label Label2 
      Height          =   1140
      Left            =   105
      TabIndex        =   11
      Top             =   630
      Width           =   13335
      BackColor       =   8421504
      Size            =   "23521;2011"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1140
      Left            =   195
      TabIndex        =   10
      Top             =   735
      Width           =   13335
      BackColor       =   0
      Size            =   "23521;2011"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   5805
      Left            =   195
      TabIndex        =   15
      Top             =   2010
      Width           =   13335
      BackColor       =   0
      Size            =   "23521;10239"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_Material_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytStatus       As Byte

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdExcel_Click()
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Please wait.... exporting records to excel.")
     Call clsPrintMenu.MaterialList.ExportMaterialToExcel(hflxViewMaterial)
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
End Sub
'***********************temporary only
Private Sub cmdTest_Click()
     Call clsDB.SqlServer(PrintMenuDb)
     
     Dim strSQL As String
     Dim rs As Recordset
     
     strSQL = "Select * from MonthlyStock4 Where Trans_Date = '2004/12/01'"
     
     Set rs = clsDB.GetRecordSet(strSQL, False)
     
     Set hflxViewMaterial.DataSource = rs
     
     Dim lngLoop As Long
     Dim rs1 As Recordset, rs2 As Recordset
     
     hflxViewMaterial.Cols = hflxViewMaterial.Cols + 3
     rs.MoveFirst
     With hflxViewMaterial
     
          For lngLoop = 1 To rs.RecordCount - 1
               
               Set rs1 = clsDB.GetRecordSet("Select qty from GetStockBalance('" & Format(.TextMatrix(lngLoop, 9), "yyyy/mm/dd") & "') Where ItemId = '" & .TextMatrix(lngLoop, 3) & "'")
               
               Set rs2 = clsDB.GetRecordSet("Select Qty from GetStockBalance('" & Format(.TextMatrix(lngLoop, 10), "yyyy/mm/dd") & "') Where ItemId = '" & .TextMatrix(lngLoop, 3) & "'")
               
               If Not rs1.EOF Then
               .TextMatrix(lngLoop, 11) = pfvarIs_Null(rs1.Fields("qty").Value)
               End If
               If Not rs2.EOF Then
               .TextMatrix(lngLoop, 12) = pfvarIs_Null(rs2.Fields("qty").Value)
               End If
          Next
     End With
End Sub
'*********************************

Private Sub cmdView_Click()
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Please wait.... loading records.")
     Call subLoadMaterialRecord
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
     
     Call clsPrintMenu.psubLoadDivision(cboDivision, True)
     Call clsPrintMenu.MaterialList.LoadLocation(cboLocation)
     bytStatus = 0
End Sub

Private Sub hflxViewMaterial_Click()
     If hflxViewMaterial.Row = 1 Then hflxViewMaterial.Sort = 1
End Sub

Private Sub optActive_Click()
     bytStatus = 1
End Sub

Private Sub optAll_Click()
     bytStatus = 0
End Sub

Private Sub optDisuse_Click()
     bytStatus = 2
End Sub

Private Sub txtDescription_Change()
     txtItemCode.Text = ""
End Sub

Private Sub txtItemCode_Change()
     txtDescription.Text = ""
End Sub

Private Sub subLoadMaterialRecord()
     Dim lngRecordcount As Long
     
On Error GoTo lnErrMsg

     lngRecordcount = clsPrintMenu.MaterialList.GetMaterialList( _
                                        hflxViewMaterial, cboLocation, txtItemCode, txtDescription, cboDivision, bytStatus)
     lblStatus.Caption = "Selected  " & lngRecordcount & "  records."

     Call subGridInitialize
     Exit Sub
lnErrMsg:
     MsgBox Err.Description, vbCritical
End Sub

Private Sub subGridInitialize()
     Dim intCol As Integer
     
     With hflxViewMaterial
          For intCol = 0 To .Cols - 1
              .Row = 0
              .Col = intCol
              .CellAlignment = flexAlignCenterCenter
              .ColWidth(intCol) = Choose(intCol + 1, 1200, 5500, 2000, 1000, 2200, 1500, 1500, 2000, 1300, 1500, 1000, 1200, 1000, 1800, 1800, 1800)
              .RowHeight(0) = 350
          Next
          .ColAlignment(1) = flexAlignLeftCenter
          .ColAlignment(2) = flexAlignLeftCenter
     End With
End Sub
