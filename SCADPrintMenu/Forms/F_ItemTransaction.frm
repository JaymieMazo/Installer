VERSION 5.00
Object = "{19470658-574F-4873-8267-0CABE72F8F30}#1.0#0"; "OsenControls.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form F_ItemTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Transaction List"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12120
   Icon            =   "F_ItemTransaction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   12120
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   345
      Left            =   2235
      TabIndex        =   0
      Top             =   1080
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   609
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
      CurrentDate     =   38266
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   7905
      TabIndex        =   1
      Top             =   1125
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "F_ItemTransaction.frx":0CCA
      PICN            =   "F_ItemTransaction.frx":0CE6
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
      Left            =   10320
      TabIndex        =   2
      Top             =   1125
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "F_ItemTransaction.frx":1282
      PICN            =   "F_ItemTransaction.frx":129E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxTransactionList 
      Height          =   5520
      Left            =   150
      TabIndex        =   3
      Top             =   1830
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   9737
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   16746632
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
         Size            =   9
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpUntil 
      Height          =   345
      Left            =   4500
      TabIndex        =   4
      Top             =   1080
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   609
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
      CurrentDate     =   38266
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4230
      TabIndex        =   16
      Top             =   630
      Width           =   165
   End
   Begin MSForms.TextBox txtItemCode2 
      Height          =   315
      Left            =   4500
      TabIndex        =   15
      Top             =   660
      Width           =   1875
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3307;556"
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
      Height          =   390
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12135
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Item Transaction List"
      Size            =   "21405;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCode1 
      Height          =   315
      Left            =   2235
      TabIndex        =   11
      Top             =   660
      Width           =   1875
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3307;556"
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
      Left            =   1185
      TabIndex        =   10
      Top             =   720
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
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   6585
      TabIndex        =   9
      Top             =   720
      Width           =   750
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "DIVISION"
      Size            =   "1323;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   7395
      TabIndex        =   8
      Top             =   660
      Width           =   4140
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "7302;556"
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
      Height          =   5505
      Left            =   225
      TabIndex        =   7
      Top             =   1935
      Width           =   11820
      BackColor       =   0
      Size            =   "20849;9710"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   4230
      TabIndex        =   6
      Top             =   1050
      Width           =   165
   End
   Begin MSForms.Label Label9 
      Height          =   210
      Left            =   450
      TabIndex        =   5
      Top             =   1170
      Width           =   1635
      ForeColor       =   -2147483634
      VariousPropertyBits=   276824083
      Caption         =   "TRANSACTION DATE"
      Size            =   "2884;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   1140
      Left            =   150
      TabIndex        =   12
      Top             =   525
      Width           =   11805
      BackColor       =   8421504
      Size            =   "20823;2011"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1140
      Left            =   210
      TabIndex        =   14
      Top             =   600
      Width           =   11820
      BackColor       =   -2147483630
      Size            =   "20849;2011"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_ItemTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdView_Click()
     'Call subGridInitialize
     Call subGetTransaction
     Call subGridInitialize
     DoEvents
'     Call subGetOnhand
End Sub

Private Sub Form_Load()
     dtpFrom.Value = DateAdd("m", -1, pfstrServerDate)
     dtpUntil.Value = pfstrServerDate
     Call psubLoadDivision(cboDivision)
End Sub

Private Sub subGetTransaction()
     Dim strSQL As String
     Dim objTransaction As Object
     
     strSQL = " SELECT ItemDetailsView.*, " _
            & " (Select SUM(ItemShipmentToTransactionView.Qty) " _
                & " From ItemShipmentToTransactionView " _
                & " Where ItemShipmentToTransactionView.TransactedDate >=  '2004/12/07' " _
                & " And ItemShipmentToTransactionView.TransactedDate <= '2005/01/07' " _
                & " And ItemShipmentToTransactionView.ItemId=ItemDetailsView.ItemId " _
                & " And ItemShipmentToTransactionView.Qty > 0) As Consumption, " _
            & " From ItemDetailsView " _
            & " Where ItemDetailsView.DivisionId = " & pfstrQt(pfstrGetDivisionID(cboDivision.Text)) _
            & " ORDER BY     ItemDetailsView.ItemId"
     'Debug.Print strSQL
     Set objTransaction = GetRecordSet(strSQL)
     Set hflxTransactionList.DataSource = objTransaction
End Sub

Private Sub subGridInitialize()
     Dim bytCol As Byte
     
     With hflxTransactionList
          .Cols = 8
          .RowHeight(0) = 300
'          .TextMatrix(0, 7) = "OnHand"
          For bytCol = 0 To 7
               .ColWidth(bytCol) = Choose(bytCol + 1, 1000, 2500, 1000, 1300, 1300, 1300, 1500, 1500)
               .ColAlignment(bytCol) = Choose(bytCol + 1, 4, 1, 4, 4, 7, 4, 7, 7)
               .Row = 0: .Col = bytCol
               .CellAlignment = 4
          Next
     End With
End Sub

Private Sub subGetOnhand()
     Dim lngRow As Long
     
     With hflxTransactionList
          For lngRow = 1 To .Rows - 1
               .TextMatrix(lngRow, 7) = pfvarStockBalance(.TextMatrix(lngRow, 0), cboDivision.Text, dtpUntil.Value)
          Next
     End With
End Sub
