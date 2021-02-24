VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_MachineInventoryRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Inventory Record"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10545
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   1140
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
      MICON           =   "F_MachineInventoryRecord.frx":0000
      PICN            =   "F_MachineInventoryRecord.frx":001C
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
      Left            =   6960
      TabIndex        =   1
      Top             =   660
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
      Format          =   114294785
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   8535
      TabIndex        =   2
      Top             =   660
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
      Format          =   114294785
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdExcel 
      Height          =   375
      Left            =   7695
      TabIndex        =   3
      Top             =   1140
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
      MICON           =   "F_MachineInventoryRecord.frx":05B8
      PICN            =   "F_MachineInventoryRecord.frx":05D4
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
      Left            =   9030
      TabIndex        =   4
      Top             =   1140
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
      MICON           =   "F_MachineInventoryRecord.frx":0B6E
      PICN            =   "F_MachineInventoryRecord.frx":0B8A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxMachineInvRecord 
      Height          =   5160
      Left            =   135
      TabIndex        =   5
      Top             =   2295
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9102
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
   Begin MSForms.TextBox txtMachineName 
      Height          =   315
      Left            =   1395
      TabIndex        =   18
      Top             =   1680
      Width           =   4815
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "8493;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label10 
      Height          =   390
      Left            =   480
      TabIndex        =   17
      Top             =   1640
      Width           =   840
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "MACHINE NAME"
      Size            =   "1482;688"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtMachineID 
      Height          =   315
      Left            =   4080
      TabIndex        =   16
      Top             =   675
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
   Begin MSForms.Label Label9 
      Height          =   270
      Left            =   3000
      TabIndex        =   15
      Top             =   705
      Width           =   960
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "MACHINE ID"
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
      Left            =   6450
      TabIndex        =   12
      Top             =   690
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
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   8340
      TabIndex        =   11
      Top             =   645
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
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1395
      TabIndex        =   10
      Top             =   675
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
   Begin MSForms.Label Label6 
      Height          =   270
      Left            =   405
      TabIndex        =   9
      Top             =   705
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
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   0
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
   Begin MSForms.Label lblDescription 
      Height          =   525
      Left            =   1395
      TabIndex        =   7
      Top             =   1065
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
   Begin MSForms.Label Label3 
      Height          =   5055
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   10215
      BackColor       =   0
      Size            =   "18018;8916"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   1605
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   10200
      BackColor       =   8421504
      Size            =   "17992;2831"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1725
      Left            =   255
      TabIndex        =   14
      Top             =   480
      Width           =   10200
      BackColor       =   0
      Size            =   "17992;3043"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_MachineInventoryRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
     Call connecttoserver
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please Wait.....exporting to excel.")
    Call clsPrintMenu.MachinInvRecord.ExportToExcel(hflxMachineInvRecord)
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdView_Click()
    Call connecttoserver
    If txtItemCode = "" And txtMachineID = "" And txtMachineName = "" Then
        MsgBox "Must Input at least 2 search keys!", vbInformation, pstrMessage
        Exit Sub
    Else
    End If
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait.... loading records")
    hflxMachineInvRecord.Clear
    hflxMachineInvRecord.Rows = 2
    hflxMachineInvRecord.Refresh
    Call subFormatGrid
    Call subViewRecord
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub subFormatGrid()
    Dim intLoop As Integer
    
    With hflxMachineInvRecord
        .Cols = 14
        .RowHeight(0) = 450
        For intLoop = 0 To 13
'               .TextMatrix(0, intLoop) = Choose(intLoop + 1, "TRANSACTION DATE", "TRANSID", "ITEM ID", "REQ QTY", "MACHINE ID", "MACHINE NAME", "REMARKS")
'               .ColWidth(intLoop) = Choose(intLoop + 1, 1500, 1500, 1200, 1200, 1200, 1800, 3000)
               .TextMatrix(0, intLoop) = Choose(intLoop + 1, "REQUEST DATE", "TRANSID", "ITEM ID", "PARTS NAME", "REQ QTY", "REQUESTED CATEGORY", "MACHINE NAME", "MACHINE ID", "IN", "OUT", "AC", "REQUEST DATE HISTORY", "REMARKS", "REVIEW DATE")
               .ColWidth(intLoop) = Choose(intLoop + 1, 1500, 1500, 1200, 1800, 1200, 1500, 1800, 1200, 1200, 1200, 1200, 1500, 3000, 1200)
               .Col = intLoop: .Row = 0
               .CellFontBold = True
               .CellAlignment = flexAlignCenterCenter
               .ColAlignment(intLoop) = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub subViewRecord()
    Dim objMachineInvRecord          As Object
    Dim lngSeqNo              As Long
    
On Error GoTo lnError
     Set objMachineInvRecord = clsPrintMenu.MachinInvRecord.GetMachineInventoryRecord(txtItemCode, txtMachineID, txtMachineName, dtpFrom.Value, dtpTo.Value)
     With objMachineInvRecord
          Do While Not .EOF
               lngSeqNo = lngSeqNo + 1
               hflxMachineInvRecord.Rows = lngSeqNo + 1
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 0) = .Fields("TransactedDate").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 1) = .Fields("TransId").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 2) = .Fields("ItemId").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 3) = .Fields("Description").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 4) = .Fields("Qty").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 6) = .Fields("MachineName").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 7) = .Fields("MachineID").Value
               hflxMachineInvRecord.TextMatrix(lngSeqNo, 12) = pfvarIs_Null(.Fields("Remarks").Value)
               
               .MoveNext
          Loop
     End With
     Set objMachineInvRecord = Nothing
     Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description & _
    Chr(13) & "Please try to view again.", vbExclamation, pstrMessage
End Sub


Private Sub Form_Load()
    Call connecttoserver
    Call clsDB.SQLServer(PrintMenuDb, App.Path & "\Print.ini")
    dtpFrom.Value = DateAdd("m", -1, Format(Date, "yyyy/mm/dd"))
    dtpTo.Value = Format(Date, "yyyy/mm/dd")
    
    Call subFormatGrid
    Call disconnecttoserver
End Sub
