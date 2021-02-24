VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_WH_Record 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warehouse"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "F_WH_Record.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11700
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   5580
      TabIndex        =   5
      ToolTipText     =   "view daily item records"
      Top             =   960
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
      MICON           =   "F_WH_Record.frx":0CCA
      PICN            =   "F_WH_Record.frx":0CE6
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
      Left            =   1635
      TabIndex        =   3
      Top             =   1530
      Width           =   1380
      _ExtentX        =   2434
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
      Format          =   41943041
      CurrentDate     =   38212
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxViewWHRecord 
      Height          =   5190
      Left            =   105
      TabIndex        =   14
      Top             =   2430
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   9155
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   3495
      TabIndex        =   4
      Top             =   1530
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
      Format          =   41943041
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdAC 
      Height          =   375
      Left            =   6870
      TabIndex        =   6
      ToolTipText     =   "view item records not including AC"
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Not AC"
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
      MICON           =   "F_WH_Record.frx":1282
      PICN            =   "F_WH_Record.frx":129E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdTotal 
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "view total of item records"
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&TOTAL"
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
      MICON           =   "F_WH_Record.frx":1838
      PICN            =   "F_WH_Record.frx":1854
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdPrint 
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      ToolTipText     =   "export to excel"
      Top             =   1470
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
      MICON           =   "F_WH_Record.frx":1DF0
      PICN            =   "F_WH_Record.frx":1E0C
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
      Left            =   10095
      TabIndex        =   11
      Top             =   1470
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
      MICON           =   "F_WH_Record.frx":23A6
      PICN            =   "F_WH_Record.frx":23C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdByDivision 
      Height          =   375
      Left            =   5580
      TabIndex        =   8
      ToolTipText     =   "view consumption per requester/division"
      Top             =   1470
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "BY &REQ"
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
      MICON           =   "F_WH_Record.frx":295E
      PICN            =   "F_WH_Record.frx":297A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdDetails 
      Height          =   375
      Left            =   6870
      TabIndex        =   9
      ToolTipText     =   "view consumption details"
      Top             =   1470
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&DETAILS"
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
      MICON           =   "F_WH_Record.frx":2F14
      PICN            =   "F_WH_Record.frx":2F30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   720
      Width           =   285
   End
   Begin MSForms.TextBox txtItemCodeUntil 
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   765
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
   Begin MSForms.Label Label8 
      Height          =   390
      Left            =   -15
      TabIndex        =   21
      Top             =   90
      Width           =   11715
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Warehouse Record"
      Size            =   "20664;688"
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
      TabIndex        =   19
      Top             =   2160
      Width           =   11385
      ForeColor       =   -2147483634
      BackColor       =   8807750
      Size            =   "20082;476"
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
      Left            =   300
      TabIndex        =   18
      Top             =   1215
      Width           =   960
      ForeColor       =   -2147483634
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
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   1635
      TabIndex        =   2
      Top             =   1155
      Width           =   3195
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "5636;556"
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
      TabIndex        =   17
      Top             =   1560
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "SHIPPED DATE"
      Size            =   "2064;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   3165
      TabIndex        =   16
      Top             =   1515
      Width           =   330
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "582;635"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1635
      TabIndex        =   0
      Top             =   765
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
   Begin MSForms.Label Label3 
      Height          =   270
      Left            =   300
      TabIndex        =   15
      Top             =   840
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
      Height          =   1350
      Left            =   105
      TabIndex        =   13
      Top             =   630
      Width           =   11385
      BackColor       =   8421504
      Size            =   "20082;2381"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1350
      Left            =   195
      TabIndex        =   12
      Top             =   735
      Width           =   11385
      BackColor       =   0
      Size            =   "20082;2381"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   5475
      Left            =   180
      TabIndex        =   20
      Top             =   2250
      Width           =   11415
      BackColor       =   0
      Size            =   "20135;9657"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_WH_Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdByDivision_Click()
    Call connecttoserver

    Call psubShowStatMsg("Please wait... loading consumption per requester division.")
    Screen.MousePointer = vbHourglass
    
    Call clsPrintMenu.WHRecord.GetConsumptionPerDivision( _
                  hflxViewWHRecord, dtpFrom.Value, dtpTo.Value, cboDivision, txtItemCode, txtItemCodeUntil)
                  
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDetails_Click()
    Call connecttoserver

    Call psubShowStatMsg("Please wait... loading consumption details.")
    Screen.MousePointer = vbHourglass
    Call clsPrintMenu.WHRecord.GetConsumptionDetails( _
               hflxViewWHRecord, dtpFrom.Value, dtpTo.Value, cboDivision, txtItemCode, txtItemCodeUntil)
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdPrint_Click()
    Call connecttoserver


    Dim lngRow As Long, lngCol As Long

    If hflxViewWHRecord.TextMatrix(1, 0) = "" Then
        MsgBox "Nothing to Print!", vbExclamation, pstrMessage
        Exit Sub
    End If
    Call psubShowStatMsg("Please wait...writing in excel")
    Screen.MousePointer = vbHourglass
    Call clsPrintMenu.Utility.OpenExcel                   '  psubOpenExcel
    
    '---write to excel
    With hflxViewWHRecord
          For lngRow = 0 To .Rows - 1
               For lngCol = 0 To .Cols - 1
                    If lngRow = 0 Then clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, lngCol + 1).Font.Bold = True
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, lngCol + 1) = .TextMatrix(lngRow, lngCol)
               Next lngCol
          Next lngRow
          
    Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, hflxViewWHRecord.Cols)
    clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
    clsPrintMenu.Utility.ExcelApp.Visible = True
    Call clsPrintMenu.Utility.CloseExcel
    End With
'    clsPrintMenu.Utility.ExcelWkSheet.Range("A1:H1").HorizontalAlignment = xlCenter
    
    Call clsPrintMenu.Utility.CloseExcel
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdAC_Click()
    Call connecttoserver

    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait... loading records")
    lblStatus.Caption = "VIEW NO AC"
    
    If fblnLoadRecords = False Then
        Call psubHideStatMsg
        Exit Sub
    End If
    Call subGetWHAC
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdTotal_Click()
    Call connecttoserver

    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait... loading records")
    If fblnLoadRecords = False Then
        Call psubHideStatMsg
        Exit Sub
    End If
    Call subGetWHTotal
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdView_Click()
    Call connecttoserver

    Screen.MousePointer = vbHourglass
    
    Call psubShowStatMsg("Please wait.... loading records")
    lblStatus.Caption = "VIEW RECORDS"
    DoEvents
    '--- Get Records and check if no error or not empty
    If fblnLoadRecords = False Then
        Call psubHideStatMsg
        Exit Sub
    End If
    Call subRefreshGrid
    Call subViewWHRecords
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

'--- load details of inventory record
Private Function fblnLoadRecords() As Boolean
    Dim lngSeqNo         As Long
    Dim adoTransView     As Object
    Dim strFields        As String
    
On Error GoTo lnError
    
    With hflxViewWHRecord
        .Clear
        .Rows = 2
        Call subFormatGrid
    End With
    
    Set adoTransView = clsPrintMenu.WHRecord.GetWHRecord(dtpFrom.Value, dtpTo.Value, cboDivision, txtItemCode, txtItemCodeUntil)
    
    '--- Delete all records first
    clsPrintMenu.WHRecord.DeleteWHRecordWT
    With adoTransView
          Do While Not .EOF
               lngSeqNo = lngSeqNo + 1
               Select Case .Fields("TransactionTypeId").Value
                    Case 1
                         strFields = "(SeqNo,TransDate,ItemId,ItemType,Description,InQty,Division)"
                    Case 3
                         strFields = "(SeqNo,TransDate,ItemId,ItemType,Description,OutQty,Division)"
                    Case 4
                         strFields = "(SeqNo,TransDate,ItemId,ItemType,Description,AC,Division)"
               End Select
               '--- save to warehouse record dummy table
               Call clsPrintMenu.WHRecord.SaveToWHRecordWT(strFields, lngSeqNo, pfvarIs_Null(.Fields("TransactedDate").Value) _
                                        , pfvarIs_Null(.Fields("ItemID").Value), pfvarIs_Null(.Fields("ItemType").Value), pfvarIs_Null(.Fields("Description").Value) _
                                        , .Fields("Qty").Value, pfvarIs_Null(.Fields("DivisionID").Value))
               .MoveNext
          Loop
    End With
    fblnLoadRecords = True
    Exit Function
lnError:
    MsgBox Err.Number & "-" & Err.Description & Chr(13) _
    & "Please try to view again.", vbExclamation, pstrMessage
End Function

Private Sub subViewWHRecords()
    Dim adoRSView        As Object
    Dim dblRunBalance    As Double, dblLastBalance As Double
    Dim lngRow           As Long
    Dim strItemId        As String
    
    Set adoRSView = clsPrintMenu.WHRecord.GetWHRecordsInWT
    
    dblLastBalance = clsPrintMenu.pfvarStockBalance(txtItemCode, cboDivision, DateAdd("d", -1, dtpFrom.Value))
    dblRunBalance = dblLastBalance
    lngRow = 0
    
    With hflxViewWHRecord
          .Cols = 9
          Do While Not adoRSView.EOF
                lngRow = lngRow + 1
               .Rows = lngRow + 1
               .TextMatrix(lngRow, 0) = adoRSView.Fields("TransDate").Value
               .TextMatrix(lngRow, 1) = adoRSView.Fields("ItemId").Value
               .TextMatrix(lngRow, 2) = adoRSView.Fields("Description").Value
               .TextMatrix(lngRow, 3) = adoRSView.Fields("ItemType").Value
               .TextMatrix(lngRow, 4) = adoRSView.Fields("Division").Value
               .TextMatrix(lngRow, 5) = pfvarIs_Null(adoRSView.Fields("InQty").Value, False)
               .TextMatrix(lngRow, 6) = pfvarIs_Null(adoRSView.Fields("OutQty").Value, False)
               .TextMatrix(lngRow, 7) = pfvarIs_Null(adoRSView.Fields("AcQty").Value, False)
               '--- check if actual count
               If adoRSView.Fields("AcQty").Value > 0 Then
                    dblRunBalance = adoRSView.Fields("ACQty") + _
                                    pfvarIs_Null(adoRSView.Fields("InQty").Value, False) - _
                                    pfvarIs_Null(adoRSView.Fields("OutQty").Value, False)
               Else
                    If strItemId <> adoRSView.Fields("ItemId").Value Then
                         dblLastBalance = clsPrintMenu.pfvarStockBalance( _
                                             adoRSView.Fields("ItemId").Value, cboDivision, _
                                             DateAdd("d", -1, adoRSView.Fields("TransDate").Value) _
                                          )
                         dblRunBalance = dblLastBalance
                    End If
                    dblRunBalance = dblRunBalance + _
                                   pfvarIs_Null(adoRSView.Fields("InQty").Value, False) - _
                                   pfvarIs_Null(adoRSView.Fields("OutQty").Value, False)
               End If
               .TextMatrix(lngRow, 8) = dblRunBalance
               strItemId = adoRSView.Fields("ItemId").Value
               adoRSView.MoveNext
          Loop
    End With
    Call subFormatGrid
    DoEvents
    Call subFormatGrid
    Set adoRSView = Nothing
End Sub
'---view warehouse record with actual count
Private Sub subGetWHAC()
    Dim adoRSView        As Object
    Dim intLoop          As Integer
    Dim dblLastBalance   As Double, dblRunBalance As Double
    
    Call subRefreshGrid

    Set adoRSView = clsPrintMenu.WHRecord.GetACWHRecordsInWT
    
    If adoRSView.EOF Then
        hflxViewWHRecord.Clear
        hflxViewWHRecord.Rows = 2
        Exit Sub
    End If
    Set hflxViewWHRecord.DataSource = adoRSView
    DoEvents
    Call subFormatGrid
    '--- Compute the onhand quantity from the grid
    With hflxViewWHRecord
          For intLoop = 1 To .Rows - 1
               If .TextMatrix(intLoop, 1) <> .TextMatrix(intLoop - 1, 1) Then
                    dblLastBalance = clsPrintMenu.pfvarStockBalance( _
                                        .TextMatrix(intLoop, 1), cboDivision, DateAdd("d", -1, .TextMatrix(intLoop, 0)))
                    dblRunBalance = dblLastBalance
               End If
               dblRunBalance = dblRunBalance + _
                                        pfvarIs_Null(.TextMatrix(intLoop, 5), False) - _
                                        pfvarIs_Null(.TextMatrix(intLoop, 6), False)
               .TextMatrix(intLoop, 7) = dblRunBalance
          Next intLoop
    End With
    Set adoRSView = Nothing
End Sub

Private Sub subGetWHTotal()
    Dim adoRSTotal       As Object
    Dim intLoop          As Integer
    
    lblStatus.Caption = "VIEW TOTAL"
    Call subRefreshGrid
    
    Set adoRSTotal = clsPrintMenu.WHRecord.GetTotalWHRecordsInWT
    
    If adoRSTotal.EOF Then
        hflxViewWHRecord.Clear
        hflxViewWHRecord.Rows = 2
        Exit Sub
    End If
    
    Set hflxViewWHRecord.DataSource = adoRSTotal
    Call subFormatGrid("Total")
    
    '--- Compute the onhand quantity from the grid
    With hflxViewWHRecord
          For intLoop = 1 To .Rows - 1
               .TextMatrix(intLoop, 6) = clsPrintMenu.pfvarStockBalance(.TextMatrix(intLoop, 0), cboDivision, dtpTo.Value)
          Next intLoop
    End With
    Set adoRSTotal = Nothing
End Sub

Private Sub subRefreshGrid()
    With hflxViewWHRecord
        .Clear
        .Cols = 2
        .Rows = 2
    End With
End Sub

Private Sub Form_Load()
    Call connecttoserver
    Call clsDB.SQLServer(PrintMenuDb, App.Path & "\Print.ini")
    dtpFrom.Value = DateAdd("m", -1, Format(Date, "yyyy/mm/dd"))
    dtpTo.Value = Format(Date, "yyyy/mm/dd")
    Call subFormatGrid
    Call clsPrintMenu.psubLoadDivision(cboDivision)
    Call disconnecttoserver
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Call connecttoserver
     Call clsPrintMenu.Utility.CloseExcel
     Call disconnecttoserver
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtItemCodeUntil.SetFocus
End Sub

Private Sub txtItemCode_LostFocus()
    txtItemCode.Text = UCase(txtItemCode.Text)
    If txtItemCodeUntil.Text = "" Then txtItemCodeUntil.Text = txtItemCode.Text
End Sub

Private Sub txtItemCodeUntil_GotFocus()
    With txtItemCodeUntil
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub subFormatGrid(Optional ByVal strType As String)
    Dim intCol As Integer
        
    Select Case strType
        Case ""
            With hflxViewWHRecord
                If lblStatus.Caption = "VIEW NO AC" Then
                    .Cols = 8
                    .TextMatrix(0, 7) = "ON HAND"
                    Debug.Print .Cols
                Else
                    .Cols = 9
                    .TextMatrix(0, 7) = "AC"
                    .TextMatrix(0, 8) = "ON HAND"
                    Debug.Print .Cols
                End If
                .RowHeight(0) = 450
                For intCol = 0 To 6
                    .TextMatrix(0, intCol) = Choose(intCol + 1, "TRANSACTION DATE", "ITEM ID", "DESCRIPTION", "ITEM TYPE", "DIVISION", "IN QTY", "OUT QTY")
                    .ColWidth(intCol) = Choose(intCol + 1, 1400, 1100, 3000, 1100, 1500, 1000, 1000)
                Next
            End With
        Case "Total"
lnTotal:
            With hflxViewWHRecord
                .Cols = 7
                .RowHeight(0) = 450
                For intCol = 0 To 6
                    .TextMatrix(0, intCol) = Choose(intCol + 1, "ITEM ID", "DESCRIPTION", "ITEM TYPE", "DIVISION", "IN QTY", "OUT QTY", "ON HAND")
                    .ColWidth(intCol) = Choose(intCol + 1, 1100, 3000, 1100, 2000, 1000, 1000, 1000)
                Next
            End With
    End Select
    
    With hflxViewWHRecord
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = flexAlignCenterCenter
            .ColAlignment(intCol) = flexAlignCenterCenter
            .Row = 0
            .Col = intCol
            .CellFontBold = True
        Next
    End With
End Sub
