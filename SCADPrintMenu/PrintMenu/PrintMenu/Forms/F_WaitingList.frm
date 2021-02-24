VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_WaitingList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Waiting List"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14250
   Icon            =   "F_WaitingList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14250
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   345
      Left            =   1710
      TabIndex        =   15
      Top             =   1650
      Width           =   1545
      _ExtentX        =   2725
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
      Format          =   23592961
      CurrentDate     =   38266
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   9075
      TabIndex        =   9
      Top             =   1020
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "F_WaitingList.frx":0CCA
      PICN            =   "F_WaitingList.frx":0CE6
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
      Left            =   10710
      TabIndex        =   10
      Top             =   1020
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "F_WaitingList.frx":1282
      PICN            =   "F_WaitingList.frx":129E
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
      Left            =   12360
      TabIndex        =   11
      Top             =   1620
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "F_WaitingList.frx":183A
      PICN            =   "F_WaitingList.frx":1856
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxWaitingList 
      Height          =   5520
      Left            =   120
      TabIndex        =   12
      Top             =   2325
      Width           =   13950
      _ExtentX        =   24606
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
   Begin MSComCtl2.DTPicker dtpUntil 
      Height          =   345
      Left            =   3570
      TabIndex        =   16
      Top             =   1650
      Width           =   1545
      _ExtentX        =   2725
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
      Format          =   23592961
      CurrentDate     =   38266
   End
   Begin OsenXPCntrl.OsenXPButton oxpExcel 
      Height          =   375
      Left            =   12360
      TabIndex        =   21
      Top             =   1020
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "F_WaitingList.frx":1DF2
      PICN            =   "F_WaitingList.frx":1E0E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox txtPONo 
      Height          =   315
      Left            =   6060
      TabIndex        =   20
      Top             =   1680
      Width           =   2505
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "4419;556"
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
      Height          =   210
      Left            =   5490
      TabIndex        =   19
      Top             =   1710
      Width           =   510
      ForeColor       =   -2147483634
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
   Begin MSForms.Label Label9 
      Height          =   420
      Left            =   390
      TabIndex        =   18
      Top             =   1680
      Width           =   1080
      ForeColor       =   -2147483634
      VariousPropertyBits=   276824083
      Caption         =   "ORDER DATE"
      Size            =   "1905;741"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
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
      Left            =   3330
      TabIndex        =   17
      Top             =   1650
      Width           =   165
   End
   Begin MSForms.TextBox txtSuppierName 
      Height          =   315
      Left            =   3720
      TabIndex        =   14
      Top             =   1275
      Width           =   4845
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "8546;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Image imgUncheck 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   405
      Picture         =   "F_WaitingList.frx":23A8
      Top             =   7785
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgCheck 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   0
      Picture         =   "F_WaitingList.frx":272A
      Top             =   7785
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSForms.Label Label7 
      Height          =   5505
      Left            =   195
      TabIndex        =   13
      Top             =   2460
      Width           =   13950
      BackColor       =   0
      Size            =   "24606;9710"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   5025
      TabIndex        =   8
      Top             =   885
      Width           =   3570
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "6297;556"
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
      Left            =   4260
      TabIndex        =   7
      Top             =   915
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
   Begin MSForms.ComboBox cboSupplierCode 
      Height          =   315
      Left            =   1725
      TabIndex        =   6
      Top             =   1275
      Width           =   1785
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   3
      Size            =   "3149;556"
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
   Begin MSForms.Label Label4 
      Height          =   270
      Left            =   495
      TabIndex        =   5
      Top             =   1290
      Width           =   960
      ForeColor       =   -2147483634
      BackColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "SUPPLIER"
      Size            =   "1693;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   270
      Left            =   495
      TabIndex        =   4
      Top             =   915
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
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1725
      TabIndex        =   3
      Top             =   885
      Width           =   2205
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3889;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   1470
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   13935
      BackColor       =   8421504
      Size            =   "24580;2593"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label8 
      Height          =   390
      Left            =   -30
      TabIndex        =   0
      Top             =   135
      Width           =   14265
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Waiting List"
      Size            =   "25162;688"
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
      Height          =   1470
      Left            =   195
      TabIndex        =   2
      Top             =   765
      Width           =   13950
      BackColor       =   -2147483630
      Size            =   "24606;2593"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_WaitingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                                                                                                                                                                                      Option Explicit

Private Sub cboDivision_LostFocus()
     If cboDivision.Text = "" Then cboDivision.Text = "All"
End Sub

Private Sub cboSupplierCode_Click()
    If cboSupplierCode.Text = "" Then
        txtSuppierName = ""
        Exit Sub
    End If
    '--- get the supplier name
    txtSuppierName = clsPrintMenu.WaitingList.GetSupplierName(cboSupplierCode.Text)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call subGetWaiting
    Call subPrintRR
End Sub

Private Sub subPrintRR()
     Dim adoVirtualSource     As New ADODB.Recordset
     
     Set adoVirtualSource = clsDB.GetRecordSet(clsPrintMenu.WaitingList.SQLWaitingList, False)
     
     Set DR_WaitingList.DataSource = adoVirtualSource
     
     With DR_WaitingList.Sections("Section1")
          .Controls("txtItemId").DataField = "ItemId"
          .Controls("txtDescription").DataField = "Description"
          .Controls("txtDivCode").DataField = "DivCode"
          .Controls("txtPoNo").DataField = "PoNo"
          .Controls("txtSeqNo").DataField = "SeqNo"
          .Controls("txtETA").DataField = "ETA"
          '.Controls("txtQtyExpected").DataField = "QtyExpected"
          .Controls("txtQtyExpected").DataField = "OrderQty"
          .Controls("txtQtyWaiting").DataField = "QtyWaiting"
     End With
     
     DR_WaitingList.Orientation = rptOrientLandscape
     DR_WaitingList.Show vbModal
     
     Set adoVirtualSource = Nothing
End Sub

Private Sub subGetWaiting()
     Dim intRow As Integer
     
     Call clsPrintMenu.WaitingList.DeleteWaiting
     With hflxWaitingList
              For intRow = 1 To .Rows - 1
                    .Col = 0
                    .Row = intRow
                    If .CellPicture = imgCheck.Picture Then
                         Call clsPrintMenu.WaitingList.SavePrintWaitingList(.TextMatrix(intRow, 2), _
                                               .TextMatrix(intRow, 1), .TextMatrix(intRow, 3), _
                                               .TextMatrix(intRow, 4), .TextMatrix(intRow, 5), _
                                               .TextMatrix(intRow, 6), .TextMatrix(intRow, 7), _
                                               .TextMatrix(intRow, 8), .TextMatrix(intRow, 9), _
                                               .TextMatrix(intRow, 10), .TextMatrix(intRow, 11), _
                                               .TextMatrix(intRow, 12))
                    End If
              Next intRow
     End With
End Sub

Private Sub cmdView_Click()
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait.... loading records.")
    hflxWaitingList.Clear
    hflxWaitingList.Rows = 2
    Call subFormatGrid
    '--- Loads the records
    Call subLoadRecords
    oxpExcel.Enabled = True
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    '--- Load first all the division name and supplier id
    Call clsDB.SQLServer(PrintMenuDb, App.Path & "\Print.ini")
    Screen.MousePointer = vbHourglass
 
    Call psubShowStatMsg("Please wait.... loading records.")
    Call clsPrintMenu.psubLoadDivision(cboDivision, True)
    
    Call clsPrintMenu.WaitingList.LoadSupplierID(cboSupplierCode)
    
    Call subFormatGrid
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
End Sub

Private Sub subLoadRecords()
    Dim objWaitingList  As Object
    Dim objExcessWaiting As Object
    Dim lngSeqNo        As Long

On Error GoTo lnError
    Set objWaitingList = clsPrintMenu.WaitingList.GetWaitingList( _
                                   txtItemCode, cboDivision, txtPoNo, cboSupplierCode, txtSuppierName, dtpFrom, dtpUntil)
     
    If objWaitingList.EOF Then _
          MsgBox "No record found.", vbExclamation, App.EXEName: Exit Sub
          
    Call clsPrintMenu.WaitingList.DeleteWaiting
    
    With objWaitingList
        Do Until .EOF
            lngSeqNo = lngSeqNo + 1
               With hflxWaitingList
                    .Rows = lngSeqNo + 1
                    .TextMatrix(lngSeqNo, 1) = objWaitingList.Fields("PoNo").Value
                    .TextMatrix(lngSeqNo, 2) = objWaitingList.Fields("PoDetailSeq").Value
                    .TextMatrix(lngSeqNo, 3) = pfvarIs_Null(objWaitingList.Fields("ItemId").Value)
                    .TextMatrix(lngSeqNo, 4) = pfvarIs_Null(objWaitingList.Fields("Description").Value)
                    .TextMatrix(lngSeqNo, 5) = pfvarIs_Null(objWaitingList.Fields("Division").Value)
                    .TextMatrix(lngSeqNo, 6) = pfvarIs_Null(objWaitingList.Fields("IssuedDate").Value)
                    .TextMatrix(lngSeqNo, 7) = pfvarIs_Null(objWaitingList.Fields("FtryDate").Value)
                    .TextMatrix(lngSeqNo, 8) = pfvarIs_Null(objWaitingList.Fields("OrderQty").Value, False)
                    .TextMatrix(lngSeqNo, 9) = pfvarIs_Null(objWaitingList.Fields("QtyExpected").Value, False)
                    .TextMatrix(lngSeqNo, 10) = pfvarIs_Null(objWaitingList.Fields("QtyOK").Value, False)
                    .TextMatrix(lngSeqNo, 11) = pfvarIs_Null(objWaitingList.Fields("Waiting").Value, False)
                    .TextMatrix(lngSeqNo, 12) = pfvarIs_Null(objWaitingList.Fields("SupplierName").Value)
                    .Col = 0: .Row = lngSeqNo
                     Set .CellPicture = imgUncheck.Picture
                    .CellPictureAlignment = flexAlignCenterCenter
                    .RowHeight(lngSeqNo) = 450
                    .TopRow = lngSeqNo
               End With
               DoEvents
            .MoveNext
        Loop
    End With
    Set objWaitingList = Nothing
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub

Private Sub subFormatGrid()
    
    Dim intCol  As Integer
    
    With hflxWaitingList
        .RowHeight(0) = 450
        '.Cols = 12
        .Cols = 13
        For intCol = 0 To .Cols - 1
            .Row = 0
            .Col = intCol
            .CellFontBold = True
            .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, intCol) = Choose(intCol + 1, "PRINT", "PO NO", "POSEQ", "ITEM ID", "DESCRIPTION", "DIVISION", _
                    "ORDER DATE", "FTRY DATE", "ORDER QTY", "QTY EXPECTED", "RECEIVED QTY", "WAITING QTY", "SUPPLIER NAME")
            .ColWidth(intCol) = Choose(intCol + 1, 650, 1100, 1200, 1100, 2500, 900, 1100, 1100, 1100, 1100, 1100, 1100, 2500)
            If .Col = 4 Or .Col = 12 Then
                .ColAlignment(intCol) = flexAlignLeftCenter
            Else
                .ColAlignment(intCol) = flexAlignCenterCenter
            End If
        Next
     End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Call clsDB.DisConnectPrintMenu
End Sub

Private Sub hflxWaitingList_Click()
    With hflxWaitingList
        If .Row = 1 Then
            .Sort = flexSortStringNoCaseAscending
        End If
        '--- change picture if clicked ---------------------------
        If .Col = 0 And .CellPicture = imgCheck.Picture Then
            Set .CellPicture = imgUncheck.Picture
        ElseIf .Col = 0 And .CellPicture = imgUncheck.Picture Then
            Set .CellPicture = imgCheck.Picture
        End If
        '---------------------------------------------------------
        .CellPictureAlignment = flexAlignCenterCenter
    End With
End Sub

Private Sub oxpExcel_Click()

'=============================Added Export to Excel Function=======================
'===================================Ardie 06/24/09 ================================

     Dim lngLoop As Long
     Dim bytCol As Byte

     Call psubShowStatMsg("Writing to excel....")
     Call clsPrintMenu.Utility.OpenExcel
     
     With clsPrintMenu.Utility.ExcelWkSheet
            
            For lngLoop = 0 To hflxWaitingList.Rows - 1
            
                For bytCol = 1 To 8
                    .Cells(lngLoop + 1, bytCol).NumberFormat = "@"
                Next
                    
                    .Cells(lngLoop + 1, 1) = hflxWaitingList.TextMatrix(lngLoop, 1)
                    .Cells(lngLoop + 1, 2) = hflxWaitingList.TextMatrix(lngLoop, 3)
                    .Cells(lngLoop + 1, 3) = hflxWaitingList.TextMatrix(lngLoop, 4)
                    .Cells(lngLoop + 1, 4) = hflxWaitingList.TextMatrix(lngLoop, 7)
                    .Cells(lngLoop + 1, 5) = hflxWaitingList.TextMatrix(lngLoop, 8)  ''Ardie 06/25/09
                    .Cells(lngLoop + 1, 6) = hflxWaitingList.TextMatrix(lngLoop, 9)
                    .Cells(lngLoop + 1, 7) = hflxWaitingList.TextMatrix(lngLoop, 10)
                    .Cells(lngLoop + 1, 8) = hflxWaitingList.TextMatrix(lngLoop, 11)
                    .Cells(lngLoop + 1, 9) = hflxWaitingList.TextMatrix(lngLoop, 12)
            Next
            
            clsPrintMenu.Utility.ExcelApp.Visible = True
            Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, 9)
            clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
            
            Call clsPrintMenu.Utility.CloseExcel
            
     End With
     
     Call clsPrintMenu.Utility.CloseExcel
     Call psubHideStatMsg

End Sub

Private Sub txtItemCode_LostFocus()
    txtItemCode.Text = UCase(txtItemCode.Text)
End Sub

