VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPendingWorkOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Work Order Summary"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkForklift 
      BackColor       =   &H00404000&
      Caption         =   "ForkLift"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   5760
      TabIndex        =   19
      Top             =   120
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetail 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6324
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   16
      FixedCols       =   0
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483645
      GridColorFixed  =   8421504
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   16
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   2325
      TabIndex        =   5
      Top             =   1995
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84672513
      CurrentDate     =   42731
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   4170
      TabIndex        =   6
      Top             =   1995
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   84672513
      CurrentDate     =   42731
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Work Category:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   2100
   End
   Begin MSForms.ComboBox cboWorkcategory 
      Height          =   330
      Left            =   2280
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   2100
   End
   Begin MSForms.ComboBox cboStatus 
      Height          =   330
      Left            =   2280
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   1110
   End
   Begin VB.Label lblTotalRecord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   1215
      TabIndex        =   15
      Top             =   6240
      Width           =   2685
   End
   Begin MSForms.ComboBox cboAbbvr 
      Height          =   330
      Left            =   2280
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abbreviated Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   2100
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   2100
   End
   Begin MSForms.ComboBox cboType 
      Height          =   330
      Left            =   2280
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Company:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2100
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   2280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Verdana"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   3810
      TabIndex        =   8
      Top             =   2040
      Width           =   300
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Receive Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2115
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   645
      Left            =   9960
      TabIndex        =   4
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "2672;1138"
      Picture         =   "frmPendingWorkOrder.frx":0000
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   645
      Left            =   8400
      TabIndex        =   3
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "2672;1138"
      Picture         =   "frmPendingWorkOrder.frx":1052
      Accelerator     =   83
      MouseIcon       =   "frmPendingWorkOrder.frx":20A4
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   645
      Left            =   11520
      TabIndex        =   2
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "2672;1138"
      Picture         =   "frmPendingWorkOrder.frx":30F6
      Accelerator     =   67
      MouseIcon       =   "frmPendingWorkOrder.frx":4148
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Data. Please Wait."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   315
      TabIndex        =   1
      Top             =   3690
      Width           =   14820
   End
End
Attribute VB_Name = "frmPendingWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCompany_Click()
    Call Connect
    Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0), , , , IIf((chkForklift = 1), True, False))
    Call LoadDataToCombo(cboWorkcategory, "MainCategories", cboCompany.Column(0))
    Call Disconnect
End Sub

Private Sub cboType_Click()
    Call Connect
    If cboCompany.Column(0) = "004" And cboType.Text = "MACHINE" Then
        Call LoadDataToCombo(cboAbbvr, "AbbreviatedMachines")
    Else
        Call LoadDataToCombo(cboAbbvr, "AbbreviatedTypes", cboCompany.Column(0), cboType.Column(0))
    End If
    Call Disconnect
End Sub

Private Sub chkForklift_Click()
    Call Connect
    If cboCompany <> "" Then
        Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0), , , , IIf((chkForklift = 1), True, False))
    End If
     Call LoadDataToCombo(cboWorkcategory, "MainCategories", , , , , IIf((chkForklift = 1), True, False))
    Call Disconnect
End Sub

Private Sub cmdClear_Click()
Call Connect
    Call subFormatGrid(flxDetail, "pending")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboStatus, "Status")
Call Disconnect
    Me.cboType.Clear
    Me.cboAbbvr.Clear
    Me.cboWorkcategory.Clear
    Me.lblTotalRecord.Caption = 0
End Sub

Private Sub cmdExcel_Click()
    If flxDetail.TextMatrix(1, 0) = "" Then Exit Sub
    FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
        Else
            MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
    Dim rsData As New ADODB.Recordset
    Dim strSQLwhere As String
    Dim lngRecCnt As Long
    Dim strActionTaken As String
    
    Dim i As Long
    Dim c As Long
    'On Error GoTo ErrHndlr
    
    If cboCompany.Text = "" Then Exit Sub
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    Me.flxDetail.Visible = False
    
Call Connect
DoEvents
    strSQLwhere = ""
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
   
'    If Me.cboCompany.Text <> "" Then
'        strSQLwhere = strSQLwhere & " AND CompanyID = '" & Me.cboCompany.Column(0) & "'"
'    Else
'        MsgBox "Please choose Company", vbOKOnly, "Information"
'        GoTo eExit
'    End If
    
    Set rsData = cls_GetDetails.pfLoadPending(cboAbbvr.Text, dtFrom, dtTo, _
                                            IIf(IsNull(cboCompany) Or cboCompany.Text = "ALL", "", cboCompany), cboType.Text, _
                                            IIf(IsNull(cboStatus), 0, cboStatus), IIf(chkForklift = 1, True, False), cboWorkcategory.Text)
    
    
    If rsData.EOF Then
        MsgBox "No Record found!", vbOKOnly + vbInformation, "System Information"
        Call subFormatGrid(flxDetail, "pending")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    With flxDetail
        rsData.MoveLast
        lngRecCnt = rsData.RecordCount
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rsData.MoveFirst
        For i = 1 To lngRecCnt
            DoEvents
            For c = 0 To .Cols - 1
                DoEvents
                .TextMatrix(i, c) = pfvarNoValue(rsData.Fields(c).Value)
                .Row = i
                .Col = c
                .CellAlignment = flexAlignLeftCenter
            Next c
            rsData.MoveNext
        Next i
        .MergeCells = flexMergeFree
        .MergeCol(3) = True
        .MergeCol(4) = True
    End With
Call Disconnect
    
eExit:
    flxDetail.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsData = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "System Error"
    GoTo eExit
End Sub

Private Sub Form_Load()
    Call Connect
    Me.dtFrom.Value = Date
    Me.dtTo.Value = Date
    Call subFormatGrid(flxDetail, "pending")
    Call LoadDataToCombo(cboCompany, "Companies")
    
    Call LoadDataToCombo(cboStatus, "Status")
    Call Disconnect
End Sub

Private Function exportExcel() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol As Long
    Dim i As Long
    Dim curWO As String
    Dim curCell As Long
    Dim blnJmp As Boolean
    
    Dim vntDup(35) As Variant
    Dim curWOStartRow As Integer
    Dim curWOEndRow As Integer
    Dim curWORowCtr As Integer
    Dim isSameWO As Boolean
    Dim strCurExcelCol As String
    isSameWO = False
    On Error GoTo ErrExcel
    exportExcel = False
    FM_Main.Enabled = False
    flxDetail.Visible = False
    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets("Sheet1")
        With xlSheet
            .Name = "SUMMARY"
            .Range("A1:L1").Merge
            
            .Range("A1").Formula = "WORK ORDER SUMMARY"
            .Range("A1").Font.Name = "Arial Narrow"
            .Range("A1").Font.Size = 20
            .Range("A1").Font.Bold = True
            .Range("A2:L2").Merge
            
            .Range("A2").Formula = "COMPANY : " & IIf(Me.cboCompany.Text = "", "ALL", Me.cboCompany.Text)
            .Range("A2").Font.Name = "Arial Narrow"
            .Range("A2").Font.Size = 10
            .Range("A2").Font.Bold = True
            .Range("A3:L3").Merge
            
            .Range("A3").Formula = "DATE : " & Format(Me.dtFrom.Value, "MMMM DD, YYYY") & " - " & Format(Me.dtTo.Value, "MMMM DD, YYYY")
            .Range("A3").Font.Name = "Arial Narrow"
            .Range("A3").Font.Size = 10
            .Range("A3").Font.Bold = True
           
                    .Range("A" & 6).Formula = "DATE OF REQUEST"
                    .Range("B" & 6).Formula = "DATE RECEIVED"
                    .Range("C" & 6).Formula = "CONTROL NO."
                    .Range("D" & 6).Formula = "DEPARTMENT"
                    .Range("E" & 6).Formula = "SECTION"
                    .Range("F" & 6).Formula = "LOCATION"
                    .Range("G" & 6).Formula = "MACHINE NO."
                    .Range("H" & 6).Formula = "MACHINE NAME"
                    .Range("I" & 6).Formula = "PROBLEM"
                    .Range("J" & 6).Formula = "REQUESTOR"
                    .Range("K" & 6).Formula = "ALTERNATIVE CONTACT PERSON"
                    .Range("L" & 6).Formula = "TEAM LEADER"
                    .Range("M" & 6).Formula = "DATE STARTED"
                    .Range("N" & 6).Formula = "DATE FINISHED"
                    .Range("O" & 6).Formula = "ACTION TAKEN"
                    .Range("P" & 6).Formula = "STATUS"
         End With
                    
            With xlSheet.Range("A6:P6")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 6 ' 50
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 10
                .Font.Bold = True
                .EntireColumn.AutoFit
                .Columns(5).ColumnWidth = 50
            End With
           
            curWORowCtr = 0
            RowCtr = 7
            For intloop = 1 To flxDetail.Rows - 1
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
                Me.Refresh
            With xlSheet
                If flxDetail.TextMatrix(intloop, 0) = curWO Then
                       isSameWO = True
                Else
                        isSameWO = False
                End If
                '----
                 .Range("A" & RowCtr).Formula = IIf(isSameWO, "", flxDetail.TextMatrix(intloop, 0))
                If isSameWO Then
                        .Range("A" & RowCtr).Formula = ""
                        .Range("A" & RowCtr - 1 & ":" & "A" & RowCtr).Merge
                        If flxDetail.TextMatrix(intloop - 1, 1) = flxDetail.TextMatrix(intloop, 1) Then
                                .Range("B" & RowCtr).Formula = ""
                                .Range("B" & RowCtr - 1 & ":" & "B" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 2) = flxDetail.TextMatrix(intloop, 2) Then
                                .Range("C" & RowCtr).Formula = ""
                                .Range("C" & RowCtr - 1 & ":" & "C" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 3) = flxDetail.TextMatrix(intloop, 3) Then
                                .Range("D" & RowCtr).Formula = ""
                                .Range("D" & RowCtr - 1 & ":" & "D" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 4) = flxDetail.TextMatrix(intloop, 4) Then
                                .Range("E" & RowCtr).Formula = ""
                                .Range("E" & RowCtr - 1 & ":" & "E" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 5) = flxDetail.TextMatrix(intloop, 5) Then
                                .Range("F" & RowCtr).Formula = ""
                                .Range("F" & RowCtr - 1 & ":" & "F" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 6) = flxDetail.TextMatrix(intloop, 6) Then
                                .Range("G" & RowCtr).Formula = ""
                                .Range("G" & RowCtr - 1 & ":" & "G" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 7) = flxDetail.TextMatrix(intloop, 7) Then
                                .Range("H" & RowCtr).Formula = ""
                                .Range("H" & RowCtr - 1 & ":" & "H" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 8) Then
                                .Range("I" & RowCtr).Formula = ""
                                .Range("I" & RowCtr - 1 & ":" & "I" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 9) Then
                                .Range("J" & RowCtr).Formula = ""
                                .Range("J" & RowCtr - 1 & ":" & "J" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 10) Then
                                .Range("K" & RowCtr).Formula = ""
                                .Range("K" & RowCtr - 1 & ":" & "K" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 11) Then
                                .Range("L" & RowCtr).Formula = ""
                                .Range("L" & RowCtr - 1 & ":" & "L" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 12) Then
                                .Range("M" & RowCtr).Formula = ""
                                .Range("M" & RowCtr - 1 & ":" & "M" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 13) Then
                                .Range("N" & RowCtr).Formula = ""
                                .Range("N" & RowCtr - 1 & ":" & "N" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 14) Then
                                .Range("O" & RowCtr).Formula = ""
                                .Range("O" & RowCtr - 1 & ":" & "O" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 15) Then
                                .Range("P" & RowCtr).Formula = ""
                                .Range("P" & RowCtr - 1 & ":" & "P" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If

                Else
defVal:
                
                        .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
                        .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
                        .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
                        .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
                        .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
                        .Range("F" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 5)
                        .Range("G" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 6)
                        .Range("H" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 7)
                        .Range("I" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 8)
                        .Range("J" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 9)
                        .Range("K" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 10)
                        .Range("L" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 11)
                        .Range("M" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 12)
                        .Range("N" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 13)
                        .Range("O" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 14)
                        .Range("P" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 15)
                End If
jmp:
            

                curWO = flxDetail.TextMatrix(intloop, 0)
                
                '----------------
                  
                
                '-Insert row
                If flxDetail.Rows - 1 <> 1 Then
                    .Range("A" & RowCtr).WrapText = True
                    .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                    RowCtr = RowCtr + 1
                End If
                    '-
                End With
            Next intloop
            
            '--- Excel Format -----------
            
            
            With xlSheet
                lblMessage.Caption = "Formatting Spreadsheet.."
                .Columns("A:P").EntireColumn.AutoFit
                With .Range("A6:P" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        .Borders.Weight = xlThin
                        '.VerticalAlignment = xlCenter
                        
                        '-Borders
'                        For i = 7 To 12
'                            .Borders(i).Weight = xlThin
'                        Next i
                End With
            End With
        exportExcel = True
        xlApp.Visible = True
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
        
        
        Set xlApp = Nothing
        Exit Function
        
ErrExcel:
        MsgBox Err.Description, vbOKOnly + vbCritical, "System Error"
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        exportExcel = False
End Function
