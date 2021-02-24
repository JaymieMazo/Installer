VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMachineItemMasterList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmMachineItemMasterList"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMachineItem 
      Height          =   4980
      Left            =   30
      TabIndex        =   0
      Top             =   2640
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   8784
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   36
      FixedCols       =   0
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
      GridColorFixed  =   8421504
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      _Band(0).Cols   =   36
   End
   Begin MSForms.ComboBox cboSection 
      Height          =   330
      Left            =   2055
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   720
      Width           =   3240
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5715;582"
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
   Begin MSForms.ComboBox cboType 
      Height          =   330
      Left            =   2055
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1050
      Width           =   3240
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5715;582"
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
   Begin MSForms.ComboBox cboDepartment 
      Height          =   330
      Left            =   2055
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   390
      Width           =   3240
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5715;582"
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
   Begin MSForms.ComboBox cboStatus 
      Height          =   330
      Left            =   2055
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3240
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5715;582"
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
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   2055
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   3240
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5715;582"
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
   Begin MSForms.TextBox txtMachineCtrlNo 
      Height          =   330
      Left            =   2055
      TabIndex        =   7
      Top             =   1800
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5715;582"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Section:"
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
      Left            =   210
      TabIndex        =   18
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label5 
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
      Left            =   210
      TabIndex        =   16
      Top             =   1050
      Width           =   1875
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department:"
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
      Left            =   210
      TabIndex        =   15
      Top             =   390
      Width           =   1875
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
      Left            =   210
      TabIndex        =   12
      Top             =   1380
      Width           =   1875
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   885
      Left            =   6660
      TabIndex        =   10
      Top             =   120
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "1826;1561"
      Picture         =   "frmMachineItemMasterList.frx":0000
      Accelerator     =   67
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   885
      Left            =   7830
      TabIndex        =   9
      Top             =   120
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "1826;1561"
      Picture         =   "frmMachineItemMasterList.frx":1052
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
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
      Left            =   210
      TabIndex        =   6
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MACHINE CTRL #:"
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
      Left            =   210
      TabIndex        =   5
      Top             =   1800
      Width           =   1875
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   885
      Left            =   5505
      TabIndex        =   4
      Top             =   105
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "1826;1561"
      Picture         =   "frmMachineItemMasterList.frx":20A4
      Accelerator     =   83
      MouseIcon       =   "frmMachineItemMasterList.frx":30F6
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Left            =   1260
      TabIndex        =   3
      Top             =   7710
      Width           =   2685
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL RECORD:"
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
      Left            =   60
      TabIndex        =   2
      Top             =   7710
      Width           =   1245
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
      Left            =   0
      TabIndex        =   1
      Top             =   3510
      Width           =   14820
   End
End
Attribute VB_Name = "frmMachineItemMasterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Call subFormatGrid(flxMachineItem, "machineitem")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboType, "Types")
    Call LoadDataToCombo(cboStatus, "Status")
    
    Me.cboDepartment.Clear
    Me.cboSection.Clear
    Me.txtMachineCtrlNo.Text = ""
    Me.lblTotalRecord.Caption = 0
End Sub

Private Sub cmdExcel_Click()
     If flxMachineItem.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "WODataExtractionTool"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub
Private Function exportExcel() As Boolean
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i, c, lngRecCnt As Long
    On Error GoTo ErrExcel

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets("Sheet1")
    
    exportExcel = False
    FM_Main.Enabled = False
    flxMachineItem.Visible = False
    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    
    
    With flxMachineItem
        For i = 0 To .Rows - 1
        lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & i + 1 & " out of " & flxMachineItem.Rows - 1 & " row/s)"
            For c = 0 To .Cols - 1
                xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", _
                                                "G", "H", "I", "J", "K", "L", _
                                                "M", "N", "O", "P", "Q", "R", _
                                                "S", "T", "U", "V", "W", "X", _
                                                "Y", "Z", "AA", "AB", "AC", "AD", _
                                                "AE", "AF", "AG", "AH", "AI", "AJ", "AK") & i + 1).Formula = .TextMatrix(i, c)
            Next c
        Next i
    
    End With
    
'
    With xlSheet
        .Columns("A:AK").EntireColumn.AutoFit
        .Cells.RowHeight = 15
        .Range("A1:AK1").Interior.ColorIndex = 37
        .Range("A1:AK1").Interior.Pattern = 1
        With .Range("A1:AK" & flxMachineItem.Rows)
                        .HorizontalAlignment = -4108
                        
                        '-Borders
                        For i = 7 To 12
                            .Borders(i).Weight = 2
                            .Borders(i).LineStyle = 1
                        Next i
                End With
    End With
    exportExcel = True
    xlApp.Visible = True
   
    
ErrExit:
    flxMachineItem.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    Set xlApp = Nothing
    
    Exit Function
    
ErrExcel:
    exportExcel = False
    MsgBox Err.Description, vbCritical, "ERROR -" & Err.Number
    GoTo ErrExit
    
    
End Function

Private Sub cmdSearch_Click()
    Dim rs As New ADODB.Recordset
    Dim strSQLWhere As String
    Dim lngRecCnt As Long
    Dim i As Long
    Dim c As Long
    'On Error GoTo ErrHndlr
    '
    'strSQLWhere = " WHERE Status = " & Me.cboStatus.Column(0)
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    
    If Me.cboCompany.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND CompanyID = '" & Me.cboCompany.Column(0) & "'"
    End If
    If Me.cboDepartment.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND DepartmentID = '" & Me.cboDepartment.Column(0) & "'"
    End If
    If Me.cboSection.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND SectionID = '" & Me.cboSection.Column(0) & "'"
    End If
    If Me.cboType.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND TypeID= '" & Me.cboType.Column(0) & "'"
    End If
    If Me.cboStatus.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND Status = " & Me.cboStatus.Column(0)
    End If
    If Me.txtMachineCtrlNo.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND MachineItemNo = '" & Me.txtMachineCtrlNo.Text & "'"
    End If
    
    Set rs = cls_GetDetails.pfLoadMachineMasterlist(Mid(strSQLWhere, 5))
    
    If rs.EOF Then
'        MsgBox "No Record found!"
        Call subFormatGrid(flxMachineItem, "machineitem")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    
    With flxMachineItem
        .Visible = False
        Me.Refresh
        rs.MoveLast
        lngRecCnt = rs.RecordCount
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rs.MoveFirst
        For i = 1 To lngRecCnt
                'Debug.Print "------------>  " & i
                .TextMatrix(i, 0) = pfvarNoValue(rs.Fields(0).Value)
                .TextMatrix(i, 1) = cls_GetDetails.pfGetsubdetail("Types", "TypeName", "TypeID", rs.Fields(1).Value)
                .TextMatrix(i, 2) = pfvarNoValue(rs.Fields(2).Value)
                .TextMatrix(i, 3) = cls_GetDetails.pfGetsubdetail("Companies", "CompanyName", "CompanyID", rs.Fields(3).Value)
                .TextMatrix(i, 4) = cls_GetDetails.pfGetsubdetail("Departments", "DepartmentName", "DepartmentID", pfvarNoValue(rs.Fields(4).Value), "CompanyID", rs.Fields(3).Value)
                .TextMatrix(i, 5) = cls_GetDetails.pfGetsubdetail("Sections", "SectionName", "SectionID", pfvarNoValue(rs.Fields(5).Value), "DepartmentID", pfvarNoValue(rs.Fields(4).Value))
                .TextMatrix(i, 6) = pfvarNoValue(rs.Fields(6).Value)
                .TextMatrix(i, 7) = pfvarNoValue(rs.Fields(7).Value)
                .TextMatrix(i, 8) = pfvarNoValue(rs.Fields(8).Value)
                .TextMatrix(i, 9) = pfvarNoValue(rs.Fields(9).Value)
                .TextMatrix(i, 10) = pfvarNoValue(rs.Fields(10).Value)
                .TextMatrix(i, 11) = pfvarNoValue(rs.Fields(11).Value)
                .TextMatrix(i, 12) = pfvarNoValue(rs.Fields(12).Value)
                .TextMatrix(i, 13) = cls_GetDetails.pfGetsubdetail("Units", "UnitName", "UnitID", pfvarNoValue(rs.Fields(13).Value))
                .TextMatrix(i, 14) = pfvarNoValue(rs.Fields(14).Value)
                .TextMatrix(i, 15) = pfvarNoValue(rs.Fields(15).Value)
                .TextMatrix(i, 16) = pfvarNoValue(rs.Fields(16).Value)
                .TextMatrix(i, 18) = pfvarNoValue(rs.Fields(17).Value)
                .TextMatrix(i, 19) = pfvarNoValue(rs.Fields(19).Value)
                .TextMatrix(i, 20) = pfvarNoValue(rs.Fields(20).Value)
                .TextMatrix(i, 21) = pfvarNoValue(rs.Fields(21).Value)
                .TextMatrix(i, 22) = pfvarNoValue(rs.Fields(22).Value)
                .TextMatrix(i, 23) = pfvarNoValue(rs.Fields(23).Value)
                .TextMatrix(i, 24) = pfvarNoValue(rs.Fields(24).Value)
                .TextMatrix(i, 25) = pfvarNoValue(rs.Fields(25).Value)
                .TextMatrix(i, 26) = pfvarNoValue(rs.Fields(26).Value)
                .TextMatrix(i, 27) = pfvarNoValue(rs.Fields(27).Value)
                .TextMatrix(i, 28) = cls_GetDetails.pfGetsubdetail("Status", "Status", "StatusID", rs.Fields(28).Value)
                .TextMatrix(i, 29) = pfvarNoValue(rs.Fields(29).Value)
                .TextMatrix(i, 30) = pfvarNoValue(rs.Fields(30).Value)
                .TextMatrix(i, 31) = pfvarNoValue(rs.Fields(31).Value)
                .TextMatrix(i, 32) = pfvarNoValue(rs.Fields(32).Value)
                .TextMatrix(i, 33) = pfvarNoValue(rs.Fields(33).Value)
                .TextMatrix(i, 34) = pfvarNoValue(rs.Fields(34).Value)
                .TextMatrix(i, 35) = pfvarNoValue(rs.Fields(35).Value)
                .Row = i
                                                                                                                                                                                                                                               .Col = c
                .CellAlignment = flexAlignLeftCenter
           
            rs.MoveNext
        Next i
    End With
    
    
    
    
eExit:
    flxMachineItem.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rs = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo eExit
End Sub

Private Sub Form_Load()
    Call subFormatGrid(flxMachineItem, "machineitem")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    
    Call LoadDataToCombo(cboType, "Types")
    Call LoadDataToCombo(cboStatus, "Status")
    
End Sub

Private Sub cboCompany_click()
    Call LoadDataToCombo(cboDepartment, "Departments", cboCompany.Column(0))
End Sub
Private Sub cboDepartment_Click()
    Call LoadDataToCombo(cboSection, "Sections", cboDepartment.Column(0))
End Sub


