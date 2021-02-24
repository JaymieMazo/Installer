VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEmployeeMasterList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee MasterList"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmployee 
      Height          =   5040
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   8890
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   6
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
      _Band(0).Cols   =   6
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
      TabIndex        =   11
      Top             =   3510
      Width           =   14820
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
      Left            =   120
      TabIndex        =   10
      Top             =   6090
      Width           =   1245
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
      Left            =   1290
      TabIndex        =   9
      Top             =   6090
      Width           =   2685
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   885
      Left            =   5160
      TabIndex        =   8
      Top             =   60
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "1826;1561"
      Picture         =   "frmEmployeeMasterList.frx":0000
      Accelerator     =   83
      MouseIcon       =   "frmEmployeeMasterList.frx":1052
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Employee Name:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   390
      Width           =   1605
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1605
   End
   Begin MSForms.TextBox txtEmployeeName 
      Height          =   330
      Left            =   1590
      TabIndex        =   4
      Top             =   390
      Width           =   3255
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5741;582"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1605
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   1605
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
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
   Begin MSForms.CommandButton cmdExcel 
      Height          =   885
      Left            =   7470
      TabIndex        =   1
      Top             =   60
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "1826;1561"
      Picture         =   "frmEmployeeMasterList.frx":20A4
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   885
      Left            =   6300
      TabIndex        =   0
      Top             =   60
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "1826;1561"
      Picture         =   "frmEmployeeMasterList.frx":30F6
      Accelerator     =   67
      MouseIcon       =   "frmEmployeeMasterList.frx":4148
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmEmployeeMasterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit


Private Sub cmdExcel_Click()
    If flxEmployee.TextMatrix(1, 0) = "" Then Exit Sub
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
    flxEmployee.Visible = False
    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    
    
    With flxEmployee
        For i = 0 To .Rows - 1
        lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & i + 1 & " out of " & flxEmployee.Rows - 1 & " row/s)"
            For c = 0 To .Cols - 1
                xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F") & i + 1).Formula = .TextMatrix(i, c)
            Next c
        Next i
    
    End With
    
'
    With xlSheet
        .Columns("A:F").EntireColumn.AutoFit
        .Cells.RowHeight = 15
        .Range("A1:F1").Interior.ColorIndex = 37
        .Range("A1:F1").Interior.Pattern = 1
        With .Range("A1:F" & flxEmployee.Rows)
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
    flxEmployee.Visible = True
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
    On Error GoTo ErrHndlr
    
   strSQLWhere = " WHERE EmployeeName like '%" & Me.txtEmployeeName.Text & "%'"
   
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    
    If Me.cboCompany.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND CompanyName = '" & Me.cboCompany.Text & "'"
    End If


    Set rs = cls_GetDetails.pfLoadEmployeeMasterlist(strSQLWhere)
    
    If rs.EOF Then
'        MsgBox "No Record found!"
        Call subFormatGrid(flxEmployee, "employee")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    
    With flxEmployee
        .Visible = False
        rs.MoveLast
        lngRecCnt = rs.RecordCount
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rs.MoveFirst
        For i = 1 To lngRecCnt
            For c = 0 To .Cols - 1
                .TextMatrix(i, c) = pfvarNoValue(rs.Fields(c).Value)
                .Row = i
                .Col = c
                .CellAlignment = flexAlignLeftCenter
            Next c
            rs.MoveNext
        Next i
    End With
    
eExit:
    flxEmployee.Visible = True
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
    Call subFormatGrid(flxEmployee, "employee")
    Call LoadDataToCombo(cboCompany, "Companies")
End Sub

Private Sub cmdClear_Click()
    Call subFormatGrid(flxEmployee, "employee")
    Call LoadDataToCombo(cboCompany, "Companies")
    Me.lblTotalRecord.Caption = 0
    Me.txtEmployeeName.Text = ""
    
End Sub

