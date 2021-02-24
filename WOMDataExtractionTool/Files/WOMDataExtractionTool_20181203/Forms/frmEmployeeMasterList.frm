VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEmployeeMasterList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmployee 
      Height          =   5040
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   10755
      _ExtentX        =   18971
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
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   10620
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
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
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   6120
      Width           =   2775
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   645
      Left            =   5880
      TabIndex        =   8
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "2672;1138"
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
      Left            =   120
      TabIndex        =   6
      Top             =   510
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1605
   End
   Begin MSForms.TextBox txtEmployeeName 
      Height          =   285
      Left            =   1830
      TabIndex        =   4
      Top             =   510
      Width           =   3855
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "6800;503"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1605
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   1845
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   3840
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "6773;582"
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
      Height          =   645
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "2672;1138"
      Picture         =   "frmEmployeeMasterList.frx":20A4
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   645
      Left            =   9240
      TabIndex        =   0
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "2672;1138"
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
'        If exportExcel = True Then
'            MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
'        Else
        If exportLibre(flxEmployee, "Employees", lblmessage) = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
        Else
            MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub

Private Function exportExcel() As Boolean
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim i, c, lngRecCnt As Long
    On Error GoTo ErrExcel

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets("Sheet1")
        xlSheet.Name = "EMPLOYEES"
    
    exportExcel = False
    FM_Main.Enabled = False
    flxEmployee.Visible = False
    lblmessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    
    
    With flxEmployee
        For i = 0 To .Rows - 1
        lblmessage.Caption = "Please Wait. Exporting Data to Excel.. (" & i + 1 & " out of " & flxEmployee.Rows - 1 & " row/s)"
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
    Dim strSQLwhere As String
    Dim lngRecCnt As Long
    Dim i As Long
    Dim c As Long
    On Error GoTo ErrHndlr
    
Call Connect
   strSQLwhere = " WHERE EmployeeName like '%" & Me.txtEmployeeName.Text & "%'"
   
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    
    If Me.cboCompany.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND CompanyName = '" & Me.cboCompany.Text & "'"
    End If


    Set rs = cls_GetDetails.pfLoadEmployeeMasterlist(strSQLwhere)
    
    If rs.EOF Then
'        MsgBox "No Record found!"
        Call subFormatGrid(flxEmployee, "employee")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    lblmessage.Caption = "Please Wait. Loading Data.."
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
Call Disconnect

eExit:
    flxEmployee.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rs = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "System Error"
    GoTo eExit
End Sub

Private Sub Form_Load()
Call Connect
    Call subFormatGrid(flxEmployee, "employee")
    Call LoadDataToCombo(cboCompany, "Companies")
Call Disconnect
End Sub

Private Sub cmdClear_Click()
Call Connect
    Call subFormatGrid(flxEmployee, "employee")
    Call LoadDataToCombo(cboCompany, "Companies")
    Me.lblTotalRecord.Caption = 0
    Me.txtEmployeeName.Text = ""
Call Disconnect
End Sub
