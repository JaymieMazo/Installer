VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMachineControlStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Machine Control Status"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   17595
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdExcel 
      Left            =   14670
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetail 
      Height          =   6090
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   17520
      _ExtentX        =   30903
      _ExtentY        =   10742
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   24
      FixedCols       =   0
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483645
      GridColorFixed  =   8421504
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
      _Band(0).Cols   =   24
   End
   Begin MSForms.ComboBox cboSection 
      Height          =   330
      Left            =   1800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   765
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
   Begin MSForms.CommandButton cmdSearch 
      Height          =   645
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "2672;1138"
      Picture         =   "frmMachineControlStatus.frx":0000
      Accelerator     =   83
      MouseIcon       =   "frmMachineControlStatus.frx":1052
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
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   17460
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   645
      Left            =   8280
      TabIndex        =   8
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "2672;1138"
      Picture         =   "frmMachineControlStatus.frx":20A4
      Accelerator     =   67
      MouseIcon       =   "frmMachineControlStatus.frx":30F6
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Left            =   75
      TabIndex        =   7
      Top             =   7965
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
      Left            =   1245
      TabIndex        =   6
      Top             =   7965
      Width           =   2685
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   645
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "2672;1138"
      Picture         =   "frmMachineControlStatus.frx":4148
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboStatus 
      Height          =   330
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1395
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
      Left            =   1800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
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
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
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
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
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
      Left            =   225
      TabIndex        =   13
      Top             =   1080
      Width           =   1605
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
      Left            =   225
      TabIndex        =   11
      Top             =   135
      Width           =   1605
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
      Left            =   225
      TabIndex        =   10
      Top             =   450
      Width           =   1605
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
      Left            =   225
      TabIndex        =   9
      Top             =   1395
      Width           =   1605
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
      Left            =   225
      TabIndex        =   16
      Top             =   765
      Width           =   1605
   End
End
Attribute VB_Name = "frmMachineControlStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCompany_Click()
Call Connect
    Call LoadDataToCombo(cboDepartment, "Departments", cboCompany.Column(0))
    Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0))
Call Disconnect
End Sub

Private Sub cboDepartment_Click()
Call Connect
    Call LoadDataToCombo(cboSection, "Sections", cboDepartment.Column(0))
Call Disconnect
End Sub

Private Sub cmdClear_Click()
Call Connect
    Call subFormatGrid(flxDetail, "machinecontrol")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    'Call LoadDataToCombo(cboType, "Types")
    Me.cboStatus.Clear
    Call subLoadDataToCombo(cboStatus)
Call Disconnect
    Me.cboType.Clear
    Me.cboDepartment.Clear
    Me.cboSection.Clear
    Me.lblTotalRecord.Caption = 0
End Sub

Private Sub cmdExcel_Click()
        If flxDetail.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
'        If exportExcel = True Then
'            MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
'        Else
        If exportLibre(flxDetail, "MACHINE CONTROL STATUS", lblMessage) = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
        Else
             MsgBox " An error occured. Data NOT successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub
Private Function exportExcel() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol As Long
    Dim i As Long
    
    'On Error GoTo ErrExcel
    exportExcel = False
    FM_Main.Enabled = False
    flxDetail.Visible = False
    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets("Sheet1")
        xlSheet.Name = "STATUS"
        With xlSheet
        
            .Range("A1:X1").Merge
            .Range("A1").Formula = "MACHINE CONTROL STATUS"
            .Range("A1").Font.Name = "Arial Narrow"
            .Range("A1").Font.Size = 20
            .Range("A1").Font.Bold = True
            .Range("A2:X2").Merge
            .Range("A2").Formula = "COMPANY : " & Me.cboCompany.Text
            .Range("A2").Font.Name = "Arial Narrow"
            .Range("A2").Font.Size = 10
            .Range("A2").Font.Bold = True
            .Range("A3:X3").Merge
            .Range("A3").Formula = "MAINT. DEPARTMENT  : " & IIf(Me.cboDepartment.Text = "", "ALL", Me.cboDepartment.Text)
            .Range("A3").Font.Name = "Arial Narrow"
            .Range("A3").Font.Size = 10
            .Range("A3").Font.Bold = True
           
           
                    .Range("A" & 6).Formula = "STATUS"
                    .Range("B" & 6).Formula = "MACHINE ITEM NO"
                    .Range("C" & 6).Formula = "MACHINE NAME"
                    .Range("D" & 6).Formula = "COMPANY"
                    .Range("E" & 6).Formula = "DEPARTMENT"
                    .Range("F" & 6).Formula = "SECTION"
                    .Range("G" & 6).Formula = "MAKER"
                    .Range("H" & 6).Formula = "TYPENAME"
                    .Range("I" & 6).Formula = "LOCATION"
                    .Range("J" & 6).Formula = "LINE"
                    .Range("K" & 6).Formula = "CAPACITY"
                    .Range("L" & 6).Formula = "FIXED ASSET NO"
                    .Range("M" & 6).Formula = "PREVENTIVE MAINTENANCE"
                    .Range("N" & 6).Formula = "ENGINE MODEL"
                    .Range("O" & 6).Formula = "ENGINE SERIAL NO"
                    .Range("P" & 6).Formula = "TRANSMISSION"
                    .Range("Q" & 6).Formula = "MAST TYPE"
                    .Range("R" & 6).Formula = "ATTACHMENT TYPE"
                    .Range("S" & 6).Formula = "FRONT TIRE"
                    .Range("T" & 6).Formula = "FRONT TIRE HOLES"
                    .Range("U" & 6).Formula = "ACQUISITION AMOUNT"
                    .Range("V" & 6).Formula = "DATE OF ACQUISITION"
                    .Range("W" & 6).Formula = "DATE OF DISPOSAL"
                    .Range("X" & 6).Formula = "REMARKS"
                
                  
            End With
            With xlSheet.Range("A6:X6")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 6 ' 50
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 10
                .Font.Bold = True
                .EntireColumn.AutoFit
            End With
            
            
           
            
            RowCtr = 7

            For intloop = 1 To flxDetail.Rows - 1
            
                
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    
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
                    .Range("Q" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
                    .Range("R" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 17)
                    .Range("S" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 18)
                    .Range("T" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 19)
                    .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 20)
                    .Range("V" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 21)
                    .Range("W" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 22)
                    .Range("X" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)

                    
                    .Range("B" & RowCtr).Select
                    Selection.NumberFormatLocal = "@"
                    
                    '-Insert row
                    If flxDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
            
            '--- Excel Format -----------
            
            
            With xlSheet
                lblMessage.Caption = "Formatting Spreadsheet.."
                .Columns("A:X").EntireColumn.AutoFit
                With .Range("A6:X" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        
                        '-Borders
                        For i = 7 To 11
                            .Borders(i).Weight = 2
                            .Borders(i).LineStyle = 1
                        Next i
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
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        exportExcel = False

End Function

Private Sub cmdSearch_Click()
    Dim rs As New ADODB.Recordset
    Dim strSQLwhere As String
    Dim lngRecCnt As Long
    Dim i As Long
    Dim c As Long
    On Error GoTo ErrHndlr
Call Connect
    strSQLwhere = " WHERE WorkStatus = " & Me.cboStatus.Column(0)
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    
    If Me.cboCompany.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND CompanyName = '" & Me.cboCompany.Text & "'"
    End If
    If Me.cboDepartment.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND DepartmentName = '" & Me.cboDepartment.Text & "'"
    End If
    If Me.cboSection.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND SectionName = '" & Me.cboSection.Text & "'"
    End If
    If Me.cboType.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND TypeName = '" & Me.cboType.Text & "'"
    End If
    
    Set rs = cls_GetDetails.pfLoadMachineControlStatus(strSQLwhere)
    
    If rs.EOF Then
'        MsgBox "No Record found!"
        Call subFormatGrid(flxDetail, "machinecontrol")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    
    
    With flxDetail
        .Visible = False
        rs.MoveLast
        lngRecCnt = rs.RecordCount
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rs.MoveFirst
'        For i = 1 To lngRecCnt
'            For c = 0 To .Cols - 1
'                .TextMatrix(i, c) = pfvarNoValue(rs.Fields(c).Value)
'                .Row = i
'                .Col = c
'                .CellAlignment = flexAlignLeftCenter
'            Next c
'            rs.MoveNext
'        Next i
        i = 1
        While i <= lngRecCnt
            While c < .Cols - 1
            
                If c = 12 Then
                    .TextMatrix(i, c) = GetPreventiveMaintenance(pfvarNoValue(rs.Fields(c).Value))
                Else
                    .TextMatrix(i, c) = pfvarNoValue(rs.Fields(c).Value)
                End If
                '.TextMatrix(i, c) = pfvarNoValue(rs.Fields(c).Value)
                .Row = i
                .Col = c
                .CellAlignment = flexAlignLeftCenter
                c = c + 1
            Wend
            c = 0
            i = i + 1
            rs.MoveNext
        Wend
    End With
Call Disconnect
eExit:
    flxDetail.Visible = True
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
    Call subFormatGrid(flxDetail, "machinecontrol")
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboType, "Types")
    Call subLoadDataToCombo(cboStatus)
Call Disconnect
End Sub

Private Sub subLoadDataToCombo(cbo As Object)
    Dim i As Long
    With cbo
        For i = 0 To 4
            .AddItem i + 1
            .Column(1, i) = Choose(i + 1, "Dispose", "Transfer", "InActive", "Active", "Not Existing")
        Next i
        cbo.ListIndex = 0
    End With

End Sub
