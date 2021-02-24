VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMaintenance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintenance"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdExcel 
      Left            =   14880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxmaintenance 
      Height          =   4650
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   8202
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   19
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
      _Band(0).Cols   =   19
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   2445
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
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
      Format          =   83230721
      CurrentDate     =   42731
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   300
      Left            =   2160
      TabIndex        =   2
      Top             =   2445
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
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
      Format          =   83230721
      CurrentDate     =   42731
   End
   Begin MSForms.OptionButton optReceived 
      Height          =   300
      Left            =   90
      TabIndex        =   21
      Top             =   2040
      Width           =   1875
      BackColor       =   4210688
      ForeColor       =   -2147483643
      DisplayStyle    =   5
      Size            =   "3307;529"
      Value           =   "0"
      Caption         =   "Received Date:"
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton optRequest 
      Height          =   300
      Left            =   1920
      TabIndex        =   20
      Top             =   2040
      Width           =   1875
      BackColor       =   4210688
      ForeColor       =   -2147483643
      DisplayStyle    =   5
      Size            =   "3307;529"
      Value           =   "0"
      Caption         =   "Request Date:"
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label6 
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
      Left            =   90
      TabIndex        =   19
      Top             =   1590
      Width           =   1605
   End
   Begin MSForms.ComboBox cboCategory 
      Height          =   330
      Left            =   1770
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1590
      Width           =   3195
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5644;582"
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
      Left            =   1230
      TabIndex        =   17
      Top             =   7590
      Width           =   2685
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
      Left            =   60
      TabIndex        =   16
      Top             =   7590
      Width           =   1245
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   1770
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   3200
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5644;582"
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
      Left            =   1770
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   3200
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5644;582"
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
      Left            =   1770
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3195
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5644;582"
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
   Begin MSForms.ComboBox cboSection 
      Height          =   330
      Left            =   1770
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Width           =   3200
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5644;582"
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
      TabIndex        =   7
      Top             =   4440
      Width           =   14460
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   645
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "2672;1138"
      Picture         =   "frmMaintenance.frx":0000
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
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "2672;1138"
      Picture         =   "frmMaintenance.frx":1052
      Accelerator     =   83
      MouseIcon       =   "frmMaintenance.frx":20A4
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
      Left            =   8160
      TabIndex        =   4
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "2672;1138"
      Picture         =   "frmMaintenance.frx":30F6
      Accelerator     =   67
      MouseIcon       =   "frmMaintenance.frx":4148
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Left            =   1800
      TabIndex        =   3
      Top             =   2445
      Width           =   300
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
      Left            =   90
      TabIndex        =   15
      Top             =   880
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
      Left            =   90
      TabIndex        =   14
      Top             =   510
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
      Left            =   90
      TabIndex        =   13
      Top             =   150
      Width           =   1605
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
      Left            =   90
      TabIndex        =   12
      Top             =   1230
      Width           =   1605
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCompany_Click()
Call Connect
    Call LoadDataToCombo(cboDepartment, "Departments", cboCompany.Column(0))
    Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0))
    Call LoadDataToCombo(cboCategory, "MainCategories", cboCompany.Column(0))
Call Disconnect
End Sub

Private Sub cboDepartment_Click()
Call Connect
    Call LoadDataToCombo(cboSection, "Sections", cboDepartment.Column(0))
Call Disconnect
End Sub

Private Sub cmdClear_Click()
Call Connect
    Call subFormatGrid(flxmaintenance, "maintenance")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboType, "Types")
Call Disconnect
    Me.cboType.Clear
    Me.cboDepartment.Clear
    Me.cboSection.Clear
    Me.cboCategory.Clear
    Me.lblTotalRecord.Caption = 0
End Sub

Private Sub cmdExcel_Click()
    If flxmaintenance.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
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
    Dim curWO As String
    Dim curCell As Long
    Dim blnJmp As Boolean
    
    
    On Error GoTo ErrExcel
    exportExcel = False
    FM_Main.Enabled = False
    flxmaintenance.Visible = False
    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets("Sheet1")
        With xlSheet
            .Name = "LOG SHEET"
            .Range("A1:V1").Merge
            .Range("A1").Formula = "WORK ORDER LOGSHEET AND HISTORY"
            .Range("A1").Font.Name = "Arial Narrow"
            .Range("A1").Font.Size = 20
            .Range("A1").Font.Bold = True
            .Range("A2:T2").Merge
            .Range("A2").Formula = "COMPANY : " & Me.cboCompany.Text
            .Range("A2").Font.Name = "Arial Narrow"
            .Range("A2").Font.Size = 10
            .Range("A2").Font.Bold = True
            .Range("A3:T3").Merge
            .Range("A3").Formula = "MAINT. DEPARTMENT  : " & IIf(Me.cboDepartment.Text = "", "ALL", Me.cboDepartment.Text)
            .Range("A3").Font.Name = "Arial Narrow"
            .Range("A3").Font.Size = 10
            .Range("A3").Font.Bold = True
            .Range("A4:T4").Merge
            .Range("A4").Formula = "DATE : " & Format(Me.dtFrom.Value, "MMMM DD, YYYY") & " - " & Format(Me.dtTo.Value, "MMMM DD, YYYY")
            .Range("A4").Font.Name = "Arial Narrow"
            .Range("A4").Font.Size = 10
            .Range("A4").Font.Bold = True
           
                    .Range("A" & 6).Formula = "WO CONTROL NO."
                    .Range("B" & 6).Formula = "COMPANY"
                    .Range("C" & 6).Formula = "DEPARTMENT"
                    .Range("D" & 6).Formula = "SECTION"
                    .Range("E" & 6).Formula = "LINE"
                    .Range("F" & 6).Formula = "CONTROL NO."
                    .Range("G" & 6).Formula = "EQUIPMENT NAME"
                    .Range("H" & 6).Formula = "DATE"
                    .Range("I" & 6).Formula = "TIME"
                    .Range("J" & 6).Formula = "DATE"
                    .Range("K" & 6).Formula = "TIME"
                    .Range("L" & 6).Formula = "DATE"
                    .Range("M" & 6).Formula = "TIME"
                    .Range("N" & 6).Formula = "RESPOND TIME IN MINUTE"
                    .Range("O" & 6).Formula = "ITEM CODE"
                    .Range("P" & 6).Formula = "MATERIAL NAME"
                    .Range("Q" & 6).Formula = "QTY"
                    .Range("R" & 6).Formula = "UNIT COST"
                    .Range("S" & 6).Formula = "TOTAL COST"
                    .Range("T" & 6).Formula = "PREPARED BY"
                    .Range("U" & 6).Formula = "STATUS"
                    .Range("V" & 6).Formula = "REMARKS"
                                     
                    With .Range("O5:S5")
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Interior.ColorIndex = 50
                        .Interior.Pattern = xlSolid
                        .Font.Name = "Arial Narrow"
                        .Font.Size = 10
                        .Font.Bold = True
                        For i = 7 To 11
                            .Borders(i).Weight = xlThin
                        Next i
                        .Merge
                        .Formula = "MATERIAL USED"
                    End With
                    With .Range("L5:M5")
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Interior.ColorIndex = 6 ' 50
                        .Interior.Pattern = xlSolid
                        .Font.Name = "Arial Narrow"
                        .Font.Size = 10
                        .Font.Bold = True
                        For i = 7 To 11
                            .Borders(i).Weight = xlThin
                        Next i
                        .Merge
                        .Formula = "FINISHED"
                    End With
                    With .Range("H5:I5")
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Interior.ColorIndex = 6 ' 50
                        .Interior.Pattern = xlSolid
                        .Font.Name = "Arial Narrow"
                        .Font.Size = 10
                        .Font.Bold = True
                        For i = 7 To 11
                            .Borders(i).Weight = xlThin
                        Next i
                        .Merge
                        .Formula = "RECEIVED"
                    End With
                    
                    With .Range("J5:K5")
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Interior.ColorIndex = 6 ' 50
                        .Interior.Pattern = xlSolid
                        .Font.Name = "Arial Narrow"
                        .Font.Size = 10
                        .Font.Bold = True
                        For i = 7 To 11
                            .Borders(i).Weight = xlThin
                        Next i
                        .Merge
                        .Formula = "REQUEST"
                    End With
                  
                  
                    
                    
            End With
            With xlSheet.Range("A6:V6")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 6 ' 50
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 10
                .Font.Bold = True
                
'                For i = 7 To 11
'                    .Borders(i).Weight = xlMedium
'                Next i
                .EntireColumn.AutoFit
            End With
            
            
           
            'ActiveWindow.ScrollRow = 2
            RowCtr = 7
            For intloop = 1 To flxmaintenance.Rows - 1
               
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxmaintenance.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    'ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 1
                    If curWO = flxmaintenance.TextMatrix(intloop, 0) Then
                        
'                        GoTo jmp
                    End If
                    .Range("A" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 0)
                    .Range("B" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 2)
                    .Range("D" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 4)
                    .Range("F" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 5)
                    .Range("G" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 6)
                    .Range("H" & RowCtr).Formula = Left(flxmaintenance.TextMatrix(intloop, 7), 10)
                    .Range("I" & RowCtr).Formula = IIf(Len(flxmaintenance.TextMatrix(intloop, 7)) <= 10, "", Right(flxmaintenance.TextMatrix(intloop, 7), 8))
                    .Range("J" & RowCtr).Formula = Left(flxmaintenance.TextMatrix(intloop, 8), 10)
                    .Range("K" & RowCtr).Formula = IIf(Len(flxmaintenance.TextMatrix(intloop, 8)) <= 10, "", Right(flxmaintenance.TextMatrix(intloop, 8), 8))
                    .Range("L" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 9)
                    .Range("M" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 10)
                    .Range("N" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 11)
                    .Range("O" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 12)
                    .Range("P" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 13)
                    .Range("Q" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 14)
                    .Range("R" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 15)
                    .Range("S" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 16)
                    .Range("T" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 17)
                    .Range("U" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 18)
                  
jmp:
                 
                    curWO = flxmaintenance.TextMatrix(intloop, 0)
                    '----
                    
                    '-Insert row
                    If flxmaintenance.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
                
                    
            Next intloop
            
            '--- Excel Format -----------
            
            
            With xlSheet
                
                .Columns("A:V").EntireColumn.AutoFit
                With .Range("A6:V" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        
                        '-Borders
                        For i = 7 To 12
                            .Borders(i).Weight = 2
                            .Borders(i).LineStyle = 1
                        Next i
                End With
                  .Columns("I:I").NumberFormatLocal = "h:mm;@"
                  .Columns("K:K").NumberFormatLocal = "h:mm;@"
                  .Columns("M:M").NumberFormatLocal = "h:mm;@"
        End With
        Set xlSheet = xlBook.Sheets("Sheet2")
            xlSheet.Name = "SUMMARY"
        With xlSheet
            .Range("A1:J1").Merge
            .Range("A1").Formula = "SUMMARY OF DAILY WORK SCHEDULE"
            .Range("A1").Font.Name = "Arial Narrow"
            .Range("A1").Font.Size = 20
            .Range("A1").Font.Bold = True
            .Range("A2:J2").Merge
            .Range("A2").Formula = "COMPANY : " & Me.cboCompany.Text
            .Range("A2").Font.Name = "Arial Narrow"
            .Range("A2").Font.Size = 10
            .Range("A2").Font.Bold = True
            .Range("A3:J3").Merge
            .Range("A3").Formula = "MAINT. DEPARTMENT  : " & IIf(Me.cboDepartment.Text = "", "ALL", Me.cboDepartment.Text)
            .Range("A3").Font.Name = "Arial Narrow"
            .Range("A3").Font.Size = 10
            .Range("A3").Font.Bold = True
            .Range("A4:J4").Merge
            .Range("A4").Formula = "DATE : " & Format(Me.dtFrom.Value, "MMMM DD, YYYY") & " - " & Format(Me.dtTo.Value, "MMMM DD, YYYY")
            .Range("A4").Font.Name = "Arial Narrow"
            .Range("A4").Font.Size = 10
            .Range("A4").Font.Bold = True
                    .Range("A" & 6).Formula = "DATE"
                    .Range("B" & 6).Formula = "PREPARED BY"
                    .Range("C" & 6).Formula = "WORK ORDER NO."
                    .Range("D" & 6).Formula = "DEPARTMENT"
                    .Range("E" & 6).Formula = "SECTION "
                    .Range("F" & 6).Formula = "LINE"
                    .Range("G" & 6).Formula = "MACHINE CONTROL NO."
                    .Range("H" & 6).Formula = "MACHINE NAME"
                    .Range("I" & 6).Formula = "HOURS"
                    .Range("J" & 6).Formula = "REMARKS"
                 
                                     
                    With .Range("D5:F5")
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Interior.ColorIndex = 50
                        .Interior.Pattern = xlSolid
                        .Font.Name = "Arial Narrow"
                        .Font.Size = 10
                        .Font.Bold = True
                        For i = 7 To 11
                            .Borders(i).Weight = xlThin
                        Next i
                        .Merge
                        .Formula = "LOCATION"
                    End With
            With xlSheet.Range("A6:J6")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 6 ' 50
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 10
                .Font.Bold = True
                
'                For i = 7 To 11
'                    .Borders(i).Weight = xlMedium
'                Next i
                .EntireColumn.AutoFit
            End With
            RowCtr = 7
            For intloop = 1 To flxmaintenance.Rows - 1
               
                lblMessage.Caption = "Creating Summary.. (" & intloop & " out of " & flxmaintenance.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    'ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + 1
                    If curWO = flxmaintenance.TextMatrix(intloop, 0) Then
                        
'                        GoTo jmp
                    End If
                    .Range("A" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 7)
                    .Range("B" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 15)
                    .Range("C" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 0)
                    .Range("D" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 2)
                    .Range("E" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 3)
                    .Range("F" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 4)
                    .Range("G" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 5)
                    .Range("H" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 6)
                    .Range("I" & RowCtr).Formula = IIf(Len(flxmaintenance.TextMatrix(intloop, 7)) <= 10, "", Right(flxmaintenance.TextMatrix(intloop, 7), 8))
                    '.Range("J" & RowCtr).Formula = flxmaintenance.TextMatrix(intloop, 18)
                    
                  

                 
                    curWO = flxmaintenance.TextMatrix(intloop, 0)
                    '----
                    
                    '-Insert row
                    If flxmaintenance.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
                
                    
            Next intloop
            
            '--- Excel Format -----------
            
            
            With xlSheet
                lblMessage.Caption = "Formatting Spreadsheet.."
                .Columns("A:J").EntireColumn.AutoFit
                With .Range("A6:J" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        
                        '-Borders
                        For i = 7 To 12
                            .Borders(i).Weight = 2
                            .Borders(i).LineStyle = 1
                        Next i
                End With
                  .Columns("J:J").NumberFormatLocal = "h:mm;@"
            End With
        End With
        
        
        exportExcel = True
        xlApp.Visible = True
        flxmaintenance.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
  
        Set xlApp = Nothing
        Exit Function
        
ErrExcel:
        flxmaintenance.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        exportExcel = False
End Function

Private Sub cmdSearch_Click()

    If Me.cboCompany.Text = "" Then
        MsgBox "Please select company!", vbInformation, "System Information"
        Exit Sub
    ElseIf Me.cboType.Text = "" Then
        MsgBox "Please select type/category!", vbInformation, "System Information"
        Exit Sub
    End If
    
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    Me.flxmaintenance.Visible = False
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    Call LoadFlex
End Sub

Private Sub Form_Load()
Call Connect
    Me.dtFrom.Value = Date
    Me.dtTo.Value = Date
    Me.optReceived.Value = True
    Call subFormatGrid(flxmaintenance, "maintenance")
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboType, "Types")
    Call LoadDataToCombo(cboCategory, "MainCategories")
Call Disconnect
End Sub

Private Sub LoadFlex()
   Dim rsHistory As New ADODB.Recordset
    Dim strSQLwhere As String
    Dim lngRecCnt As Long
    Dim strActionTaken As String
    Dim rsCol As Integer
    
    Dim i As Long
    Dim c As Long
    On Error GoTo ErrHndlr
    
Call Connect
    strSQLwhere = ""
    
    If Me.cboCompany.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND CompanyID = '" & Me.cboCompany.Column(0) & "'"
    End If
    
    If Me.cboDepartment.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND DepartmentID = " & Me.cboDepartment.Column(0)
    End If
    
    If Me.cboSection.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND SectionID = " & Me.cboSection.Column(0)
    End If
    
    If Me.cboType.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND TypeID = " & Me.cboType.Column(0)
    End If
    
    If Me.cboCategory.Text <> "" Then
        strSQLwhere = strSQLwhere & " AND MainCategoryID = " & Me.cboCategory.Column(0)
    End If
 
    
    Set rsHistory = cls_GetDetails.pfLoadMaintenance(Me.dtFrom.Value, Me.dtTo.Value, strSQLwhere, GetServerName(cboCompany.Column(0)))
    
    If rsHistory.EOF Then
        MsgBox "No Record found!", vbInformation, "System Information"
        Call subFormatGrid(flxmaintenance, "maintenance")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    With flxmaintenance
        rsHistory.MoveLast
        lngRecCnt = rsHistory.RecordCount
        rsCol = rsHistory.Fields.Count - 2
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rsHistory.MoveFirst
        For i = 1 To lngRecCnt
            DoEvents
            For c = 0 To rsCol - 1
                DoEvents
                   If pfvarNoValue(rsHistory.Fields(14).Value) <> "" And c = 24 Then
                    .TextMatrix(i, c) = "Finished"
                   Else
                        .TextMatrix(i, c) = pfvarNoValue(rsHistory.Fields(c).Value)
                   End If
                .Row = i
                .Col = c
                .CellAlignment = flexAlignLeftCenter
            Next c
            rsHistory.MoveNext
        Next i
    End With
Call Disconnect
    
    
eExit:
    flxmaintenance.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsHistory = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "System Error"
    GoTo eExit
End Sub

