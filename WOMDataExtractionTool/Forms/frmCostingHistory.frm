VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCostingHistory 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Costing and History"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   18885
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   270
      TabIndex        =   10
      Top             =   540
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
      Format          =   21495809
      CurrentDate     =   42731
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   540
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
      Format          =   21495809
      CurrentDate     =   42731
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetail 
      Height          =   5460
      Left            =   0
      TabIndex        =   4
      Top             =   2655
      Width           =   18840
      _ExtentX        =   33232
      _ExtentY        =   9631
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   29
      FixedCols       =   0
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      BackColorSel    =   4210688
      ForeColorSel    =   16777215
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
      _Band(0).Cols   =   29
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
      Left            =   90
      TabIndex        =   16
      Top             =   8280
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
      Left            =   1260
      TabIndex        =   15
      Top             =   8280
      Width           =   2685
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
      TabIndex        =   14
      Top             =   5535
      Width           =   14820
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Received Date"
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
      Left            =   315
      TabIndex        =   13
      Top             =   90
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   870
      Left            =   90
      Top             =   225
      Width           =   3660
   End
   Begin VB.Label Label1 
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
      TabIndex        =   12
      Top             =   585
      Width           =   300
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   885
      Left            =   6075
      TabIndex        =   9
      Top             =   180
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "1826;1561"
      Picture         =   "frmCostingHistory.frx":0000
      Accelerator     =   67
      MouseIcon       =   "frmCostingHistory.frx":1052
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   1620
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1170
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
      Left            =   1620
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1485
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
   Begin MSForms.ComboBox cboSection 
      Height          =   330
      Left            =   1620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
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
   Begin VB.Label Label2 
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
      Left            =   45
      TabIndex        =   3
      Top             =   1800
      Width           =   1605
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   885
      Left            =   3915
      TabIndex        =   1
      Top             =   180
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "1826;1561"
      Picture         =   "frmCostingHistory.frx":20A4
      Accelerator     =   83
      MouseIcon       =   "frmCostingHistory.frx":30F6
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
      Left            =   4995
      TabIndex        =   0
      Top             =   180
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "1826;1561"
      Picture         =   "frmCostingHistory.frx":4148
      Accelerator     =   69
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
      Left            =   45
      TabIndex        =   8
      Top             =   1170
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
      Left            =   45
      TabIndex        =   6
      Top             =   1485
      Width           =   1605
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      Height          =   870
      Left            =   90
      Top             =   225
      Width           =   3675
   End
End
Attribute VB_Name = "frmCostingHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSM As Object
Private Function exportExcel() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol As Long
    Dim i As Long
    
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
        
            .Range("A1:I1").Merge
            .Range("A1").Formula = "MAINTENANCE WORK ORDER CONTROL NUMBER LOGSHEET"
            .Range("A1").Font.Name = "Arial Narrow"
            .Range("A1").Font.Size = 12
            .Range("A1").Font.Bold = True
            .Range("A2:I2").Merge
            .Range("A2").Formula = "COMPANY : " & Me.cboCompany.Text
            .Range("A2").Font.Name = "Arial Narrow"
            .Range("A2").Font.Size = 10
            .Range("A2").Font.Bold = True
            .Range("A3:I3").Merge
            .Range("A3").Formula = "MAINT. SECTION  : " & Me.cboDepartment.Text
            .Range("A3").Font.Name = "Arial Narrow"
            .Range("A3").Font.Size = 10
            .Range("A3").Font.Bold = True
            .Range("A4:I4").Merge
            .Range("A4").Formula = "DATE : " & Format(Me.dtFrom.Value, "MMMM DD, YYYY") & " - " & Format(Me.dtTo.Value, "MMMM DD, YYYY")
            .Range("A4").Font.Name = "Arial Narrow"
            .Range("A4").Font.Size = 10
            .Range("A4").Font.Bold = True
           
                    .Range("A" & 6).Formula = "WO CONTROL NO."
                    .Range("B" & 6).Formula = "DEPARTMENT"
                    .Range("C" & 6).Formula = "SECTION"
                    .Range("D" & 6).Formula = "LINE"
                    .Range("E" & 6).Formula = "CONTROL NO."
                    .Range("F" & 6).Formula = "EQUIPMENT NAME"
                    .Range("G" & 6).Formula = "WORK CATEGORY"
                    .Range("H" & 6).Formula = "MACHINE CLASSIFICATION"
                    .Range("I" & 6).Formula = "PART OF MACHINE"
                    .Range("J" & 6).Formula = "CONDITION/PROBLEM"
                    .Range("K" & 6).Formula = "DATE"
                    .Range("L" & 6).Formula = "TIME"
                    .Range("M" & 6).Formula = "DATE"
                    .Range("N" & 6).Formula = "TIME"
                    .Range("O" & 6).Formula = "DATE"
                    .Range("P" & 6).Formula = "TIME"
                    .Range("Q" & 6).Formula = "DATE"
                    .Range("R" & 6).Formula = "TIME"
                    .Range("S" & 6).Formula = "RESPOND TIME IN MINUTE"
                    .Range("T" & 6).Formula = "ACTION TAKEN"
                    .Range("U" & 6).Formula = "ITEM CODE"
                    .Range("V" & 6).Formula = "MATERIAL NAME"
                    .Range("W" & 6).Formula = "QTY"
                    .Range("X" & 6).Formula = "UNIT COST"
                    .Range("Y" & 6).Formula = "TOTAL COST"
                    .Range("Z" & 6).Formula = "TOTAL EXPENSES"
                    .Range("AA" & 6).Formula = "PREPARED BY"
                    .Range("AB" & 6).Formula = "STATUS"
                    .Range("AC" & 6).Formula = "REMARKS"
                    .Range("AD" & 6).Formula = "# OF MANPOWER AFFECTED OF BREAKDOWN"
                    .Range("AE" & 6).Formula = "TOTAL MINUTES OF BREAKDOWN (DOWNTIME)"
                    .Range("AF" & 6).Formula = "TOTAL MANHOUR LOSS (BREAKDOWN)"
                    .Range("AG" & 6).Formula = "TARGET DATE"
                    .Range("AH" & 6).Formula = "TARGET TIME"
                                     
                    With .Range("U5:Z5")
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
                    With .Range("Q5:R5")
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
                    With .Range("O5:P5")
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
                        .Formula = "STARTED"
                    End With
                    With .Range("M5:N5")
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
                        .Formula = "RESPOND"
                    End With
                    With .Range("K5:L5")
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
                  
                    .Columns("L:L").NumberFormatLocal = "h:mm:ss;@"
                    
                    
            End With
            With xlSheet.Range("A6:AH6")
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
                    .Range("K" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 10), 10)
                    .Range("L" & RowCtr).Formula = IIf(Len(flxDetail.TextMatrix(intloop, 10)) <= 10, "", Right(flxDetail.TextMatrix(intloop, 10), 8))
                    .Range("M" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 11), 10)
                    .Range("N" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 11), 8)
                    .Range("O" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 12), 10)
                    .Range("P" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 12), 8)
                    .Range("Q" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 13), 10)
                    .Range("R" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 13), 8)
                    .Range("S" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 14)
                    .Range("T" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 15)
                    .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
                    .Range("V" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 17)
                    .Range("W" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 18)
                    .Range("X" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 19)
                    .Range("Y" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 20)
                    .Range("Z" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 21)
                    .Range("AA" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 22)
                    .Range("AB" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)
                    .Range("AC" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 24)
                    .Range("AD" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 25)
                    .Range("AE" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 26)
                    .Range("AF" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 27)
                    .Range("AG" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 28), 10)
                    .Range("AH" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 28), 8)
                    
                    
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
                .Columns("A:AH").EntireColumn.AutoFit
                With .Range("A6:AH" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        
                        '-Borders
                        For i = 7 To 12
                            .Borders(i).Weight = xlThin
                            .Borders(i).LineStyle = xlContinuous
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
Private Sub cboCompany_Click()
    Call LoadDataToCombo(cboDepartment, "Departments", cboCompany.Column(0))
    cboSection.Clear
End Sub

Private Sub cboDepartment_Click()
    Call LoadDataToCombo(cboSection, "Sections", cboDepartment.Column(0))
End Sub

Private Sub cmdClear_Click()
    Call subFormatGrid(flxDetail, "costing")
    Call LoadDataToCombo(cboCompany, "Companies")
    Me.cboSection.Clear
    Me.cboDepartment.Clear
    Me.lblTotalRecord.Caption = 0
End Sub

Private Sub cmdExcel_Click()
       If flxDetail.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "WODataExtractionTool"
'        ElseIf exportLibre = True Then
'            MsgBox "Report Succesfully saved to LibreOffice!", vbInformation, "WODataExtractionTool"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub



Private Function Love(objMe As Object, objYou As Object)
    
End Function


Private Sub cmdSearch_Click()
    If Me.cboCompany.Text = "" Or Me.cboDepartment.Text = "" Then
        MsgBox "Please complete the fields."
        Exit Sub
    End If
    Dim rsCosting As New ADODB.Recordset
    Dim strSQLWhere As String
    Dim lngRecCnt As Long
    Dim i As Long
    Dim c As Long
    On Error GoTo ErrHndlr
    
    
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    
    strSQLWhere = ""
    strSQLWhere = strSQLWhere & " AND CostingAndHistoryView.CompanyID = '" & Me.cboCompany.Column(0) & "'"
    strSQLWhere = strSQLWhere & " AND CostingAndHistoryView.DepartmentName = '" & Me.cboDepartment.Text & "'"
    If Me.cboSection.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND CostingAndHistoryView.SectionName = '" & Me.cboSection.Text & "'"
    End If
    
    
    Set rsCosting = cls_GetDetails.pfLoadCosting(Me.dtFrom.Value, Me.dtTo.Value, strSQLWhere)
    
    If rsCosting.EOF Then
        MsgBox "No Record found!"
        Call subFormatGrid(flxDetail, "costing")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    
    With flxDetail
        .Visible = False
        rsCosting.MoveLast
        lngRecCnt = rsCosting.RecordCount
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rsCosting.MoveFirst
        For i = 1 To lngRecCnt
            For c = 0 To .Cols - 1
                .TextMatrix(i, c) = pfvarNoValue(rsCosting.Fields(c).Value)
                .Row = i
                .Col = c
                .CellAlignment = flexAlignLeftCenter
            Next c
            rsCosting.MoveNext
        Next i
    End With
    
eExit:
    flxDetail.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsCosting = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo eExit
End Sub

Private Sub Form_Load()
    Call subFormatGrid(flxDetail, "costing")
    Call LoadDataToCombo(cboCompany, "Companies")
    
End Sub


