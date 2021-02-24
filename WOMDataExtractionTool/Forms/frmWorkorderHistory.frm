VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmWorkorderAndCostingHistory 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workorder And Costing History"
   ClientHeight    =   11505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11505
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   315
      TabIndex        =   10
      Top             =   675
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
      Format          =   86048769
      CurrentDate     =   42731
   End
   Begin MSComDlg.CommonDialog cdExcel 
      Left            =   14670
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetail 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   4110
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   11880
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   30
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
      _Band(0).Cols   =   30
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   2205
      TabIndex        =   12
      Top             =   675
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
      Format          =   86048769
      CurrentDate     =   42731
   End
   Begin MSForms.TextBox txtMachineCtrl 
      Height          =   315
      Left            =   1660
      TabIndex        =   25
      Top             =   3240
      Width           =   3225
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5689;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboWorkcategory 
      Height          =   330
      Left            =   1665
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2930
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
   Begin MSForms.OptionButton optFinished 
      Height          =   300
      Left            =   1920
      TabIndex        =   21
      Top             =   180
      Width           =   1785
      BackColor       =   4210688
      ForeColor       =   -2147483643
      DisplayStyle    =   5
      Size            =   "3149;529"
      Value           =   "0"
      Caption         =   "Finished Date"
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton optReceived 
      Height          =   300
      Left            =   180
      TabIndex        =   20
      Top             =   180
      Width           =   1755
      BackColor       =   4210688
      ForeColor       =   -2147483643
      DisplayStyle    =   5
      Size            =   "3096;529"
      Value           =   "0"
      Caption         =   "Received Date"
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   1665
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1350
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
      Left            =   1665
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1665
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
      Left            =   1665
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1980
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
      Left            =   1665
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2295
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
      Left            =   1665
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2610
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
      Left            =   5190
      TabIndex        =   17
      Top             =   270
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "1826;1561"
      Accelerator     =   69
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
      TabIndex        =   16
      Top             =   11040
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
      Left            =   105
      TabIndex        =   15
      Top             =   11055
      Width           =   1245
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   885
      Left            =   6270
      TabIndex        =   14
      Top             =   270
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "1826;1561"
      Accelerator     =   67
      MouseIcon       =   "frmWorkorderHistory.frx":0000
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Left            =   90
      TabIndex        =   13
      Top             =   2610
      Width           =   1605
   End
   Begin VB.Label Label4 
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
      Left            =   1845
      TabIndex        =   11
      Top             =   720
      Width           =   300
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
      TabIndex        =   7
      Top             =   1665
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
      TabIndex        =   5
      Top             =   1350
      Width           =   1605
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
      Left            =   45
      TabIndex        =   3
      Top             =   7155
      Width           =   14820
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
      TabIndex        =   2
      Top             =   2295
      Width           =   1605
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   885
      Left            =   4110
      TabIndex        =   1
      Top             =   270
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "1826;1561"
      Picture         =   "frmWorkorderHistory.frx":1052
      Accelerator     =   83
      MouseIcon       =   "frmWorkorderHistory.frx":20A4
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   870
      Left            =   135
      Top             =   360
      Width           =   3660
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000A&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404000&
      BorderWidth     =   5
      Height          =   870
      Left            =   135
      Top             =   360
      Width           =   3675
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
      TabIndex        =   19
      Top             =   1980
      Width           =   1605
   End
   Begin VB.Label Label5 
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
      TabIndex        =   23
      Top             =   2930
      Width           =   1605
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Machine ctrl:"
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
      TabIndex        =   24
      Top             =   3240
      Width           =   1605
   End
End
Attribute VB_Name = "frmWorkorderAndCostingHistory"
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
        
            .Range("A1:I1").Merge
            '.Range("A1").Formula = "MAINTENANCE WORK ORDER CONTROL NUMBER LOGSHEET"
            .Range("A1").Formula = "WORK ORDER AND COSTING HISTORY"
            .Range("A1").Font.Name = "Arial Narrow"
            .Range("A1").Font.Size = 12
            .Range("A1").Font.Bold = True
            .Range("A2:I2").Merge
            .Range("A2").Formula = "COMPANY : " & Me.cboCompany.Text
            .Range("A2").Font.Name = "Arial Narrow"
            .Range("A2").Font.Size = 10
            .Range("A2").Font.Bold = True
            .Range("A3:I3").Merge
            .Range("A3").Formula = "MAINT. DEPARTMENT  : " & Me.cboDepartment.Text
            .Range("A3").Font.Name = "Arial Narrow"
            .Range("A3").Font.Size = 10
            .Range("A3").Font.Bold = True
            .Range("A4:I4").Merge
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
                    .Range("H" & 6).Formula = "WORK CATEGORY"
                    .Range("I" & 6).Formula = "MACHINE CLASSIFICATION"
                    .Range("J" & 6).Formula = "PART OF MACHINE"
                    .Range("K" & 6).Formula = "MachineProblemFound"
                    .Range("L" & 6).Formula = "CONDITION/PROBLEM"
                    .Range("M" & 6).Formula = "DATE"
                    .Range("N" & 6).Formula = "TIME"
                    .Range("O" & 6).Formula = "DATE"
                    .Range("P" & 6).Formula = "TIME"
                    .Range("Q" & 6).Formula = "DATE"
                    .Range("R" & 6).Formula = "TIME"
                    .Range("S" & 6).Formula = "DATE"
                    .Range("T" & 6).Formula = "TIME"
                    .Range("U" & 6).Formula = "RESPOND TIME IN MINUTE"
                    .Range("V" & 6).Formula = "ACTION TAKEN"
                    .Range("W" & 6).Formula = "ITEM CODE"
                    .Range("X" & 6).Formula = "MATERIAL NAME"
                    .Range("Y" & 6).Formula = "QTY"
                    .Range("Z" & 6).Formula = "CURRENCY UNIT"
                    .Range("AA" & 6).Formula = "UNIT COST"
                    .Range("AB" & 6).Formula = "TOTAL COST"
                    .Range("AC" & 6).Formula = "PREPARED BY"
                    .Range("AD" & 6).Formula = "STATUS"
                    .Range("AE" & 6).Formula = "REMARKS"
                    .Range("AF" & 6).Formula = "# OF MANPOWER AFFECTED OF BREAKDOWN"
                    .Range("AG" & 6).Formula = "TOTAL MINUTES OF BREAKDOWN (DOWNTIME)"
                    .Range("AH" & 6).Formula = "TOTAL MANHOUR LOSS (BREAKDOWN)"
                    .Range("AI" & 6).Formula = "TARGET DATE"
                    .Range("AJ" & 6).Formula = "TARGET TIME"
                                     
                    With .Range("W5:AB5")
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
                    With .Range("S5:T5")
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
                        .Formula = "STARTED"
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
                        .Formula = "RESPOND"
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
                        .Formula = "RECEIVED"
                    End With
                  
                    .Columns("M:M").NumberFormatLocal = "h:mm:ss;@"
                    
                    
            End With
            With xlSheet.Range("A6:AI6")
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
            
            
           
            curWORowCtr = 0
            RowCtr = 7
     '    xlApp.Visible = True
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
                        If flxDetail.TextMatrix(intloop - 1, 9) = flxDetail.TextMatrix(intloop, 9) Then
                                .Range("J" & RowCtr).Formula = ""
                                .Range("J" & RowCtr - 1 & ":" & "J" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 10) = flxDetail.TextMatrix(intloop, 10) Then
                                .Range("K" & RowCtr).Formula = ""
                                .Range("K" & RowCtr - 1 & ":" & "K" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 11) = flxDetail.TextMatrix(intloop, 11) Then
                                .Range("L" & RowCtr).Formula = ""
                                .Range("L" & RowCtr - 1 & ":" & "L" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 11) = flxDetail.TextMatrix(intloop, 11) Then
                                .Range("M" & RowCtr).Formula = ""
                                .Range("M" & RowCtr - 1 & ":" & "M" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 12) = flxDetail.TextMatrix(intloop, 12) Then
                                .Range("N" & RowCtr).Formula = ""
                                .Range("N" & RowCtr - 1 & ":" & "N" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 12) = flxDetail.TextMatrix(intloop, 12) Then
                                .Range("O" & RowCtr).Formula = ""
                                .Range("O" & RowCtr - 1 & ":" & "O" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 13) = flxDetail.TextMatrix(intloop, 13) Then
                                .Range("P" & RowCtr).Formula = ""
                                .Range("P" & RowCtr - 1 & ":" & "P" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 13) = flxDetail.TextMatrix(intloop, 13) Then
                                .Range("Q" & RowCtr).Formula = ""
                                .Range("Q" & RowCtr - 1 & ":" & "Q" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 14) = flxDetail.TextMatrix(intloop, 14) Then
                                .Range("R" & RowCtr).Formula = ""
                                .Range("R" & RowCtr - 1 & ":" & "R" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 14) = flxDetail.TextMatrix(intloop, 14) Then
                                .Range("S" & RowCtr).Formula = ""
                                .Range("S" & RowCtr - 1 & ":" & "S" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 15) = flxDetail.TextMatrix(intloop, 15) Then
                                .Range("T" & RowCtr).Formula = ""
                                .Range("T" & RowCtr - 1 & ":" & "T" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 16) = flxDetail.TextMatrix(intloop, 16) Then
                                .Range("U" & RowCtr).Formula = ""
                                .Range("U" & RowCtr - 1 & ":" & "U" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                       
'                        If flxDetail.TextMatrix(intloop - 1, 22) = flxDetail.TextMatrix(intloop, 22) Then
'                                .Range("AA" & RowCtr).Formula = ""
'                                .Range("AA" & RowCtr - 1 & ":" & "AA" & RowCtr).Merge
'                        Else
'                                GoTo defVal
'                        End If
                        If flxDetail.TextMatrix(intloop - 1, 23) = flxDetail.TextMatrix(intloop, 23) Then
                                .Range("AB" & RowCtr).Formula = ""
                                .Range("AB" & RowCtr - 1 & ":" & "AB" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 24) = flxDetail.TextMatrix(intloop, 24) Then
                                .Range("AC" & RowCtr).Formula = ""
                                .Range("AC" & RowCtr - 1 & ":" & "AC" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 25) = flxDetail.TextMatrix(intloop, 25) Then
                                .Range("AD" & RowCtr).Formula = ""
                                .Range("AD" & RowCtr - 1 & ":" & "AD" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 26) = flxDetail.TextMatrix(intloop, 26) Then
                                .Range("AE" & RowCtr).Formula = ""
                                .Range("AE" & RowCtr - 1 & ":" & "AE" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 27) = flxDetail.TextMatrix(intloop, 27) Then
                                .Range("AF" & RowCtr).Formula = ""
                                .Range("AF" & RowCtr - 1 & ":" & "AF" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 28) = flxDetail.TextMatrix(intloop, 28) Then
                                .Range("AG" & RowCtr).Formula = ""
                                .Range("AG" & RowCtr - 1 & ":" & "AG" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 29) = flxDetail.TextMatrix(intloop, 29) Then
                                .Range("AH" & RowCtr).Formula = ""
                                .Range("AH" & RowCtr - 1 & ":" & "AH" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 29) = flxDetail.TextMatrix(intloop, 29) Then
                                .Range("AI" & RowCtr).Formula = ""
                                .Range("AI" & RowCtr - 1 & ":" & "AI" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
                        If flxDetail.TextMatrix(intloop - 1, 30) = flxDetail.TextMatrix(intloop, 30) Then
                                .Range("Aj" & RowCtr).Formula = ""
                                .Range("Aj" & RowCtr - 1 & ":" & "Aj" & RowCtr).Merge
                        Else
                                GoTo defVal
                        End If
'                         If flxDetail.TextMatrix(intloop - 1, 17) = flxDetail.TextMatrix(intloop, 17) Then
'                                .Range("V" & RowCtr).Formula = ""
'                                .Range("V" & RowCtr - 1 & ":" & "V" & RowCtr).Merge
'                        Else
'                                GoTo defVal
'                        End If
'                        If flxDetail.TextMatrix(intloop - 1, 18) = flxDetail.TextMatrix(intloop, 18) Then
'                                .Range("W" & RowCtr).Formula = ""
'                                .Range("W" & RowCtr - 1 & ":" & "W" & RowCtr).Merge
'                        Else
'                                GoTo defVal
'                        End If
                Else
defVal:
                
'                        .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
'                        .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
'                        .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
'                        .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
'                        .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
'                        .Range("F" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 5)
'                        .Range("G" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 6)
'                        .Range("H" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 7)
'                        .Range("I" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 8)
'                        .Range("J" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 9)
'                        .Range("K" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 10)
'                        .Range("L" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 11), 10)
'                        .Range("M" & RowCtr).Formula = IIf(Len(flxDetail.TextMatrix(intloop, 11)) <= 10, "", Right(flxDetail.TextMatrix(intloop, 11), 8))
'                        .Range("N" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 12), 10)
'                        .Range("O" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 12), 8)
'                        .Range("P" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 13), 10)
'                        .Range("Q" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 13), 8)
'                        .Range("R" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 14), 10)
'                        .Range("S" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 14), 8)
'                        .Range("T" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 15)
'                        .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
'
'                        .Range("AB" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)
'                        .Range("AC" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 24)
'                        .Range("AD" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 25)
'                        .Range("AE" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 26)
'                        .Range("AF" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 27)
'                        .Range("AG" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 28)
'                        .Range("AH" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 29), 10)
'                        .Range("AI" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 29), 8)
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
                        .Range("M" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 12), 10)
                        .Range("N" & RowCtr).Formula = IIf(Len(flxDetail.TextMatrix(intloop, 12)) <= 10, "", Right(flxDetail.TextMatrix(intloop, 12), 8))
                        .Range("O" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 13), 10)
                        .Range("P" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 13), 8)
                        .Range("Q" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 14), 10)
                        .Range("R" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 14), 8)
                        .Range("S" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 15), 10)
                        .Range("T" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 15), 8)
                        .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
                        
                        .Range("AB" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)
                        .Range("AC" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 24)
                        .Range("AD" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 25)
                        .Range("AE" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 26)
                        .Range("AF" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 27)
                        .Range("AG" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 28)
                        .Range("AH" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 29)
                        .Range("AI" & RowCtr).Formula = Left(flxDetail.TextMatrix(intloop, 30), 10)
                        .Range("AI" & RowCtr).Formula = Right(flxDetail.TextMatrix(intloop, 30), 8)
                End If
                '.Range("B" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 1) = flxDetail.TextMatrix(intloop, 1), "", flxDetail.TextMatrix(intloop, 1))
                '.Range("C" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 2) = flxDetail.TextMatrix(intloop, 2), "", flxDetail.TextMatrix(intloop, 2))
                '.Range("D" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 3) = flxDetail.TextMatrix(intloop, 3), "", flxDetail.TextMatrix(intloop, 3))
'                .Range("E" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 4) = flxDetail.TextMatrix(intloop, 4), "", flxDetail.TextMatrix(intloop, 4))
'                .Range("F" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 5) = flxDetail.TextMatrix(intloop, 5), "", flxDetail.TextMatrix(intloop, 5))
'                .Range("G" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 6) = flxDetail.TextMatrix(intloop, 6), "", flxDetail.TextMatrix(intloop, 6))
'                .Range("H" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 7) = flxDetail.TextMatrix(intloop, 7), "", flxDetail.TextMatrix(intloop, 7))
'                .Range("I" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 8), "", flxDetail.TextMatrix(intloop, 8))
'                .Range("J" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 9) = flxDetail.TextMatrix(intloop, 9), "", flxDetail.TextMatrix(intloop, 9))
'                .Range("K" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 10) = flxDetail.TextMatrix(intloop, 10), "", flxDetail.TextMatrix(intloop, 10))
'                .Range("L" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 11) = flxDetail.TextMatrix(intloop, 11), "", Left(flxDetail.TextMatrix(intloop, 11), 10))
'                .Range("M" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 11) = flxDetail.TextMatrix(intloop, 11), "", IIf(Len(flxDetail.TextMatrix(intloop, 11)) <= 10, "", Right(flxDetail.TextMatrix(intloop, 11), 8)))
                '.Range("N" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 12) = flxDetail.TextMatrix(intloop, 12), "", Left(flxDetail.TextMatrix(intloop, 12), 10))
              '  .Range("O" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 12) = flxDetail.TextMatrix(intloop, 12), "", Right(flxDetail.TextMatrix(intloop, 12), 8))
               ' .Range("P" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 13) = flxDetail.TextMatrix(intloop, 13), "", Left(flxDetail.TextMatrix(intloop, 13), 10))
             '   .Range("Q" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 13) = flxDetail.TextMatrix(intloop, 13), "", Right(flxDetail.TextMatrix(intloop, 13), 8))
                '.Range("R" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 14) = flxDetail.TextMatrix(intloop, 14), "", Left(flxDetail.TextMatrix(intloop, 14), 10))
               ' .Range("S" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 14) = flxDetail.TextMatrix(intloop, 14), "", Right(flxDetail.TextMatrix(intloop, 14), 8))
               ' .Range("T" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 15) = flxDetail.TextMatrix(intloop, 15), "", flxDetail.TextMatrix(intloop, 15))
               ' .Range("U" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 16) = flxDetail.TextMatrix(intloop, 16), "", flxDetail.TextMatrix(intloop, 16))

               ' .Range("V" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 17) = flxDetail.TextMatrix(intloop, 17), "", flxDetail.TextMatrix(intloop, 17))
               ' .Range("W" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 18) = flxDetail.TextMatrix(intloop, 18), "", flxDetail.TextMatrix(intloop, 18))
                .Range("V" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 17)
                .Range("W" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 18)
                .Range("X" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 19)
                .Range("Y" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 20)
                .Range("Z" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 21)
                .Range("AA" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 22)
               '.Range("AA" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 22) = flxDetail.TextMatrix(intloop, 22), "", flxDetail.TextMatrix(intloop, 22))
'                .Range("AB" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 23) = flxDetail.TextMatrix(intloop, 23), "", flxDetail.TextMatrix(intloop, 23))
'                .Range("AC" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 24) = flxDetail.TextMatrix(intloop, 24), "", flxDetail.TextMatrix(intloop, 24))
'                .Range("AD" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 25) = flxDetail.TextMatrix(intloop, 25), "", flxDetail.TextMatrix(intloop, 25))
'                .Range("AE" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 26) = flxDetail.TextMatrix(intloop, 26), "", flxDetail.TextMatrix(intloop, 26))
'                .Range("AF" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 27) = flxDetail.TextMatrix(intloop, 27), "", flxDetail.TextMatrix(intloop, 27))
'                .Range("AG" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 28) = flxDetail.TextMatrix(intloop, 28), "", flxDetail.TextMatrix(intloop, 28))
'                .Range("AH" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 29) = flxDetail.TextMatrix(intloop, 29), "", Left(flxDetail.TextMatrix(intloop, 29), 10))
'                .Range("AI" & RowCtr).Formula = IIf(isSameWO And flxDetail.TextMatrix(intloop - 1, 29) = flxDetail.TextMatrix(intloop, 29), "", Right(flxDetail.TextMatrix(intloop, 29), 8))
jmp:
            
'                For i = 0 To flxDetail.Cols - 1
'                        vntDup(i) = flxDetail.TextMatrix(intloop, i)
'                        Debug.Print vntDup(i)
'                Next i
                curWO = flxDetail.TextMatrix(intloop, 0)
                
                '----------------
                  
                
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
                .Columns("A:Aj").EntireColumn.AutoFit
                With .Range("A6:Aj" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        
                        '-Borders
                        For i = 7 To 12
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
'    Dim xlApp       As Excel.Application
'    Dim xlBook      As Excel.Workbook
'    Dim xlSheet     As Excel.Worksheet
'
'    Dim strNewFile As String
'    Dim intloop As Long
'    Dim curCol As Long
'    Dim i As Long
'
'    On Error GoTo ErrExcel
'    exportExcel = False
'    FM_Main.Enabled = False
'    flxDetail.Visible = False
'    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
'    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
'    Set xlApp = CreateObject("Excel.Application")
'    Set xlBook = xlApp.Workbooks.Add
'    Set xlSheet = xlBook.Sheets("Sheet1")
'        With xlSheet
'
'            .Range("A1:I1").Merge
'            .Rows("2:2").WrapText = True
'            With .Range("A1")
'                .Font.Name = "Arial Narrow"
'                .Font.Size = 14
'                .Formula = "WORKODER HISTORY (" & Me.dtFrom.Value & " - " & Me.dtTo.Value & ")"
'                .HorizontalAlignment = xlCenter
'                .Font.Bold = True
'            End With
'
'                    .Range("A" & 2).Formula = "NO."
'                    .Range("B" & 2).Formula = "RECIEVED DATE"
'                    .Range("C" & 2).Formula = "COMPANY NAME"
'                    .Range("D" & 2).Formula = "DEPARTMENT NAME"
'                    .Range("E" & 2).Formula = "TYPE NAME"
'                    .Range("F" & 2).Formula = "WORK CATEGORY NAME"
'                    .Range("G" & 2).Formula = "WORK ORDER CONTROL #"
'                    .Range("H" & 2).Formula = "MACHINE ITEM #"
'                    .Range("I" & 2).Formula = "STATUS"
'
'
'            End With
'            With xlSheet.Range("A2:I2")
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlCenter
'                .Interior.ColorIndex = 35
'                .Interior.Pattern = xlSolid
'                .Font.Name = "Arial Narrow"
'                .Font.Size = 9
'                .Font.Bold = True
'                For i = 7 To 11
'                    .Borders(i).Weight = xlMedium
'                Next i
'            End With
'            RowCtr = 3
'            For intloop = 1 To flxDetail.Rows - 1
'
'                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
'                Me.Refresh
'                With xlSheet
'                    .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
'                    .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
'                    .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
'                    .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
'                    .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
'                    .Range("F" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 5)
'                    .Range("G" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 6)
'                    .Range("H" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 7)
'                    .Range("I" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 8)
'
'                    '-Insert row
'                    If flxDetail.Rows - 1 <> 1 Then
'                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
'                        RowCtr = RowCtr + 1
'                    End If
'                    '-
'                End With
'            Next intloop
'
'            '--- Excel Format -----------
'
'
'            With xlSheet
'                lblMessage.Caption = "Formatting Spreadsheet.."
'                .Rows("2:" & RowCtr - 1).EntireRow.AutoFit
'                With .Range("A3:I" & RowCtr - 1)
'                        .HorizontalAlignment = xlCenter
'                        '.VerticalAlignment = xlCenter
'                        .WrapText = True
'                        '-Borders
'                        For i = 7 To 12
'                            .Borders(i).Weight = xlThin
'                            .Borders(i).LineStyle = xlContinuous
'                        Next i
'                End With
'                .Range("B:B").NumberFormatLocal = "yyyy/m/d"
'        End With
'        exportExcel = True
'        xlApp.Visible = True
'        flxDetail.Visible = True
'        FM_Main.StatusBar1.Panels(3).Text = ""
'        FM_Main.Enabled = True
'
'        Set xlApp = Nothing
'        Exit Function
'
'ErrExcel:
'        flxDetail.Visible = True
'        FM_Main.StatusBar1.Panels(3).Text = ""
'        FM_Main.Enabled = True
'        exportExcel = False
End Function

Private Sub cboCompany_click()
    Call LoadDataToCombo(cboDepartment, "Departments", cboCompany.Column(0))
End Sub

Private Sub cboDepartment_Click()
    Call LoadDataToCombo(cboSection, "Sections", cboDepartment.Column(0))
End Sub


 

Private Sub cmdClear_Click()
    Call subFormatGrid(flxDetail, "costing")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboType, "Types")
    
    Call LoadDataToCombo(cboStatus, "Status")
    Call LoadDataToCombo(cboWorkcategory, "WorkCategory")
    
    Me.cboDepartment.Clear
    Me.cboSection.Clear
    Me.lblTotalRecord.Caption = 0
End Sub

Private Sub cmdExcel_Click()
     If flxDetail.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "WODataExtractionTool"
        ElseIf exportLibre = True Then
            MsgBox "Report Succesfully saved to LibreOffice!", vbInformation, "WODataExtractionTool"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub
Private Function MakePropertyValue(propName, propVal) As Object
    Dim oPropValue As Object
    Set oPropValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    oPropValue.Name = propName
    oPropValue.Value = propVal
    Set MakePropertyValue = oPropValue
End Function


Private Function exportLibre() As Boolean
     Dim oDoc As Object, _
        oDesk As Object, _
        oSheet As Object, _
        oPar(1) As Object, _
        CellProp As Object, _
        CellStyle As Object, _
        NewStyle As Object, _
        oRange As Object, _
        oColumns As Object, _
        PageStyles As Object, _
        NewPageStyle As Object, _
        StdPage As Object, _
        basicBorder As Object, _
        oBorder As Object
    Dim oCol As Long, oRow As Long
    Dim Charts As Object
    Dim Chart As Object
    Dim Rect As Object
    Dim oChartRange As Object
    Dim RangeAddress(0) As Object
    
    exportLibre = False
    oRow = 0
    
    Set oSM = CreateObject("com.sun.star.ServiceManager")
    Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
    Set oPar(0) = MakePropertyValue("Hidden", True)
    Set oPar(1) = MakePropertyValue("Overwrite", True)
    Set oDoc = oDesk.LoadComponentFromURL("private:factory/scalc", "_blank", 0, oPar)
    Set oSheet = oDoc.Sheets.getByIndex("0")
    Set CellProp = oDoc.StyleFamilies.getByName("CellStyles")
    Set NewStyle = oDoc.createInstance("com.sun.star.style.CellStyle")
    Call CellProp.InsertbyName("MyStyle", NewStyle)
    NewStyle.ParentStyle = "Default"
    Set CellStyle = CellProp.getByName("MyStyle")
    Set PageStyles = oDoc.StyleFamilies.getByName("PageStyles")
    Set StdPage = PageStyles.getByName("Default")
    Set basicBorder = oDoc.Bridge_GetStruct("com.sun.star.table.BorderLine")
    
    basicBorder.Color = RGB(0, 0, 0)
    basicBorder.InnerLineWidth = 0
    basicBorder.OuterLineWidth = 11
    basicBorder.LineDistance = 0
    
'        With StdPage
'            .FooterIsOn = False
'            .HeaderIsOn = False
'            .IsLandscape = False
'            .Width = 29700
'            .Height = 21000
'            .LeftMargin = 1000
'            .RightMargin = 1000
'            .TopMargin = 1000
'            .BottomMargin = 1000
'        End With
        '-Header--
        
            FM_Main.Enabled = False
            flxDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Libre Office.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Libre Office.."
        oSheet.getCellByPosition(0, oRow).String = "WORKODER HISTORY (" & Me.dtFrom.Value & " - " & Me.dtTo.Value & ")"
         
                
        For oCol = 0 To flxDetail.Cols - 1
           With oSheet
                With CellStyle
                    .CharWeight = 300
                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .CellBackColor = RGB(255, 184, 0)
                End With
                .getCellByPosition(oCol, 1).CellStyle = "MyStyle"
                .getCellByPosition(oCol, 1).String = flxDetail.TextMatrix(0, oCol)
                .getCellByPosition(oCol, 1).HoriJustify = 2
           End With
            Set oBorder = oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder
            oBorder.LeftLine = basicBorder
            oBorder.Topline = basicBorder
            oBorder.RightLine = basicBorder
            oBorder.BottomLine = basicBorder
            oBorder.VerticalLine = basicBorder
            oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder = oBorder
        Next oCol
    
        '-Content--
        For oRow = 0 To flxDetail.Rows - 1
            DoEvents
            
            lblMessage.Caption = "Exporting Data : " & oRow & " rows out of " & flxDetail.Rows - 1 & " rows"
            For oCol = 0 To flxDetail.Cols - 1
               With oSheet
                    .getCellByPosition(oCol, oRow + 1).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .getCellByPosition(oCol, oRow + 1).HoriJustify = 2
                    flxDetail.Row = oRow
                    flxDetail.Col = oCol
                    .getCellByPosition(oCol, oRow + 1).String = flxDetail.TextMatrix(oRow, oCol)
                    If flxDetail.CellBackColor = &H8080FF Then
                        .getCellByPosition(oCol, oRow + 1).CellBackColor = RGB(128, 128, 255)
                    End If
               End With
            Next oCol
                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder
                oBorder.LeftLine = basicBorder
                oBorder.Topline = basicBorder
                oBorder.RightLine = basicBorder
                oBorder.BottomLine = basicBorder
                oBorder.VerticalLine = basicBorder
                oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder = oBorder
        Next oRow
     
        
        
        Call oDoc.storeToURL("file:///C:/Exported.xls", oPar)
        Set oPar(0) = MakePropertyValue("Hidden", False)
        Set oDoc = oDesk.LoadComponentFromURL("file:///C:/Exported.xls", "_blank", 0, oPar)
        exportLibre = True
    
    
        Set oSM = Nothing
        Set oDesk = Nothing
        Set oDoc = Nothing
        Set oSheet = Nothing
        Set oPar(1) = Nothing
        Set CellProp = Nothing
        Set CellStyle = Nothing
        Set NewStyle = Nothing
        Set oRange = Nothing
        Set oColumns = Nothing
        Set PageStyles = Nothing
        Set NewPageStyle = Nothing
        Set StdPage = Nothing
        exportLibre = True
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        Exit Function
ErrLibre:
        exportLibre = False
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
End Function
Private Sub cmdSearch_Click()
    Dim rsHistory As New ADODB.Recordset
    Dim strSQLWhere As String
    Dim lngRecCnt As Long
    Dim strActionTaken As String
    
    Dim i As Long
    Dim c As Long
    'On Error GoTo ErrHndlr
    
    strSQLWhere = ""
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
   
    If Me.cboCompany.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND CompanyID = '" & Me.cboCompany.Value & "'"
    End If
    If Me.cboDepartment.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND DepartmentID = '" & Me.cboDepartment.Value & "'"
    End If
    If Me.cboSection.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND SectionID = '" & Me.cboSection.Value & "'"
    End If
    If Me.cboType.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND TypeID = '" & Me.cboType.Value & "'"
    End If
    If Me.cboWorkcategory.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND WorkCategoryID = '" & Me.cboWorkcategory.Value & "'"
    End If
    If Me.txtMachineCtrl.Text <> "" Then
        strSQLWhere = strSQLWhere & " AND ControlNo like '%" & Me.txtMachineCtrl.Text & "%'"
    End If
    If Me.cboStatus.Text <> "" Then
       ' If Me.cboStatus.Text = "Finished" Or Me.cboStatus.Text = "FINISHED" Then
            strSQLWhere = strSQLWhere & " AND Status =  '" & Me.cboStatus.Text & "'"
       ' Else
    
       strSQLWhere = strSQLWhere & " AND WOControlNo LIKE '" & chkAVB(Me.cboDepartment.Text) & "%'"
            '   strSQLWhere = strSQLWhere & " AND Status = '" & Me.cboStatus.Text & "'"
        'End If
    End If
    
    Set rsHistory = cls_GetDetails.pfLoadHistory(Me.dtFrom.Value, Me.dtTo.Value, strSQLWhere, IIf(optReceived, "Received", "Finished"))
    
    If rsHistory.EOF Then
        MsgBox "No Record found!"
        Call subFormatGrid(flxDetail, "costing")
        Me.lblTotalRecord.Caption = 0
        GoTo eExit
    End If
    
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    
    With flxDetail
        .Visible = False
        rsHistory.MoveLast
        lngRecCnt = rsHistory.RecordCount
        Me.lblTotalRecord.Caption = lngRecCnt
        .Rows = lngRecCnt + 1
        rsHistory.MoveFirst
        For i = 1 To lngRecCnt
            For c = 0 To .Cols - 1
            
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
    
    
    
    
eExit:
    flxDetail.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsHistory = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo eExit
End Sub

Public Function chkAVB(ByVal dptname As String) As String
Select Case dptname
    Case "GLASS"
     chkAVB = "M1-GL "
    Case "STEEL"
     chkAVB = "M1-SL"
    Case "WINDOW"
     chkAVB = "M1-WI "
    Case "STRUCTURAL"
     chkAVB = "M1-ST "
    Case "SAWMILL"
     chkAVB = "M1-SW "
    Case "HAGARA"
     chkAVB = "M1-HA "
    Case "PREPARATION"
     chkAVB = "M1-PR "
    Case "PREPARATION 2"
     chkAVB = "M1-PR2 "
End Select
End Function

Private Sub Form_Load()
Me.dtFrom.Value = Date
Me.dtTo.Value = Date
Me.optReceived.Value = True


    Call subFormatGrid(flxDetail, "costing")
    
    Call LoadDataToCombo(cboCompany, "Companies")
    
    Call LoadDataToCombo(cboType, "Types")
    Call LoadDataToCombo(cboStatus, "Status")
     Call LoadDataToCombo(cboWorkcategory, "WorkCategory")
    
End Sub

Private Sub optFinished_Click()
    cboStatus.Enabled = False
    cboStatus.ListIndex = 5
End Sub

Private Sub optReceived_Click()
    cboStatus.Enabled = True
    
End Sub

Private Sub TextBox1_Change()

End Sub
