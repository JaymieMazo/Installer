VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAccomplishment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monthly Accomplishment Report"
   ClientHeight    =   7575
   ClientLeft      =   195
   ClientTop       =   330
   ClientWidth     =   17760
   Icon            =   "frmAccomplishment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtMonth 
      Height          =   330
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   85458945
      CurrentDate     =   42991
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAccomlishment 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   17535
      _ExtentX        =   30930
      _ExtentY        =   9340
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   -2147483645
      GridColorFixed  =   8421504
      HighLight       =   0
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   2160
      Width           =   17475
      _ExtentX        =   30824
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSForms.ComboBox cboAbbr 
      Height          =   285
      Left            =   8760
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2160
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "3810;503"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Century Gothic"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin VB.Label lblAbbr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6600
      TabIndex        =   11
      Top             =   120
      Width           =   2085
   End
   Begin MSForms.CommandButton cmdExport 
      Height          =   645
      Left            =   16080
      TabIndex        =   10
      Top             =   600
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "2672;1138"
      Picture         =   "frmAccomplishment.frx":1042
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdGetData 
      Height          =   645
      Left            =   14400
      TabIndex        =   9
      Top             =   600
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "2672;1138"
      Picture         =   "frmAccomplishment.frx":2094
      Accelerator     =   83
      MouseIcon       =   "frmAccomplishment.frx":30E6
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Caption         =   "Select Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   8
      Top             =   980
      Width           =   1725
   End
   Begin MSForms.Label lblPleaseWait 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   17535
      Caption         =   "MONTHLY ACCOMPLISHMENT"
      Size            =   "30930;1720"
      BorderStyle     =   1
      FontName        =   "Century Gothic"
      FontEffects     =   1073741825
      FontHeight      =   525
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   530
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1725
   End
   Begin MSForms.ComboBox cboType 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   525
      Width           =   4440
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "7832;503"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Century Gothic"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   4440
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "7832;503"
      ColumnCount     =   2
      cColumnInfo     =   2
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Century Gothic"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      Object.Width           =   "0;3527"
   End
End
Attribute VB_Name = "frmAccomplishment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCompany_Click()
Call Connect
    Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0))
Call Disconnect
    If cboCompany.Column(0) = "002" Then
        Me.cboAbbr.Visible = True
        Me.lblAbbr.Visible = True
    Else
        Me.cboAbbr.Visible = False
        Me.lblAbbr.Visible = False
    End If
End Sub

Private Sub cboType_Click()
Call Connect
    Call LoadDataToCombo(cboAbbr, "AbbreviatedTypes", cboCompany.Column(0), cboType.Column(0))
Call Disconnect
End Sub

Private Sub cmdExport_Click()
        If flxAccomlishment.TextMatrix(2, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub
Private Function exportExcel() As Boolean
        
'         Dim xlApp As Object
'         Dim xlBook As Object
'         Dim xlSheet As Object
         Dim xlApp As Excel.Application
         Dim xlBook As Excel.Workbook
         Dim xlSheet As Excel.Worksheet
         Dim i, c, lngRecCnt As Long
         Dim strLastCol, strLastCol2 As String
         Dim strCol1 As String
         Dim strCol2 As String
         Dim mnth(12) As Variant
         
         FM_Main.MousePointer = vbCustom
         For i = 1 To 12
                mnth(i) = Choose(i, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
         Next i
         
         On Error GoTo ErrExcel
        
         Set xlApp = CreateObject("Excel.Application")
         Set xlBook = xlApp.Workbooks.Add
         Set xlSheet = xlBook.Sheets("Sheet1")
             xlSheet.Name = "ACCOMPLISHMENTS"
         exportExcel = False
         FM_Main.Enabled = False
         flxAccomlishment.Visible = False
         lblPleaseWait.Caption = "Please Wait. Exporting Data to Excel.."
         FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
'        xlApp.Visible = True
         DoEvents
        With flxAccomlishment
                For i = 0 To .Rows - 1
                lblPleaseWait.Caption = "Please Wait. Exporting Data to Excel..  "
                DoEvents
                        For c = 0 To .Cols - 1
                                 xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ") & i + 1).Formula = .TextMatrix(i, c)
                        Next c
                Next i
                For i = 0 To .Rows
                        For c = 0 To .Cols - 1
                                  xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ") & i + 1).Interior.ColorIndex = IIf(c Mod 2, 36, 0)
                                xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ") & 2).Interior.ColorIndex = IIf(c Mod 2, 36, 15)
                                 strLastCol = Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")
                        Next c
                Next i
                For i = .Rows To .Rows
                        For c = 1 To .Cols - 1
                                 xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ") & i + 1).FormulaR1C1 = "=SUM(R[-" & .Row - 1 & "]C:R[-1]C)"
                        Next c
                Next i
                For i = 2 To .Rows
                        For c = .Cols To .Cols
                                With xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ") & i + 1)
                                        .FormulaR1C1 = "=RC[-1]/RC[-2]"
                                        .NumberFormatLocal = "0%"
                                End With
                        Next c
                Next i
                For c = 0 To .Cols
                        strLastCol2 = Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                                                                                "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                                                                                "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", _
                                                                                "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ")
                Next c
        End With
                lblPleaseWait.Caption = "Formatting Spreadsheet..  "
         With xlSheet
                'HEADER ROWS
                
                .Range("A1:" & strLastCol2 & "2").Interior.Pattern = 1
               
                .Range("A1:" & strLastCol2 & "2").Orientation = 90
                .Range("A1:" & strLastCol2 & "1").Interior.ColorIndex = 15
                .Range(strLastCol2 & "1").Formula = "PERCENTAGE"
                
                .Range("A2").HorizontalAlignment = xlCenter
                .Range("A2").Formula = "DEPARTMENTS"
                .Range("A2").Font.Bold = True
                .Range("A2").Font.Size = 12
                .Range("A2").Orientation = 0
             
             
                'TOTAL (last row)
                .Range("A" & flxAccomlishment.Rows + 1 & ":" & strLastCol2 & flxAccomlishment.Rows + 1).Interior.ColorIndex = 15
                .Range("A" & flxAccomlishment.Rows + 1 & ":" & strLastCol2 & flxAccomlishment.Rows + 1).Font.Bold = True
                .Range("A" & flxAccomlishment.Rows + 1).Formula = "TOTAL"
                
                
                'DETAILS
                .Columns("A:" & strLastCol2).EntireColumn.AutoFit
                .Columns("A:" & strLastCol2).EntireRow.AutoFit
                With .Range("A1:" & strLastCol2 & flxAccomlishment.Rows + 1)
                             .HorizontalAlignment = -4108
                             .EntireColumn.AutoFit
                             .Font.Name = "Arial"
                             '-Borders
                             For i = 7 To 12
                                 .Borders(i).Weight = 2
                                 .Borders(i).LineStyle = 1
                             Next i
                             .Range("1:1").Insert Shift:=xlDown
                             .Range("1:1").Insert Shift:=xlDown
                             .Range("1:1").Insert Shift:=xlDown
                             .Range("1:1").Insert Shift:=xlDown
                            
                End With
                
                'TITLE
                With .Range("A2:" & strLastCol2 & "3")
                        .Merge
                        .HorizontalAlignment = -4108
                        .Font.Name = "Arial"
                        .Font.Size = 11
                        .Interior.ColorIndex = 15
                End With
                .Range("A1").Font.Name = "Arial"
                .Range("A1").Formula = UCase(Me.cboCompany.Text)
                .Range("A1").Font.Bold = True
                
                .Range("A2").Formula = UCase("MONTHLY " & Me.cboType.Text & " ACCOMPLISHMENT (" & mnth(Month(Me.dtMonth.Value)) & " " & Year(Me.dtMonth.Value) & ")")
                .Range("A2").Font.Bold = True
                
                'SUMMARY
                
                
                
                
                'FOOTER
                .Range("A" & flxAccomlishment.Rows + 11).Formula = "Prepared by:"
                .Range("A" & flxAccomlishment.Rows + 14).Formula = "OS"
                
                .Range("M" & flxAccomlishment.Rows + 11).Formula = "Reviewed by:"
                .Range("M" & flxAccomlishment.Rows + 14).Formula = " ASV / SV"
                
                .Range("AI" & flxAccomlishment.Rows + 11).Formula = "Noted by:"
                .Range("AI" & flxAccomlishment.Rows + 14).Formula = "DHT"
                
                
                
                 .Columns(strLastCol2).ColumnWidth = 8
                'Page Setup
                With .PageSetup
                        .LeftMargin = Application.InchesToPoints(0)
                        .RightMargin = Application.InchesToPoints(0)
                        .TopMargin = Application.InchesToPoints(0)
                        .BottomMargin = Application.InchesToPoints(0)
                        .HeaderMargin = Application.InchesToPoints(0)
                        .FooterMargin = Application.InchesToPoints(0)
                        .Orientation = xlLandscape
                        .PaperSize = xlPaperA3
                        .Zoom = 90
                        .PrintErrors = xlPrintErrorsDisplayed
                         
                End With
               
         End With
         exportExcel = True
         xlApp.Visible = True
        
         
ErrExit:
        lblPleaseWait.Caption = "MONTHLY ACCOMPLISHMENT"
         flxAccomlishment.Visible = True
         FM_Main.StatusBar1.Panels(3).Text = ""
         FM_Main.Enabled = True
         Set xlApp = Nothing
         
         Exit Function
         
ErrExcel:
         exportExcel = False
         MsgBox Err.Description, vbCritical, "ERROR -" & Err.Number
         GoTo ErrExit
    
End Function

Private Sub cmdGetData_Click()
        Dim i, ii As Integer
        Dim rs As Object
        Dim where(4) As Variant
        Dim nFlxCol As Integer
        Dim strDate As String
        Dim intTotalSched As Long
        Dim intTotalAccomp As Long
        
        intTotalSched = 0
        intTotalAccomp = 0
        FM_Main.MousePointer = vbCustom
        Me.cmdGetData.Enabled = False
        Me.cmdExport.Enabled = False
        Me.lblPleaseWait.Caption = "PLEASE WAIT..."
        FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
        Me.Refresh
        Call Connect
        If Me.cboCompany.Text = "" Or Me.cboType.Text = "" Then GoTo exitt
        Call M_Tool.setUpAccomplishmentHeader(DaysInMonth(0, Me.dtMonth.Value), Me.flxAccomlishment)
        where(0) = Me.cboType.Column(0)
        where(1) = Me.cboCompany.Column(0)
        where(2) = DaysInMonth(1, Me.dtMonth.Value)
        where(3) = DaysInMonth(0, Me.dtMonth.Value)
        Set rs = cls_GetDetails.pfLoadAccomplishments(where)
        If rs.EOF Then
                MsgBox "No record Found!", vbInformation, "System Information"
                GoTo exitt
        End If
        With Me.flxAccomlishment
                .Visible = 0
                Me.Refresh
                rs.MoveLast
                rs.MoveFirst
                .Rows = rs.RecordCount + 2
                Me.ProgressBar1.Value = 0
                Me.ProgressBar1.Max = (.Cols - 2) * (.Rows - 2)
                For i = 2 To .Rows - 1
                        For ii = 0 To 0
                                .Col = ii
                                .Row = i
                                .ColWidth(.Col) = 1500
                                .CellAlignment = flexAlignCenterCenter
                                .TextMatrix(i, ii) = rs.Fields(2).Value
                        Next ii
                        For ii = 1 To .Cols - 3
                                .Col = ii
                                .Row = i
                                .ColWidth(.Col) = 425
                                .CellAlignment = flexAlignCenterCenter
                                If ii Mod 2 Then
                                        .TextMatrix(i, ii) = GetDetails(0, .TextMatrix(0, ii), rs.Fields(1).Value, intTotalSched, cboAbbr.Text)
                                Else
                                        .TextMatrix(i, ii) = GetDetails(1, .TextMatrix(0, ii), rs.Fields(1).Value, intTotalAccomp, cboAbbr.Text)
                                End If
                                Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
                        Next ii
                        For ii = .Cols - 1 To .Cols
                                .TextMatrix(i, ii - 1) = IIf(ii Mod 2, intTotalAccomp, intTotalSched)
                        Next ii
                        rs.MoveNext
                        intTotalSched = 0
                        intTotalAccomp = 0
                Next i
                DoEvents
                .Visible = 1
        End With
        Call Disconnect
exitt:
        DoEvents
        FM_Main.StatusBar1.Panels(3).Text = " "
        Me.lblPleaseWait.Caption = "MONTHLY ACCOMPLISHMENT"
        Me.cmdExport.Enabled = True
        Me.cmdGetData.Enabled = True
        FM_Main.MousePointer = vbDefault
        FM_Main.Enabled = True
End Sub

Private Function GetDetails(ByVal intRecFin As Integer, ByVal strLoopDate As String, ByVal intDeptID As Integer, ByRef lngTotal As Long, Optional strAbbvrName As String) As Long
    Dim rs As Object
    DoEvents
    Set rs = cls_GetDetails.pfLoadAccomplishmentsDetails(strLoopDate, Me.cboType.Column(0), Me.cboCompany.Column(0), intDeptID, intRecFin, strAbbvrName)
    GetDetails = IIf(rs.EOF, 0, rs(IIf(intRecFin = 0, "SCHEDULED", "FINISHED")).Value)
    lngTotal = lngTotal + GetDetails
End Function

Public Function DaysInMonth(ByVal firstlast As Boolean, Optional dtmDate As Variant) As Date
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    DaysInMonth = DateSerial(Year(dtmDate), IIf(firstlast, Month(dtmDate), Month(dtmDate) + 1), IIf(firstlast, 1, 0))
    
End Function

Private Sub Form_Load()
Call Connect
    DoEvents
    dtMonth = Date
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboType, "Types")
Call Disconnect
    Me.cboAbbr.Visible = False
    Me.lblAbbr.Visible = False
End Sub
