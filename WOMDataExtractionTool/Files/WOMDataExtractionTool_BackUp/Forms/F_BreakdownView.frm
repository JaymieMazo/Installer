VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form F_BreakdownView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Breakdown View"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   17025
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBreakdownDetails 
      Height          =   4305
      Left            =   0
      TabIndex        =   0
      Top             =   945
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   7594
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   11627568
      ForeColorFixed  =   16777215
      BackColorSel    =   14787665
      FocusRect       =   2
      HighLight       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   330
      Left            =   1890
      TabIndex        =   4
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   16646147
      CurrentDate     =   42258
      MinDate         =   40179
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   330
      Left            =   3690
      TabIndex        =   5
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   16646147
      CurrentDate     =   42258
      MinDate         =   40179
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBreakdownSummary 
      Height          =   4080
      Left            =   0
      TabIndex        =   11
      Top             =   5850
      Width           =   17000
      _ExtentX        =   29977
      _ExtentY        =   7197
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   11627568
      ForeColorFixed  =   16777215
      BackColorSel    =   14787665
      FocusRect       =   2
      HighLight       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin MSForms.Label Label5 
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   90
      Width           =   1935
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "DATE:"
      Size            =   "3413;582"
      BorderColor     =   0
      BorderStyle     =   1
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   255
      Left            =   5310
      TabIndex        =   10
      Top             =   585
      Width           =   600
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "' ' '"
      Size            =   "1058;450"
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboType 
      Height          =   330
      Left            =   1890
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   540
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
   Begin MSForms.Label Label1 
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   540
      Width           =   1935
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "TYPE:"
      Size            =   "3413;582"
      BorderColor     =   0
      BorderStyle     =   1
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   5490
      Width           =   16965
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "SUMMARY"
      Size            =   "29924;582"
      BorderColor     =   0
      BorderStyle     =   1
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   315
      Left            =   3330
      TabIndex        =   6
      Top             =   90
      Width           =   375
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "~"
      Size            =   "661;556"
      BorderColor     =   0
      BorderStyle     =   1
      FontName        =   "Verdana"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtRecordCount 
      Height          =   315
      Left            =   1935
      TabIndex        =   3
      Top             =   10035
      Width           =   1095
      VariousPropertyBits=   746604571
      ForeColor       =   0
      BorderStyle     =   1
      Size            =   "1931;556"
      Value           =   "0"
      SpecialEffect   =   0
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontHeight      =   225
      FontCharSet     =   128
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Top             =   10035
      Width           =   1935
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   " TOTAL RECORD :"
      Size            =   "3413;556"
      BorderColor     =   0
      BorderStyle     =   1
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontHeight      =   225
      FontCharSet     =   128
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdExtract 
      Height          =   525
      Left            =   3105
      TabIndex        =   1
      Top             =   10035
      Width           =   1590
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "EXTRACT"
      Size            =   "2805;926"
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "F_BreakdownView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()

    Dim rsBreakdown As Object
    Dim rsBreakdownSummary As Object
    Dim i As Long
    Dim lngRecCnt As Long
    Dim lngCol As Long
     
    On Error GoTo Err_Handlr
    
    Set rsBreakdown = cls_GetDetails.pfGetBreakdown(Me.dtFrom.Value, Me.dtTo.Value, Me.cboType.Column(0))
    With flxBreakdownDetails
        If rsBreakdown.EOF Then
            MsgBox "No Record Found!"
            Call subFormatGrid(flxBreakdownDetails, "breakdown")
            Call subFormatGrid(flxBreakdownSummary, "summary")
            Exit Sub
        End If
        FM_Main.MousePointer = vbHourglass
        .Visible = False
        rsBreakdown.MoveLast
        lngRecCnt = rsBreakdown.RecordCount
        .Rows = lngRecCnt + 1
        rsBreakdown.MoveFirst
        For i = 1 To lngRecCnt
            For lngCol = 0 To .Cols - 1
                .TextMatrix(i, lngCol) = Choose(lngCol + 1, _
                                                                rsBreakdown("ReceivedDate"), _
                                                                rsBreakdown("DepartmentName"), _
                                                                rsBreakdown("SectionName"), _
                                                                rsBreakdown("Received"), _
                                                                rsBreakdown("Finished"))
            Next lngCol
            rsBreakdown.MoveNext
        Next i
        .Visible = True
        FM_Main.MousePointer = vbDefault
    End With
    
    Set rsBreakdownSummary = cls_GetDetails.pfGetBreakdownSummary(Me.dtFrom.Value, Me.dtTo.Value, Me.cboType.Column(0))
    With flxBreakdownSummary
        If rsBreakdownSummary.EOF Then
            MsgBox "No Record Found!"
            Call subFormatGrid(flxBreakdownDetails, "breakdown")
            Call subFormatGrid(flxBreakdownSummary, "summary")
            Exit Sub
        End If
        FM_Main.MousePointer = vbHourglass
        .Visible = False
        rsBreakdownSummary.MoveLast
        lngRecCnt = rsBreakdownSummary.RecordCount
        .Rows = lngRecCnt + 1
        rsBreakdownSummary.MoveFirst
        For i = 1 To lngRecCnt
            For lngCol = 0 To .Cols - 1
                .TextMatrix(i, lngCol) = Choose(lngCol + 1, _
                                                                rsBreakdownSummary("DepartmentName"), _
                                                                rsBreakdownSummary("SectionName"), _
                                                                rsBreakdownSummary("Received"), _
                                                                rsBreakdownSummary("Finished"), _
                                                                rsBreakdownSummary("FinishedWOFromPendingWO"), _
                                                                rsBreakdownSummary("FinishedOnTheSucceedingMonth"), _
                                                                rsBreakdownSummary("Cancelled"), _
                                                                rsBreakdownSummary("Turnover"), _
                                                                rsBreakdownSummary("WaitingParts"), _
                                                                rsBreakdownSummary("ForSchedule"), _
                                                                rsBreakdownSummary("ForConfirmation"))
            Next lngCol
            rsBreakdownSummary.MoveNext
        Next i
        .Visible = True
        FM_Main.MousePointer = vbDefault
    End With
Exit Sub
Err_Handlr:
    MsgBox "Error searching records." & vbCrLf & Err.Number & ": " & Err.Description
    Set rsBreakdownSummary = Nothing
    flxBreakdownSummary.Visible = True
    Set rsBreakdown = Nothing
    flxBreakdownDetails.Visible = True
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.dtFrom.Value = Now
    Me.dtTo.Value = Now
    
    Call subFormatGrid(flxBreakdownDetails, "breakdown")
    Call subFormatGrid(flxBreakdownSummary, "summary")
    Call LoadDataToCombo(cboType, "Types")
End Sub




