VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBreakDown 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Breakdown"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   17010
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Height          =   1275
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   16995
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
         Left            =   5400
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   1935
         TabIndex        =   7
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Format          =   83689473
         CurrentDate     =   42731
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   285
         Left            =   3825
         TabIndex        =   8
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Format          =   83689473
         CurrentDate     =   42731
      End
      Begin MSComCtl2.DTPicker dtsearchDate 
         Height          =   420
         Left            =   8775
         TabIndex        =   9
         Top             =   1050
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   83689473
         CurrentDate     =   42731
      End
      Begin VB.Label Label6 
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
         Left            =   360
         TabIndex        =   19
         Top             =   550
         Width           =   1515
      End
      Begin MSForms.ComboBox cboCompany 
         Height          =   285
         Left            =   1935
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   555
         Width           =   3360
         VariousPropertyBits=   746608667
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "5927;503"
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
         Left            =   3435
         TabIndex        =   16
         Top             =   180
         Width           =   300
      End
      Begin MSForms.CommandButton cmdExcel 
         Height          =   645
         Left            =   8640
         TabIndex        =   15
         Top             =   240
         Width           =   1515
         ForeColor       =   16777215
         BackColor       =   4210688
         Caption         =   "EXTRACT"
         Size            =   "2672;1138"
         Picture         =   "frmBreakDown.frx":0000
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
         Left            =   6960
         TabIndex        =   14
         Top             =   240
         Width           =   1515
         ForeColor       =   16777215
         BackColor       =   4210688
         Caption         =   "SEARCH"
         Size            =   "2672;1138"
         Picture         =   "frmBreakDown.frx":1052
         Accelerator     =   83
         MouseIcon       =   "frmBreakDown.frx":20A4
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
         Caption         =   "RecievedDate:"
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
         Left            =   360
         TabIndex        =   13
         Top             =   180
         Width           =   1515
      End
      Begin MSForms.ComboBox cboType 
         Height          =   285
         Left            =   1935
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   900
         Width           =   3360
         VariousPropertyBits=   746608667
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "5927;503"
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
         Left            =   360
         TabIndex        =   11
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Work Order Of The Day:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   420
         Left            =   7200
         TabIndex        =   10
         Top             =   1050
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.TextBox txtPreviousPending 
      Height          =   375
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   6750
      Visible         =   0   'False
      Width           =   1770
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSummaryDetail 
      Height          =   5265
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   16965
      _ExtentX        =   29924
      _ExtentY        =   9287
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   15
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
      _Band(0).Cols   =   15
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSummaryDetail_forkLift_Battery 
      Height          =   1635
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   2884
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   8
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
      _Band(0).Cols   =   8
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSummaryDetail_forkLift_Diesel 
      Height          =   1635
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   16995
      _ExtentX        =   29977
      _ExtentY        =   2884
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   8
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
      _Band(0).Cols   =   8
   End
   Begin VB.Label Label4 
      Caption         =   "Previous pending"
      Height          =   285
      Left            =   105
      TabIndex        =   5
      Top             =   6795
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label lblMessage2 
      Alignment       =   2  'Center
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
      TabIndex        =   1
      Top             =   4080
      Width           =   14820
   End
End
Attribute VB_Name = "frmBreakDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSM As Object

Private Function MakePropertyValue(propName, propVal) As Object
    Dim oPropValue As Object
    Set oPropValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    oPropValue.Name = propName
    oPropValue.Value = propVal
    Set MakePropertyValue = oPropValue
End Function


'Private Function exportLibre() As Boolean
'     Dim oDoc As Object, _
'        oDesk As Object, _
'        oSheet As Object, _
'        oPar(1) As Object, _
'        CellProp As Object, _
'        CellStyle As Object, _
'        NewStyle As Object, _
'        oRange As Object, _
'        oColumns As Object, _
'        PageStyles As Object, _
'        NewPageStyle As Object, _
'        StdPage As Object, _
'        basicBorder As Object, _
'        oBorder As Object
'    Dim oCol As Long, oRow As Long, nRow As Long
'    Dim Charts As Object
'    Dim Chart As Object
'    Dim Rect As Object
'    Dim oChartRange As Object
'    Dim RangeAddress(0) As Object
'
'    exportLibre = False
'    oRow = 0
'
'    Set oSM = CreateObject("com.sun.star.ServiceManager")
'    Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
'    Set oPar(0) = MakePropertyValue("Hidden", True)
'    Set oPar(1) = MakePropertyValue("Overwrite", True)
'    Set oDoc = oDesk.LoadComponentFromURL("private:factory/scalc", "_blank", 0, oPar)
'    Set oSheet = oDoc.Sheets.getByIndex("0")
'    Set CellProp = oDoc.StyleFamilies.getByName("CellStyles")
'    Set NewStyle = oDoc.createInstance("com.sun.star.style.CellStyle")
'    Call CellProp.InsertbyName("MyStyle", NewStyle)
'    NewStyle.ParentStyle = "Default"
'    Set CellStyle = CellProp.getByName("MyStyle")
'    Set PageStyles = oDoc.StyleFamilies.getByName("PageStyles")
'    Set StdPage = PageStyles.getByName("Default")
'    Set basicBorder = oDoc.Bridge_GetStruct("com.sun.star.table.BorderLine")
'
'    basicBorder.Color = RGB(0, 0, 0)
'    basicBorder.InnerLineWidth = 0
'    basicBorder.OuterLineWidth = 11
'    basicBorder.LineDistance = 0
'
''        With StdPage
''            .FooterIsOn = False
''            .HeaderIsOn = False
''            .IsLandscape = False
''            .Width = 29700
''            .Height = 21000
''            .LeftMargin = 1000
''            .RightMargin = 1000
''            .TopMargin = 1000
''            .BottomMargin = 1000
''        End With
'        '-Header--
'
'            FM_Main.Enabled = False
'
'              lblMessage2.Caption = "Please Wait. Exporting Data to Excel.."
'            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
'        oSheet.getCellByPosition(0, oRow).String = "MACHINE BREAKDOWN WORK ORDER (" & Format(Me.dtFrom, "mmmm dd,yyyy") & " - " & Format(Me.dtTo, "mmmm dd,yyyy") & ")"
'
'
'        For oCol = 0 To flxDetail.Cols - 1
'           With oSheet
'                With CellStyle
'                    .CharWeight = 300
'                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
'                    .CellBackColor = RGB(255, 184, 0)
'                End With
'                .getCellByPosition(oCol, 1).CellStyle = "MyStyle"
'                .getCellByPosition(oCol, 1).String = flxDetail.TextMatrix(0, oCol)
'                .getCellByPosition(oCol, 1).HoriJustify = 2
'           End With
'            Set oBorder = oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder
'            oBorder.LeftLine = basicBorder
'            oBorder.Topline = basicBorder
'            oBorder.RightLine = basicBorder
'            oBorder.BottomLine = basicBorder
'            oBorder.VerticalLine = basicBorder
'            oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder = oBorder
'        Next oCol
'
'
'        '-Content--
'        For oRow = 0 To flxDetail.Rows - 1
'            DoEvents
'
'            lblMessage1.Caption = "Exporting Data : " & oRow & " rows out of " & flxDetail.Rows - 1 & " rows"
'            For oCol = 0 To flxDetail.Cols - 1
'               With oSheet
'                    .getCellByPosition(oCol, oRow + 1).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
'                    .getCellByPosition(oCol, oRow + 1).HoriJustify = 2
'                    flxDetail.Row = oRow
'                    flxDetail.Col = oCol
'                    .getCellByPosition(oCol, oRow + 1).String = flxDetail.TextMatrix(oRow, oCol)
'                    If flxDetail.CellBackColor = &H8080FF Then
'                        .getCellByPosition(oCol, oRow + 1).CellBackColor = RGB(128, 128, 255)
'                    End If
'               End With
'            Next oCol
'                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder
'                oBorder.LeftLine = basicBorder
'                oBorder.Topline = basicBorder
'                oBorder.RightLine = basicBorder
'                oBorder.BottomLine = basicBorder
'                oBorder.VerticalLine = basicBorder
'                oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder = oBorder
'        Next oRow
'
'
'
'        For oCol = 0 To flxSummaryDetail.Cols - 1
'           With oSheet
'                With CellStyle
'                    .CharWeight = 300
'                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
'                    .CellBackColor = RGB(255, 184, 0)
'                End With
'                .getCellByPosition(oCol, oRow + 2).CellStyle = "MyStyle"
'                .getCellByPosition(oCol, oRow + 2).String = flxSummaryDetail.TextMatrix(0, oCol)
'                .getCellByPosition(oCol, oRow + 2).HoriJustify = 2
'           End With
'            Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 2, flxSummaryDetail.Cols - 1, oRow + 2).TableBorder
'            oBorder.LeftLine = basicBorder
'            oBorder.Topline = basicBorder
'            oBorder.RightLine = basicBorder
'            oBorder.BottomLine = basicBorder
'            oBorder.VerticalLine = basicBorder
'            oSheet.getCellRangeByPosition(0, oRow + 2, flxSummaryDetail.Cols - 1, oRow + 2).TableBorder = oBorder
'        Next oCol
'
'
'        '-Content--
'        For oRow = 0 To flxSummaryDetail.Rows - 1
'            DoEvents
'
'            'lblMessage.Caption = "Exporting Data : " & oRow & " rows out of " & flxSummaryDetail.Rows - 1 & " rows"
'            For oCol = 0 To flxSummaryDetail.Cols - 1
'               With oSheet
'                    .getCellByPosition(oCol, oRow + 3).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
'                    .getCellByPosition(oCol, oRow + 3).HoriJustify = 2
'                    flxSummaryDetail.Row = oRow
'                    flxSummaryDetail.Col = oCol
'                    .getCellByPosition(oCol, oRow + 7).String = flxSummaryDetail.TextMatrix(oRow, oCol)
'                    If flxSummaryDetail.CellBackColor = &H8080FF Then
'                        .getCellByPosition(oCol, oRow + 3).CellBackColor = RGB(128, 128, 255)
'                    End If
'               End With
'            Next oCol
'                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 3, flxSummaryDetail.Cols - 1, oRow + 3).TableBorder
'                oBorder.LeftLine = basicBorder
'                oBorder.Topline = basicBorder
'                oBorder.RightLine = basicBorder
'                oBorder.BottomLine = basicBorder
'                oBorder.VerticalLine = basicBorder
'                oSheet.getCellRangeByPosition(0, oRow + 3, flxSummaryDetail.Cols - 1, oRow + 3).TableBorder = oBorder
'        Next oRow
'
'
'
'        Call oDoc.storeToURL("file:///C:/Exported.xls", oPar)
'        Set oPar(0) = MakePropertyValue("Hidden", False)
'        Set oDoc = oDesk.LoadComponentFromURL("file:///C:/Exported.xls", "_blank", 0, oPar)
'        exportLibre = True
'
'
'        Set oSM = Nothing
'        Set oDesk = Nothing
'        Set oDoc = Nothing
'        Set oSheet = Nothing
'        Set oPar(1) = Nothing
'        Set CellProp = Nothing
'        Set CellStyle = Nothing
'        Set NewStyle = Nothing
'        Set oRange = Nothing
'        Set oColumns = Nothing
'        Set PageStyles = Nothing
'        Set NewPageStyle = Nothing
'        Set StdPage = Nothing
'        exportLibre = True
'        flxDetail.Visible = True
'        FM_Main.StatusBar1.Panels(3).Text = ""
'        FM_Main.Enabled = True
'        Exit Function
'ErrLibre:
'        exportLibre = False
'        flxDetail.Visible = True
'        FM_Main.StatusBar1.Panels(3).Text = ""
'        FM_Main.Enabled = True
'
'End Function
'

Private Sub cboCompany_Click()
Call Connect
    Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0), , Me.Name)
Call Disconnect
End Sub

Private Sub cboType_Click()
        If Me.cboType.Text = "ForkLift" Then
              '  Me.Label3.Visible = True
               ' Me.txtPreviousPending.Visible = True
              '  Me.Label4.Visible = True
              '  Me.dtsearchDate.Visible = True
               ' Me.dtsearchDate.Value = Me.dtTo.Value
                
        Else
                Me.Label3.Visible = False
                Me.dtsearchDate.Visible = False
                Me.txtPreviousPending.Visible = False
                Me.Label4.Visible = False
                
        End If
        
End Sub

Private Sub chkForklift_Click()
    If chkForklift.Value = 1 Then
        cboCompany.Enabled = False
        cboType.Enabled = False
        cboCompany.Clear
        cboType.Clear
        lblMessage2.Caption = "Please Wait. Loading Data.."
        FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    Else
        cboCompany.Enabled = True
        cboType.Enabled = True
    Call Connect
        Call LoadDataToCombo(cboCompany, "Companies")
    Call Disconnect
    End If
End Sub

Private Sub cmdExcel_Click()
        If flxSummaryDetail.TextMatrix(1, 0) = "" And _
                flxSummaryDetail_forkLift_Battery.TextMatrix(1, 0) = "" And _
                flxSummaryDetail_forkLift_Diesel.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If Me.chkForklift.Value = 1 Then
                If exportExcel_forklift = True Then
                        MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
                Else
                        MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
                End If
        Else
                 If exportExcel = True Then
                        MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
                Else
                        MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
                End If
        End If
        
        FM_Main.MousePointer = vbDefault
End Sub


Private Function exportExcel_forklift() As Boolean
        
        Dim clsExcelTools As cls_ExcelTools
        Dim strFileName, excelCol As String
        Dim i, iCol As Integer
        Dim excelRow As Integer
        
        On Error GoTo lneError
        exportExcel_forklift = False
        Set clsExcelTools = New cls_ExcelTools
        
'         'Call FileCopy(App.Path & "\Reports\Machine Breakdown Report.xls", strNewFile)
'            strNewFile = Environ$("WINDIR") & "\Temp\"
'            Set xlApp = CreateObject("Excel.Application")
'            Set xlBook = xlApp.Workbooks.Add
'            'Set xlBook = xlApp.Workbooks.Open(strNewFile)
        
       ' If clsExcelTools.fblnOpenTemplate("\\wkn-appserver\access$\References\WOMDET\SYSTEM DEFINED.xls") Then
        If clsExcelTools.fblnOpenTemplate("\\wkn-appserver\access$\References\WOMDET\Summary format - FORKLIFT.xls") Then
        
                With clsExcelTools.xlSheet
                        excelRow = 7
                        For i = 1 To Me.flxSummaryDetail_forkLift_Diesel.Rows - 1
                                
                                For iCol = 3 To 7
                                        Select Case Me.flxSummaryDetail_forkLift_Diesel.TextMatrix(i, 1) 'Company Column
                                                '--HTI
                                                Case "002"
                                                        excelCol = "B"
                                                '--WKN
                                                Case "004"
                                                        excelCol = "D"
                                                '--SCAD
                                                Case "001"
                                                        excelCol = "F"
                                                 '--PV
                                                Case "000"
                                                        excelCol = "H"
                                                 '--HRD
                                                Case "003"
                                                        excelCol = "J"
                                                 '--MGTC
                                                Case "005"
                                                        excelCol = "J"
                                        End Select
                                        excelRow = excelRow + 1
                                        .Range(excelCol & excelRow) = Me.flxSummaryDetail_forkLift_Diesel.TextMatrix(i, iCol)
                                         If iCol = 7 Then excelRow = 7
                                Next iCol
                        Next i
                        excelRow = 17
                        For i = 1 To Me.flxSummaryDetail_forkLift_Battery.Rows - 1
                                For iCol = 3 To 7
                                        Select Case Me.flxSummaryDetail_forkLift_Battery.TextMatrix(i, 1) 'Company Column
                                                '--HTI
                                                Case "002"
                                                        excelCol = "B"
                                                '--WKN
                                                Case "004"
                                                        excelCol = "D"
                                                '--SCAD
                                                Case "001"
                                                        excelCol = "F"
                                                 '--PV
                                                Case "000"
                                                        excelCol = "H"
                                                 '--HRD
                                                Case "003"
                                                        excelCol = "J"
                                                 '--MGTC
                                                Case "005"
                                                        excelCol = "J"
                                        End Select
                                        excelRow = excelRow + 1
                                        .Range(excelCol & excelRow) = Me.flxSummaryDetail_forkLift_Battery.TextMatrix(i, iCol)
                                         If iCol = 7 Then excelRow = 17
                                Next iCol
                        Next i
                  .Range("B3") = Me.dtTo.Value
'                        .Range("A31") = .Range("A31") & " " & Me.dtsearchDate.Value - 1
'                        .Range("A32") = txtPreviousPending.Text
'                        .Range("B31") = .Range("B31") & " " & Me.dtsearchDate.Value
'                        .Range("C31") = .Range("C31") & " " & Me.dtsearchDate.Value
'                        .Range("D31") = .Range("D31") & " " & Me.dtsearchDate.Value
'                        .Range("C23") = .Range("C23") & Me.dtFrom.Value & " - " & Me.dtTo.Value
                        .Range("D37") = .Range("D37") & Me.dtTo.Value
                End With
        End If
        exportExcel_forklift = True
        Set clsExcelTools = Nothing
        
Exit Function
        
lneError:
    Set clsExcelTools = Nothing
    exportExcel_forklift = False
End Function

Private Function exportExcel() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol1, curCol2 As Long
    Dim i As Long
    On Error GoTo ErrSave
        If flxSummaryDetail.TextMatrix(1, 0) = "" Then Exit Function
        exportExcel = False
            FM_Main.Enabled = False
          
            flxSummaryDetail.Visible = False

            lblMessage2.Caption = "Please Wait. Exporting Data to Spreadsheet.."
            
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Spreadsheet.."
        
            'Call FileCopy(App.Path & "\Reports\Machine Breakdown Report.xls", strNewFile)
            
            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Add
            'Set xlBook = xlApp.Workbooks.Open(strNewFile)
            Set xlSheet = xlBook.Sheets("Sheet1")
                xlSheet.Name = "MACHINE BREAKDOWN"
            RowCtr = 2
        
            
            'For SUMMARY REPORT--
            curCol2 = RowCtr
            RowCtr = RowCtr
            With xlSheet
                    .Range("A" & RowCtr).Formula = "COMPANY "
                    .Range("B" & RowCtr).Formula = "DEPARTMENT "
                    .Range("C" & RowCtr).Formula = "BACKLOG"
                    .Range("D" & RowCtr).Formula = "RECEIVED"
                    .Range("E" & RowCtr).Formula = "FINISHED"
                    .Range("F" & RowCtr).Formula = "FINISHED WO FROM PENDINGWO"
                    .Range("G" & RowCtr).Formula = "FINISHED ON THE SUCCEEDING MONTH"
                    .Range("H" & RowCtr).Formula = "CANCELLED"
                    .Range("I" & RowCtr).Formula = "TURNOVER"
                    .Range("J" & RowCtr).Formula = "WAITING PARTS"
                    .Range("K" & RowCtr).Formula = "FOR SCHEDULE"
                    .Range("L" & RowCtr).Formula = "FOR CONFIRMATION/ONGOING"
                    .Range("M" & RowCtr).Formula = "TOTAL UNFINISHED " & vbCrLf & "(" & Format(Me.dtFrom, "mmmm dd,yyyy") & " - " & Format(Me.dtTo, "mmmm dd,yyyy") & ")"
                    .Range("N" & RowCtr).Formula = "TOTAL UNFINISHED" & vbCrLf & "BACKLOGS FROM Previous"
                    .Range("O" & RowCtr).Formula = "TOTAL UNFINISHED "
            End With
            RowCtr = RowCtr + 1
            For intloop = 1 To flxSummaryDetail.Rows - 1
                lblMessage2.Caption = "Please Wait. Exporting Data to Spreadsheet.. (" & intloop & " out of " & flxSummaryDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    .Range("A" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 0)
                    .Range("B" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 2)
                    .Range("D" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 4)
                    .Range("F" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 5)
                    .Range("G" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 6)
                    .Range("H" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 7)
                    .Range("I" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 8)
                    .Range("J" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 9)
                    .Range("K" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 10)
                    .Range("L" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 11)
                    .Range("M" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 12)
                    .Range("N" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 13)
                    .Range("O" & RowCtr).Formula = flxSummaryDetail.TextMatrix(intloop, 14)
                    '-Insert row
                    If flxSummaryDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
             lblMessage2.Caption = ""
            '------
            '--- Excel Format -----------
            With xlSheet
                lblMessage2.Caption = "Formatting Spreadsheet.."
                '-borders

                With .Range("A" & curCol2 & ":O" & RowCtr - 1)
                        .Columns(1).ColumnWidth = 30
                        .Columns("B:O").ColumnWidth = 20
                        .HorizontalAlignment = xlCenter
                        .Font.Bold = False
                        '.WrapText = True
                        For i = 7 To 11
                            With .Borders(i)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .ColorIndex = xlAutomatic
                            End With
                        Next i
                End With
            End With
                '-style
            'first Table
            With xlSheet
                .Rows("3:3").EntireRow.AutoFit
                .Rows(curCol2 & ":" & curCol2).EntireRow.AutoFit
            
            'Second Table
            With .Range("A" & curCol2 & ":O" & curCol2)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 35
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 9
                .Font.Bold = True
                .WrapText = True
                For i = 7 To 11
                    .Borders(i).Weight = xlMedium
                Next i
            End With
            .Rows(curCol2 & ":" & curCol2).RowHeight = 27.75
    
                '--Prepared By:--------
                With .Range("A" & RowCtr + 2)
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .FormulaR1C1 = "Prepared by:"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 11
                End With
                '--Underline----------
                With .Range("A" & RowCtr + 4)
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With

                '--Maint.Staff--
                With .Range("A" & RowCtr + 5)
                    .HorizontalAlignment = xlLeft
                    .FormulaR1C1 = "Maintenance OS"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 11
                End With
                '-Reviewed by: ----
                With .Range("D" & RowCtr + 2)
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .FormulaR1C1 = "Reviewed by:"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 11
                End With
                '--Underline----------
                With .Range("D" & RowCtr + 4 & ":G" & RowCtr + 4)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--MAINT. TL/ASV/ SV--
                With .Range("D" & RowCtr + 5 & ":G" & RowCtr + 5)
                    .Merge
                    .HorizontalAlignment = xlLeft
                    .FormulaR1C1 = "Maintenance TL/Maintenance ASV/Maintenance SV"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 11
                End With
                '-Noted by: ----
                With .Range("J" & RowCtr + 2)
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    .FormulaR1C1 = "Noted by:"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 11
                End With
                '--Underline----------
                With .Range("J" & RowCtr + 4 & ":K" & RowCtr + 4)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--MAINT. DH--
                With .Range("J" & RowCtr + 5 & ":K" & RowCtr + 5)
                    .Merge
                    .HorizontalAlignment = xlLeft
                    .FormulaR1C1 = "Maintenance DH"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 11
                End With
                .Columns("A:O").EntireColumn.AutoFit
                With .Range("A1:I1")
                    .Merge
                    .Range("A1").Formula = Me.cboType.Text & " BREAKDOWN WORK ORDER (" & Format(Me.dtFrom, "mmmm dd,yyyy") & " - " & Format(Me.dtTo, "mmmm dd,yyyy") & ")"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 15
                    .Font.Bold = True
                End With
            End With
                        

        exportExcel = True
        xlApp.Visible = True
        flxSummaryDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True

        Set xlApp = Nothing
        
        Exit Function
         
ErrSave:
    exportExcel = False
        MsgBox Err.Number & " " & Err.Description, vbCritical, "WODataExtractionTool"
        flxSummaryDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
End Function

Private Sub cmdSearch_Click()
    If Me.cboCompany.Text = "" And chkForklift.Value <> 1 Then
        MsgBox "Please select Company", vbOKOnly + vbInformation, "System Information"
        Exit Sub
    ElseIf Me.cboCompany.Text <> "" And cboType.Text = "" And chkForklift.Value <> 1 Then
        MsgBox "Please select Types", vbOKOnly + vbInformation, "System Information"
        Exit Sub
    End If
    Call LoadFlexBreakdownView
End Sub

Private Sub dtsearchDate_LostFocus()
    dtTo.Value = dtsearchDate.Value
End Sub

Private Sub dtTo_LostFocus()
    dtsearchDate.Value = dtTo.Value
End Sub

Private Sub Form_Load()
Call Connect
    Me.dtFrom.Value = Date
    Me.dtTo.Value = Date
    Call subFormatGrid(flxSummaryDetail, "summary")
    Call LoadDataToCombo(cboCompany, "Companies")
Call Disconnect
End Sub

Private Sub LoadFlexBreakdownView()
    Dim rsFlex1 As ADODB.Recordset
    Dim rsFlex2 As ADODB.Recordset
    Dim rsFlex_battery As ADODB.Recordset
    Dim rsFlex_diesel As ADODB.Recordset
    Dim lngLoop, i, c As Long
    Dim lngrow As Long
    Dim lngNo As Long
    Dim d1, d2, d3 As Date
    Dim lngRecCnt As Long
    Dim intPrevPending As Integer
    
    Call Connect
        FM_Main.MousePointer = vbCustom
        FM_Main.Enabled = False
        Me.Refresh
        d1 = Format(Me.dtFrom.Value, "YYYY/MM/DD")
        d2 = Format(Me.dtTo.Value, "YYYY/MM/DD")
        d3 = Format(Me.dtsearchDate.Value, "YYYY/MM/DD")
        flxSummaryDetail_forkLift_Battery.Visible = False
        flxSummaryDetail_forkLift_Diesel.Visible = False
        flxSummaryDetail.Visible = False
        Me.Refresh
        
    If frmBreakDown.cboType.Text <> "FORKLIFT" And chkForklift.Value <> 1 Then
        Set rsFlex2 = cls_GetDetails.pfLoadBreakdown(d1, d2, cboCompany.Column(0), cboType.Column(0))
        If rsFlex2.EOF Then
            MsgBox "No Record found!", vbOKOnly + vbInformation, "System Information"
            Call subFormatGrid(flxSummaryDetail, "summary")
            GoTo LDExit
        End If
       
    Else
        Set rsFlex_battery = cls_GetDetails.pfLoadBreakdown_forklift(d1, d2, "BATTERY", d2)
        Set rsFlex_diesel = cls_GetDetails.pfLoadBreakdown_forklift(d1, d2, "DIESEL", d2)
        If rsFlex_battery.EOF Or rsFlex_diesel.EOF Then
            MsgBox "No Record found!", vbOKOnly + vbInformation, "System Information"
            Call subFormatGrid(flxSummaryDetail_forkLift_Battery, "summary_forklift")
            Call subFormatGrid(flxSummaryDetail_forkLift_Diesel, "summary_forklift")
            GoTo LDExit
        End If
    End If
    
    
    
    lblMessage2.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
            
    Call subFormatGrid(flxSummaryDetail_forkLift_Battery, IIf(chkForklift = 1, "summary_forklift", "summary"))
    Call subFormatGrid(flxSummaryDetail_forkLift_Diesel, IIf(chkForklift = 1, "summary_forklift", "summary"))
    Call subFormatGrid(flxSummaryDetail, "summary")
        
        
        Select Case chkForklift.Value 'Me.cboType.Text
                Case 1 '"FORKLIFT" '--------------------FORKKLIFT
                         '---------FORKKLIFT DIESEL
                        With flxSummaryDetail_forkLift_Diesel
                            DoEvents
                            intPrevPending = 0
                            .Visible = False
                            rsFlex_diesel.MoveLast
                            lngRecCnt = rsFlex_diesel.RecordCount
                            rsFlex_diesel.MoveFirst
                            .Rows = lngRecCnt + 1
                                        For i = 1 To lngRecCnt
                                                For c = 0 To .Cols - 1
                                                        .TextMatrix(i, c) = Choose(c + 1, pfvarNoValue(rsFlex_diesel.Fields("TYPE").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("CompanyID").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("CompanyName").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("FOR SCHEDULE").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("WAITING PARTS").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("ON GOING").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("FINISHED REPAIR").Value), _
                                                                                    pfvarNoValue(rsFlex_diesel.Fields("NEW BREAKDOWN").Value))
                                                Next c
                                                'intPrevPending = intPrevPending + rsFlex_diesel.Fields("PREVIOUS PENDING").Value
                                               ' Debug.Print intPrevPending
                                                rsFlex_diesel.MoveNext
                                        Next i
                            .Visible = True
                        End With
                        '---------FORKKLIFT BATTERY
                        With flxSummaryDetail_forkLift_Battery
                            DoEvents
                            .Visible = False
                            rsFlex_battery.MoveLast
                            lngRecCnt = rsFlex_battery.RecordCount
                            rsFlex_battery.MoveFirst
                            .Rows = lngRecCnt + 1
                                        For i = 1 To lngRecCnt
                                                For c = 0 To .Cols - 1
                                                        .TextMatrix(i, c) = Choose(c + 1, pfvarNoValue(rsFlex_battery.Fields("TYPE").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("CompanyID").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("CompanyName").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("FOR SCHEDULE").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("WAITING PARTS").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("ON GOING").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("FINISHED REPAIR").Value), _
                                                                                    pfvarNoValue(rsFlex_battery.Fields("NEW BREAKDOWN").Value))
                                                Next c
                                               ' intPrevPending = intPrevPending + rsFlex_battery.Fields("PREVIOUS PENDING").Value
                                               ' Debug.Print intPrevPending
                                                rsFlex_battery.MoveNext
                                        Next i
                            .Visible = True
                        End With
                       ' txtPreviousPending.Text = intPrevPending
                        
                        
                Case Else '------------------NOT FORKLIFT
                        With flxSummaryDetail
                            DoEvents
                            .Visible = False
                            rsFlex2.MoveLast
                            lngRecCnt = rsFlex2.RecordCount
                            rsFlex2.MoveFirst
                            .Rows = lngRecCnt + 1
                                        For i = 1 To lngRecCnt
                                                For c = 0 To .Cols - 1
                                                        .TextMatrix(i, c) = Choose(c + 1, pfvarNoValue(rsFlex2.Fields("CompanyName").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("DepartmentName").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("BackLog").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("RECEIVED").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("FINISHED").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("FINISHEDWOFROMPENDINGWO").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("FINISHEDONTHESUCCEEDINGMONTH").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("CANCELLED").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("TURNOVER").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("WAITINGPARTS").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("FORSCHEDULE").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("FORCONFIRMATION").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("TotalUnFinishedJO").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("BackLog").Value) - pfvarNoValue(rsFlex2.Fields("FINISHEDWOFROMPENDINGWO").Value), _
                                                                                    pfvarNoValue(rsFlex2.Fields("TotalUnFinishedJO").Value) + _
                                                                                    (pfvarNoValue(rsFlex2.Fields("BackLog").Value) - pfvarNoValue(rsFlex2.Fields("FINISHEDWOFROMPENDINGWO").Value)))
                                                Next c
                                                rsFlex2.MoveNext
                                        Next i
                            .Visible = True
                        End With
        End Select
        Call Disconnect
        FM_Main.StatusBar1.Panels(3).Text = ""
       
LDExit:
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsFlex1 = Nothing
    Set rsFlex2 = Nothing
    Exit Sub
LDErr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo LDExit
End Sub
