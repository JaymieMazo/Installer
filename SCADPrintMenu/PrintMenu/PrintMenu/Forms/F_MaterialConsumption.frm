VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_MaterialConsumption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Consumption and Turn-over "
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   Icon            =   "F_MaterialConsumption.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   11730
   Begin VB.ComboBox cboDivision 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "F_MaterialConsumption.frx":0CCA
      Left            =   8340
      List            =   "F_MaterialConsumption.frx":0CDD
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   765
      Width           =   2430
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/MM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   1635
      TabIndex        =   17
      Top             =   1170
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   -2147483624
      CustomFormat    =   "yyyy/MM"
      Format          =   23658499
      UpDown          =   -1  'True
      CurrentDate     =   38718
   End
   Begin VB.OptionButton optMatTurnover 
      BackColor       =   &H8000000C&
      Caption         =   "MATERIAL TURN-OVER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   5040
      TabIndex        =   16
      Top             =   1200
      Width           =   2460
   End
   Begin VB.OptionButton optMatConsumption 
      BackColor       =   &H8000000C&
      Caption         =   "MATERIAL CONSUMPTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   5040
      TabIndex        =   15
      Top             =   795
      Value           =   -1  'True
      Width           =   2460
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   7500
      TabIndex        =   0
      ToolTipText     =   "view data"
      Top             =   1215
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&VIEW"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "F_MaterialConsumption.frx":0D12
      PICN            =   "F_MaterialConsumption.frx":0D2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxMatConsumption 
      Height          =   5190
      Left            =   105
      TabIndex        =   1
      Top             =   2190
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   9155
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483647
      BackColorBkg    =   11049333
      GridColor       =   -2147483633
      WordWrap        =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   375
      Left            =   10110
      TabIndex        =   2
      ToolTipText     =   "exit "
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&CLOSE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "F_MaterialConsumption.frx":12CA
      PICN            =   "F_MaterialConsumption.frx":12E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdExcel 
      Height          =   375
      Left            =   8805
      TabIndex        =   3
      ToolTipText     =   "exporting data to excel"
      Top             =   1200
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&EXCEL"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "F_MaterialConsumption.frx":1882
      PICN            =   "F_MaterialConsumption.frx":189E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpTo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/MM"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   330
      Left            =   3375
      TabIndex        =   18
      Top             =   1155
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   -2147483624
      CustomFormat    =   "yyyy/MM"
      Format          =   23658499
      UpDown          =   -1  'True
      CurrentDate     =   38718
      MaxDate         =   2958435
   End
   Begin MSForms.Label Label6 
      Height          =   210
      Left            =   7575
      TabIndex        =   20
      Top             =   840
      Width           =   705
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "DIVISION"
      Size            =   "1244;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   270
      Left            =   315
      TabIndex        =   11
      Top             =   855
      Width           =   960
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "ITEM CODE"
      Size            =   "1693;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1650
      TabIndex        =   10
      Top             =   780
      Width           =   1455
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2566;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   3195
      TabIndex        =   9
      Top             =   1215
      Width           =   330
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "582;635"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   330
      TabIndex        =   8
      Top             =   1260
      Width           =   1170
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "SHIPPED DATE"
      Size            =   "2064;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblStatus 
      Height          =   270
      Left            =   120
      TabIndex        =   7
      Top             =   1935
      Width           =   11385
      ForeColor       =   -2147483634
      BackColor       =   8807750
      Size            =   "20082;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label8 
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   105
      Width           =   11715
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Material Consumption and Material Turn-over Monitoring"
      Size            =   "20664;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCodeUntil 
      Height          =   315
      Left            =   3375
      TabIndex        =   5
      Top             =   780
      Width           =   1455
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2566;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3135
      TabIndex        =   4
      Top             =   735
      Width           =   285
   End
   Begin MSForms.Label Label7 
      Height          =   5475
      Left            =   195
      TabIndex        =   14
      Top             =   2025
      Width           =   11415
      BackColor       =   0
      Size            =   "20135;9657"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   1110
      Left            =   120
      TabIndex        =   12
      Top             =   630
      Width           =   11385
      BackColor       =   8421504
      Size            =   "20082;1958"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   1080
      Left            =   210
      TabIndex        =   13
      Top             =   750
      Width           =   11385
      BackColor       =   0
      Size            =   "20082;1905"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_MaterialConsumption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************Developed by Ariel Balisi
'****************March 2006
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExcel_Click()
    Call psubShowStatMsg("Exporting data to excel...Please wait.")
    Call ExportToExcel
    Call psubHideStatMsg
End Sub

Private Sub cmdView_Click()
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait... Loading data.")
    cmdClose.Enabled = False
    Call subFormatGrid
    Call subLoadDataToGrid
    Call psubHideStatMsg
    cmdClose.Enabled = True
    Screen.MousePointer = vbDefault
    cmdExcel.Enabled = hflxMatConsumption.TextMatrix(1, 0) <> ""
End Sub
Public Sub subFormatGrid()
Dim objGetMonthYear As Object
Dim intCol As Integer
Dim strDate As String
         
With hflxMatConsumption
    .Rows = 2
    .Cols = 3
    .TextMatrix(0, 0) = "Item Code"
    .TextMatrix(0, 1) = "ItemTypeId"
    .TextMatrix(0, 2) = "Description"
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 3500
    
    intCol = 2

    strDate = Format(DateAdd("m", -1, dtpFrom.Value), "yyyy/MM")
    Do Until strDate = Format(dtpTo.Value, "yyyy/MM")
        hflxMatConsumption.Cols = hflxMatConsumption.Cols + 1
        intCol = intCol + 1
        strDate = Format(DateAdd("m", 1, strDate), "yyyy/MM")
        .TextMatrix(0, intCol) = strDate
    Loop
    
    For intCol = 0 To 2
        .Col = intCol: .Row = 0
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .RowHeight(0) = 400
    Next
    For intCol = 3 To .Cols - 1
        .Col = intCol: .Row = 0
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter
        .ColWidth(intCol) = 1000
        .RowHeight(0) = 400
    Next
End With
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Format(DateAdd("m", -5, clsDB.ServerDate), "yyyy/MM")
    dtpTo.Value = Format(clsDB.ServerDate, "yyyy/MM")
    cboDivision.ListIndex = 0
    Call subFormatGrid
End Sub

Public Sub subLoadDataToGrid()
Dim intRow, intCol As Integer
Dim strDate As String
Dim dblPreMonthBalance, _
    dblEndMonthBalace, _
    dblTotalMonthBalance, _
    dblMatConsumption, _
    dblMatTurnover As Double
Dim objGetMaterialData As Object
                       
Set objGetMaterialData = clsPrintMenu.MatConsumption.fGetMaterialConsumption(dtpFrom.Value, DateAdd("d", -1, DateAdd("m", 1, dtpTo.Value)), cboDivision.Text, txtItemCode, txtItemCodeUntil)

With hflxMatConsumption
    If objGetMaterialData.EOF Then
        MsgBox "No record found!"
        For intCol = 0 To 2
            .TextMatrix(1, intCol) = ""
        Next
        Exit Sub
    End If
    intRow = 0
    objGetMaterialData.MoveFirst
    Do Until objGetMaterialData.EOF
        DoEvents
        intRow = intRow + 1
        .Rows = intRow + 1
        .TextMatrix(intRow, 0) = objGetMaterialData.Fields("ItemId").Value
        .TextMatrix(intRow, 1) = objGetMaterialData.Fields("ItemTypeId").Value
        .TextMatrix(intRow, 2) = objGetMaterialData.Fields("Description").Value
        
        intCol = 2
        
        If optMatConsumption.Value = True Then
            'for material consumption
            strDate = DateAdd("m", -1, dtpFrom.Value)
            Do Until strDate = dtpTo.Value
                intCol = intCol + 1
                strDate = DateAdd("m", 1, strDate)
                .TextMatrix(intRow, intCol) = clsPrintMenu.MatConsumption.fGetMatConsumption(strDate, DateAdd("d", -1, DateAdd("m", 1, strDate)), objGetMaterialData.Fields("ItemId").Value)
            Loop
        ElseIf optMatTurnover.Value = True Then
            ' for material turnover
            strDate = DateAdd("m", -1, dtpFrom.Value)
            Do Until strDate = dtpTo.Value
                intCol = intCol + 1
                strDate = DateAdd("m", 1, strDate)
                dblPreMonthBalance = clsPrintMenu.pfvarStockBalance(objGetMaterialData.Fields("ItemId").Value, cboDivision.Text, DateAdd("d", -1, strDate))
                dblEndMonthBalace = clsPrintMenu.pfvarStockBalance(objGetMaterialData.Fields("ItemId").Value, cboDivision.Text, DateAdd("d", -1, DateAdd("m", 1, strDate)))
                dblTotalMonthBalance = ((dblPreMonthBalance + dblEndMonthBalace) / 2)
                dblMatConsumption = clsPrintMenu.MatConsumption.fGetMatConsumption(strDate, DateAdd("d", -1, DateAdd("m", 1, strDate)), .TextMatrix(intRow, 0))
                
                If dblMatConsumption = 0 Or dblTotalMonthBalance = 0 Then
                    .TextMatrix(intRow, intCol) = 0
                Else
                    .TextMatrix(intRow, intCol) = Round((dblMatConsumption / dblTotalMonthBalance), 3)
                End If
            Loop
        End If
        objGetMaterialData.MoveNext
    Loop
End With
End Sub
Public Sub ExportToExcel()
    Dim bytCol      As Byte, _
        lngRow      As Long, _
        strRange    As String
    
On Error GoTo lnError
     '---open excel application
    Call clsPrintMenu.Utility.OpenExcel
    With hflxMatConsumption
        For lngRow = 0 To .Rows - 1
            For bytCol = 0 To .Cols - 1
                strRange = "A" & lngRow + 1
                clsPrintMenu.Utility.ExcelWkSheet.Range("A1:" & strRange & "").NumberFormat = "@"
                clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1) = .TextMatrix(lngRow, bytCol)
                If lngRow = 0 Then clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1).Font.Bold = True
            Next
            Call psubShowStatMsg(lngRow & "/" & .Rows - 1, 4)
        Next
        
        strRange = "A" & lngRow
        clsPrintMenu.Utility.ExcelWkSheet.Range("A1:" & strRange & "").HorizontalAlignment = flexAlignRightCenter
        clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
        Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, .Cols, 33)
        clsPrintMenu.Utility.ExcelApp.Visible = True
        Call clsPrintMenu.Utility.CloseExcel
    End With
    Call psubHideStatMsg
    GoTo lnCleanUp
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
lnCleanUp:
     Call clsPrintMenu.Utility.CloseExcel
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtItemCodeUntil.SetFocus
End Sub

Private Sub txtItemCode_LostFocus()
    If txtItemCodeUntil.Text = "" Then txtItemCodeUntil.Text = txtItemCode.Text
End Sub

Private Sub txtItemCodeUntil_GotFocus()
    With txtItemCodeUntil
          .SelStart = 0
          .SelLength = Len(.Text)
     End With
End Sub

Private Sub txtItemCodeUntil_LostFocus()
    If txtItemCode <> "" And txtItemCodeUntil.Text = "" Then txtItemCodeUntil.Text = txtItemCode.Text
End Sub
