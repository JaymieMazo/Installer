VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_POData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PO Data"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   Icon            =   "F_POData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   13065
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   780
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   120979457
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   3660
      TabIndex        =   1
      Top             =   780
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   120979457
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   375
      Left            =   11310
      TabIndex        =   4
      Top             =   765
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "F_POData.frx":0CCA
      PICN            =   "F_POData.frx":0CE6
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
      Left            =   6765
      TabIndex        =   3
      Top             =   765
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&EXCEL"
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
      MICON           =   "F_POData.frx":1282
      PICN            =   "F_POData.frx":129E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdView 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   765
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "F_POData.frx":1838
      PICN            =   "F_POData.frx":1854
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxPOData 
      Height          =   5520
      Left            =   150
      TabIndex        =   8
      Top             =   1440
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   9737
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   13536000
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   135
      Width           =   13080
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Purchase Order Data"
      Size            =   "23072;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   5535
      Left            =   255
      TabIndex        =   9
      Top             =   1545
      Width           =   12660
      BackColor       =   0
      Size            =   "22331;9763"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   1395
      TabIndex        =   6
      Top             =   825
      Width           =   540
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DATE"
      Size            =   "952;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   3450
      TabIndex        =   5
      Top             =   765
      Width           =   165
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "291;635"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   645
      Left            =   150
      TabIndex        =   10
      Top             =   630
      Width           =   12660
      BackColor       =   8421504
      Size            =   "22331;1138"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   645
      Left            =   270
      TabIndex        =   7
      Top             =   735
      Width           =   12645
      BackColor       =   0
      Size            =   "22304;1138"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_POData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExcel_Click()
    Call connecttoserver
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please Wait.....exporting to excel.")
    Call clsPrintMenu.POData.ExportToExcel(hflxPOData)
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub Form_Load()
    Call connecttoserver
    dtpFrom.Value = Format(Date, "yyyy/mm/dd")
    dtpTo.Value = Format(Date, "yyyy/mm/dd")
    Call disconnecttoserver
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Call connecttoserver
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait... loading records")
    hflxPOData.Clear
    Call subFormatGrid
    hflxPOData.Rows = 2
    Call subLoadRecords
    Call subFormatGrid
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

'--- Sorts the Clicked Column
Private Sub hflxPOData_Click()
    With hflxPOData
        If .Row = 1 Then
            .Sort = flexSortStringNoCaseAscending
        End If
    End With
End Sub

Private Sub subLoadRecords()
    Dim objPOData      As Object
    Dim intCol         As Integer
    Dim lngRow         As Long
    
On Error GoTo lnError

    Set objPOData = clsPrintMenu.POData.GetPOData(dtpFrom, dtpTo)
    If objPOData.EOF Then
        MsgBox "No Record Found!", vbExclamation, pstrMessage
        Exit Sub
    End If
    
    With objPOData
        hflxPOData.Clear
        hflxPOData.Cols = .Fields.Count
        hflxPOData.Rows = 2
        
        For intCol = 0 To .Fields.Count - 1
            hflxPOData.TextMatrix(0, intCol) = .Fields(intCol).Name
        Next
         DoEvents
        Do Until .EOF
            lngRow = lngRow + 1
            For intCol = 0 To .Fields.Count - 1
                hflxPOData.TextMatrix(lngRow, intCol) = pfvarIs_Null(.Fields(intCol).Value)
            Next
            hflxPOData.Rows = hflxPOData.Rows + 1
            .MoveNext
        Loop
    End With
    hflxPOData.Rows = hflxPOData.Rows - 1
    Set objPOData = Nothing
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub

Private Sub subFormatGrid()
    Dim intCols As Integer
    
    With hflxPOData
        .RowHeight(0) = 450
        
        For intCols = 0 To .Cols - 1
            .Row = 0
            .Col = intCols
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
            If intCols = 2 Then
                .ColAlignment(intCols) = flexAlignLeftCenter
            Else
                .ColAlignment(intCols) = flexAlignCenterCenter
            End If
        Next
        
        .ColWidth(2) = 3000
        For intCols = 3 To .Cols - 1
            .ColWidth(intCols) = 1500
        Next
    End With
End Sub

