VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form F_InvCheck_Waiting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Check Waiting"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "F_InvCheck_Waiting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8445
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxChkWaiting 
      Height          =   2310
      Left            =   120
      TabIndex        =   0
      Top             =   1365
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   4075
      _Version        =   393216
      BackColor       =   -2147483624
      Rows            =   3
      Cols            =   5
      BackColorFixed  =   7368816
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483647
      BackColorBkg    =   11049333
      GridColor       =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
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
   Begin MSForms.TextBox txtLeadTime 
      Height          =   315
      Left            =   6390
      TabIndex        =   10
      Top             =   645
      Width           =   1380
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      Size            =   "2434;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   5235
      TabIndex        =   9
      Top             =   705
      Width           =   1020
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "LEAD TIME"
      Size            =   "1799;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtDate 
      Height          =   315
      Left            =   6390
      TabIndex        =   8
      Top             =   255
      Width           =   1380
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      Size            =   "2434;556"
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
      Height          =   270
      Left            =   5235
      TabIndex        =   7
      Top             =   315
      Width           =   1260
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DATE TODAY"
      Size            =   "2222;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtAVG 
      Height          =   315
      Left            =   1410
      TabIndex        =   6
      Top             =   645
      Width           =   1380
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2434;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   270
      Left            =   540
      TabIndex        =   5
      Top             =   705
      Width           =   885
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "AVERAGE"
      Size            =   "1561;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemID 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Top             =   270
      Width           =   1380
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      Size            =   "2434;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   270
      Left            =   540
      TabIndex        =   1
      Top             =   330
      Width           =   885
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "ITEM ID"
      Size            =   "1561;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   1065
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   8160
      BackColor       =   8421504
      Size            =   "14393;1879"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   1065
      Left            =   210
      TabIndex        =   4
      Top             =   195
      Width           =   8160
      BackColor       =   0
      Size            =   "14393;1879"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label6 
      Height          =   2325
      Left            =   210
      TabIndex        =   11
      Top             =   1470
      Width           =   8160
      BackColor       =   0
      Size            =   "14393;4101"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_InvCheck_Waiting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With F_InvCheck
        txtItemId.Text = .txtItemCode.Text
        txtDate.Text = Format(Date, "yyyy/mm/dd")
        txtLeadTime.Text = clsPrintMenu.InventoryCheck.GetLeadTime(txtItemId.Text)
        txtAVG.Text = clsPrintMenu.InventoryCheck.AVGCons(.dtpFrom.Value, .dtpTo.Value, .txtItemCode.Text, .cboDivision.Text)
    End With
    Call subFormatGrid
    Call subLoadWaiting
End Sub

Private Sub txtAVG_GotFocus()
     With txtAVG
          .SelStart = 0
          .SelLength = Len(txtAVG.Text)
     End With
End Sub

Private Sub txtAVG_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then Call subLoadWaiting
End Sub

Private Sub subLoadWaiting()
Dim adoRSWaiting    As Object
    
On Error GoTo lnError
    Set adoRSWaiting = clsPrintMenu.InventoryCheck.GetWaitingDetails(txtItemId.Text, F_InvCheck.cboDivision.Text)
    With hflxChkWaiting
        .Clear
        .Rows = 3
        Call subFormatGrid
        '--- Onhand Qty
        .TextMatrix(1, 1) = clsPrintMenu.pfvarStockBalance(txtItemId.Text, F_InvCheck.cboDivision, clsDB.ServerDate)
        '____IQCQty
        .TextMatrix(2, 1) = clsPrintMenu.InventoryCheck.GetIQCQty(txtItemId, F_InvCheck.cboDivision)
        '--- Days of Onhand and IQCQty
        If Val(txtAVG) > 0 Then
            .TextMatrix(1, 4) = CLng(Val(.TextMatrix(1, 1)) / CDbl(txtAVG.Text))
            .TextMatrix(2, 4) = CLng(Val(.TextMatrix(2, 1)) / CDbl(txtAVG.Text)) + CLng(.TextMatrix(.Rows - 2, 4))
        Else
            .TextMatrix(1, 4) = 0
            .TextMatrix(2, 4) = 0
        End If
        '--- Consume date of Onhand and IQCQty
        .TextMatrix(1, 3) = Format(CDate(txtDate.Text) + Val(.TextMatrix(1, 4)), "yyyy/mm/dd")
        .TextMatrix(2, 3) = Format(CDate(txtDate.Text) + Val(.TextMatrix(.Rows - 1, 4)), "yyyy/mm/dd")
        
        '____Waiting
        If adoRSWaiting.EOF Then Exit Sub
        
        Do Until adoRSWaiting.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "WAITING " & .Rows - 3
            '--- Waiting Qty
            .TextMatrix(.Rows - 1, 1) = pfvarIs_Null(adoRSWaiting.Fields("Waiting").Value, False)
            '--- ETA Date
            .TextMatrix(.Rows - 1, 2) = pfvarIs_Null(adoRSWaiting.Fields("FtryDate").Value)
            '--- Days of Waiting
            If Val(txtAVG) > 0 Then
                .TextMatrix(.Rows - 1, 4) = CLng(Val(.TextMatrix(.Rows - 1, 1)) / CDbl(txtAVG.Text) + Val(.TextMatrix(.Rows - 2, 4)))
            Else
                .TextMatrix(.Rows - 1, 4) = Val(.TextMatrix(.Rows - 1, 4))
            End If
            '--- Consume date of Waiting
            .TextMatrix(.Rows - 1, 3) = Format(CDate(txtDate.Text) + Val(.TextMatrix(.Rows - 1, 4)), "yyyy/mm/dd")
            
            .Col = 0
            .Row = .Rows - 1
            .CellFontBold = True
            
            adoRSWaiting.MoveNext
        Loop
    End With
    Set adoRSWaiting = Nothing
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub

Private Sub subFormatGrid()
    Dim bytLoop As Byte
    
    With hflxChkWaiting
        For bytLoop = 0 To .Cols - 1
            .Row = 0
            .Col = bytLoop
            .CellAlignment = flexAlignCenterCenter
            .CellFontBold = True
            .TextMatrix(0, bytLoop) = Choose(bytLoop + 1, "", "INVENTORY", "FACTORY DATE", "TOTAL CONSUME DATE", "TOTAL DAYS")
            .ColWidth(bytLoop) = Choose(bytLoop + 1, 1200, 1500, 1700, 2300, 1400)
        Next
        .TextMatrix(1, 0) = "ON HAND"
        .Row = 1: .Col = 0
        .CellFontBold = True
        
        .TextMatrix(2, 0) = "IQCQty"
        .Row = 2: .Col = 0
        .CellFontBold = True
    End With
End Sub

