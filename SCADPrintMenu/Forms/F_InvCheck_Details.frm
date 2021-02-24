VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form F_InvCheck_Details 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Check Details"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "F_InvCheck_Details.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7815
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxConsOrdData 
      Height          =   3240
      Left            =   135
      TabIndex        =   29
      Top             =   2985
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5715
      _Version        =   393216
      BackColor       =   -2147483624
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
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxConsAverage 
      Height          =   3240
      Left            =   4035
      TabIndex        =   31
      Top             =   2985
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   5715
      _Version        =   393216
      BackColor       =   -2147483624
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
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.Label Label23 
      Height          =   255
      Left            =   4035
      TabIndex        =   34
      Top             =   2745
      Width           =   3555
      ForeColor       =   16777215
      BackColor       =   -2147483647
      Caption         =   "Consume Order Data (Average)"
      Size            =   "6271;450"
      BorderStyle     =   1
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label22 
      Height          =   255
      Left            =   135
      TabIndex        =   33
      Top             =   2745
      Width           =   3555
      ForeColor       =   16777215
      BackColor       =   -2147483647
      Caption         =   "Consume Order Data"
      Size            =   "6271;450"
      BorderStyle     =   1
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label21 
      Height          =   3480
      Left            =   4140
      TabIndex        =   32
      Top             =   2865
      Width           =   3555
      BackColor       =   0
      Size            =   "6271;6138"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label20 
      Height          =   3480
      Left            =   240
      TabIndex        =   30
      Top             =   2865
      Width           =   3555
      BackColor       =   0
      Size            =   "6271;6138"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label19 
      Height          =   270
      Left            =   4245
      TabIndex        =   28
      Top             =   1740
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Safety Stock"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtPlusSS 
      Height          =   315
      Left            =   4245
      TabIndex        =   27
      Top             =   1980
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
   Begin MSForms.Label Label18 
      Height          =   240
      Left            =   3885
      TabIndex        =   26
      Top             =   1980
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "+"
      Size            =   "503;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label17 
      Height          =   270
      Left            =   2385
      TabIndex        =   25
      Top             =   1740
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Consume Until"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtPlusConsUntil 
      Height          =   315
      Left            =   2385
      TabIndex        =   24
      Top             =   1980
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
   Begin MSForms.Label Label16 
      Height          =   270
      Left            =   1905
      TabIndex        =   23
      Top             =   1995
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "="
      Size            =   "503;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label15 
      Height          =   270
      Left            =   330
      TabIndex        =   22
      Top             =   1740
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Order Point"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtOrderPoint 
      Height          =   315
      Left            =   330
      TabIndex        =   21
      Top             =   1980
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
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   4260
      X2              =   5595
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4110
      X2              =   4275
      Y1              =   1515
      Y2              =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   4035
      X2              =   4125
      Y1              =   1425
      Y2              =   1530
   End
   Begin MSForms.Label Label14 
      Height          =   240
      Left            =   5640
      TabIndex        =   20
      Top             =   1245
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "x"
      Size            =   "503;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label13 
      Height          =   270
      Left            =   5880
      TabIndex        =   19
      Top             =   1005
      Width           =   1560
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Standard Deviation"
      Size            =   "2752;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtStDev 
      Height          =   315
      Left            =   5940
      TabIndex        =   18
      Top             =   1245
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
   Begin MSForms.Label Label12 
      Height          =   270
      Left            =   4245
      TabIndex        =   17
      Top             =   1005
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Lead Time"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtSQR_LT 
      Height          =   315
      Left            =   4245
      TabIndex        =   16
      Top             =   1245
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
   Begin MSForms.Label Label11 
      Height          =   240
      Left            =   3870
      TabIndex        =   15
      Top             =   1260
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "x"
      Size            =   "503;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   270
      Left            =   2385
      TabIndex        =   14
      Top             =   1005
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Safety Variable"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtSafetyVar 
      Height          =   315
      Left            =   2385
      TabIndex        =   13
      Top             =   1245
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
   Begin MSForms.Label Label8 
      Height          =   270
      Left            =   1905
      TabIndex        =   12
      Top             =   1260
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "="
      Size            =   "503;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label7 
      Height          =   270
      Left            =   330
      TabIndex        =   11
      Top             =   1005
      Width           =   1500
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Safety Stock"
      Size            =   "2646;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtSS 
      Height          =   315
      Left            =   330
      TabIndex        =   10
      Top             =   1245
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
   Begin MSForms.Label Label6 
      Height          =   270
      Left            =   4245
      TabIndex        =   9
      Top             =   285
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Lead Time"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtLeadTime 
      Height          =   315
      Left            =   4245
      TabIndex        =   8
      Top             =   525
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
      Height          =   240
      Left            =   3870
      TabIndex        =   7
      Top             =   540
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "x"
      Size            =   "503;423"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   270
      Left            =   2385
      TabIndex        =   6
      Top             =   285
      Width           =   1380
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Average"
      Size            =   "2434;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtAVG 
      Height          =   315
      Left            =   2385
      TabIndex        =   5
      Top             =   525
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
   Begin MSForms.Label Label1 
      Height          =   270
      Left            =   1905
      TabIndex        =   4
      Top             =   540
      Width           =   285
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "="
      Size            =   "503;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label3 
      Height          =   270
      Left            =   330
      TabIndex        =   3
      Top             =   285
      Width           =   1500
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Consume Until LT"
      Size            =   "2646;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtConsUntilLT 
      Height          =   315
      Left            =   330
      TabIndex        =   2
      Top             =   525
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
   Begin MSForms.Label Label2 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7485
      BackColor       =   8421504
      Size            =   "13203;4154"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   2355
      Left            =   225
      TabIndex        =   1
      Top             =   240
      Width           =   7485
      BackColor       =   0
      Size            =   "13203;4154"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_InvCheck_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call connecttoserver
     Dim strdteFrom       As String, _
          strdteUntil     As String, _
          strItemId       As String, _
          strDivision     As String

    '--- Get data from F_InvCheck form
    strdteFrom = Format(F_InvCheck.dtpFrom.Value, "yyyy/mm/dd")
    strdteUntil = Format(F_InvCheck.dtpTo.Value, "yyyy/mm/dd")
    strItemId = F_InvCheck.txtItemCode.Text
    strDivision = F_InvCheck.cboDivision.Text
    '--- Get the needed Values from functions
    txtAVG.Text = Math.Round(clsPrintMenu.InventoryCheck.AVGCons(strdteFrom, strdteUntil, strItemId, strDivision), 2)
    txtLeadTime.Text = clsPrintMenu.InventoryCheck.GetLeadTime(strItemId)
    txtSafetyVar.Text = Math.Round(F_InvCheck.cboSafetyVar.Text, 2)
    txtSQR_LT.Text = txtLeadTime.Text
    txtStDev.Text = Math.Round(clsPrintMenu.InventoryCheck.GetStDev(strItemId, strdteFrom, strdteUntil, strDivision), 2)
    '--- Compute the values for inventory check
    txtConsUntilLT.Text = Math.Round(Val(txtAVG.Text) * Val(txtLeadTime), 2)
    txtSS.Text = Math.Round(Val(txtSafetyVar.Text) * Sqr(Val(txtSQR_LT.Text)) * Val(txtStDev.Text), 2)
    txtPlusConsUntil.Text = txtConsUntilLT.Text
    txtPlusSS.Text = txtSS.Text
    txtOrderPoint.Text = Math.Round(Val(txtPlusConsUntil.Text) + Val(txtPlusSS.Text), 2)
    '--- Load consumption data
    Call clsPrintMenu.InventoryCheck.LoadConsumptionData(hflxConsOrdData, strItemId, strdteFrom, strdteUntil, strDivision, False)
    Call clsPrintMenu.InventoryCheck.LoadConsumptionData(hflxConsAverage, strItemId, strdteFrom, strdteUntil, strDivision, True)
    Call subFormatGrids
    Call disconnecttoserver
End Sub

'--- Format the 2 flexgrid
Private Sub subFormatGrids()
     With hflxConsAverage
         .Cols = 2
         .ColWidth(0) = 1500
         .ColWidth(1) = 1500
         .TextMatrix(0, 0) = "Trans Date"
         .TextMatrix(0, 1) = "Consumption"
     End With
     With hflxConsOrdData
         .Cols = 2
         .ColWidth(0) = 1500
         .ColWidth(1) = 1500
         .TextMatrix(0, 0) = "Trans Date"
         .TextMatrix(0, 1) = "Consumption"
     End With
End Sub

