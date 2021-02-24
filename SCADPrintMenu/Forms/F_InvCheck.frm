VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_InvCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "F_InvCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7020
   Begin VB.PictureBox picLoading 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   570
      ScaleHeight     =   495
      ScaleWidth      =   5850
      TabIndex        =   40
      Top             =   3705
      Visible         =   0   'False
      Width           =   5880
      Begin VB.Timer tmrBlink 
         Left            =   0
         Top             =   0
      End
      Begin MSForms.Label lblMessage 
         Height          =   315
         Left            =   375
         TabIndex        =   41
         Top             =   105
         Width           =   5070
         ForeColor       =   65535
         VariousPropertyBits=   8388627
         Caption         =   "LOADING DATA...........  PLEASE WAIT"
         Size            =   "8943;556"
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin OsenXPCntrl.OsenXPButton cmdWaiting 
      Height          =   345
      Left            =   225
      TabIndex        =   4
      Top             =   2250
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Waiting"
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
      MICON           =   "F_InvCheck.frx":0CCA
      PICN            =   "F_InvCheck.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   315
      Left            =   2220
      TabIndex        =   5
      Top             =   4635
      Width           =   1320
      _ExtentX        =   2328
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
      Format          =   122224641
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   315
      Left            =   3795
      TabIndex        =   6
      Top             =   4635
      Width           =   1320
      _ExtentX        =   2328
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
      Format          =   122224641
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdDetail 
      Height          =   345
      Left            =   5310
      TabIndex        =   7
      Top             =   4620
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Detail"
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
      MICON           =   "F_InvCheck.frx":1280
      PICN            =   "F_InvCheck.frx":129C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   420
      Left            =   5310
      TabIndex        =   10
      Top             =   6255
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   741
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
      MICON           =   "F_InvCheck.frx":1836
      PICN            =   "F_InvCheck.frx":1852
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
      Height          =   345
      Left            =   5520
      TabIndex        =   3
      Top             =   1215
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "F_InvCheck.frx":1DEE
      PICN            =   "F_InvCheck.frx":1E0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label22 
      BackColor       =   &H80000010&
      Caption         =   "QtyExpected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   660
      TabIndex        =   45
      Top             =   3030
      Width           =   1140
   End
   Begin MSForms.TextBox txtQtyExpected 
      Height          =   315
      Left            =   2010
      TabIndex        =   44
      Top             =   2985
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
   Begin MSForms.TextBox txtIQCQty 
      Height          =   315
      Left            =   2010
      TabIndex        =   43
      Top             =   2625
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
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "IQCQty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   750
      TabIndex        =   42
      Top             =   2700
      Width           =   615
   End
   Begin MSForms.Label Label20 
      Height          =   270
      Left            =   3090
      TabIndex        =   39
      Top             =   825
      Width           =   750
      ForeColor       =   -2147483634
      VariousPropertyBits=   8388627
      Caption         =   "DIVISION"
      Size            =   "1323;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   3825
      TabIndex        =   2
      Top             =   795
      Width           =   2775
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "4895;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboAverage 
      Height          =   315
      Left            =   1890
      TabIndex        =   9
      Top             =   6360
      Width           =   675
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1191;556"
      ListWidth       =   1164
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "1"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
      Object.Width           =   "1058"
   End
   Begin MSForms.Label Label19 
      Height          =   270
      Left            =   330
      TabIndex        =   38
      Top             =   6075
      Width           =   1605
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "SAFETY VARIABLE"
      Size            =   "2831;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label18 
      Height          =   270
      Left            =   330
      TabIndex        =   37
      Top             =   6405
      Width           =   1560
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "AVERAGE OF"
      Size            =   "2752;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboSafetyVar 
      Height          =   315
      Left            =   1890
      TabIndex        =   8
      Top             =   6030
      Width           =   675
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "1191;556"
      ListWidth       =   1164
      cColumnInfo     =   1
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      Value           =   "2"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
      Object.Width           =   "1058"
   End
   Begin MSForms.OptionButton optAllAC 
      Height          =   300
      Left            =   1380
      TabIndex        =   35
      Top             =   5640
      Width           =   3705
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   16777215
      DisplayStyle    =   5
      Size            =   "6535;529"
      Value           =   "0"
      Caption         =   "INCLUDE ACTUAL COUNT  ( ALL )"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton optMinusAC 
      Height          =   300
      Left            =   1380
      TabIndex        =   34
      Top             =   5370
      Width           =   3705
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   16777215
      DisplayStyle    =   5
      Size            =   "6535;529"
      Value           =   "0"
      Caption         =   "INCLUDE ACTUAL COUNT  ( MINUS ONLY )"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.OptionButton optNoAC 
      Height          =   360
      Left            =   1380
      TabIndex        =   33
      Top             =   5115
      Width           =   3810
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   16777215
      DisplayStyle    =   5
      Size            =   "6720;635"
      Value           =   "1"
      Caption         =   "NOT INCLUDE ACTUAL COUNT AND RETURN"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label16 
      Height          =   270
      Left            =   330
      TabIndex        =   32
      Top             =   5085
      Width           =   1140
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "OUTGOING"
      Size            =   "2011;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label15 
      Height          =   270
      Left            =   330
      TabIndex        =   31
      Top             =   4680
      Width           =   2055
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "Compute from Log of"
      Size            =   "3625;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label14 
      Height          =   360
      Left            =   3600
      TabIndex        =   30
      Top             =   4620
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
   Begin MSForms.Label Label13 
      Height          =   315
      Left            =   225
      TabIndex        =   29
      Top             =   4215
      Width           =   3675
      ForeColor       =   -2147483634
      BackColor       =   8807750
      Caption         =   "Conditions for Compute Order Point"
      Size            =   "6482;556"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtWaiting 
      Height          =   315
      Left            =   2010
      TabIndex        =   25
      Top             =   2265
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
   Begin MSForms.TextBox txtOrderPoint 
      Height          =   315
      Left            =   5175
      TabIndex        =   24
      Top             =   2625
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
   Begin MSForms.Label Label9 
      Height          =   270
      Left            =   3405
      TabIndex        =   23
      Top             =   2655
      Width           =   1755
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "ORDER POINT"
      Size            =   "3096;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtSafetyStock 
      Height          =   315
      Left            =   5175
      TabIndex        =   22
      Top             =   2250
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
   Begin MSForms.Label Label8 
      Height          =   270
      Left            =   3405
      TabIndex        =   21
      Top             =   2280
      Width           =   1755
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "SAFETY STOCK"
      Size            =   "3096;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtConsumeLT 
      Height          =   315
      Left            =   5175
      TabIndex        =   20
      Top             =   1890
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
   Begin MSForms.Label Label5 
      Height          =   270
      Left            =   3405
      TabIndex        =   19
      Top             =   1920
      Width           =   1755
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "CONSUME UNTIL LT"
      Size            =   "3096;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtTotal 
      Height          =   315
      Left            =   2010
      TabIndex        =   18
      Top             =   3345
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
      Height          =   405
      Left            =   225
      TabIndex        =   17
      Top             =   3285
      Width           =   1590
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "ON HAND + WAITING + IQCQty"
      Size            =   "2805;714"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtOnhand 
      Height          =   315
      Left            =   2010
      TabIndex        =   16
      Top             =   1905
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
      Left            =   240
      TabIndex        =   15
      Top             =   1935
      Width           =   1575
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "ON HAND"
      Size            =   "2778;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      X1              =   165
      X2              =   6705
      Y1              =   1755
      Y2              =   1755
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   450
      Left            =   1395
      TabIndex        =   14
      Top             =   1170
      Width           =   4020
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      Size            =   "7091;794"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   270
      Left            =   270
      TabIndex        =   13
      Top             =   1215
      Width           =   1140
      ForeColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DESCRIPTION"
      Size            =   "2011;476"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   105
      Width           =   7020
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Inventory Check"
      Size            =   "12382;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemCode 
      Height          =   315
      Left            =   1395
      TabIndex        =   1
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
   Begin MSForms.Label Label6 
      Height          =   270
      Left            =   270
      TabIndex        =   0
      Top             =   810
      Width           =   960
      ForeColor       =   16777215
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
   Begin MSForms.Label Label2 
      Height          =   3240
      Left            =   165
      TabIndex        =   12
      Top             =   660
      Width           =   6585
      BackColor       =   8421504
      Size            =   "11615;5715"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label10 
      Height          =   3240
      Left            =   270
      TabIndex        =   26
      Top             =   795
      Width           =   6585
      BackColor       =   0
      Size            =   "11615;5715"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label17 
      Height          =   885
      Left            =   1290
      TabIndex        =   36
      Top             =   5070
      Width           =   4305
      VariousPropertyBits=   8388627
      Size            =   "7594;1561"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label11 
      Height          =   2670
      Left            =   165
      TabIndex        =   27
      Top             =   4185
      Width           =   6585
      BackColor       =   8421504
      Size            =   "11615;4710"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label12 
      Height          =   2670
      Left            =   270
      TabIndex        =   28
      Top             =   4275
      Width           =   6585
      BackColor       =   0
      Size            =   "11615;4710"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_InvCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdWaiting_Click()
    Call connecttoserver
    Screen.MousePointer = vbHourglass
    F_InvCheck_Waiting.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Call connecttoserver
    Dim intLoop As Integer
    
    dtpFrom.Value = Format(DateAdd("m", -1, Date), "yyyy/mm/dd")
    dtpTo.Value = Format(Date, "yyyy/mm/dd")
    
    Call clsPrintMenu.psubLoadDivision(cboDivision)
    
    For intLoop = 1 To 10
        cboAverage.AddItem intLoop
        cboSafetyVar.AddItem intLoop
    Next
    clsPrintMenu.InventoryCheck.OptionAC = 0
    Call disconnecttoserver
End Sub

Private Sub cmdDetail_Click()
    Call connecttoserver
    Screen.MousePointer = vbHourglass
    Call psubShowStatMsg("Please wait....loading records")
    F_InvCheck_Details.Show
    Call psubHideStatMsg
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Call connecttoserver
    Dim lngLeadTime As Long
On Error GoTo lnErrMsg
    Screen.MousePointer = vbHourglass
    lngLeadTime = clsPrintMenu.InventoryCheck.GetLeadTime(txtItemCode)
    txtOnhand.Text = clsPrintMenu.pfvarStockBalance(txtItemCode, cboDivision, clsDB.ServerDate)
    txtQtyExpected = clsPrintMenu.InventoryCheck.GetQtyExpected(txtItemCode, cboDivision)
    txtWaiting.Text = clsPrintMenu.InventoryCheck.GetWaitingQty(txtItemCode, cboDivision.Text) + _
                    clsPrintMenu.InventoryCheck.GetExcessWaitingQty(txtItemCode.Text, cboDivision.Text)
    txtIQCQty.Text = clsPrintMenu.InventoryCheck.GetIQCQty(txtItemCode, cboDivision.Text)
    txtTotal.Text = CDbl(txtOnhand.Text) + CDbl(txtWaiting.Text) + Val(txtIQCQty.Text)
    txtConsumeLT.Text = Math.Round(CDbl(clsPrintMenu.InventoryCheck.AVGCons(dtpFrom.Value, dtpTo.Value, txtItemCode, cboDivision.Text)) _
                        * CDbl(lngLeadTime), 2)
    txtSafetyStock.Text = Math.Round(CInt(cboSafetyVar.Text) * Math.Sqr(CDbl(lngLeadTime)) * _
                          clsPrintMenu.InventoryCheck.GetStDev(txtItemCode, dtpFrom.Value, dtpTo.Value, cboDivision), 2)
    txtOrderPoint.Text = Math.Round(Val(txtConsumeLT.Text) + Val(txtSafetyStock.Text), 2)
    Screen.MousePointer = vbDefault
    Call disconnecttoserver
    Exit Sub
    
lnErrMsg:
    Screen.MousePointer = Default
    MsgBox Err.Description, vbCritical, "Error Message"
    Call subClearTxtBoxes
    
End Sub
Private Sub optAllAC_Click()
    Call connecttoserver
    clsPrintMenu.InventoryCheck.OptionAC = 2
    Call disconnecttoserver
    
End Sub

Private Sub optMinusAC_Click()
    Call connecttoserver
    clsPrintMenu.InventoryCheck.OptionAC = 1
    Call disconnecttoserver
End Sub

Private Sub optNoAC_Click()
    Call connecttoserver
    clsPrintMenu.InventoryCheck.OptionAC = 0
    Call disconnecttoserver
End Sub

Private Sub txtItemCode_Change()
    Call subClearTxtBoxes
End Sub

Private Sub txtItemCode_LostFocus()
Call connecttoserver
    txtDescription.Text = clsPrintMenu.pfstrGetItemDescription(txtItemCode)
    If txtDescription.Text = "" And txtItemCode.Text <> "" Then
          MsgBox "Invalid Item ID!", vbExclamation, pstrMessage
          txtItemCode.SelStart = 0
          txtItemCode.SelLength = Len(txtItemCode.Text)
          txtItemCode.SetFocus
    End If
    txtItemCode.Text = UCase(txtItemCode.Text)
Call disconnecttoserver
End Sub

'--- Clear textboxes
Private Sub subClearTxtBoxes()
    Dim objFormObjects  As Object
    
    For Each objFormObjects In Me
        If TypeOf objFormObjects Is MSForms.TextBox Then
            If objFormObjects.Name <> "txtItemCode" Then _
                objFormObjects.Text = ""
        End If
    Next
End Sub



