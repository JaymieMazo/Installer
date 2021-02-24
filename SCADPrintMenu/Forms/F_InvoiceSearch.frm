VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{331187EF-B4B5-4368-9ACE-9E4E2FACD921}#1.0#0"; "OsenControls.ocx"
Begin VB.Form F_InvoiceSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invoice Search"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14985
   Icon            =   "F_InvoiceSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboSupplierCateg 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000015&
      Height          =   315
      ItemData        =   "F_InvoiceSearch.frx":0CCA
      Left            =   1095
      List            =   "F_InvoiceSearch.frx":0CD7
      TabIndex        =   32
      Top             =   1815
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hflxPOSearch 
      Height          =   4290
      Left            =   0
      TabIndex        =   15
      Top             =   3240
      Width           =   14805
      _ExtentX        =   26114
      _ExtentY        =   7567
      _Version        =   393216
      BackColor       =   -2147483624
      FixedCols       =   0
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483634
      BackColorSel    =   16746632
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
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpODFrom 
      Height          =   315
      Left            =   7965
      TabIndex        =   6
      Top             =   825
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
      Format          =   85917697
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker dtpODTo 
      Height          =   315
      Left            =   9735
      TabIndex        =   7
      Top             =   825
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
      Format          =   85917697
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdClose 
      Height          =   375
      Left            =   13560
      TabIndex        =   14
      Top             =   2670
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
      MICON           =   "F_InvoiceSearch.frx":0CEF
      PICN            =   "F_InvoiceSearch.frx":0D0B
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
      Left            =   9900
      TabIndex        =   11
      Top             =   2640
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
      MICON           =   "F_InvoiceSearch.frx":12A7
      PICN            =   "F_InvoiceSearch.frx":12C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdTotal 
      Height          =   375
      Left            =   11100
      TabIndex        =   12
      ToolTipText     =   "View total delivery per invoice"
      Top             =   2670
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&TOTAL"
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
      MICON           =   "F_InvoiceSearch.frx":185F
      PICN            =   "F_InvoiceSearch.frx":187B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpInvFrom 
      Height          =   315
      Left            =   11460
      TabIndex        =   8
      Top             =   825
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
      Format          =   85917697
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker dtpInvTo 
      Height          =   315
      Left            =   13200
      TabIndex        =   9
      Top             =   825
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
      Format          =   85917697
      CurrentDate     =   38212
   End
   Begin OsenXPCntrl.OsenXPButton cmdFTotal 
      Height          =   375
      Left            =   12300
      TabIndex        =   13
      ToolTipText     =   "View over all total delivery "
      Top             =   2670
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&FTOTAL"
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
      MICON           =   "F_InvoiceSearch.frx":1E17
      PICN            =   "F_InvoiceSearch.frx":1E33
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTEncFrom 
      Height          =   315
      Left            =   7965
      TabIndex        =   34
      Top             =   1440
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
      Format          =   85917697
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker DTEncTo 
      Height          =   315
      Left            =   9735
      TabIndex        =   35
      Top             =   1440
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
      Format          =   85917697
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtInvoiceFrm 
      Height          =   315
      Left            =   11460
      TabIndex        =   38
      Top             =   1440
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
      Format          =   85917697
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker DTInvoiceTo 
      Height          =   315
      Left            =   13200
      TabIndex        =   39
      Top             =   1440
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
      Format          =   85917697
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtETDFrom 
      Height          =   315
      Left            =   7965
      TabIndex        =   42
      Top             =   2055
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
      Format          =   85917697
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker dtETDTo 
      Height          =   315
      Left            =   9720
      TabIndex        =   44
      Top             =   2055
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
      Format          =   85917697
      CurrentDate     =   38212
   End
   Begin MSComCtl2.DTPicker dtActFrom 
      Height          =   315
      Left            =   11325
      TabIndex        =   46
      Top             =   2055
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
      Format          =   85917697
      CurrentDate     =   38257.5461458333
   End
   Begin MSComCtl2.DTPicker dtActTo 
      Height          =   315
      Left            =   13080
      TabIndex        =   47
      Top             =   2055
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
      Format          =   85917697
      CurrentDate     =   38212
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   14280
      Top             =   6960
   End
   Begin MSForms.Label Label23 
      Height          =   840
      Left            =   11280
      TabIndex        =   49
      Top             =   1845
      Width           =   1155
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ACTUAL ETD"
      Size            =   "2037;1482"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label22 
      Height          =   360
      Left            =   12840
      TabIndex        =   48
      Top             =   2040
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
   Begin MSForms.Label Label21 
      Height          =   360
      Left            =   9480
      TabIndex        =   45
      Top             =   2040
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
   Begin MSForms.Label Label20 
      Height          =   210
      Left            =   7920
      TabIndex        =   43
      Top             =   1845
      Width           =   315
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ETD"
      Size            =   "556;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label19 
      Height          =   180
      Left            =   11280
      TabIndex        =   41
      Top             =   1200
      Width           =   1140
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "INVOICE DATE"
      Size            =   "2011;317"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label18 
      Height          =   360
      Left            =   12960
      TabIndex        =   40
      Top             =   1395
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
   Begin MSForms.Label Label17 
      Height          =   210
      Left            =   7920
      TabIndex        =   37
      Top             =   1200
      Width           =   1215
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ENCODED DATE"
      Size            =   "2143;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label16 
      Height          =   360
      Left            =   9495
      TabIndex        =   36
      Top             =   1455
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
   Begin MSForms.CheckBox productionPurchaseCheck 
      Height          =   360
      Left            =   2400
      TabIndex        =   33
      ToolTipText     =   "View complete PO only"
      Top             =   1800
      Width           =   2235
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   -2147483634
      DisplayStyle    =   4
      Size            =   "3942;635"
      Value           =   "0"
      Caption         =   "Production Purchase?"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUP. CATEG."
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
      Height          =   210
      Left            =   75
      TabIndex        =   31
      Top             =   1875
      Width           =   1005
   End
   Begin MSForms.TextBox txtItemTypeId 
      Height          =   315
      Left            =   6270
      TabIndex        =   30
      Top             =   600
      Width           =   1410
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2487;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label14 
      Height          =   210
      Left            =   5220
      TabIndex        =   29
      Top             =   660
      Width           =   1125
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ITEM TYPE ID"
      Size            =   "1984;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label13 
      Height          =   210
      Left            =   90
      TabIndex        =   28
      Top             =   1080
      Width           =   705
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "DIVISION"
      Size            =   "1244;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboDivision 
      Height          =   315
      Left            =   795
      TabIndex        =   2
      Top             =   1020
      Width           =   1605
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "2831;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "MS UI Gothic"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   128
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label12 
      Height          =   210
      Left            =   2445
      TabIndex        =   27
      Top             =   660
      Width           =   930
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "INVOICE NO"
      Size            =   "1640;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtInvoiceNo 
      Height          =   315
      Left            =   3375
      TabIndex        =   1
      Top             =   600
      Width           =   1770
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "3122;556"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label11 
      Height          =   360
      Left            =   12960
      TabIndex        =   26
      Top             =   825
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
   Begin MSForms.Label Label10 
      Height          =   210
      Left            =   11280
      TabIndex        =   25
      Top             =   600
      Width           =   1245
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "DELIVERY DATE"
      Size            =   "2196;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox chkComplete 
      Height          =   330
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "View complete PO only"
      Top             =   1800
      Width           =   1215
      VariousPropertyBits=   1015023635
      BackColor       =   -2147483633
      ForeColor       =   -2147483634
      DisplayStyle    =   4
      Size            =   "2143;582"
      Value           =   "0"
      Caption         =   "Complete"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   315
      Left            =   2430
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   5295
      VariousPropertyBits=   746604575
      BackColor       =   -2147483624
      Size            =   "9340;556"
      SpecialEffect   =   3
      FontName        =   "MS UI Gothic"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   128
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboSupplier 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   1005
      Width           =   4320
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      DisplayStyle    =   7
      Size            =   "7620;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   3
      FontName        =   "MS UI Gothic"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   128
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label9 
      Height          =   210
      Left            =   2430
      TabIndex        =   24
      Top             =   1080
      Width           =   780
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "SUPPLIER"
      Size            =   "1376;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtItemId 
      Height          =   315
      Left            =   780
      TabIndex        =   4
      Top             =   1440
      Width           =   1620
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2857;556"
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
      Height          =   210
      Left            =   105
      TabIndex        =   23
      Top             =   1530
      Width           =   600
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ITEM ID"
      Size            =   "1058;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtPoNo 
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   600
      Width           =   1590
      VariousPropertyBits=   746604571
      BackColor       =   -2147483624
      Size            =   "2805;556"
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
      Height          =   210
      Left            =   240
      TabIndex        =   22
      Top             =   690
      Width           =   510
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "PO NO"
      Size            =   "900;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label5 
      Height          =   210
      Left            =   7920
      TabIndex        =   18
      Top             =   600
      Width           =   1005
      ForeColor       =   16777215
      VariousPropertyBits=   276824083
      Caption         =   "ORDER DATE"
      Size            =   "1773;370"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4 
      Height          =   360
      Left            =   9495
      TabIndex        =   19
      Top             =   840
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
      Height          =   2625
      Left            =   0
      TabIndex        =   20
      Top             =   480
      Width           =   14805
      BackColor       =   8421504
      Size            =   "26114;4630"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label7 
      Height          =   390
      Left            =   -30
      TabIndex        =   16
      Top             =   0
      Width           =   15000
      ForeColor       =   -2147483634
      BackColor       =   4210752
      Caption         =   "Invoice/Delivery Search"
      Size            =   "26458;688"
      BorderStyle     =   1
      FontName        =   "Times New Roman"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   2415
      Left            =   15
      TabIndex        =   21
      Top             =   720
      Width           =   14925
      BackColor       =   0
      Size            =   "26326;4260"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   4425
      Left            =   0
      TabIndex        =   17
      Top             =   3240
      Width           =   14940
      BackColor       =   0
      Size            =   "26352;7805"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "F_InvoiceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnComplete As Boolean
Dim blnProdPurchase As Boolean
Dim intValidation As Integer
Private Sub Timer1_Timer()
    If hflxPOSearch.Rows > 1 Then
        intValidation = 1
    Else
        intValidation = 0
    End If
End Sub

Private Sub cboDivision_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdFTotal_Click()
    Call connecttoserver
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Please wait...")
     blnComplete = chkComplete.Value
     blnProdPurchase = productionPurchaseCheck.Value
     
     Call clsPrintMenu.InvoiceSearch.GetInvoiceSearch( _
                                        hflxPOSearch, GetSupCategory(cboSupplierCateg.Text), txtItemId, cboSupplier, cboDivision, txtPoNo, txtInvoiceNo, _
                                        dtpODFrom, dtpODTo, dtpInvFrom, dtpInvTo, txtItemTypeId, DTEncFrom, DTEncTo, dtInvoiceFrm, DTInvoiceTo, dtETDFrom, dtETDTo, dtActFrom, dtActTo, , True, blnComplete, blnProdPurchase)
     Call subInitGrid
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdTotal_Click()
    Call connecttoserver
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Please wait...")
     blnComplete = chkComplete.Value
     blnProdPurchase = productionPurchaseCheck.Value
      
     Call clsPrintMenu.InvoiceSearch.GetInvoiceSearch( _
                                        hflxPOSearch, GetSupCategory(cboSupplierCateg.Text), txtItemId, cboSupplier, cboDivision, txtPoNo, txtInvoiceNo, _
                                        dtpODFrom, dtpODTo, dtpInvFrom, dtpInvTo, txtItemTypeId, DTEncFrom, DTEncTo, dtInvoiceFrm, DTInvoiceTo, dtETDFrom, dtETDTo, dtActFrom, dtActTo, True, , blnComplete, blnProdPurchase)
     Call subInitGrid
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
    Call disconnecttoserver
End Sub

Private Sub cmdView_Click()
     Call connecttoserver
     Screen.MousePointer = vbHourglass
     Call psubShowStatMsg("Please wait... loading data")
     blnComplete = chkComplete.Value
     blnProdPurchase = productionPurchaseCheck.Value
     Call clsPrintMenu.InvoiceSearch.GetInvoiceSearch( _
                                        hflxPOSearch, GetSupCategory(cboSupplierCateg.Text), txtItemId, cboSupplier, cboDivision, txtPoNo, txtInvoiceNo, _
                    dtpODFrom, dtpODTo, dtpInvFrom, dtpInvTo, txtItemTypeId, DTEncFrom, DTEncTo, dtInvoiceFrm, DTInvoiceTo, dtETDFrom, dtETDTo, dtActFrom, dtActTo, , , blnComplete, blnProdPurchase)
     Call subInitGrid
     Call psubHideStatMsg
     Screen.MousePointer = vbDefault
     Call disconnecttoserver
End Sub

Private Sub dtpFrom_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub dtpTo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub dtpODFrom_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
     Call connecttoserver
     Call clsPrintMenu.psubLoadSupplier(cboSupplier)
     Call clsPrintMenu.psubLoadDivision(cboDivision)
     cboSupplierCateg.ListIndex = 0
     Call disconnecttoserver
End Sub


Private Sub hflxPOSearch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If intValidation = 1 Then
        If Button = 2 And hflxPOSearch.Rows > 2 And hflxPOSearch.TextMatrix(1, 0) <> "" Then
           PopupMenu F_PopMenu.mnuInvoiceSearch
        End If
    End If
End Sub


Private Sub txtInvoiceNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtItemId_Change()
     txtDescription = ""
End Sub

Private Sub txtItemID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Call connecttoserver
     If KeyCode = vbKeyReturn Then
          txtDescription.Text = clsPrintMenu.pfstrGetItemDescription(txtItemId)
          SendKeys "{tab}"
     End If
     Call disconnecttoserver
End Sub

Private Sub txtPoNo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub subInitGrid()
     Dim intCol As Integer

     With hflxPOSearch
          .RowHeight(0) = 300
          .Row = 0
          For intCol = 0 To .Cols - 1
               .Col = intCol
               .ColWidth(intCol) = 1200
               .CellAlignment = 4
               .CellFontBold = True
          Next
     End With
End Sub
Private Function GetSupCategory(strSupCateg As String) As Integer
    Select Case strSupCateg
        Case "ALL"
            GetSupCategory = 3
        Case "LOCAL"
            GetSupCategory = 0
        Case "IMPORT"
            GetSupCategory = 1
    End Select
End Function




