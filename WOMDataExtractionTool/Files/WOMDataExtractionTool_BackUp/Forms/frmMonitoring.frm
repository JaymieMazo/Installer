VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMonitoring 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PRS Monitoring"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox MaskDate 
      Height          =   315
      Left            =   7560
      TabIndex        =   28
      Top             =   8880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483644
      MaxLength       =   10
      Mask            =   "####/##/##"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cboDepartment 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMonitoring.frx":0000
      Left            =   2040
      List            =   "frmMonitoring.frx":002F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   3240
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmMonitoring.frx":00F0
      Left            =   2040
      List            =   "frmMonitoring.frx":00FA
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3240
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmMonitoring.frx":010D
      Left            =   2040
      List            =   "frmMonitoring.frx":011D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   3240
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      ItemData        =   "frmMonitoring.frx":0162
      Left            =   2040
      List            =   "frmMonitoring.frx":016F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   3240
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   5280
      TabIndex        =   15
      Top             =   8850
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMonitoring 
      Height          =   6090
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   10742
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   30
      FixedCols       =   0
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   285
      Left            =   8280
      TabIndex        =   8
      Top             =   1080
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
      Format          =   122224641
      CurrentDate     =   42731
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   285
      Left            =   10080
      TabIndex        =   9
      Top             =   1080
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
      Format          =   122224641
      CurrentDate     =   42731
   End
   Begin MSForms.TextBox txtPRSNo 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5715;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRS :"
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
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   1800
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
      Left            =   9720
      TabIndex        =   26
      Top             =   1080
      Width           =   300
   End
   Begin MSForms.TextBox txtInCharge 
      Height          =   315
      Left            =   8280
      TabIndex        =   10
      Top             =   1560
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5715;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtMachineName 
      Height          =   315
      Left            =   8280
      TabIndex        =   7
      Top             =   600
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5715;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtMachineNo 
      Height          =   315
      Left            =   8280
      TabIndex        =   6
      Top             =   120
      Width           =   3240
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      Size            =   "5715;556"
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS :"
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
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   1800
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IN - CHARGE :"
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
      Left            =   5400
      TabIndex        =   24
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE OF MAKING PRS :"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MACHINE NAME :"
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
      Left            =   5400
      TabIndex        =   22
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MACHINE CONTROL NO. :"
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
      Left            =   5400
      TabIndex        =   21
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEPARTMENT :"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TYPE :"
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
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COMPANY :"
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
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   1800
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   885
      Left            =   11760
      TabIndex        =   11
      Top             =   120
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "1826;1561"
      Picture         =   "frmMonitoring.frx":01D7
      Accelerator     =   83
      MouseIcon       =   "frmMonitoring.frx":1229
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   885
      Left            =   14160
      TabIndex        =   13
      Top             =   120
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "1826;1561"
      Picture         =   "frmMonitoring.frx":227B
      Accelerator     =   67
      MouseIcon       =   "frmMonitoring.frx":32CD
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   885
      Left            =   12960
      TabIndex        =   12
      Top             =   120
      Width           =   1035
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "1826;1561"
      Picture         =   "frmMonitoring.frx":431F
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
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
      Left            =   0
      TabIndex        =   17
      Top             =   5565
      Width           =   14820
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
      Left            =   1290
      TabIndex        =   14
      Top             =   8880
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
      Left            =   120
      TabIndex        =   16
      Top             =   8880
      Width           =   1245
   End
End
Attribute VB_Name = "frmMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub flxMonitoring_Click()
    With flxMonitoring
        If (.Col = 2 Or .Col = 5 Or .Col = 11 Or .Col = 23 Or .Col = 24 Or .Col = 26) And lblTotalRecord.Caption <> 0 And .TextMatrix(.Row, .Col) <> "N/A" Then
            txtEdit.Text = .TextMatrix(.Row, .Col)
            txtEdit.Visible = True
            txtEdit.Width = .CellWidth
            txtEdit.Height = .CellHeight
            txtEdit.Top = .CellTop + .Top
            txtEdit.Left = .CellLeft + .Left
            txtEdit.SetFocus
        ElseIf (.Col = 14) And lblTotalRecord.Caption <> 0 And .TextMatrix(.Row, .Col) <> "N/A" Then
            MaskDate = IIf(.TextMatrix(.Row, .Col) = "", "____/__/__", Format(.TextMatrix(.Row, .Col), "YYYY/MM/DD"))
            MaskDate.Visible = True
            MaskDate.Width = .CellWidth
            MaskDate.Height = .CellHeight
            MaskDate.Top = .CellTop + .Top
            MaskDate.Left = .CellLeft + .Left
            MaskDate.SetFocus
        Else
            txtEdit.Visible = False
            MaskDate.Visible = False
        End If
    End With
End Sub

Private Sub flxMonitoring_Scroll()
txtEdit.Visible = False
End Sub

Private Sub cboCompany_Click()

    If cboCompany.ListIndex = 0 Then
        cboDepartment.Enabled = True
    Else
        cboDepartment.Enabled = False
        cboDepartment.ListIndex = -1
    End If

End Sub

Private Sub MaskDate_KeyPress(KeyAscii As Integer)

Dim rs As New ADODB.Recordset
Dim r As Integer
Dim strSQL As String
Dim InsertSQL As String
Dim UpdateSQL As String
Dim strCompanyID As String
Dim strPRS As String
Dim strTagNo As String
Dim strRBP As String
Dim strInCharge As String
Dim strReceivedBy As String
Dim strRemarks As String

    If KeyAscii = 13 Then 'ENTER
        With flxMonitoring
         .TextMatrix(.Row, .Col) = IIf(IsDate(Me.MaskDate) = False, "", Me.MaskDate.Text)
        End With
        
        MaskDate.Visible = False
        r = flxMonitoring.Row
        
        strSQL = ""
        
        If cboCompany.ListIndex = 0 Then 'WK
            strCompanyID = "40"
        ElseIf cboCompany.ListIndex = 1 Then 'SC
            strCompanyID = "10"
        ElseIf cboCompany.ListIndex = 2 Then 'HTI
            strCompanyID = "20"
        End If
        
        strPRS = flxMonitoring.TextMatrix(r, 1)
        strTagNo = flxMonitoring.TextMatrix(r, 2)
        strRBP = flxMonitoring.TextMatrix(r, 14)
        strInCharge = flxMonitoring.TextMatrix(r, 23)
        strReceivedBy = flxMonitoring.TextMatrix(r, 24)
        strRemarks = flxMonitoring.TextMatrix(r, 26)
        
        strSQL = " WHERE CompanyID = '" & strCompanyID & "' "
        strSQL = strSQL & " AND PurchaseRequestNo = '" & strPRS & "' "
            
        Set rs = cls_GetDetails.pfLoadPRSAdditionalData(strSQL)
        
        'RECEIVED BY PURCHASING
        
        If rs.RecordCount = 0 Then 'Insert
            
                InsertSQL = ""
                InsertSQL = "INSERT INTO WOMDE_PRSAdditionalData"
                InsertSQL = InsertSQL & " (CompanyID,PurchaseRequestNo,TagNo,ReceivedByPurchasing,InCharge,ReceivedBy,Remarks)"
                InsertSQL = InsertSQL & " VALUES"
                InsertSQL = InsertSQL & " ('" & strCompanyID & "', '" & strPRS & "',"
                
                If strTagNo = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strTagNo & "',"
                End If
                If strRBP = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strRBP & "',"
                End If
                If strInCharge = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strInCharge & "',"
                End If
                If strReceivedBy = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strReceivedBy & "',"
                End If
                If strRemarks = "" Then
                    InsertSQL = InsertSQL & " NULL)"
                Else
                    InsertSQL = InsertSQL & " '" & strRemarks & "')"
                End If
                
                GetRecordSet (InsertSQL)
                
        Else ' Update
        
                UpdateSQL = ""
                UpdateSQL = "UPDATE WOMDE_PRSAdditionalData"

                If strRBP = "" Then
                    UpdateSQL = UpdateSQL & " SET ReceivedByPurchasing = NULL"
                Else
                    UpdateSQL = UpdateSQL & " SET ReceivedByPurchasing = '" & strRBP & "'"
                End If

                UpdateSQL = UpdateSQL & " WHERE CompanyID = '" & strCompanyID & "' AND PurchaseRequestNo = '" & strPRS & "' "
                
                GetRecordSet (UpdateSQL)
                
        End If
        MsgBox "Data Saved !", vbInformation, "Work Order Data Extraction Tool"
    End If
    
End Sub

Private Sub MaskDate_GotFocus()
   MaskDate.SelStart = 0
   MaskDate.SelLength = Len(MaskDate)
End Sub

Private Sub MaskDate_LostFocus()
MaskDate.Visible = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

Dim rs As New ADODB.Recordset
Dim r As Integer
Dim intSeq As Integer
Dim strSQL As String
Dim InsertSQL As String
Dim UpdateSQL As String
Dim strCompanyID As String
Dim strPRS As String
Dim strTagNo As String
Dim strRBP As String
Dim strInCharge As String
Dim strReceivedBy As String
Dim strRemarks As String
Dim strNewItem As String
Dim strMCN As String
Dim strMN As String
Dim strES As String

    If KeyAscii = 13 Then 'ENTER
    
        With flxMonitoring
         .TextMatrix(.Row, .Col) = IIf(Me.txtEdit = "", "", Me.txtEdit.Text)
        End With
        
        txtEdit.Visible = False
        
        r = flxMonitoring.Row
        
        strSQL = ""
        
        If cboCompany.ListIndex = 0 Then 'WK
            strCompanyID = "40"
        ElseIf cboCompany.ListIndex = 1 Then 'SC
            strCompanyID = "10"
        ElseIf cboCompany.ListIndex = 2 Then 'HTI
            strCompanyID = "20"
        End If
        
        strPRS = flxMonitoring.TextMatrix(r, 1)
        strTagNo = flxMonitoring.TextMatrix(r, 2)
        strRBP = flxMonitoring.TextMatrix(r, 14)
        strInCharge = flxMonitoring.TextMatrix(r, 23)
        strReceivedBy = flxMonitoring.TextMatrix(r, 24)
        strRemarks = flxMonitoring.TextMatrix(r, 26)
        intSeq = flxMonitoring.TextMatrix(r, 29)
        strNewItem = flxMonitoring.TextMatrix(r, 5)
        strMCN = flxMonitoring.TextMatrix(r, 9)
        strMN = flxMonitoring.TextMatrix(r, 10)
        strES = flxMonitoring.TextMatrix(r, 11)
        
        If flxMonitoring.Col = 5 Then 'NEWITEMCODE
        
            strSQL = " WHERE CompanyID = '" & strCompanyID & "' "
            strSQL = strSQL & " AND PurchaseRequestNo = '" & strPRS & "' AND PurchaseRequestDetailSeq = " & intSeq & ""
                
            Set rs = cls_GetDetails.pfLoadPRSNewItem(strSQL)
            
            If rs.RecordCount = 0 Then 'Insert
            
                InsertSQL = ""
                InsertSQL = "INSERT INTO WOMDE_PRSNewItem"
                InsertSQL = InsertSQL & " (CompanyID,PurchaseRequestNo,PurchaseRequestDetailSeq,NewItemId)"
                InsertSQL = InsertSQL & " VALUES"
                InsertSQL = InsertSQL & " ('" & strCompanyID & "', '" & strPRS & "', " & intSeq & ","
                
                If strNewItem = "" Then
                    InsertSQL = InsertSQL & " NULL)"
                Else
                    InsertSQL = InsertSQL & " '" & strNewItem & "')"
                End If
                
                GetRecordSet (InsertSQL)
                
            Else ' Update
            
                UpdateSQL = ""
                UpdateSQL = "UPDATE WOMDE_PRSNewItem"
                
                If strNewItem = "" Then
                    UpdateSQL = UpdateSQL & " SET NewItemId = NULL"
                Else
                    UpdateSQL = UpdateSQL & " SET NewItemId = '" & strNewItem & "'"
                End If
                
                UpdateSQL = UpdateSQL & " WHERE CompanyID = '" & strCompanyID & "' AND PurchaseRequestNo = '" & strPRS & "' AND PurchaseRequestDetailSeq = " & intSeq & ""
                
                GetRecordSet (UpdateSQL)
                
            End If
            
        ElseIf flxMonitoring.Col = 11 Then 'EQUIPMENTSTATUS
        
            strSQL = " WHERE MachineItemNo = '" & strMCN & "'"
            
            Set rs = cls_GetDetails.pfLoadMachineStatus(strSQL)
            
            If rs.RecordCount = 0 Then 'Insert
            
                InsertSQL = ""
                InsertSQL = "INSERT INTO WOMDE_MachineStatus"
                InsertSQL = InsertSQL & " (MachineItemNo,MachineName,EquipmentStatus)"
                InsertSQL = InsertSQL & " VALUES"
                InsertSQL = InsertSQL & " ('" & strMCN & "', '" & strMN & "',"
                
                If strES = "" Then
                    InsertSQL = InsertSQL & " NULL)"
                Else
                    InsertSQL = InsertSQL & " '" & strES & "')"
                End If
                
                GetRecordSet (InsertSQL)
                
            Else ' Update
            
                UpdateSQL = ""
                UpdateSQL = "UPDATE WOMDE_MachineStatus"
                
                If strES = "" Then
                    UpdateSQL = UpdateSQL & " SET EquipmentStatus = NULL"
                Else
                    UpdateSQL = UpdateSQL & " SET EquipmentStatus = '" & strES & "'"
                End If
                
                UpdateSQL = UpdateSQL & " WHERE MachineItemNo = '" & strMCN & "'"
                
                GetRecordSet (UpdateSQL)
                
            End If
            
        Else 'TAGNO,RBP,INCHARGE,RECEIVEDBY,REMARKS
                        
            strSQL = " WHERE CompanyID = '" & strCompanyID & "' "
            strSQL = strSQL & " AND PurchaseRequestNo = '" & strPRS & "' "
                
            Set rs = cls_GetDetails.pfLoadPRSAdditionalData(strSQL)
            
            If rs.RecordCount = 0 Then 'Insert
            
                InsertSQL = ""
                InsertSQL = "INSERT INTO WOMDE_PRSAdditionalData"
                InsertSQL = InsertSQL & " (CompanyID,PurchaseRequestNo,TagNo,ReceivedByPurchasing,InCharge,ReceivedBy,Remarks)"
                InsertSQL = InsertSQL & " VALUES"
                InsertSQL = InsertSQL & " ('" & strCompanyID & "', '" & strPRS & "',"
                
                If strTagNo = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strTagNo & "',"
                End If
                If strRBP = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strRBP & "',"
                End If
                If strInCharge = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strInCharge & "',"
                End If
                If strReceivedBy = "" Then
                    InsertSQL = InsertSQL & " NULL,"
                Else
                    InsertSQL = InsertSQL & " '" & strReceivedBy & "',"
                End If
                If strRemarks = "" Then
                    InsertSQL = InsertSQL & " NULL)"
                Else
                    InsertSQL = InsertSQL & " '" & strRemarks & "')"
                End If
                
                GetRecordSet (InsertSQL)
                
            Else ' Update
            
                UpdateSQL = ""
                UpdateSQL = "UPDATE WOMDE_PRSAdditionalData"
                
                If flxMonitoring.Col = 2 Then 'TAGNO
                    If strTagNo = "" Then
                        UpdateSQL = UpdateSQL & " SET TagNo = NULL"
                    Else
                        UpdateSQL = UpdateSQL & " SET TagNo = '" & strTagNo & "'"
                    End If
                ElseIf flxMonitoring.Col = 23 Then ' INCHARGE
                    If strInCharge = "" Then
                        UpdateSQL = UpdateSQL & " SET InCharge = NULL"
                    Else
                        UpdateSQL = UpdateSQL & " SET InCharge = '" & strInCharge & "'"
                    End If
                ElseIf flxMonitoring.Col = 24 Then ' RECEIVEDBY
                    If strReceivedBy = "" Then
                        UpdateSQL = UpdateSQL & " SET ReceivedBy = NULL"
                    Else
                        UpdateSQL = UpdateSQL & " SET ReceivedBy = '" & strReceivedBy & "'"
                    End If
                ElseIf flxMonitoring.Col = 26 Then ' REMARKS
                    If strRemarks = "" Then
                        UpdateSQL = UpdateSQL & " SET Remarks = NULL"
                    Else
                        UpdateSQL = UpdateSQL & " SET Remarks = '" & strRemarks & "'"
                    End If
                End If
                
                UpdateSQL = UpdateSQL & " WHERE CompanyID = '" & strCompanyID & "' AND PurchaseRequestNo = '" & strPRS & "' "
                
                GetRecordSet (UpdateSQL)
                
            End If
            
        End If
        MsgBox "Data Saved !", vbInformation, "Work Order Data Extraction Tool"
    End If
    
End Sub

Private Sub txtEdit_GotFocus()
   txtEdit.SelStart = 0
   txtEdit.SelLength = Len(txtEdit)
End Sub

Private Sub txtEdit_LostFocus()
txtEdit.Visible = False
End Sub

Private Sub Form_Load()

Call subFormatGrid(flxMonitoring, "prsmonitoring")
lblTotalRecord.Caption = 0
txtEdit.Visible = False
MaskDate.Visible = False
dtFrom.Value = Now()
dtTo.Value = Now()

End Sub

Private Sub txtPRSNo_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case (KeyAscii)
        Case 48 To 57
        Case 8
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmdSearch_Click()

    Dim rs As New ADODB.Recordset
    Dim strSQLwhere As String
    Dim lngRecCnt As Long
    Dim i As Long
    Dim c As Long
    Dim x As Long
    On Error GoTo ErrHndlr
    
    If cboCompany.ListIndex < 0 Then
        MsgBox "Please select a Company !", vbInformation, "Work Order Data Extraction Tool"
    ElseIf txtPRSNo = "" And cboDepartment = "" And cboType = "" And cboStatus = "" And txtMachineNo = "" And txtMachineName = "" And txtInCharge = "" Then
        MsgBox "Please enter atleast one more value for search !", vbInformation, "Work Order Data Extraction Tool"
    ElseIf dtFrom.Value > dtTo.Value Then
        MsgBox "Date to must be greater than or equal to Date from !", vbInformation, "Work Order Data Extraction Tool"
    Else
    
        strSQLwhere = ""
        
        If dtFrom.Value = dtTo.Value Then
            strSQLwhere = strSQLwhere & "Where IssuedDate = '" & Format(dtFrom.Value, "YYYY/MM/DD") & "' "
        Else
            strSQLwhere = strSQLwhere & "Where IssuedDate >= '" & Format(dtFrom.Value, "YYYY/MM/DD") & "' "
            strSQLwhere = strSQLwhere & "AND IssuedDate <= '" & Format(dtTo.Value, "YYYY/MM/DD") & "' "
        End If
        
        If txtPRSNo.Text <> "" Then
            strSQLwhere = strSQLwhere & "AND PurchaseRequest.PurchaseRequestNo = '" & txtPRSNo.Text & "'"
        End If
        
        If cboType.Text <> "" Then
            If cboType.ListIndex = 0 Then
                strSQLwhere = strSQLwhere & "AND IsImport = 'Local' "
            Else
                strSQLwhere = strSQLwhere & "AND IsImport = 'Import' "
            End If
        End If

        If cboStatus <> "" Then
            If cboCompany.ListIndex = 0 Or cboCompany.ListIndex = 1 Then
                If cboStatus.ListIndex = 0 Then 'Complete
                    strSQLwhere = strSQLwhere & "AND Qty = QtyReceived AND QtyReceived IS NOT  NULL AND CancelledDate IS NULL "
                ElseIf cboStatus.ListIndex = 1 Then 'With Balance
                    strSQLwhere = strSQLwhere & "AND Qty <> QtyReceived  AND QtyReceived IS NOT  NULL AND CancelledDate IS NULL "
                ElseIf cboStatus.ListIndex = 2 Then 'Cancelled
                    strSQLwhere = strSQLwhere & "AND CancelledDate  IS NOT  NULL "
                Else 'Not Yet Received
                    strSQLwhere = strSQLwhere & " AND QtyReceived IS NULL AND CancelledDate IS NULL "
                End If
            Else
                If cboStatus.ListIndex = 0 Then 'Complete
                    strSQLwhere = strSQLwhere & "AND Qty = QtyReceived AND QtyReceived IS NOT  NULL AND Cancelled = 0 "
                ElseIf cboStatus.ListIndex = 1 Then 'With Balance
                    strSQLwhere = strSQLwhere & "AND Qty <> QtyReceived  AND QtyReceived IS NOT  NULL AND Cancelled = 0 "
                ElseIf cboStatus.ListIndex = 2 Then 'Cancelled
                    strSQLwhere = strSQLwhere & "AND Cancelled = 1 "
                Else 'Not Yet Received
                    strSQLwhere = strSQLwhere & " AND QtyReceived IS NULL AND Cancelled = 0 "
                End If
            End If
        End If
        
        If cboDepartment <> "" Then
            If cboDepartment.ListIndex = 0 Then 'J
                strSQLwhere = strSQLwhere & "AND Division = 'J' "
            ElseIf cboDepartment.ListIndex = 1 Then 'P
                strSQLwhere = strSQLwhere & "AND Division = 'P' "
            ElseIf cboDepartment.ListIndex = 2 Then 'S
                strSQLwhere = strSQLwhere & "AND Division = 'S' "
            ElseIf cboDepartment.ListIndex = 3 Then 'H
                strSQLwhere = strSQLwhere & "AND Division = 'H' "
            ElseIf cboDepartment.ListIndex = 4 Then 'RC
                strSQLwhere = strSQLwhere & "AND Division = 'RC' "
            ElseIf cboDepartment.ListIndex = 5 Then 'WT
                strSQLwhere = strSQLwhere & "AND Division = 'WT' "
            ElseIf cboDepartment.ListIndex = 6 Then 'L
                strSQLwhere = strSQLwhere & "AND Division = 'L' "
            ElseIf cboDepartment.ListIndex = 7 Then 'M
                strSQLwhere = strSQLwhere & "AND Division = 'M' "
            ElseIf cboDepartment.ListIndex = 8 Then 'LG
                strSQLwhere = strSQLwhere & "AND Division = 'LG' "
            ElseIf cboDepartment.ListIndex = 9 Then 'PM
                strSQLwhere = strSQLwhere & "AND Division = 'PM' "
            ElseIf cboDepartment.ListIndex = 10 Then 'C
                strSQLwhere = strSQLwhere & "AND Division = 'C' "
            ElseIf cboDepartment.ListIndex = 11 Then 'HK
                strSQLwhere = strSQLwhere & "AND Division = 'HK' "
            Else 'FL
                strSQLwhere = strSQLwhere & "AND Division = 'FL' "
            End If
        End If
        
        If txtMachineNo.Text <> "" Then
            strSQLwhere = strSQLwhere & "AND WorkOrder.MachineItemNo LIKE '%" & txtMachineNo.Text & "%' "
        End If
        
        If txtMachineName.Text <> "" Then
             strSQLwhere = strSQLwhere & "AND WorkOrder.MachineName LIKE '%" & txtMachineName.Text & "%' "
        End If
        
        If txtInCharge.Text <> "" Then
             strSQLwhere = strSQLwhere & "AND LeaderIncharge LIKE '%" & txtInCharge.Text & "%' "
        End If
        
        FM_Main.MousePointer = vbCustom
        FM_Main.Enabled = False
        
        If cboCompany.ListIndex = 0 Then 'Wukong
            Set rs = cls_GetDetails.pfLoadPRSMonitoring(40, "wkn-appserver", strSQLwhere)
        ElseIf cboCompany.ListIndex = 1 Then 'Scad
            Set rs = cls_GetDetails.pfLoadPRSMonitoring(10, "a-sv17", strSQLwhere)
        ElseIf cboCompany.ListIndex = 2 Then 'Hti
            Set rs = cls_GetDetails.pfLoadPRSMonitoring(20, "jd-004", strSQLwhere)
        End If
        
        If rs.EOF Then
            MsgBox "No Record found!", vbInformation, "Work Order Data Extraction Tool"
            Call subFormatGrid(flxMonitoring, "prsmonitoring")
            Me.lblTotalRecord.Caption = 0
            GoTo eExit
        End If
        
        lblMessage.Caption = "Please Wait. Loading Data.."
        FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
        
        With flxMonitoring
            .Visible = False
            rs.MoveLast
            lngRecCnt = rs.RecordCount
            Me.lblTotalRecord.Caption = lngRecCnt
            .Rows = lngRecCnt + 1
            rs.MoveFirst
            x = 0
            For i = 1 To lngRecCnt
                For c = 0 To .Cols - 1
'                    If c = 2 Or c = 5 Or c = 14 Or c = 23 Or c = 24 Or c = 26 Then
'                        .TextMatrix(i, c) = ""
                        
                    If c = 5 Then
                        If Not (IsNull(rs.Fields(5).Value)) Then
                            .TextMatrix(i, c) = rs.Fields(5).Value
                        Else
                            If .TextMatrix(i, 4) = "" Then
                                .TextMatrix(i, c) = "N/A"
                            Else
                                .TextMatrix(i, c) = ""
                            End If
                        End If
                        x = x + 1
                    ElseIf c = 11 Then
                        If Not (IsNull(rs.Fields(11).Value)) Then
                            .TextMatrix(i, c) = rs.Fields(11).Value
                        Else
                            If cboCompany.ListIndex = 0 Then
                                If .TextMatrix(i, 9) = "" Or (rs.Fields(27).Value = "P" And (rs.Fields(28).Value = "MFA" Or rs.Fields(28).Value = "MFB" Or rs.Fields(28).Value = "MFC")) Then
                                    .TextMatrix(i, c) = "N/A"
                                Else
                                    .TextMatrix(i, c) = ""
                                End If
                            Else
                                If .TextMatrix(i, 9) = "" Then
                                    .TextMatrix(i, c) = "N/A"
                                Else
                                    .TextMatrix(i, c) = ""
                                End If
                            End If
                        End If
                         x = x + 1
                    ElseIf c = 25 Then
                        If cboCompany.ListIndex = 0 Or cboCompany.ListIndex = 1 Then
                        
                            If Not (IsNull(rs.Fields(26).Value)) Then
                                .TextMatrix(i, c) = "Cancelled"
                            ElseIf IsNull(rs.Fields(18).Value) Then
                                .TextMatrix(i, c) = "Not Yet Received"
                            ElseIf rs.Fields(17).Value <> rs.Fields(18).Value And Not (IsNull(rs.Fields(17).Value)) Then
                                .TextMatrix(i, c) = "With Balance"
                            ElseIf rs.Fields(17).Value = rs.Fields(18).Value And Not (IsNull(rs.Fields(17).Value)) Then
                                .TextMatrix(i, c) = "Completely Delivered"
                            End If
                            
                        ElseIf cboCompany.ListIndex = 2 Then
                        
                            If rs.Fields(26).Value = 1 Then
                                .TextMatrix(i, c) = "Cancelled"
                            ElseIf IsNull(rs.Fields(18).Value) Then
                                .TextMatrix(i, c) = "Not Yet Received"
                            ElseIf rs.Fields(17).Value <> rs.Fields(18).Value And Not (IsNull(rs.Fields(17).Value)) Then
                                .TextMatrix(i, c) = "With Balance"
                            ElseIf rs.Fields(17).Value = rs.Fields(18).Value And Not (IsNull(rs.Fields(17).Value)) Then
                                .TextMatrix(i, c) = "Completely Delivered"
                            End If
                            
                        End If
                        
                    ElseIf c = 27 Then
                        If cboCompany.ListIndex = 0 Then
                            .TextMatrix(i, c) = pfvarNoValue(rs.Fields(27).Value) 'Division
                        Else
                            .TextMatrix(i, c) = "N/A"
                        End If
                        
                    ElseIf c = 28 Then
                        If cboCompany.ListIndex = 0 Then
                            .TextMatrix(i, c) = pfvarNoValue(rs.Fields(28).Value) 'FinalDestination
                        Else
                            .TextMatrix(i, c) = "N/A"
                        End If
                        
                    ElseIf c = 29 Then
                        If cboCompany.ListIndex = 2 Then
                            .TextMatrix(i, c) = pfvarNoValue(rs.Fields(27).Value) 'Seq
                        Else
                            .TextMatrix(i, c) = pfvarNoValue(rs.Fields(29).Value) 'Seq
                        End If
                        
                    Else
                        .TextMatrix(i, c) = pfvarNoValue(rs.Fields(x).Value)
                        x = x + 1
                    End If
                        .Row = i
                    .Col = c
                    .CellAlignment = flexAlignLeftCenter
                    If .TextMatrix(i, c) = "Import" Then
                        .CellForeColor = &HFF&
                    ElseIf .TextMatrix(i, c) = "Local" Then
                        .CellForeColor = &H80000008
                    End If
                Next c
                x = 0
                rs.MoveNext
            Next i
        End With
    End If
eExit:
    flxMonitoring.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rs = Nothing
    Exit Sub
ErrHndlr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo eExit
End Sub

Private Sub cmdExcel_Click()
    If flxMonitoring.TextMatrix(1, 0) = "" Then Exit Sub
        FM_Main.MousePointer = vbCustom
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "WODataExtractionTool"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        FM_Main.MousePointer = vbDefault
End Sub

Private Function exportExcel() As Boolean
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i, c, lngRecCnt As Long
    On Error GoTo ErrExcel

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets("Sheet1")
    
    exportExcel = False
    FM_Main.Enabled = False
    flxMonitoring.Visible = False
    lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
    FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
    
    
    With flxMonitoring
        For i = 0 To .Rows - 1
        lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & i + 1 & " out of " & flxMonitoring.Rows - 1 & " row/s)"
            For c = 0 To 26 '.Cols - 1
                xlSheet.Range(Choose(c + 1, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA") & i + 1).Formula = .TextMatrix(i, c)
            Next c
        Next i
    
    End With
    
'
    With xlSheet
        .Columns("A:AA").EntireColumn.AutoFit
        .Cells.RowHeight = 15
        .Range("A1:AA1").Interior.ColorIndex = 37
        .Range("A1:AA1").Interior.Pattern = 1
        With .Range("A1:AA" & flxMonitoring.Rows)
                        .HorizontalAlignment = -4108
                        
                        '-Borders
                        For i = 7 To 12
                            .Borders(i).Weight = 2
                            .Borders(i).LineStyle = 1
                        Next i
                End With
    End With
    exportExcel = True
    xlApp.Visible = True
   
    
ErrExit:
    flxMonitoring.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    Set xlApp = Nothing
    
    Exit Function
    
ErrExcel:
    exportExcel = False
    MsgBox Err.Description, vbCritical, "ERROR -" & Err.Number
    GoTo ErrExit
    
    
End Function

Private Sub cmdClear_Click()

Call subFormatGrid(flxMonitoring, "prsmonitoring")
lblTotalRecord.Caption = 0
cboType.ListIndex = -1
cboDepartment.ListIndex = -1
cboCompany.ListIndex = -1
cboStatus.ListIndex = -1
txtMachineNo.Text = ""
txtMachineName.Text = ""
txtInCharge.Text = ""
dtFrom.Value = Now()
dtTo.Value = Now()
txtPRSNo.Text = ""

End Sub
