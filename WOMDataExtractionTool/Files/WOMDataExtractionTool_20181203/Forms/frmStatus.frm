VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmStatus 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "STATUS"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cdExcel 
      Left            =   14760
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetail 
      Height          =   4665
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8229
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   27
      BackColorFixed  =   4210688
      ForeColorFixed  =   16777215
      BackColorSel    =   16777215
      ForeColorSel    =   0
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
      _Band(0).Cols   =   27
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   1815
      TabIndex        =   11
      Top             =   1530
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   85590017
      CurrentDate     =   42731
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   3660
      TabIndex        =   12
      Top             =   1530
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   85590017
      CurrentDate     =   42731
   End
   Begin MSForms.ComboBox cboStatus 
      Height          =   330
      Left            =   1650
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1125
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status:"
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
      Left            =   90
      TabIndex        =   16
      Top             =   1125
      Width           =   1500
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Receive Date:"
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
      Left            =   90
      TabIndex        =   14
      Top             =   1530
      Width           =   1665
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
      Left            =   3300
      TabIndex        =   13
      Top             =   1575
      Width           =   300
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   645
      Left            =   8160
      TabIndex        =   10
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "CLEAR"
      Size            =   "2672;1138"
      Picture         =   "frmStatus.frx":0000
      Accelerator     =   67
      MouseIcon       =   "frmStatus.frx":1052
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cboCompany 
      Height          =   330
      Left            =   1650
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   135
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
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
   Begin MSForms.ComboBox cboDepartment 
      Height          =   330
      Left            =   1650
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   810
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
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
   Begin MSForms.ComboBox cboType 
      Height          =   330
      Left            =   1650
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   3330
      VariousPropertyBits=   746608667
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "5874;582"
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
   Begin VB.Label Label3 
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
      Left            =   90
      TabIndex        =   9
      Top             =   135
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department:"
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
      Left            =   90
      TabIndex        =   6
      Top             =   810
      Width           =   1500
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   645
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "SEARCH"
      Size            =   "2672;1138"
      Picture         =   "frmStatus.frx":20A4
      Accelerator     =   83
      MouseIcon       =   "frmStatus.frx":30F6
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   645
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   1515
      ForeColor       =   16777215
      BackColor       =   4210688
      Caption         =   "EXTRACT"
      Size            =   "2672;1138"
      Picture         =   "frmStatus.frx":4148
      Accelerator     =   69
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
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   1500
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
      Left            =   195
      TabIndex        =   2
      Top             =   4050
      Width           =   14820
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSM As Object
Private Function exportExcel(strType As String) As Boolean
        Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol As Long
    Dim i As Long
    Dim curWO As String
    Dim isSameWO As Boolean
    
    
     On Error GoTo ErrExcel
        
        If flxDetail.TextMatrix(1, 0) = "" Then Exit Function
        'strNewFile = CommonDialogSave(cdExcel, "MachineStatusReport")
        If blnCancel = True Then
            'Exit Sub
        End If
        exportExcel = False
            FM_Main.Enabled = False
            flxDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        
            
            Set xlApp = CreateObject("Excel.Application")
             Set xlBook = xlApp.Workbooks.Add
           Set xlSheet = xlBook.Sheets("Sheet1")
           
        Select Case strType
          Case ""
            With xlSheet
                .Range("A1:U2").Merge
                With .Range("A1")
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 14
                    .Formula = "DAILY REPORT FOR THE STATUS OF " & Me.cboType.Text
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                End With
                        
                    .Range("V2").Formula = "DATE:"
                    
                    .Range("A" & 4).Formula = "NO."
                    
                    .Range("B" & 4).Formula = "COMPANY "
                    .Range("C" & 4).Formula = "DEPARTMENT"
                    
                    .Range("D" & 4).Formula = "DATE OF W.O. "
                    .Range("E" & 4).Formula = "WORK CATEGORY"
                    .Range("F" & 4).Formula = "SECTION"
                    .Range("G" & 4).Formula = "LINE"
                    .Range("H" & 4).Formula = "PERSON IN-CHARGED / TL"
                    .Range("I" & 4).Formula = "W.O. #"
                    .Range("J" & 4).Formula = "EQPT. CONTROL NO."
                    .Range("K" & 4).Formula = "MACHINE NAME"
                    .Range("L" & 4).Formula = "TYPE OF REQUEST"
                    .Range("M" & 4).Formula = "SPECIFIC TROUBLE"
                    .Range("N" & 4).Formula = "STATUS"
                    .Range("O" & 4).Formula = "PARTS NEEDED"
                    .Range("P" & 4).Formula = "DATE OF MAKING MRS / MACHINE PARTS FOR REQUEST"
                    .Range("Q" & 4).Formula = "DATE OF MAKING PRS"
                    .Range("R" & 4).Formula = "PRS #"
                    .Range("S" & 4).Formula = "PO #"
                    .Range("T" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PRS)"
                    .Range("U" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PURCHASING)"
                    .Range("V" & 4).Formula = "DATE OF ACTUAL RECEIVING OF ITEM"
                    .Range("W" & 4).Formula = "DATE FINISHED"
                    .Range("X" & 4).Formula = "REMARKS"
                    
                    
            End With
            With xlSheet.Range("A4:X4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 35
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 9
                .Font.Bold = True
                For i = 7 To 11
                    .Borders(i).Weight = xlMedium
                Next i
            End With
            RowCtr = 5
            For intloop = 1 To flxDetail.Rows - 1
                
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                        
                         If flxDetail.TextMatrix(intloop, 8) = curWO Then
                               isSameWO = True
                        Else
                                isSameWO = False
                        End If
                        '----
                         .Range("I" & RowCtr).Formula = IIf(isSameWO, "", flxDetail.TextMatrix(intloop, 8))
                        If isSameWO Then
                                If flxDetail.TextMatrix(intloop - 1, 0) = flxDetail.TextMatrix(intloop, 0) Then
                                        .Range("A" & RowCtr).Formula = ""
                                        .Range("A" & RowCtr - 1 & ":" & "A" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 1) = flxDetail.TextMatrix(intloop, 1) Then
                                        .Range("B" & RowCtr).Formula = ""
                                        .Range("B" & RowCtr - 1 & ":" & "B" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 2) = flxDetail.TextMatrix(intloop, 2) Then
                                        .Range("C" & RowCtr).Formula = ""
                                        .Range("C" & RowCtr - 1 & ":" & "C" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 3) = flxDetail.TextMatrix(intloop, 3) Then
                                        .Range("D" & RowCtr).Formula = ""
                                        .Range("D" & RowCtr - 1 & ":" & "D" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 4) = flxDetail.TextMatrix(intloop, 4) Then
                                        .Range("E" & RowCtr).Formula = ""
                                        .Range("E" & RowCtr - 1 & ":" & "E" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 5) = flxDetail.TextMatrix(intloop, 5) Then
                                        .Range("F" & RowCtr).Formula = ""
                                        .Range("F" & RowCtr - 1 & ":" & "F" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 6) = flxDetail.TextMatrix(intloop, 6) Then
                                        .Range("G" & RowCtr).Formula = ""
                                        .Range("G" & RowCtr - 1 & ":" & "G" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 7) = flxDetail.TextMatrix(intloop, 7) Then
                                        .Range("H" & RowCtr).Formula = ""
                                        .Range("H" & RowCtr - 1 & ":" & "H" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 8) = flxDetail.TextMatrix(intloop, 8) Then
                                        .Range("I" & RowCtr).Formula = ""
                                        .Range("I" & RowCtr - 1 & ":" & "I" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 9) = flxDetail.TextMatrix(intloop, 9) Then
                                        .Range("J" & RowCtr).Formula = ""
                                        .Range("J" & RowCtr - 1 & ":" & "J" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 10) = flxDetail.TextMatrix(intloop, 10) Then
                                        .Range("K" & RowCtr).Formula = ""
                                        .Range("K" & RowCtr - 1 & ":" & "K" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 11) = flxDetail.TextMatrix(intloop, 11) Then
                                        .Range("L" & RowCtr).Formula = ""
                                        .Range("L" & RowCtr - 1 & ":" & "L" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 12) = flxDetail.TextMatrix(intloop, 12) Then
                                        .Range("M" & RowCtr).Formula = ""
                                        .Range("M" & RowCtr - 1 & ":" & "M" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 13) = flxDetail.TextMatrix(intloop, 13) Then
                                        .Range("N" & RowCtr).Formula = ""
                                        .Range("N" & RowCtr - 1 & ":" & "N" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 15) = flxDetail.TextMatrix(intloop, 15) Then
                                        .Range("P" & RowCtr).Formula = ""
                                        .Range("P" & RowCtr - 1 & ":" & "P" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 16) = flxDetail.TextMatrix(intloop, 16) Then
                                        .Range("Q" & RowCtr).Formula = ""
                                        .Range("Q" & RowCtr - 1 & ":" & "Q" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 17) = flxDetail.TextMatrix(intloop, 17) Then
                                        .Range("R" & RowCtr).Formula = ""
                                        .Range("R" & RowCtr - 1 & ":" & "R" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 18) = flxDetail.TextMatrix(intloop, 18) Then
                                        .Range("S" & RowCtr).Formula = ""
                                        .Range("S" & RowCtr - 1 & ":" & "S" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 19) = flxDetail.TextMatrix(intloop, 19) Then
                                        .Range("T" & RowCtr).Formula = ""
                                        .Range("T" & RowCtr - 1 & ":" & "T" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 20) = flxDetail.TextMatrix(intloop, 20) Then
                                        .Range("U" & RowCtr).Formula = ""
                                        .Range("U" & RowCtr - 1 & ":" & "U" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 21) = flxDetail.TextMatrix(intloop, 21) Then
                                        .Range("V" & RowCtr).Formula = ""
                                        .Range("V" & RowCtr - 1 & ":" & "V" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 22) = flxDetail.TextMatrix(intloop, 22) Then
                                        .Range("W" & RowCtr).Formula = ""
                                        .Range("W" & RowCtr - 1 & ":" & "W" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                                If flxDetail.TextMatrix(intloop - 1, 23) = flxDetail.TextMatrix(intloop, 23) Then
                                        .Range("X" & RowCtr).Formula = ""
                                        .Range("X" & RowCtr - 1 & ":" & "X" & RowCtr).Merge
                                Else
                                        GoTo defVal
                                End If
                        Else
defVal:
        
                                
                   
                    .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
                       
                    
                    .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
                    
                    
                    .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
                    .Range("F" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 5)
                    .Range("G" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 6)
                    .Range("H" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 7)
                    .Range("I" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 8)
                    .Range("J" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 9)
                    .Range("K" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 10)
                    .Range("L" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 11)
                    .Range("M" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 12)
                    .Range("N" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 13)
                   
                    .Range("P" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 15)
                    .Range("Q" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
                    .Range("R" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 17)
                    .Range("S" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 18)
                    .Range("T" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 19)
                    .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 20)
                    .Range("V" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 21)
                    .Range("W" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 22)
                    .Range("X" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)
                End If
                
                 .Range("O" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 14)
                    '-Insert row
                    curWO = flxDetail.TextMatrix(intloop, 8)
                    If flxDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
            
            '--- Excel Format -----------
            With xlSheet
                .PageSetup.PaperSize = xlPaperA3
                .PageSetup.Orientation = xlLandscape
                .PageSetup.Zoom = 54
                .PageSetup.BottomMargin = 25
                .PageSetup.FooterMargin = 0
                .PageSetup.TopMargin = 25
                .PageSetup.LeftMargin = 55
                .PageSetup.RightMargin = 55
                .PageSetup.PrintTitleRows = "$1:$4"
                .PageSetup.RightFooter = "&P"
                .Columns(14).ColumnWidth = 10
                .Columns(15).ColumnWidth = 40
                lblMessage.Caption = "Formatting Spreadsheet.."
'                .Columns("A:V").EntireColumn.AutoFit
                .Rows("4:" & RowCtr - 1).EntireRow.AutoFit
                With .Range("A5:X" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        .WrapText = True
                        '-Borders
                        For i = 7 To 12
                           
                            .Borders(i).Weight = xlThin
                            .Borders(i).LineStyle = xlContinuous
'                            With .Borders(i)
'                                .LineStyle = xlContinuous
'                                .Weight = xlThin
'                                .ColorIndex = xlAutomatic
'                            End With
                        Next i
                End With
                '-Date-
                    With .Range("W2")
                        .HorizontalAlignment = xlCenter
                        .Formula = Date & " " & Time
                    End With
                    '-
                '--Prepared By:--------
                With .Range("D" & RowCtr + 2 & ":E" & RowCtr + 3)
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "PREPARED BY:"
                    .Font.Name = "Arial Narrow"
                End With
                '--Underline----------
                With .Range("F" & RowCtr + 2 & ":G" & RowCtr + 3)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--Maint.Staff--
                With .Range("F" & RowCtr + 4 & ":G" & RowCtr + 4)
                    .HorizontalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "MAINT. STAFF"
                    .Font.Name = "Arial Narrow"
                End With
                '-Reviewed by: ----
                With .Range("P" & RowCtr + 2 & ":Q" & RowCtr + 3)
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "REVIEWED BY:"
                    .Font.Name = "Arial Narrow"
                End With
                '--Underline----------
                With .Range("R" & RowCtr + 2 & ":S" & RowCtr + 3)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--MAINT. ASV/ SV--
                With .Range("R" & RowCtr + 4 & ":S" & RowCtr + 4)
                    .HorizontalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "MAINT. ASV/ SV"
                    .Font.Name = "Arial Narrow"
                End With
                .Columns("A:X").EntireColumn.AutoFit
'                With xlSheet.PageSetup
'                    .PrintTitleRows = "$4:$4"
'                    .PrintTitleColumns = ""
'                    .Orientation = 2
'                    .RightFooter = "&P"
'                End With
            End With
            
            '======================
            
            'CASEEEEEEEEEEEEEEEEEEEEEEEEEe ELSEEEEEEee
            
            Case Else
                With xlSheet
                .Range("A1:U2").Merge
                With .Range("A1")
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 14
                    .Formula = "DAILY REPORT FOR THE STATUS OF " & Me.cboType.Text
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                End With
                        
                    .Range("V2").Formula = "DATE:"
                    
                    .Range("A" & 4).Formula = "NO."
        
                    .Range("B" & 4).Formula = "COMPANY "
                    .Range("C" & 4).Formula = "DEPARTMENT"
                    
                    .Range("D" & 4).Formula = "SECTION"
                    
                    .Range("E" & 4).Formula = "DATE OF W.O. "
                    .Range("F" & 4).Formula = "PERSON IN-CHARGED / TL"
                    .Range("G" & 4).Formula = "W.O. #"
                    .Range("H" & 4).Formula = "EQPT. CONTROL NO."
                    
                    .Range("I" & 4).Formula = "BRAND"
                    .Range("J" & 4).Formula = "MODEL"
                    
                    .Range("K" & 4).Formula = "MACHINE NAME"
                    .Range("L" & 4).Formula = "TYPE OF REQUEST"
                    .Range("M" & 4).Formula = "SPECIFIC TROUBLE"
                    .Range("N" & 4).Formula = "STATUS"
                    .Range("O" & 4).Formula = "PARTS NEEDED"
                    .Range("P" & 4).Formula = "QTY"
                    .Range("Q" & 4).Formula = "DATE OF MAKING MRS / MACHINE PARTS FOR REQUEST"
                    .Range("R" & 4).Formula = "DATE OF MAKING PRS"
                    .Range("S" & 4).Formula = "PRS #"
                    .Range("T" & 4).Formula = "PO #"
                    .Range("U" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PRS)"
                    .Range("V" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PURCHASING)"
                    .Range("W" & 4).Formula = "DATE OF ACTUAL RECEIVING OF ITEM"
                    
                    .Range("X" & 4).Formula = "SCHEDULE OF REPAIR"
                    .Range("Y" & 4).Formula = "ACTUAL REPAIR DATE"
                    .Range("Z" & 4).Formula = "DATE FINISHED"
                    
                    .Range("AA" & 4).Formula = "REMARKS"
                    
                    
            End With
            With xlSheet.Range("A4:AA4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 35
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 9
                .Font.Bold = True
                For i = 7 To 11
                    .Borders(i).Weight = xlMedium
                Next i
            End With
            RowCtr = 5
            For intloop = 1 To flxDetail.Rows - 1
                
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
                    
                    .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
                    .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
                    .Range("F" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 5)
                    .Range("G" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 6)
                    .Range("H" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 7)
                    .Range("I" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 8)
                    .Range("J" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 9)
                    .Range("K" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 10)
                    .Range("L" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 11)
                    .Range("M" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 12)
                    .Range("N" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 13)
                    .Range("O" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 14)
                    .Range("P" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 15)
                    .Range("Q" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
                    .Range("R" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 17)
                    .Range("S" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 18)
                    .Range("T" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 19)
                    .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 20)
                    .Range("V" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 21)
                    .Range("W" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 22)
                    .Range("X" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)
                    .Range("Y" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 24)
                    .Range("Z" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 25)
                    '-Insert row
                    If flxDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
            
            '--- Excel Format -----------
            With xlSheet
                .PageSetup.PaperSize = xlPaperA3
                .PageSetup.Orientation = xlLandscape
                .PageSetup.Zoom = 54
                .PageSetup.BottomMargin = 25
                .PageSetup.FooterMargin = 0
                .PageSetup.TopMargin = 25
                .PageSetup.LeftMargin = 55
                .PageSetup.RightMargin = 55
                .PageSetup.PrintTitleRows = "$1:$4"
                .PageSetup.RightFooter = "&P"
                .Columns(14).ColumnWidth = 10
                .Columns(15).ColumnWidth = 40
                lblMessage.Caption = "Formatting Spreadsheet.."
'                .Columns("A:V").EntireColumn.AutoFit
                .Rows("4:" & RowCtr - 1).EntireRow.AutoFit
                With .Range("A5:AA" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        .WrapText = True
                        '-Borders
                        For i = 7 To 12
                           
                            .Borders(i).Weight = xlThin
                            .Borders(i).LineStyle = xlContinuous
'                            With .Borders(i)
'                                .LineStyle = xlContinuous
'                                .Weight = xlThin
'                                .ColorIndex = xlAutomatic
'                            End With
                        Next i
                End With
                '-Date-
                    With .Range("W2")
                        .HorizontalAlignment = xlCenter
                        .Formula = Date & " " & Time
                    End With
                    '-
                '--Prepared By:--------
                With .Range("D" & RowCtr + 2 & ":E" & RowCtr + 3)
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "PREPARED BY:"
                    .Font.Name = "Arial Narrow"
                End With
                '--Underline----------
                With .Range("F" & RowCtr + 2 & ":G" & RowCtr + 3)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--Maint.Staff--
                With .Range("F" & RowCtr + 4 & ":G" & RowCtr + 4)
                    .HorizontalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "MAINT. STAFF"
                    .Font.Name = "Arial Narrow"
                End With
                '-Reviewed by: ----
                With .Range("P" & RowCtr + 2 & ":Q" & RowCtr + 3)
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "REVIEWED BY:"
                    .Font.Name = "Arial Narrow"
                End With
                '--Underline----------
                With .Range("R" & RowCtr + 2 & ":S" & RowCtr + 3)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--MAINT. ASV/ SV--
                With .Range("R" & RowCtr + 4 & ":S" & RowCtr + 4)
                    .HorizontalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "MAINT. ASV/ SV"
                    .Font.Name = "Arial Narrow"
                End With
                .Columns("A:AA").EntireColumn.AutoFit
'                With xlSheet.PageSetup
'                    .PrintTitleRows = "$4:$4"
'                    .PrintTitleColumns = ""
'                    .Orientation = 2
'                    .RightFooter = "&P"
'                End With
            End With
            
         End Select
'        xlBook.Save
'        xlBook.Close
'        xlApp.Quit
        exportExcel = True
        xlApp.Visible = True
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
       
        '- Open extracted report report --
        'Shell "explorer " & strNewFile, vbMaximizedFocus
        '-
        
'        Set xlSheet = Nothing
'        Set xlBook = Nothing
        Set xlApp = Nothing
        Exit Function
        
ErrExcel:
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        exportExcel = False
End Function



Private Function exportExcel_forklift() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol As Long
    Dim i As Long
    
    On Error GoTo ErrExcel
        
        If flxDetail.TextMatrix(1, 0) = "" Then Exit Function
        'strNewFile = CommonDialogSave(cdExcel, "MachineStatusReport")
        If blnCancel = True Then
            'Exit Sub
        End If
        exportExcel_forklift = False
            FM_Main.Enabled = False
            flxDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        
            
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Sheets("Sheet1")
            
        With xlSheet
            .Range("A1:U2").Merge
            With .Range("A1")
                .Font.Name = "Arial Narrow"
                .Font.Size = 14
                .Formula = "DAILY REPORT FOR THE STATUS OF " & Me.cboType.Text
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
            End With
                        
                    .Range("V2").Formula = "DATE:"
                    
                    .Range("A" & 4).Formula = "NO."
                    
                    .Range("B" & 4).Formula = "COMPANY "
                    .Range("C" & 4).Formula = "DEPARTMENT"
                    
                    .Range("D" & 4).Formula = "DATE OF W.O. "
                    .Range("E" & 4).Formula = "WORK CATEGORY"
                    .Range("F" & 4).Formula = "SECTION"
                    .Range("G" & 4).Formula = "LINE"
                    .Range("H" & 4).Formula = "PERSON IN-CHARGED / TL"
                    .Range("I" & 4).Formula = "W.O. #"
                    .Range("J" & 4).Formula = "EQPT. CONTROL NO."
                    
                    .Range("K" & 4).Formula = "BRAND"
                    .Range("L" & 4).Formula = "MODEL"
                    
                    .Range("M" & 4).Formula = "MACHINE NAME"
                    .Range("N" & 4).Formula = "TYPE OF REQUEST"
                    .Range("O" & 4).Formula = "SPECIFIC TROUBLE"
                    .Range("P" & 4).Formula = "STATUS"
                    .Range("Q" & 4).Formula = "PARTS NEEDED"
                    .Range("R" & 4).Formula = "DATE OF MAKING MRS / MACHINE PARTS FOR REQUEST"
                    .Range("S" & 4).Formula = "DATE OF MAKING PRS"
                    .Range("T" & 4).Formula = "PRS #"
                    .Range("U" & 4).Formula = "PO #"
                    .Range("V" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PRS)"
                    .Range("W" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PURCHASING)"
                    .Range("X" & 4).Formula = "DATE OF ACTUAL RECEIVING OF ITEM"
                    .Range("Y" & 4).Formula = "DATE FINISHED"
                    
                    .Range("Z" & 4).Formula = "SCHEDULE OF REPAIR"
                    .Range("AA" & 4).Formula = "ACTUAL REPAIR DATE"
                    
                    .Range("AB" & 4).Formula = "REMARKS"
                    
                    
            End With
            With xlSheet.Range("A4:AB4")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Interior.ColorIndex = 35
                .Interior.Pattern = xlSolid
                .Font.Name = "Arial Narrow"
                .Font.Size = 9
                .Font.Bold = True
                For i = 7 To 11
                    .Borders(i).Weight = xlMedium
                Next i
            End With
            RowCtr = 5
            For intloop = 1 To flxDetail.Rows - 1
                
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
                    
                    .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
                    
                    
                    .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
                    .Range("F" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 5)
                    .Range("G" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 6)
                    .Range("H" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 7)
                    .Range("I" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 8)
                    .Range("J" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 9)
                    
                    .Range("K" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 10)
                    .Range("L" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 11)
                    
                    .Range("M" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 12)
                    .Range("N" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 13)
                    .Range("O" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 14)
                    .Range("P" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 15)
                    .Range("Q" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 16)
                    .Range("R" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 17)
                    .Range("S" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 18)
                    .Range("T" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 19)
                    .Range("U" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 20)
                    .Range("V" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 21)
                    .Range("W" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 22)
                    .Range("X" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 23)
                    .Range("Y" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 24)
                    
                    .Range("Z" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 25)
                    .Range("AA" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 26)
                    
                    .Range("AB" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 27)
                    '-Insert row
                    If flxDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
            
            '--- Excel Format -----------
            With xlSheet
                .PageSetup.PaperSize = xlPaperA3
                .PageSetup.Orientation = xlLandscape
                .PageSetup.Zoom = 54
                .PageSetup.BottomMargin = 25
                .PageSetup.FooterMargin = 0
                .PageSetup.TopMargin = 25
                .PageSetup.LeftMargin = 55
                .PageSetup.RightMargin = 55
                .PageSetup.PrintTitleRows = "$1:$4"
                .PageSetup.RightFooter = "&P"
                .Columns(14).ColumnWidth = 10
                .Columns(15).ColumnWidth = 40
                lblMessage.Caption = "Formatting Spreadsheet.."
'                .Columns("A:V").EntireColumn.AutoFit
                .Rows("4:" & RowCtr - 1).EntireRow.AutoFit
                With .Range("A5:X" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        .WrapText = True
                        '-Borders
                        For i = 7 To 12
                           
                            .Borders(i).Weight = xlThin
                            .Borders(i).LineStyle = xlContinuous
'                            With .Borders(i)
'                                .LineStyle = xlContinuous
'                                .Weight = xlThin
'                                .ColorIndex = xlAutomatic
'                            End With
                        Next i
                End With
                '-Date-
                    With .Range("Y2")
                        .HorizontalAlignment = xlCenter
                        .Formula = Date & " " & Time
                    End With
                    '-
                '--Prepared By:--------
                With .Range("D" & RowCtr + 2 & ":E" & RowCtr + 3)
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "PREPARED BY:"
                    .Font.Name = "Arial Narrow"
                End With
                '--Underline----------
                With .Range("F" & RowCtr + 2 & ":G" & RowCtr + 3)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--Maint.Staff--
                With .Range("F" & RowCtr + 4 & ":G" & RowCtr + 4)
                    .HorizontalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "MAINT. STAFF"
                    .Font.Name = "Arial Narrow"
                End With
                '-Reviewed by: ----
                With .Range("P" & RowCtr + 2 & ":Q" & RowCtr + 3)
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "REVIEWED BY:"
                    .Font.Name = "Arial Narrow"
                End With
                '--Underline----------
                With .Range("R" & RowCtr + 2 & ":S" & RowCtr + 3)
                    .Merge
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                    .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                End With
                '--MAINT. ASV/ SV--
                With .Range("R" & RowCtr + 4 & ":S" & RowCtr + 4)
                    .HorizontalAlignment = xlCenter
                    .Merge
                    .FormulaR1C1 = "MAINT. ASV/ SV"
                    .Font.Name = "Arial Narrow"
                End With
                .Columns("A:AB").EntireColumn.AutoFit
'                With xlSheet.PageSetup
'                    .PrintTitleRows = "$4:$4"
'                    .PrintTitleColumns = ""
'                    .Orientation = 2
'                    .RightFooter = "&P"
'                End With
            End With
                
'        xlBook.Save
'        xlBook.Close
'        xlApp.Quit
        exportExcel_forklift = True
        xlApp.Visible = True
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
       
        '- Open extracted report report --
        'Shell "explorer " & strNewFile, vbMaximizedFocus
        '-
        
'        Set xlSheet = Nothing
'        Set xlBook = Nothing
        Set xlApp = Nothing
        Exit Function
        
ErrExcel:
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        exportExcel_forklift = False
End Function

Private Sub cmdClear_Click()
Call Connect
    Call subFormatGrid(flxDetail, "status")
    Call LoadDataToCombo(cboType, "Types")
    Call LoadDataToCombo(cboCompany, "Companies")
    Call LoadDataToCombo(cboStatus, "Status", , True)
Call Disconnect
    Me.cboType.Clear
    Me.cboDepartment.Clear
End Sub

Private Sub cmdExcel_Click()
    If flxDetail.TextMatrix(1, 0) = "" Then Exit Sub
    FM_Main.MousePointer = vbCustom
    
    If exportExcel(IIf(Me.cboType.Text = "FORKLIFT", "Forklift", "")) = True Then
        MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
'    ElseIf exportLibre = True Then
'        MsgBox "Report Succesfully saved to LibreOffice!", vbInformation, "System Information"
'    ElseIf exportExcel = True Then
'        MsgBox "Report Succesfully saved to Excel!", vbInformation, "System Information"
    Else
        MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
    End If
    FM_Main.MousePointer = vbDefault
         
End Sub

Private Function MakePropertyValue(propName, propVal) As Object
    Dim oPropValue As Object
    Set oPropValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    oPropValue.Name = propName
    oPropValue.Value = propVal
    Set MakePropertyValue = oPropValue
End Function


Private Function exportLibre() As Boolean
     Dim oDoc As Object, _
        oDesk As Object, _
        oSheet As Object, _
        oPar(1) As Object, _
        CellProp As Object, _
        CellStyle As Object, _
        NewStyle As Object, _
        oRange As Object, _
        oColumns As Object, _
        PageStyles As Object, _
        NewPageStyle As Object, _
        StdPage As Object, _
        basicBorder As Object, _
        oBorder As Object
    Dim oCol As Long, oRow As Long
    Dim Charts As Object
    Dim Chart As Object
    Dim Rect As Object
    Dim oChartRange As Object
    Dim RangeAddress(0) As Object
    
    exportLibre = False
    oRow = 0
    
    Set oSM = CreateObject("com.sun.star.ServiceManager")
    Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
    Set oPar(0) = MakePropertyValue("Hidden", True)
    Set oPar(1) = MakePropertyValue("Overwrite", True)
    Set oDoc = oDesk.LoadComponentFromURL("private:factory/scalc", "_blank", 0, oPar)
    Set oSheet = oDoc.Sheets.getByIndex("0")
    Set CellProp = oDoc.StyleFamilies.getByName("CellStyles")
    Set NewStyle = oDoc.createInstance("com.sun.star.style.CellStyle")
    Call CellProp.InsertbyName("MyStyle", NewStyle)
    NewStyle.ParentStyle = "Default"
    Set CellStyle = CellProp.getByName("MyStyle")
    Set PageStyles = oDoc.StyleFamilies.getByName("PageStyles")
    Set StdPage = PageStyles.getByName("Default")
    Set basicBorder = oDoc.Bridge_GetStruct("com.sun.star.table.BorderLine")
    
    basicBorder.Color = RGB(0, 0, 0)
    basicBorder.InnerLineWidth = 0
    basicBorder.OuterLineWidth = 11
    basicBorder.LineDistance = 0
    
'        With StdPage
'            .FooterIsOn = False
'            .HeaderIsOn = False
'            .IsLandscape = False
'            .Width = 29700
'            .Height = 21000
'            .LeftMargin = 1000
'            .RightMargin = 1000
'            .TopMargin = 1000
'            .BottomMargin = 1000
'        End With
        '-Header--
        
            FM_Main.Enabled = False
            flxDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        oSheet.getCellByPosition(0, oRow).String = "DAILY REPORT FOR THE STATUS OF MACHINE"
         
                
        For oCol = 0 To flxDetail.Cols - 1
           With oSheet
                With CellStyle
                    .CharWeight = 300
                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .CellBackColor = RGB(255, 184, 0)
                End With
                .getCellByPosition(oCol, 1).CellStyle = "MyStyle"
                .getCellByPosition(oCol, 1).String = flxDetail.TextMatrix(0, oCol)
                .getCellByPosition(oCol, 1).HoriJustify = 2
           End With
            Set oBorder = oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder
            oBorder.LeftLine = basicBorder
            oBorder.Topline = basicBorder
            oBorder.RightLine = basicBorder
            oBorder.BottomLine = basicBorder
            oBorder.VerticalLine = basicBorder
            oSheet.getCellRangeByPosition(0, 1, flxDetail.Cols - 1, 1).TableBorder = oBorder
        Next oCol
    
        '-Content--
        For oRow = 0 To flxDetail.Rows - 1
            DoEvents
            
            lblMessage.Caption = "Exporting Data : " & oRow & " rows out of " & flxDetail.Rows - 1 & " rows"
            For oCol = 0 To flxDetail.Cols - 1
               With oSheet
                    .getCellByPosition(oCol, oRow + 1).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .getCellByPosition(oCol, oRow + 1).HoriJustify = 2
                    flxDetail.Row = oRow
                    flxDetail.Col = oCol
                    .getCellByPosition(oCol, oRow + 1).String = flxDetail.TextMatrix(oRow, oCol)
                    If flxDetail.CellBackColor = &H8080FF Then
                        .getCellByPosition(oCol, oRow + 1).CellBackColor = RGB(128, 128, 255)
                    End If
               End With
            Next oCol
                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder
                oBorder.LeftLine = basicBorder
                oBorder.Topline = basicBorder
                oBorder.RightLine = basicBorder
                oBorder.BottomLine = basicBorder
                oBorder.VerticalLine = basicBorder
                oSheet.getCellRangeByPosition(0, oRow + 1, flxDetail.Cols - 1, oRow + 1).TableBorder = oBorder
        Next oRow
     
        
        
        Call oDoc.storeToURL("file:///C:/Exported.xls", oPar)
        Set oPar(0) = MakePropertyValue("Hidden", False)
        Set oDoc = oDesk.LoadComponentFromURL("file:///C:/Exported.xls", "_blank", 0, oPar)
        exportLibre = True
    
    
        Set oSM = Nothing
        Set oDesk = Nothing
        Set oDoc = Nothing
        Set oSheet = Nothing
        Set oPar(1) = Nothing
        Set CellProp = Nothing
        Set CellStyle = Nothing
        Set NewStyle = Nothing
        Set oRange = Nothing
        Set oColumns = Nothing
        Set PageStyles = Nothing
        Set NewPageStyle = Nothing
        Set StdPage = Nothing
        exportLibre = True
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        Exit Function
ErrLibre:
        exportLibre = False
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
End Function

Private Sub LoadFlexDailyReportStatus(strWhere As String, Optional strServer As String)
    Dim rsFlex As ADODB.Recordset
    Dim lngLoop, i, c As Long
    Dim lngrow As Long
    Dim lngNo As Long
    Dim lngRecCnt As Long
    
Call Connect
    
    Set rsFlex = cls_GetDetails.pfLoadStatus(Me.cboType.Column(0), strWhere, strServer)
    
    If rsFlex.EOF Then
        MsgBox "No Record found!", vbInformation, "System Information"
        Call subFormatGrid(flxDetail, "status")
        GoTo LDExit
    Else
       
        With flxDetail
            rsFlex.MoveLast
            lngRecCnt = rsFlex.RecordCount
            rsFlex.MoveFirst
            .Rows = lngRecCnt + 1
            For i = 1 To lngRecCnt
                DoEvents
                For c = 0 To .Cols - 1
                    DoEvents
                    .Row = i
                    .Col = c
                      .CellAlignment = flexAlignLeftCenter
                    .TextMatrix(i, c) = Choose(c + 1, i, _
                                                            pfvarNoValue(rsFlex.Fields("CompanyName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("DepartmentName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("DateOfWorkOrder").Value), _
                                                            pfvarNoValue(rsFlex.Fields("WorkCategory").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Section").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Line").Value), _
                                                            pfvarNoValue(rsFlex.Fields("LeaderIncharge").Value), _
                                                            pfvarNoValue(rsFlex.Fields("WorkOrderControlNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MachineItemNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MachineName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("TypeOfRequest").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MachineProblem").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Status").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Description").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsDate").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsDate").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PoNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsExpectedDelivery").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PoExpectedDelivery").Value), _
                                                            pfvarNoValue(rsFlex.Fields("ActualReceived").Value), _
                                                            pfvarNoValue(rsFlex.Fields("FinishedDate").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Remarks").Value), "", "", "", "")
                                                                
                                                                    
                                                    
                Next c
                rsFlex.MoveNext
            Next i
            .Visible = True
        End With
        FM_Main.StatusBar1.Panels(3).Text = ""
        
   End If
Call Disconnect
LDExit:
    Me.flxDetail.Visible = True
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsFlex = Nothing
    Exit Sub
LDErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "System Error"
    GoTo LDExit
End Sub
Private Sub LoadFlexDailyReportStatus_forklift(strWhere As String, strServer As String)
    Dim rsFlex As ADODB.Recordset
    Dim lngLoop, i, c As Long
    Dim lngrow As Long
    Dim lngNo As Long
    Dim lngRecCnt As Long
    
Call Connect
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    
    Set rsFlex = cls_GetDetails.pfLoadStatus(Me.cboType.Column(0), strWhere, strServer)
    
    If rsFlex.EOF Then
        MsgBox "No Record found!", vbOKOnly + vbInformation, "System Information"
        Call subFormatGrid(flxDetail, "status")
        GoTo LDExit
    Else
       
        lblMessage.Caption = "Please Wait. Loading Data.."
        FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
        With flxDetail
            .Visible = False
            rsFlex.MoveLast
            lngRecCnt = rsFlex.RecordCount
            rsFlex.MoveFirst
            .Rows = lngRecCnt + 1
            For i = 1 To lngRecCnt
                DoEvents
                For c = 0 To .Cols - 1
                    DoEvents
                    .Row = i
                    .Col = c
                    .CellAlignment = flexAlignLeftCenter
                    .TextMatrix(i, c) = Choose(c + 1, i, _
                                                            pfvarNoValue(rsFlex.Fields("CompanyName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("DepartmentName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Section").Value), _
                                                            pfvarNoValue(rsFlex.Fields("DateOfWorkOrder").Value), _
                                                            pfvarNoValue(rsFlex.Fields("LeaderIncharge").Value), _
                                                            pfvarNoValue(rsFlex.Fields("WorkOrderControlNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MachineItemNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MakerName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Model").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MachineName").Value), _
                                                            pfvarNoValue(rsFlex.Fields("TypeOfRequest").Value), _
                                                            pfvarNoValue(rsFlex.Fields("MachineProblem").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Status").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Description").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Qty").Value) & " " & pfvarNoValue(rsFlex.Fields("QtyDescription").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsDate").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsDate").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PoNo").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PrsExpectedDelivery").Value), _
                                                            pfvarNoValue(rsFlex.Fields("PoExpectedDelivery").Value), _
                                                            pfvarNoValue(rsFlex.Fields("ActualReceived").Value), "", "", _
                                                            pfvarNoValue(rsFlex.Fields("FinishedDate").Value), _
                                                            pfvarNoValue(rsFlex.Fields("Remarks").Value))
                                                            
                                                    
                Next c
                rsFlex.MoveNext
            Next i
            .Visible = True
        End With
        FM_Main.StatusBar1.Panels(3).Text = ""
        
   End If
Call Disconnect
LDExit:
    FM_Main.StatusBar1.Panels(3).Text = ""
    FM_Main.Enabled = True
    FM_Main.MousePointer = vbDefault
    Set rsFlex = Nothing
    Exit Sub
LDErr:
    MsgBox Err.Description, vbOKOnly + vbCritical, "System Error"
    GoTo LDExit
End Sub

Private Sub cboCompany_Click()
Call Connect
    Call LoadDataToCombo(cboDepartment, "Departments", cboCompany.Column(0))
    Call LoadDataToCombo(cboType, "Types", cboCompany.Column(0))
Call Disconnect
End Sub
Private Sub cmdSearch_Click()
    Dim strWhere  As String
    
    If Me.cboCompany.Text = "" Then
        MsgBox "Please Select Company", vbOKOnly + vbInformation, "System Information"
        Exit Sub
    ElseIf Me.cboType.Text = "" Then
        MsgBox "Please Select Types", vbOKOnly + vbInformation, "System Information"
        Exit Sub
    End If
    If Me.cboType.Text = "FORKLIFT" Then
        Call subFormatGrid(flxDetail, "status_forklift")
    Else
        Call subFormatGrid(flxDetail, "status")
    End If

    'Call subFormatGrid(flxDetail, "status")
    
    strWhere = ""
    If cboCompany.Text <> "" Then
        strWhere = strWhere & " AND CompanyName = '" & cboCompany.Column(1) & "'"
    End If
    If cboDepartment.Text <> "" Then
        strWhere = strWhere & " AND DepartmentId = " & cboDepartment.Column(0)
    End If
    If cboStatus.Text <> "" Then
        If cboStatus.Text = "FINISHED" Then
            strWhere = strWhere & " AND Finisheddate IS NOT NULL"
        Else
            strWhere = strWhere & " AND Statusid  = " & cboStatus.Column(0)
            strWhere = strWhere & " AND Finisheddate IS NULL"
        End If
    Else
        strWhere = strWhere & " AND StatusID IN (1,2,4)"
        strWhere = strWhere & " AND Finisheddate is  NULL"
    End If
    strWhere = strWhere & " AND CONVERT(VARCHAR(20),DateOfWorkOrder,111) >= '" & dtFrom.Value & "'"
    strWhere = strWhere & " AND CONVERT(VARCHAR(20),DateOfWorkOrder,111)  <= '" & dtTo.Value & "'"
    
    FM_Main.MousePointer = vbCustom
    FM_Main.Enabled = False
    Me.flxDetail.Visible = False
    lblMessage.Caption = "Please Wait. Loading Data.."
    FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
    
    'Call LoadFlexDailyReportStatus(strWhere, GetServerName(cboCompany.Column(0)))
    
    If Me.cboType.Text = "FORKLIFT" Then
        Call LoadFlexDailyReportStatus_forklift(strWhere, GetServerName(cboCompany.Column(0)))
    Else
        Call LoadFlexDailyReportStatus(strWhere, GetServerName(cboCompany.Column(0)))
    End If
End Sub


Private Sub Form_Load()
Call Connect
    Me.dtFrom.Value = Date
    Me.dtTo.Value = Date
    Call subFormatGrid(flxDetail, "status")
    Call LoadDataToCombo(cboType, "Types")
    Call LoadDataToCombo(cboCompany, "Companies")
   
    Call LoadDataToCombo(cboStatus, "Status", , True)
Call Disconnect
End Sub



