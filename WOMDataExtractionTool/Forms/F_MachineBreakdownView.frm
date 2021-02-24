VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form F_MachineBreakdownView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Breakdown View"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17145
   Icon            =   "F_MachineBreakdownView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   17145
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   945
      TabIndex        =   0
      Top             =   225
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16777217
      CurrentDate     =   42678
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetail 
      Height          =   3705
      Left            =   0
      TabIndex        =   1
      Top             =   675
      Width           =   17040
      _ExtentX        =   30057
      _ExtentY        =   6535
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   11627568
      ForeColorFixed  =   16777215
      ForeColorSel    =   -2147483635
      GridColorFixed  =   8421504
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
      _Band(0).Cols   =   5
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2475
      TabIndex        =   2
      Top             =   225
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16777217
      CurrentDate     =   42678
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSummaryDetail 
      Height          =   3705
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   6535
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   11627568
      ForeColorFixed  =   16777215
      ForeColorSel    =   -2147483635
      GridColorFixed  =   8421504
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
      _Band(0).Cols   =   11
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4455
      TabIndex        =   11
      Top             =   270
      Width           =   645
   End
   Begin MSForms.ComboBox cboType 
      Height          =   330
      Left            =   5085
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   225
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
   Begin VB.Label lblMessage1 
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
      Left            =   90
      TabIndex        =   9
      Top             =   2385
      Width           =   14820
   End
   Begin VB.Label lblMessage2 
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
      Left            =   90
      TabIndex        =   8
      Top             =   6930
      Width           =   14820
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   315
      TabIndex        =   7
      Top             =   315
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SUMMARY "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   4770
      Width           =   2355
   End
   Begin MSForms.CommandButton cmdSearch 
      Height          =   390
      Left            =   8415
      TabIndex        =   5
      Top             =   225
      Width           =   1605
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "SEARCH"
      Size            =   "2831;688"
      Accelerator     =   83
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExcel 
      Height          =   390
      Left            =   225
      TabIndex        =   4
      Top             =   8865
      Width           =   1605
      ForeColor       =   16777215
      BackColor       =   11627568
      Caption         =   "EXTRACT"
      Size            =   "2831;688"
      Accelerator     =   69
      FontName        =   "‚l‚r ‚oƒSƒVƒbƒN"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "F_MachineBreakdownView"
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
    Dim oCol As Long, oRow As Long, nRow As Long
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
            lblMessage1.Caption = "Please Wait. Exporting Data to Excel.."
              lblMessage2.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        oSheet.getCellByPosition(0, oRow).String = "MACHINE BREAKDOWN WORK ORDER (" & Format(DTPicker1, "mmmm dd,yyyy") & " - " & Format(DTPicker2, "mmmm dd,yyyy") & ")"
         
                
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
            
            lblMessage1.Caption = "Exporting Data : " & oRow & " rows out of " & flxDetail.Rows - 1 & " rows"
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
        
       
        
        For oCol = 0 To flxSummaryDetail.Cols - 1
           With oSheet
                With CellStyle
                    .CharWeight = 300
                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .CellBackColor = RGB(255, 184, 0)
                End With
                .getCellByPosition(oCol, oRow + 2).CellStyle = "MyStyle"
                .getCellByPosition(oCol, oRow + 2).String = flxSummaryDetail.TextMatrix(0, oCol)
                .getCellByPosition(oCol, oRow + 2).HoriJustify = 2
           End With
            Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 2, flxSummaryDetail.Cols - 1, oRow + 2).TableBorder
            oBorder.LeftLine = basicBorder
            oBorder.Topline = basicBorder
            oBorder.RightLine = basicBorder
            oBorder.BottomLine = basicBorder
            oBorder.VerticalLine = basicBorder
            oSheet.getCellRangeByPosition(0, oRow + 2, flxSummaryDetail.Cols - 1, oRow + 2).TableBorder = oBorder
        Next oCol
        
    
        '-Content--
        For oRow = 0 To flxSummaryDetail.Rows - 1
            DoEvents
            
            'lblMessage.Caption = "Exporting Data : " & oRow & " rows out of " & flxSummaryDetail.Rows - 1 & " rows"
            For oCol = 0 To flxSummaryDetail.Cols - 1
               With oSheet
                    .getCellByPosition(oCol, oRow + 3).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .getCellByPosition(oCol, oRow + 3).HoriJustify = 2
                    flxSummaryDetail.Row = oRow
                    flxSummaryDetail.Col = oCol
                    .getCellByPosition(oCol, oRow + 7).String = flxSummaryDetail.TextMatrix(oRow, oCol)
                    If flxSummaryDetail.CellBackColor = &H8080FF Then
                        .getCellByPosition(oCol, oRow + 3).CellBackColor = RGB(128, 128, 255)
                    End If
               End With
            Next oCol
                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 3, flxSummaryDetail.Cols - 1, oRow + 3).TableBorder
                oBorder.LeftLine = basicBorder
                oBorder.Topline = basicBorder
                oBorder.RightLine = basicBorder
                oBorder.BottomLine = basicBorder
                oBorder.VerticalLine = basicBorder
                oSheet.getCellRangeByPosition(0, oRow + 3, flxSummaryDetail.Cols - 1, oRow + 3).TableBorder = oBorder
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

Private Sub cmdExcel_Click()
        If flxDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "WODataExtractionTool"
        ElseIf exportLibre = True Then
            MsgBox "Report Succesfully saved to LibreOffice!", vbInformation, "WODataExtractionTool"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
         
End Sub

Private Function exportExcel() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol1, curCol2 As Long
    Dim i As Long
    
    On Error GoTo ErrSave
        If flxDetail.TextMatrix(1, 0) = "" Or flxSummaryDetail.TextMatrix(1, 0) = "" Then Exit Function
        exportExcel = False
            FM_Main.Enabled = False
            flxDetail.Visible = False
            flxSummaryDetail.Visible = False
            lblMessage1.Caption = "Please Wait. Exporting Data to Spreadsheet.."
            lblMessage2.Caption = "Please Wait. Exporting Data to Spreadsheet.."
            
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Spreadsheet.."
        
            'Call FileCopy(App.Path & "\Reports\Machine Breakdown Report.xls", strNewFile)
            
            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Add
            'Set xlBook = xlApp.Workbooks.Open(strNewFile)
            Set xlSheet = xlBook.Sheets("Sheet1")
            RowCtr = 1
            With xlSheet
                   
                    .Range("A3").Formula = "RECEIVED DATE"
                    .Range("B3").Formula = "DEPARTMENT NAME"
                    .Range("C3").Formula = "SECTION NAME"
                    .Range("D3").Formula = "RECEIVED"
                    .Range("E3").Formula = "FINISHED"
                    
            End With
            '-For breakdown report
            RowCtr = 4
            For intloop = 1 To flxDetail.Rows - 1
                
                lblMessage1.Caption = "Please Wait. Exporting Data to Spreadsheet.. (" & intloop & " out of " & flxDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    .Range("A" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 0)
                    .Range("B" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 2)
                    .Range("D" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxDetail.TextMatrix(intloop, 4)
                    '-Insert row
                    If flxDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
             lblMessage1.Caption = ""
            curCol1 = RowCtr
            
            'For SUMMARY REPORT--
            curCol2 = RowCtr + 4
            RowCtr = RowCtr + 3
            With xlSheet
                    .Range("A" & RowCtr).Formula = "DEPARTMENT NAME"
                    .Range("B" & RowCtr).Formula = "SECTION NAME"
                    .Range("C" & RowCtr).Formula = "RECEIVED"
                    .Range("D" & RowCtr).Formula = "FINISHED"
                    .Range("E" & RowCtr).Formula = "FINISHED WO FROM PENDINGWO"
                    .Range("F" & RowCtr).Formula = "FINISHED ON THE SUCCEEDING MONTH"
                    .Range("G" & RowCtr).Formula = "CANCELLED"
                    .Range("H" & RowCtr).Formula = "TURNOVER"
                    .Range("I" & RowCtr).Formula = "WAITING PARTS"
                    .Range("J" & RowCtr).Formula = "FOR SCHEDULE"
                    .Range("K" & RowCtr).Formula = "FOR CONFIRMATION"
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
                lblMessage1.Caption = "Formatting Spreadsheet.."
                lblMessage2.Caption = "Formatting Spreadsheet.."
                '-borders
                With .Range("A4:E" & curCol1 - 1)
                        .HorizontalAlignment = xlCenter
                        .Font.Bold = False
                        .WrapText = True
                        For i = 7 To 12
                            With .Borders(i)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .ColorIndex = xlAutomatic
                            End With
                        Next i
                End With
                With .Range("A" & curCol2 & ":K" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        .Font.Bold = False
                        .WrapText = True
                        For i = 7 To 12
                            With .Borders(i)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .ColorIndex = xlAutomatic
                            End With
                        Next i
                End With
                
                '-style
            'first Table
            With xlSheet
                .Rows("3:3").EntireRow.AutoFit
                .Rows(curCol2 - 1 & ":" & curCol2 - 1).EntireRow.AutoFit
            End With
            With .Range("A3:E3")
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
            'Second Table
            With .Range("A" & curCol2 - 1 & ":K" & curCol2 - 1)
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
            .Rows(curCol2 - 1 & ":" & curCol2 - 1).RowHeight = 27.75
          
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
                With .Range("A1:H1")
                    .Merge
                    .Range("A1").Formula = "MACHINE BREAKDOWN WORK ORDER (" & Format(DTPicker1, "mmmm dd,yyyy") & " - " & Format(DTPicker2, "mmmm dd,yyyy") & ")"
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 15
                    .Font.Bold = True
                End With
            End With
                        
'        xlBook.Save
'        xlBook.Close
'        xlApp.Quit
        exportExcel = True
        xlApp.Visible = True
        flxDetail.Visible = True
        flxSummaryDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
'        MsgBox "Report Succesfully imported!", vbInformation, "WODataExtractionTool"
        
        '- Open extracted report report --
        'Shell "explorer " & strNewFile, vbMaximizedFocus
        '-
        
'        Set xlSheet = Nothing
'        Set xlBook = Nothing
        Set xlApp = Nothing
        
        
       
        Exit Function
         
ErrSave:
    exportExcel = False
        MsgBox Err.Number & " " & Err.Description, vbCritical, "WODataExtractionTool"
        flxDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
End Function

Private Sub cmdSearch_Click()
    
    Call LoadFlexBreakdownView
   
End Sub

Private Sub Form_Load()
    Call subFormatGrid(flxDetail, "breakdown")
    Call subFormatGrid(flxSummaryDetail, "summary")
    Call LoadDataToCombo(cboType, "Types")
    
End Sub

Private Sub LoadFlexBreakdownView()
    Dim rsFlex1 As ADODB.Recordset
    Dim rsFlex2 As ADODB.Recordset
    Dim lngLoop, i As Long
    Dim lngrow As Long
    Dim lngNo As Long
    Dim d1, d2 As Date
    
    d1 = Format(DTPicker1.Value, "YYYY/MM/DD")
    d2 = Format(DTPicker2.Value, "YYYY/MM/DD")
    
    Set rsFlex1 = cls_GetDetails.pfLoadBreakdown1(d1, d2, cboType.Column(0))
    Set rsFlex2 = cls_GetDetails.pfLoadBreakdown2(d1, d2, cboType.Column(0))
    
    If rsFlex1.EOF Or rsFlex2.EOF Then
        MsgBox "No Record found!"
        Call subFormatGrid(flxDetail, "breakdown")
        Call subFormatGrid(flxSummaryDetail, "summary")
        Set rsFlex1 = Nothing
        Set rsFlex2 = Nothing
        Exit Sub
    Else
        flxDetail.Visible = False
        flxSummaryDetail.Visible = False
        FM_Main.Enabled = False
        lblMessage1.Caption = "Please Wait. Loading Data.."
        lblMessage2.Caption = "Please Wait. Loading Data.."
        FM_Main.StatusBar1.Panels(3).Text = "Please Wait. Loading Data.."
        With flxDetail
            rsFlex1.MoveFirst
            lngrow = 1
            Do While Not rsFlex1.EOF
                        .Rows = lngrow + 1
                        .TextMatrix(lngrow, 0) = Format(pfvarNoValue(rsFlex1.Fields("ReceivedDate").Value), "mmmm-dd-yyyy")
                        .TextMatrix(lngrow, 1) = pfvarNoValue(rsFlex1.Fields("DepartmentName").Value)
                        .TextMatrix(lngrow, 2) = pfvarNoValue(rsFlex1.Fields("SectionName").Value)
                        .TextMatrix(lngrow, 3) = pfvarNoValue(rsFlex1.Fields("Received").Value)
                        .TextMatrix(lngrow, 4) = pfvarNoValue(rsFlex1.Fields("Finished").Value)
                lngrow = lngrow + 1
                rsFlex1.MoveNext
                Loop
        End With
        With flxSummaryDetail
            rsFlex2.MoveFirst
            lngrow = 1
            Do While Not rsFlex2.EOF
                        .Rows = lngrow + 1
                        .TextMatrix(lngrow, 0) = pfvarNoValue(rsFlex2.Fields("DepartmentName").Value)
                        .TextMatrix(lngrow, 1) = pfvarNoValue(rsFlex2.Fields("SectionName").Value)
                        .TextMatrix(lngrow, 2) = pfvarNoValue(rsFlex2.Fields("RECEIVED").Value)
                        .TextMatrix(lngrow, 3) = pfvarNoValue(rsFlex2.Fields("FINISHED").Value)
                        .TextMatrix(lngrow, 4) = pfvarNoValue(rsFlex2.Fields("FINISHEDWOFROMPENDINGWO").Value)
                        .TextMatrix(lngrow, 5) = pfvarNoValue(rsFlex2.Fields("FINISHEDONTHESUCCEEDINGMONTH").Value)
                        .TextMatrix(lngrow, 6) = pfvarNoValue(rsFlex2.Fields("CANCELLED").Value)
                        .TextMatrix(lngrow, 7) = pfvarNoValue(rsFlex2.Fields("TURNOVER").Value)
                        .TextMatrix(lngrow, 8) = pfvarNoValue(rsFlex2.Fields("WAITINGPARTS").Value)
                        .TextMatrix(lngrow, 9) = pfvarNoValue(rsFlex2.Fields("FORSCHEDULE").Value)
                        .TextMatrix(lngrow, 10) = pfvarNoValue(rsFlex2.Fields("FORCONFIRMATION").Value)
                lngrow = lngrow + 1
                rsFlex2.MoveNext
                Loop
        End With
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
         flxSummaryDetail.Visible = True
        flxDetail.Visible = True
   End If
LDExit:
    
    Set rsFlex1 = Nothing
    Set rsFlex2 = Nothing
    Exit Sub
LDErr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo LDExit
End Sub

