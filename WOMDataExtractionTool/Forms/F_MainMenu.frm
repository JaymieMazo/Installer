VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form F_MachineStatusView 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Report View"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14865
   Icon            =   "F_MainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   14865
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "DAILY REPORT FOR THE STATUS OF MACHINE"
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
      Height          =   4830
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   14865
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMachineDetail 
         Height          =   3885
         Left            =   45
         TabIndex        =   1
         Top             =   315
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6853
         _Version        =   393216
         Cols            =   22
         BackColorFixed  =   11627568
         ForeColorFixed  =   16777215
         ForeColorSel    =   -2147483635
         GridColorFixed  =   8421504
         AllowUserResizing=   1
         BorderStyle     =   0
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
         _Band(0).Cols   =   22
      End
      Begin MSForms.CommandButton cmdExcel 
         Height          =   390
         Left            =   180
         TabIndex        =   3
         Top             =   4365
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
         Left            =   45
         TabIndex        =   2
         Top             =   2025
         Width           =   14820
      End
   End
   Begin MSComDlg.CommonDialog cdExcel 
      Left            =   14670
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "F_MachineStatusView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSM As Object
Private Function exportExcel() As Boolean
    Dim xlApp       As Excel.Application
    Dim xlBook      As Excel.Workbook
    Dim xlSheet     As Excel.Worksheet
    
    Dim strNewFile As String
    Dim intloop As Long
    Dim curCol As Long
    Dim i As Long
    
    On Error GoTo ErrExcel
        
        If flxMachineDetail.TextMatrix(1, 0) = "" Then Exit Function
        'strNewFile = CommonDialogSave(cdExcel, "MachineStatusReport")
        If blnCancel = True Then
            'Exit Sub
        End If
        exportExcel = False
            FM_Main.Enabled = False
            flxMachineDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        
            'Call FileCopy(App.Path & "\Reports\Machine Status Report.xls", strNewFile)
            
            Set xlApp = CreateObject("Excel.Application")
             Set xlBook = xlApp.Workbooks.Add
           ' Set xlBook = xlApp.Workbooks.Open(strNewFile)
            Set xlSheet = xlBook.Sheets("Sheet1")
            
            With xlSheet
                .Range("T1").Formula = "DEPARTMENT:"
                .Range("A1:S2").Merge
                With .Range("A1")
                    .Font.Name = "Arial Narrow"
                    .Font.Size = 14
                    .Formula = "DAILY REPORT FOR THE STATUS OF MACHINE"
                    .HorizontalAlignment = xlCenter
                    .Font.Bold = True
                End With
               
                .Range("T2").Formula = "DATE:"
                    .Range("A" & 4).Formula = "NO."
                    .Range("B" & 4).Formula = "DATE OF W.O. "
                    .Range("C" & 4).Formula = "WORK CATEGORY"
                    .Range("D" & 4).Formula = "SECTION"
                    .Range("E" & 4).Formula = "LINE"
                    .Range("F" & 4).Formula = "PERSON IN-CHARGED / TL"
                    .Range("G" & 4).Formula = "W.O. #"
                    .Range("H" & 4).Formula = "EQPT. CONTROL NO."
                    .Range("I" & 4).Formula = "MACHINE NAME"
                    .Range("J" & 4).Formula = "TYPE OF REQUEST"
                    .Range("K" & 4).Formula = "SPECIFIC TROUBLE"
                    .Range("L" & 4).Formula = "STATUS"
                    .Range("M" & 4).Formula = "PARTS NEEDED"
                    .Range("N" & 4).Formula = "DATE OF MAKING MRS / MACHINE PARTS FOR REQUEST"
                    .Range("O" & 4).Formula = "DATE OF MAKING PRS"
                    .Range("P" & 4).Formula = "PRS #"
                    .Range("Q" & 4).Formula = "PO #"
                    .Range("R" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PRS)"
                    .Range("S" & 4).Formula = "EXPECTED DATE DELIVERY (FROM PURCHASING)"
                    .Range("T" & 4).Formula = "DATE OF ACTUAL RECEIVING OF ITEM"
                    .Range("U" & 4).Formula = "DATE FINISHED"
                    .Range("V" & 4).Formula = "REMARKS"
                    
                    
            End With
            With xlSheet.Range("A4:V4")
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
            For intloop = 1 To flxMachineDetail.Rows - 1
                
                lblMessage.Caption = "Please Wait. Exporting Data to Excel.. (" & intloop & " out of " & flxMachineDetail.Rows - 1 & " row/s)"
                Me.Refresh
                With xlSheet
                    .Range("A" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 0)
                    .Range("B" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 1)
                    .Range("C" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 2)
                    .Range("D" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 3)
                    .Range("E" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 4)
                    .Range("F" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 5)
                    .Range("G" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 6)
                    .Range("H" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 7)
                    .Range("I" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 8)
                    .Range("J" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 9)
                    .Range("K" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 10)
                    .Range("L" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 11)
                    .Range("M" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 12)
                    .Range("N" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 13)
                    .Range("O" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 14)
                    .Range("P" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 15)
                    .Range("Q" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 16)
                    .Range("R" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 17)
                    .Range("S" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 18)
                    .Range("T" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 19)
                    .Range("U" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 20)
                    .Range("V" & RowCtr).Formula = flxMachineDetail.TextMatrix(intloop, 21)
                    '-Insert row
                    If flxMachineDetail.Rows - 1 <> 1 Then
                        .Rows(RowCtr + 1 & ":" & RowCtr + 1).Insert Shift:=xlUp
                        RowCtr = RowCtr + 1
                    End If
                    '-
                End With
            Next intloop
            
            '--- Excel Format -----------
            With xlSheet
                lblMessage.Caption = "Formatting Spreadsheet.."
                
                .Rows("4:" & RowCtr - 1).EntireRow.AutoFit
                With .Range("A4:V" & RowCtr - 1)
                        .HorizontalAlignment = xlCenter
                        '.VerticalAlignment = xlCenter
                        .WrapText = True
                        '-Borders
                        For i = 7 To 12
                            With .Borders(i)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .ColorIndex = xlAutomatic
                            End With
                        Next i
                End With
                '-Date-
                    With .Range("U2")
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
               
            End With
                
'        xlBook.Save
'        xlBook.Close
'        xlApp.Quit
        exportExcel = True
        xlApp.Visible = True
        flxMachineDetail.Visible = True
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
        flxMachineDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        exportExcel = False
End Function



Private Sub cmdExcel_Click()
        If flxMachineDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If exportExcel = True Then
            MsgBox "Report Succesfully saved to Excel!", vbInformation, "WODataExtractionTool"
        ElseIf exportLibre = True Then
            MsgBox "Report Succesfully saved to LibreOffice!", vbInformation, "WODataExtractionTool"
        Else
             MsgBox " An error occured. Data not successfully exported ", vbCritical, " System Error "
        End If
        
    
       
        Exit Sub
         
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
            flxMachineDetail.Visible = False
            lblMessage.Caption = "Please Wait. Exporting Data to Excel.."
            FM_Main.StatusBar1.Panels(3).Text = "Exporting Data to Excel.."
        oSheet.getCellByPosition(0, oRow).String = "DAILY REPORT FOR THE STATUS OF MACHINE"
         
                
        For oCol = 0 To flxMachineDetail.Cols - 1
           With oSheet
                With CellStyle
                    .CharWeight = 300
                    .CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .CellBackColor = RGB(255, 184, 0)
                End With
                .getCellByPosition(oCol, 1).CellStyle = "MyStyle"
                .getCellByPosition(oCol, 1).String = flxMachineDetail.TextMatrix(0, oCol)
                .getCellByPosition(oCol, 1).HoriJustify = 2
           End With
            Set oBorder = oSheet.getCellRangeByPosition(0, 1, flxMachineDetail.Cols - 1, 1).TableBorder
            oBorder.LeftLine = basicBorder
            oBorder.Topline = basicBorder
            oBorder.RightLine = basicBorder
            oBorder.BottomLine = basicBorder
            oBorder.VerticalLine = basicBorder
            oSheet.getCellRangeByPosition(0, 1, flxMachineDetail.Cols - 1, 1).TableBorder = oBorder
        Next oCol
    
        '-Content--
        For oRow = 0 To flxMachineDetail.Rows - 1
            DoEvents
            
            lblMessage.Caption = "Exporting Data : " & oRow & " rows out of " & flxMachineDetail.Rows - 1 & " rows"
            For oCol = 0 To flxMachineDetail.Cols - 1
               With oSheet
                    .getCellByPosition(oCol, oRow + 1).CharFontName = "‚l‚r ‚oƒSƒVƒbƒN"
                    .getCellByPosition(oCol, oRow + 1).HoriJustify = 2
                    flxMachineDetail.Row = oRow
                    flxMachineDetail.Col = oCol
                    .getCellByPosition(oCol, oRow + 1).String = flxMachineDetail.TextMatrix(oRow, oCol)
                    If flxMachineDetail.CellBackColor = &H8080FF Then
                        .getCellByPosition(oCol, oRow + 1).CellBackColor = RGB(128, 128, 255)
                    End If
               End With
            Next oCol
                Set oBorder = oSheet.getCellRangeByPosition(0, oRow + 1, flxMachineDetail.Cols - 1, oRow + 1).TableBorder
                oBorder.LeftLine = basicBorder
                oBorder.Topline = basicBorder
                oBorder.RightLine = basicBorder
                oBorder.BottomLine = basicBorder
                oBorder.VerticalLine = basicBorder
                oSheet.getCellRangeByPosition(0, oRow + 1, flxMachineDetail.Cols - 1, oRow + 1).TableBorder = oBorder
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
        flxMachineDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        Exit Function
ErrLibre:
        exportLibre = False
        flxMachineDetail.Visible = True
        FM_Main.StatusBar1.Panels(3).Text = ""
        FM_Main.Enabled = True
        
End Function

Private Sub LoadFlexDailyReportStatus()
    Dim rsFlex As ADODB.Recordset
    Dim lngLoop, i As Long
    Dim lngrow As Long
    Dim lngNo As Long
    
    Set rsFlex = cls_GetDetails.pfLoadData("Machine")
    
    If rsFlex.EOF Then
        MsgBox "No Record found!"
        Set rsFlex = Nothing
        Exit Sub
    Else
        With flxMachineDetail
            rsFlex.MoveFirst
            lngrow = 1
            lngNo = 1
            Do While Not rsFlex.EOF
                        .Rows = lngrow + 1
                        .CellAlignment = flexAlignCenterCenter
                        .TextMatrix(lngrow, 0) = lngNo
                        .TextMatrix(lngrow, 1) = pfvarNoValue(rsFlex.Fields("DateOfWorkOrder").Value)
                        .TextMatrix(lngrow, 2) = pfvarNoValue(rsFlex.Fields("WorkCategory").Value)
                        .TextMatrix(lngrow, 3) = pfvarNoValue(rsFlex.Fields("Section").Value)
                        .TextMatrix(lngrow, 4) = pfvarNoValue(rsFlex.Fields("Line").Value)
                        .TextMatrix(lngrow, 5) = pfvarNoValue(rsFlex.Fields("PersonInCharged").Value)
                        .TextMatrix(lngrow, 6) = pfvarNoValue(rsFlex.Fields("WorkOrderControlNo").Value)
                        .TextMatrix(lngrow, 7) = pfvarNoValue(rsFlex.Fields("MachineItemNo").Value)
                        .TextMatrix(lngrow, 8) = pfvarNoValue(rsFlex.Fields("MachineName").Value)
                        .TextMatrix(lngrow, 9) = pfvarNoValue(rsFlex.Fields("TypeOfRequest").Value)
                        .TextMatrix(lngrow, 10) = pfvarNoValue(rsFlex.Fields("ProblemFound").Value)
                        .TextMatrix(lngrow, 11) = pfvarNoValue(rsFlex.Fields("Status").Value)
                        .TextMatrix(lngrow, 12) = pfvarNoValue(rsFlex.Fields("Description").Value)
                        .TextMatrix(lngrow, 13) = pfvarNoValue(rsFlex.Fields("RequestDate").Value)
                        .TextMatrix(lngrow, 14) = pfvarNoValue(rsFlex.Fields("PrsDate").Value)
                        .TextMatrix(lngrow, 15) = pfvarNoValue(rsFlex.Fields("PrsNo").Value)
                        .TextMatrix(lngrow, 16) = pfvarNoValue(rsFlex.Fields("PoNo").Value)
                        .TextMatrix(lngrow, 17) = pfvarNoValue(rsFlex.Fields("PrsExpectedDelivery").Value)
                        .TextMatrix(lngrow, 18) = pfvarNoValue(rsFlex.Fields("PoExpectedDelivery").Value)
                        .TextMatrix(lngrow, 19) = pfvarNoValue(rsFlex.Fields("ActualReceived").Value)
                        .TextMatrix(lngrow, 20) = pfvarNoValue(rsFlex.Fields("FinishedDate").Value)
                        .TextMatrix(lngrow, 21) = pfvarNoValue(rsFlex.Fields("Remarks").Value)
                lngrow = lngrow + 1
                lngNo = lngNo + 1
                rsFlex.MoveNext
                Loop
        End With
        
   End If
LDExit:
    
    Set rsFlex = Nothing
    Exit Sub
LDErr:
    MsgBox Err.Description, vbCritical, "Work Order Data Extraction Tool"
    GoTo LDExit
End Sub


Private Sub Form_Load()
   
    Call subFormatGrid(flxMachineDetail, "status")
    Call LoadFlexDailyReportStatus
    Frame1.Visible = True
    
End Sub

