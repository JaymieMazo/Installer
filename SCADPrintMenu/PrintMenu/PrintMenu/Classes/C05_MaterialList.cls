VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C05_MaterialList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQLMaterial As String
'Updated December 2, 2005
'by Ariel Balisi
Public Function GetMaterialList(ByRef hflxGrid As Object, ByVal strLocation As String _
                           , ByVal strItemId As String, ByVal strDescription As String _
                           , ByVal strDivision As String, ByVal bytStatus As Byte) As Long
     Dim adoRsMaterial   As Object
     Dim strWhere As String

     strSQLMaterial = " SELECT * " _
                     & " FROM MaterialListView"
    
     strWhere = ""
     If strDivision <> "All" Then _
        strWhere = strWhere & " Division = " & pfstrQt(strDivision) & " AND"
            
     If strLocation <> "ALL" Then _
           strWhere = strWhere & " Location = " & pfstrQt(strLocation) & " AND"
     
     If Not strItemId = vbNullString Then _
         strWhere = strWhere & " ItemId like " & pfstrQt(strItemId & "%") & " AND"
         
     If Not strDescription = vbNullString Then _
         strWhere = strWhere & " Description like " & pfstrQt(strDescription & "%") & " AND"
     
     Select Case bytStatus
         Case 1: '--active
              strWhere = strWhere & " Disuse=0 AND"
         Case 2: '--disuse
              strWhere = strWhere & " Disuse=1 AND"
     End Select
     
     
     If strWhere <> "" Then _
         strWhere = " WHERE " & Left(strWhere, Len(strWhere) - 3)
            
     strSQLMaterial = strSQLMaterial & strWhere
    
     strSQLMaterial = strSQLMaterial & " Order by 1"
     Set adoRsMaterial = clsDB.GetRecordSet(strSQLMaterial)
     GetMaterialList = CLng(adoRsMaterial.RecordCount)
     Set hflxGrid.DataSource = adoRsMaterial
     
End Function
'------- Export Material List to Excel
Public Sub ExportMaterialToExcel(ByRef hflxGrid As Object)
     Dim bytCol As Byte, lngRow As Long
     
     Call clsPrintMenu.Utility.OpenExcel
     With hflxGrid
          For lngRow = 0 To .Rows - 1
               For bytCol = 0 To .Cols - 1
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1).NumberFormat = "@"
                    clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1) = .TextMatrix(lngRow, bytCol)
                     If lngRow = 0 Then clsPrintMenu.Utility.ExcelWkSheet.Cells(lngRow + 1, bytCol + 1).Font.Bold = True
               Next
          Next
          Call clsPrintMenu.Utility.SetCellColor(1, 1, 1, .Cols)
          clsPrintMenu.Utility.ExcelWkSheet.Columns.AutoFit
          clsPrintMenu.Utility.ExcelApp.Visible = True
          Call clsPrintMenu.Utility.CloseExcel
     End With
     Call clsPrintMenu.Utility.CloseExcel
End Sub
'--- load item locations
Public Sub LoadLocation(ByVal cboSource As Object)
     Dim rsAdoLocation   As Object
     
     strSQLMaterial = "select Location from Locations"
     Set rsAdoLocation = clsDB.GetRecordSet(strSQLMaterial)
     cboSource.Clear
     With rsAdoLocation
          cboSource.AddItem "ALL"
          Do While Not .EOF
               cboSource.AddItem .Fields("Location").Value
               .MoveNext
          Loop
          cboSource.ListIndex = 0
     End With
 End Sub
