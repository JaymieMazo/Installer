VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm F_PrintMainMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Print Menu"
   ClientHeight    =   4875
   ClientLeft      =   165
   ClientTop       =   1020
   ClientWidth     =   12765
   Icon            =   "PrintMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   4530
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   635
            MinWidth        =   441
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2006/06/20"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11695
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS UI Gothic"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuWHRecord 
      Caption         =   "Ware&house Record"
   End
   Begin VB.Menu mnuInvRecord 
      Caption         =   "&Inventory Record"
   End
   Begin VB.Menu mnuOnhand 
      Caption         =   "&OnHand Summary"
   End
   Begin VB.Menu mnuWaiting 
      Caption         =   "Waiting &List"
   End
   Begin VB.Menu mnuInvoiceSearch 
      Caption         =   "In&voice Search"
   End
   Begin VB.Menu mnuPOData 
      Caption         =   "PO &Data"
   End
   Begin VB.Menu mnuPOSearch 
      Caption         =   "PO &Search"
   End
   Begin VB.Menu mnuInvCheck 
      Caption         =   "Inventory &Check"
   End
   Begin VB.Menu mnuMaterialList 
      Caption         =   "&Material List"
   End
   Begin VB.Menu mnuSupplierItemList 
      Caption         =   "Su&pplier Item List"
   End
   Begin VB.Menu mnuMatConsumption 
      Caption         =   "&Mat'l Turn-Over"
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "F_PrintMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
    Set clsPrintMenu = Nothing
    Set clsDB = Nothing
    End
End Sub

Private Sub mnuInvCheck_Click()
    F_InvCheck.Show
End Sub

Private Sub mnuInvoiceSearch_Click()
    F_InvoiceSearch.Show
End Sub

Private Sub mnuInvRecord_Click()
    F_InvRecord.Show
End Sub

Private Sub mnuMatConsumption_Click()
    F_MaterialConsumption.Show
End Sub

Private Sub mnuMaterialList_Click()
    F_Material_List.Show
End Sub

Private Sub mnuOnhand_Click()
     F_OnhandSummary.Show
End Sub

Private Sub mnuOnhandSummary_Click()
    F_OnhandSummary.Show
End Sub

Private Sub mnuPOData_Click()
    F_POData.Show
End Sub

Private Sub mnuPOSearch_Click()
    F_POSearch.Show
End Sub

Private Sub mnuSupplierItemList_Click()
     F_SupplierItemList.Show
End Sub

Private Sub mnuWaiting_Click()
    F_WaitingList.Show
End Sub

Private Sub mnuWHRecord_Click()
    F_WH_Record.Show
End Sub

