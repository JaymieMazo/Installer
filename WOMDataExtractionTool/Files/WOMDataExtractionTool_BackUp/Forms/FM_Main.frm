VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FM_Main 
   BackColor       =   &H00404000&
   Caption         =   "Work Order Maintenance Data Extraction Tool"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   14040
   Icon            =   "FM_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "FM_Main.frx":1042
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4815
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   450
            MinWidth        =   441
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "2018/04/10"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14121
            MinWidth        =   3352
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
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStatusView 
         Caption         =   "Daily Report Status View"
      End
      Begin VB.Menu mnuBreakdown 
         Caption         =   "Breakdown Work Order View"
      End
      Begin VB.Menu mnuWorkorderHistory 
         Caption         =   "Workorder And Costing History"
      End
      Begin VB.Menu mnuMachineControlStatus 
         Caption         =   "Machine Control Status"
      End
      Begin VB.Menu mnuMaintenance 
         Caption         =   "Maintenance"
      End
      Begin VB.Menu mnuAccomplishments 
         Caption         =   "Accomplishments"
      End
   End
   Begin VB.Menu mnuMasterlist 
      Caption         =   "&Masterlist"
      Begin VB.Menu mnuMasterEmployee 
         Caption         =   "Employees"
      End
      Begin VB.Menu mnuMasterMachineItems 
         Caption         =   "Machine Items"
      End
   End
   Begin VB.Menu mnuMonitoring 
      Caption         =   "M&onitoring"
      Begin VB.Menu mnuPRS 
         Caption         =   "PRS"
      End
   End
End
Attribute VB_Name = "FM_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MDIForm_Load()
    'If App.PrevInstance = True Then MsgBox "The system is already open!", vbCritical: Exit Sub
'    Call Connect
        
End Sub



Private Sub mnuAccomplishments_Click()
        FM_Main.MousePointer = vbCustom
    frmAccomplishment.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuBreakdown_Click()
    FM_Main.MousePointer = vbCustom
    frmBreakDown.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuMaintenance_Click()
    FM_Main.MousePointer = vbCustom
    frmMaintenance.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuMasterEmployee_Click()
    FM_Main.MousePointer = vbCustom
    frmEmployeeMasterList.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuPRS_Click()
    FM_Main.MousePointer = vbCustom
    frmMonitoring.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuExit_Click()
    If MsgBox("Do really want to exit the system?", vbYesNo, "Exit") = vbYes Then
        End
    End If
End Sub


Private Sub mnuMachineControlStatus_Click()
    FM_Main.MousePointer = vbCustom
    frmMachineControlStatus.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuMasterMachineItems_Click()
    FM_Main.MousePointer = vbCustom
    frmMachineItemMasterList.Show
    FM_Main.MousePointer = vbDefault
End Sub


Private Sub mnuStatusView_Click()
    FM_Main.MousePointer = vbCustom
    frmStatus.Show
    FM_Main.MousePointer = vbDefault
End Sub

Private Sub mnuWorkorderHistory_Click()
    FM_Main.MousePointer = vbCustom
    frmWorkorderAndCostingHistory.Show
    FM_Main.MousePointer = vbDefault
End Sub
