VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1275
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   9375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSplash.frx":1042
   MousePointer    =   99  'Custom
   ScaleHeight     =   1275
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   9195
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Released date: 11, January 2018"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6660
         TabIndex        =   2
         Top             =   135
         Width           =   2400
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1170
         Picture         =   "frmSplash.frx":190C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   7815
      End
      Begin VB.Image imgLogo 
         Height          =   660
         Left            =   180
         Picture         =   "frmSplash.frx":B4C9
         Stretch         =   -1  'True
         Top             =   315
         Width           =   750
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © January 2018  HRD-SMD-SD, All Rights Reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   495
         TabIndex        =   1
         Top             =   795
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
