VERSION 5.00
Begin VB.Form F_MachineViewMenu 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Machine View Menu"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBreakdown 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Breakdown Work Order View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   495
      Width           =   2895
   End
   Begin VB.CommandButton cmdStatus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Daily Report Status View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   2895
   End
End
Attribute VB_Name = "F_MachineViewMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
