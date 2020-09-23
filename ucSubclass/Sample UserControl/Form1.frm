VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   2790
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin prjUserControl.ucSubclass ucSub 
      Height          =   1350
      Left            =   120
      Top             =   1305
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2381
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Leave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1485
      TabIndex        =   2
      Top             =   2805
      Width           =   1170
   End
   Begin VB.Label lblEnter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2805
      Width           =   1170
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  lbl.Caption = "The UserControl uses subclassing to detect mouse entry and exit. Also, the control subclasses the parent form's size and move messages."
End Sub

Private Sub ucSub_MouseEnter()
  Me.lblEnter.BackColor = RGB(0, 255, 0)
  Me.lblExit.BackColor = &H8000000F
  sb.SimpleText = "Mouse enter"
End Sub

Private Sub ucSub_MouseLeave()
  Me.lblEnter.BackColor = &H8000000F
  Me.lblExit.BackColor = RGB(0, 255, 0)
  sb.SimpleText = "Mouse leave"
End Sub

Private Sub ucSub_Status(ByVal sStatus As String)
  sb.SimpleText = sStatus
End Sub
