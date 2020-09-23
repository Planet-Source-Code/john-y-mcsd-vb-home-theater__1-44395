VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB Home Theater"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   1620
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":06E3
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   6
      Top             =   480
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2001 Fredisoft Corp."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   2490
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fredisoft@bol.com.br"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1200
      MouseIcon       =   "frmAbout.frx":0FAD
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1200
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program was developed by Frederico Machado."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   3750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version x.xx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   240
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB Home Theater"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   435
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2835
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Label2 = "Version " & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub Label4_Click()
  Shell "start mailto:fredisoft@bol.com.br"
End Sub
