VERSION 5.00
Begin VB.Form frmRate 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rate"
   ClientHeight    =   1110
   ClientLeft      =   6915
   ClientTop       =   6060
   ClientWidth     =   2055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtRate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "100"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   165
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate (0-200):"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSet_Click()
  If Val(txtRate) < 0 Then txtRate = "100"
  If txtRate = "" Then txtRate = "100"
  If Val(txtRate) > 200 Then txtRate = "100"
  SetRate AliasName, Val(txtRate)
  frmMain.SetFocus
End Sub

Private Sub Form_Load()
  XLeft = GetSetting(App.Title, "Config", "RateX", Me.left)
  YTop = GetSetting(App.Title, "Config", "RateY", Me.top)
  Me.Move XLeft, YTop
  txtRate = GetRate(AliasName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting App.Title, "Config", "RateX", Str$(Me.left)
  SaveSetting App.Title, "Config", "RateY", Str$(Me.top)
End Sub

Private Sub txtRate_GotFocus()
  txtRate.SelStart = 0
  txtRate.SelLength = Len(txtRate)
End Sub
