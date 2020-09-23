VERSION 5.00
Begin VB.Form frmPList 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Playlist"
   ClientHeight    =   3165
   ClientLeft      =   5850
   ClientTop       =   2160
   ClientWidth     =   5550
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
   ScaleHeight     =   3165
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstPath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      IntegralHeight  =   0   'False
      ItemData        =   "frmPList.frx":0000
      Left            =   1320
      List            =   "frmPList.frx":0002
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   2460
      IntegralHeight  =   0   'False
      ItemData        =   "frmPList.frx":0004
      Left            =   130
      List            =   "frmPList.frx":0006
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   140
      Width           =   5280
   End
   Begin VB.Label lblSavePList 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   2205
      TabIndex        =   9
      Top             =   2775
      Width           =   240
   End
   Begin VB.Label lblOpenPList 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1845
      TabIndex        =   8
      Top             =   2775
      Width           =   240
   End
   Begin VB.Label lblAddFile 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3735
      TabIndex        =   7
      Top             =   2805
      Width           =   375
   End
   Begin VB.Label lblListDOWN 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   2910
      Width           =   255
   End
   Begin VB.Label lblListUP 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   720
      TabIndex        =   5
      Top             =   2715
      Width           =   255
   End
   Begin VB.Label lblNewList 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   1530
      TabIndex        =   4
      Top             =   2775
      Width           =   225
   End
   Begin VB.Label lblDelFile 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   990
      TabIndex        =   3
      Top             =   2775
      Width           =   240
   End
   Begin VB.Label lblAddDir 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   2790
      Width           =   375
   End
   Begin VB.Image imgPListBar 
      Height          =   300
      Left            =   720
      Picture         =   "frmPList.frx":0008
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   5310
   End
End
Attribute VB_Name = "frmPList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MoviePath As String
Dim strLOpenPath As String
Dim strLSavePath As String

Private Sub Form_Load()
  XLeft = GetSetting(App.Title, "Config", "PListX", Me.left)
  YTop = GetSetting(App.Title, "Config", "PListY", Me.top)
  Me.Move XLeft, YTop
  strLOpenPath = GetSetting(App.Title, "Config", "LOpenPath", App.Path)
  strLSavePath = GetSetting(App.Title, "Config", "LSavePath", App.Path)
  Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If bHTEnd = False Then
    Cancel = True
  End If
  Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If lstPath.ListCount > 0 Then
    file = LTrim$(App.Path & "\plist.htp")
    Open file For Output As #1
      For i = 0 To lstPath.ListCount - 1
        Print #1, lstPath.List(i)
      Next
    Close #1
  ElseIf lstPath.ListCount = 0 Then
    On Error Resume Next
    Kill App.Path & "\plist.htp"
  End If
  SaveSetting App.Title, "Config", "PListX", Str$(Me.left)
  SaveSetting App.Title, "Config", "PListY", Str$(Me.top)
  frmMain.SetFocus
End Sub

Private Sub lblAddDir_Click()
  MoviePath = BrowseForDirectory
  temp = Dir$(MoviePath & "\*.mpg")
  While Len(temp) > 0
    lstPath.AddItem MoviePath & "\" & temp
    i = InStr(temp, ".")
    If i > 0 Then temp = left$(temp, i - 1)
    lstNames.AddItem temp
    temp = Dir$
  Wend
  temp = Dir$(MoviePath & "\*.mpeg")
  While Len(temp) > 0
    lstPath.AddItem MoviePath & "\" & temp
    i = InStr(temp, ".")
    If i > 0 Then temp = left$(temp, i - 1)
    lstNames.AddItem temp
    temp = Dir$
  Wend
  temp = Dir$(MoviePath & "\*.avi")
  While Len(temp) > 0
    lstPath.AddItem MoviePath & "\" & temp
    i = InStr(temp, ".")
    If i > 0 Then temp = left$(temp, i - 1)
    lstNames.AddItem temp
    temp = Dir$
  Wend
End Sub

Private Sub lblAddFile_Click()
  Dim CD As New clsDialog
  temp = CD.OpenDialog(frmPList, "Video Files (*.MPG;*.MPEG;*.AVI) |*.mpg;*.mpeg;*.avi|", "Add File", strLOpenPath)
  If temp = "" Then Exit Sub
  file = LTrim$(temp)
  lstPath.AddItem file
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P2$ = P1$: P$ = left$(file, j): X$ = Mid$(file, j + 1)
  If left$(file, 1) = "\" Then P2$ = left$(P1$, 2)
  b$ = X$: i = InStr(X$, ".")
  If i > 0 Then b$ = left$(X$, i - 1)
  lstNames.AddItem b$
  SaveSetting App.Title, "Config", "LOpenPath", P$
  strLOpenPath = P$
End Sub

Private Sub lblDelFile_Click()
  If lstNames.ListCount = 0 Then Exit Sub
  If lstNames.SelCount = 0 Then Exit Sub
  If lstNames.ListIndex = lstNames.ListCount - 1 Then
    temp = lstNames.ListIndex - 1
  Else
    temp = lstNames.ListIndex
  End If
  lstPath.RemoveItem lstNames.ListIndex
  lstNames.RemoveItem lstNames.ListIndex
  lstNames.ListIndex = temp
End Sub

Private Sub lblListDOWN_Click()
  ListMove 1
End Sub

Private Sub lblListUP_Click()
  ListMove -1
End Sub

Private Sub lblNewList_Click()
  lstNames.Clear
  lstPath.Clear
End Sub

Private Sub lblOpenPList_Click()
  Dim CD As New clsDialog
  Dim temp As String
  temp = CD.OpenDialog(frmPList, "Home Theater Playlist (*.HTP) |*.htp|", "Open Playlist", strLOpenPath)
  If temp = "" Then Exit Sub
  OpenPList temp
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P$ = left$(file, j)
  SaveSetting App.Title, "Config", "LOpenPath", P$
  strLOpenPath = P$
End Sub

Private Sub lblSavePList_Click()
  If lstPath.ListCount = 0 Then Exit Sub
  Dim CD As New clsDialog
  temp = CD.SaveDialog(frmPList, "Home Theater Playlist (*.HTP) |*.htp|", "Save Playlist", strLSavePath)
  If temp = "" Then Exit Sub
  If LCase(Right$(temp, 4)) <> ".htp" Then temp = temp & ".htp"
  file = LTrim$(temp)
  Open file For Output As #1
    For i = 0 To lstPath.ListCount - 1
      Print #1, lstPath.List(i)
    Next
  Close #1
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P$ = left$(file, j)
  SaveSetting App.Title, "Config", "LSavePath", P$
  strLSavePath = P$
End Sub

Public Sub lstNames_DblClick()
  If lstNames.ListCount = 0 Then Exit Sub
  frmMain.StopVideo: frmMain.CloseVideo
  frmMain.OpenVideo lstPath.List(lstNames.ListIndex)
  DoEvents
  frmMain.PlayVideo
  frmMain.SetFocus
End Sub

Public Sub PlayNext()
  frmMain.StopVideo
  frmMain.CloseVideo
  If lstNames.ListIndex = lstNames.ListCount - 1 Then Exit Sub
  frmMain.OpenVideo lstPath.List(lstNames.ListIndex + 1)
  frmMain.PlayVideo
  lstNames.ListIndex = lstNames.ListIndex + 1
End Sub

Public Sub ListMove(D)
  N = lstNames.ListIndex
  If (N + D) > 0 And (N + D) < lstNames.ListCount Then
    T1$ = lstNames.List(N): T2$ = lstPath.List(N)
    lstNames.List(N) = lstNames.List(N + D)
    lstPath.List(N) = lstPath.List(N + D)
    lstNames.List(N + D) = T1$
    lstPath.List(N + D) = T2$
    lstNames.ListIndex = N + D
  End If
End Sub

Public Sub OpenPList(file As String)
  Open file For Input As 1
    lstNames.Clear: lstPath.Clear
    GoSub LoadM3U
  Close 1
  Exit Sub
LoadM3U:
    While Not EOF(1)
        Line Input #1, AA$: a$ = LTrim$(AA$)
        GoSub AddIt
    Wend
    Return

AddIt:
    GoSub SplitPF: N = N + 1
    lstNames.AddItem b$
    lstPath.AddItem P2$ + a$
    Return

SplitPF:
    For j = Len(a$) To 1 Step -1
        If Mid$(a$, j, 1) = "\" Then Exit For
    Next
    P2$ = P1$: P$ = left$(a$, j): X$ = Mid$(a$, j + 1)
    If left$(a$, 1) = "\" Then P2$ = left$(P1$, 2)
    b$ = X$: i = InStr(X$, ".")
    If i > 0 Then b$ = left$(X$, i - 1)
    Return
End Sub

Private Sub lstNames_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then lblDelFile_Click
End Sub

Private Sub lstNames_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo errh

'Debug.Print " OLEDragDrop :: " & Data.Files.Count & Data.Files(1)
If Data.Files.Count = 0 Then
    
    Exit Sub
End If

Dim fso As New FileSystemObject ', file0 As file
Dim file As String

For i = 1 To Data.Files.Count
    'strFiles = strFiles & Data.Files(i)
    
    file = Data.Files(i)
    If fso.FileExists(file) Then
          
        If file = "" Then Exit Sub
        'file = LTrim(temp)
        lstPath.AddItem file
        For j = Len(file) To 1 Step -1
          If Mid(file, j, 1) = "\" Then Exit For
        Next
        'P2 = P1: P = left(file, j): X = Mid(file, j + 1)
        'If left(file, 1) = "\" Then P2 = left(P1, 2)
        'b = X: i = InStr(X, ".")
        'If i > 0 Then b = left(X, i - 1)
        
        lstNames.AddItem fso.GetFileName(file)
        SaveSetting App.Title, "Config", "LOpenPath", P
        strLOpenPath = P
    '.AddItem Data.Files(i)
    End If
Next i

Set fso = Nothing

End Sub
