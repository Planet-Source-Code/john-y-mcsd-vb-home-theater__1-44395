VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "VB Home Theater"
   ClientHeight    =   4680
   ClientLeft      =   600
   ClientTop       =   2640
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   Begin VB.PictureBox picVideo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Timer tmrHTheater 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   240
   End
   Begin VB.Timer tmrFile 
      Interval        =   10
      Left            =   240
      Top             =   840
   End
   Begin VB.Image imgSlider 
      Height          =   180
      Left            =   120
      Picture         =   "frmMain.frx":0BC2
      Top             =   3870
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Shape shpTime 
      BorderColor     =   &H00404040&
      Height          =   150
      Left            =   120
      Top             =   3885
      Width           =   4695
   End
   Begin VB.Label lblPlayList 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   3075
      TabIndex        =   4
      Top             =   4155
      Width           =   255
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   3405
      TabIndex        =   3
      Top             =   4155
      Width           =   255
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   3720
      TabIndex        =   2
      Top             =   4155
      Width           =   255
   End
   Begin VB.Image imgNext1 
      Height          =   300
      Left            =   2880
      Picture         =   "frmMain.frx":1154
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPrev1 
      Height          =   300
      Left            =   2400
      Picture         =   "frmMain.frx":11E8
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgStop1 
      Height          =   300
      Left            =   1920
      Picture         =   "frmMain.frx":127D
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPause1 
      Height          =   300
      Left            =   1440
      Picture         =   "frmMain.frx":1302
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPlay1 
      Height          =   300
      Left            =   240
      Picture         =   "frmMain.frx":1390
      Top             =   5040
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgNext2 
      Height          =   300
      Left            =   2880
      Picture         =   "frmMain.frx":1439
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPrev2 
      Height          =   300
      Left            =   2400
      Picture         =   "frmMain.frx":14A5
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgStop2 
      Height          =   300
      Left            =   1920
      Picture         =   "frmMain.frx":150F
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPause2 
      Height          =   300
      Left            =   1440
      Picture         =   "frmMain.frx":1569
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgPlay2 
      Height          =   300
      Left            =   240
      Picture         =   "frmMain.frx":15CA
      Top             =   4680
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Image imgNext 
      Height          =   300
      Left            =   2280
      Picture         =   "frmMain.frx":1639
      Top             =   4185
      Width           =   330
   End
   Begin VB.Image imgPrev 
      Height          =   300
      Left            =   1920
      Picture         =   "frmMain.frx":16CD
      Top             =   4185
      Width           =   330
   End
   Begin VB.Image imgStop 
      Height          =   300
      Left            =   1560
      Picture         =   "frmMain.frx":1762
      Top             =   4185
      Width           =   330
   End
   Begin VB.Image imgPause 
      Height          =   300
      Left            =   1200
      Picture         =   "frmMain.frx":17E7
      Top             =   4200
      Width           =   330
   End
   Begin VB.Image imgPlay 
      Height          =   300
      Left            =   120
      Picture         =   "frmMain.frx":1875
      Top             =   4185
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   328
      Y1              =   249
      Y2              =   249
   End
   Begin VB.Label lblCurrTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":"
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
      Height          =   300
      Left            =   4080
      TabIndex        =   1
      Top             =   4185
      Width           =   735
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuPlaylist 
         Caption         =   "&Playlist"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRate 
         Caption         =   "&Rate"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pau&se"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "P&lay"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuMax 
         Caption         =   "Ma&ximize"
         Shortcut        =   {F11}
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currpos As String
Dim CurrentTime As String
Dim TotalFrames As String
Dim TotalTime As String
Dim FramesPerSecond As String
Dim FullScreen As Boolean
Dim Paused As Boolean
Dim PLVisible As Boolean
Dim RVisible As Boolean
Dim PressKeys As String
Dim SlideFlag As Boolean
Dim IX, IY, TX, TY, FX, FY

Public Sub CloseVideo()
Dim Result As String

Result = CloseMultimedia(AliasName)

If Result = "Success" Then
tmrHTheater.Enabled = False

currpos = 0
End If
End Sub

Public Sub OpenVideo(FileName As String)
Dim typeDevice As String
Dim Result As String

If LCase(Right(FileName, 4)) = ".avi" Then
  typeDevice = "AviVideo"
Else
  typeDevice = "MPEGVideo"
End If

Result = OpenMultimedia(picVideo.hWnd, AliasName, FileName, typeDevice)

If Result = "Success" Then
    On Error Resume Next
    Width = CInt(GetSize(AliasName, "cx")) * 15
    Height = (CInt(GetSize(AliasName, "cy")) * 15) + 1650
    If frmRate.Visible = True Then frmRate.txtRate = GetRate(AliasName)
    FramesPerSecond = GetFramesPerSecond(AliasName)
    TotalFrames = GetTotalframes(AliasName)  'Get total frames
    TotalTime = GetTotalTimeByMS(AliasName) / 1000   'Get Total Time
    tmrHTheater.Enabled = True
    strFilePath = FileName
End If
End Sub

Public Sub PauseVideo()
Dim Result As String

Result = PauseMultimedia(AliasName)
End Sub

Public Sub PlayVideo()
Dim Result As String

Form_Resize
DoEvents
ResizeVideo 0, 0, 0, 0
imgSlider.Move 8: imgSlider.Visible = True


Result = PlayMultimedia(AliasName, 0, 0)
End Sub

Public Sub ResizeVideo(left As Long, top As Long, Width As Long, Height As Long)
Dim Result As String

Result = PutMultimedia(picVideo.hWnd, AliasName, left, top, Width, Height)
End Sub

Public Sub ResumeVideo()
Dim Result As String

Result = ResumeMultimedia(AliasName)
End Sub

Public Sub StopVideo()
Dim Result As String

Result = StopMultimedia(AliasName)
imgSlider.Visible = False
End Sub

Private Sub Form_Load()
  If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
    SetDefaultDevice "MPEGVideo", "mciqtz.drv"
  End If
  If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
    SetDefaultDevice "avivideo", "mciavi.drv"
  End If
  FullScreen = False
  XLeft = GetSetting(App.Title, "Config", "MainX", Me.left)
  YTop = GetSetting(App.Title, "Config", "MainY", Me.top)
  Me.Move XLeft, YTop
  Show
  Load frmPList
  If GetSetting(App.Title, "Config", "SPLS", "True") = "True" Then
    frmPList.Show
  End If
  If GetSetting(App.Title, "Config", "SRS", "True") = "True" Then
    frmRate.Show
  End If
  If Command = "" Then
    If Dir$(App.Path & "\plist.htp") <> "" Then
      frmPList.OpenPList App.Path & "\plist.htp"
    End If
  Else
    ParseCommand Command
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  bHTEnd = True
  Unload frmPList
  DoEvents
  tmrHTheater.Enabled = False
  tmrFile.Enabled = False
  SaveSetting App.Title, "Config", "MainX", Str$(Me.left)
  SaveSetting App.Title, "Config", "MainY", Str$(Me.top)
  Dim Result As String
  Result = CloseAll()
  DoEvents
  End
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Width <= 4800 Then Width = 4800
  If Height <= 5000 Then Height = 5000
  picVideo.Width = ScaleWidth
  picVideo.Height = ScaleHeight - 64
  Line1.Y1 = picVideo.Height: Line1.Y2 = Line1.Y1
  Line1.X2 = ScaleWidth
  shpTime.top = picVideo.Height + 10
  shpTime.Width = ScaleWidth - 16
  imgSlider.top = shpTime.top - 1
  lblCurrTime.left = ScaleWidth - 57
  lblCurrTime.top = shpTime.top + 20
  lblAbout.left = lblCurrTime.left - 24
  lblAbout.top = lblCurrTime.top - 2
  lblRate.left = lblAbout.left - 24
  lblRate.top = lblAbout.top
  lblPlayList.left = lblRate.left - 24
  lblPlayList.top = lblAbout.top
  imgPlay.top = lblCurrTime.top
  imgPause.top = imgPlay.top
  imgStop.top = imgPlay.top
  imgPrev.top = imgPlay.top
  imgNext.top = imgPlay.top
  If Width >= 5300 Then
    imgPause.left = 88
    imgStop.left = 120
    imgPrev.left = 152
    imgNext.left = 184
  Else
    imgPause.left = 80
    imgStop.left = 104
    imgPrev.left = 128
    imgNext.left = 152
  End If
  ResizeVideo 0, 0, 0, 0
End Sub

Private Sub imgNext_Click()
  If frmPList.lstNames.ListIndex = frmPList.lstNames.ListCount - 1 Then Exit Sub
  frmPList.lstNames.ListIndex = frmPList.lstNames.ListIndex + 1
  frmPList.lstNames_DblClick
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgNext = imgNext2
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgNext = imgNext1
End Sub

Private Sub imgPause_Click()
  PauseVideo
  Paused = True
End Sub

Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPause = imgPause2
End Sub

Private Sub imgPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPause = imgPause1
End Sub

Private Sub imgPlay_Click()
  If Paused = True Then
    ResumeVideo
    Paused = False
    Exit Sub
  End If
  If strFilePath <> "" Then
    If strFilePath = frmPList.lstPath.List(frmPList.lstNames.ListIndex) Then
      PlayVideo
    Else
      frmPList.lstNames_DblClick
    End If
  Else
    If frmPList.lstPath.ListCount = 0 Then Exit Sub
    If frmPList.lstNames.SelCount = 0 Then
      frmPList.lstNames.ListIndex = 0
      frmPList.lstNames_DblClick
    End If
  End If
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPlay = imgPlay2
End Sub

Private Sub imgPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPlay = imgPlay1
End Sub

Private Sub imgPrev_Click()
  If frmPList.lstNames.ListIndex <= 0 Then Exit Sub
  frmPList.lstNames.ListIndex = frmPList.lstNames.ListIndex - 1
  frmPList.lstNames_DblClick
End Sub

Private Sub imgPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPrev = imgPrev2
End Sub

Private Sub imgPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgPrev = imgPrev1
End Sub

Private Sub imgSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SlideFlag = False Then
    IX = X: FX = imgSlider.left
    TX = Screen.TwipsPerPixelX
    SlideFlag = True
  End If
End Sub

Private Sub imgSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SlideFlag = True Then
    pos = FX + (X - IX) / TX
    If pos < 8 Then pos = 8
    If pos > ScaleHeight - 20 Then pos = ScaleHeight - 20
    FX = pos: imgSlider.left = pos
  End If
End Sub

Private Sub imgSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim From As String
  From = Int(((imgSlider.left - 28) / ((shpTime.Width) - 70)) * TotalFrames)
  Result = StopMultimedia(AliasName)
  Result = PlayMultimedia(AliasName, From, 0)
  SlideFlag = False
End Sub

Private Sub imgStop_Click()
  StopVideo
  CloseVideo
  Caption = App.Title
  lblCurrTime = ":"
  SetFocus
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgStop = imgStop2
End Sub

Private Sub imgStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgStop = imgStop1
End Sub

Private Sub lblAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub lblPlayList_Click()
  frmPList.Show
End Sub

Private Sub lblRate_Click()
  frmRate.Show
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1
End Sub

Private Sub mnuMax_Click()
If Me.WindowState = 2 Then
    Me.WindowState = 0
ElseIf Me.WindowState = 0 Then
    Me.WindowState = 2
End If

End Sub

Private Sub mnuOptions_Click()
    If frmOptions.Visible = True Then
        frmOptions.Hide
    Else
        frmOptions.Show
    End If
End Sub

Private Sub mnuPause_Click()
    Call imgPause_Click
End Sub

Private Sub mnuPlay_Click()
    Call imgPlay_Click
End Sub

Private Sub mnuPlaylist_Click()
    If frmPList.Visible = True Then
        frmPList.Hide
    Else
        frmPList.Show
    End If
End Sub

Private Sub mnuRate_Click()
  If frmRate.Visible = True Then
    frmRate.Hide
  Else
    frmRate.Show
  End If
End Sub

Private Sub mnuStop_Click()
    Call imgStop_Click
End Sub

Private Sub picVideo_KeyDown(KeyCode As Integer, Shift As Integer)
  PressKeys = KeyCode & PressKeys
End Sub

Private Sub picVideo_KeyUp(KeyCode As Integer, Shift As Integer)
  If left$(PressKeys, 4) = "1318" Then
    If FullScreen = True Then
      WindowState = 0
      FullScreen = False
    Else
      WindowState = 2
      FullScreen = True
    End If
  End If
  PressKeys = ""
End Sub

Private Sub picVideo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo errh

''' added code John Yung MCSD 20030402
Debug.Print "cboFile_OLEDragDrop :: " & Data.Files.Count & Data.Files(1)

'If lstNames.ListCount = 0 Then Exit Sub

If Data.Files.Count = 1 Then

  frmMain.StopVideo: frmMain.CloseVideo
  frmMain.OpenVideo Data.Files(1)
  DoEvents
  frmMain.PlayVideo
  frmMain.SetFocus
    
    Exit Sub
End If
''' added code John Yung MCSD 20030402

End Sub

Private Sub tmrFile_Timer()
  strCommand = GetSetting(App.Title, "Config", "Command", "")
  If strCommand <> "" Then
    ParseCommand strCommand
    SaveSetting App.Title, "Config", "Command", ""
  End If
End Sub

Private Sub tmrHTheater_Timer()
  Dim Percent As Long

  Percent = GetPercent(AliasName)
  currpos = GetCurrentMultimediaPos(AliasName)
  CurrentTime = Val(currpos) / Val(FramesPerSecond)
  If SlideFlag = False Then
    imgSlider.left = 8 + Int((currpos / TotalFrames) * (shpTime.Width - 26))
  End If
  Dim min As Integer
  Dim sec As Integer
  min = CurrentTime \ 60
  sec = CurrentTime - (min * 60)
  If sec = "-1" Then sec = "0"
  lblCurrTime = Format$(min, "00") & ":" & Format$(sec, "00")
  Dim a As Integer
  Dim b As Integer
  a = TotalTime \ 60
  b = TotalTime - (a * 60)
  If b = "-1" Then sec = "0"
  Caption = "VB Home Theater - " & lblCurrTime & "/" & Format$(a, "00") & ":" & Format$(b, "00")
  If AreMultimediaAtEnd(AliasName, 0) = True Then
    If frmPList.lstNames.ListCount = 1 Then
      tmrHTheater.Enabled = False
      StopVideo
      CloseVideo
    End If
    frmPList.PlayNext
  End If
End Sub

Private Sub SOpenPlayFile()
  frmPList.lstPath.Clear
  frmPList.lstNames.Clear
  file = LTrim$(strFilePath)
  If left$(file, 1) = Chr$(34) Then file = Mid$(file, 2)
  If Right$(file, 1) = Chr$(34) Then file = Mid$(file, 1, Len(file) - 1)
  frmPList.lstPath.AddItem file
  For j = Len(file) To 1 Step -1
    If Mid$(file, j, 1) = "\" Then Exit For
  Next
  P2$ = P1$: P$ = left$(file, j): X$ = Mid$(file, j + 1)
  If left$(file, 1) = "\" Then P2$ = left$(P1$, 2)
  b$ = X$: i = InStr(X$, ".")
  If i > 0 Then b$ = left$(X$, i - 1)
  frmPList.lstNames.AddItem b$
  frmPList.lstNames.ListIndex = 0
  frmPList.lstNames_DblClick
End Sub

Private Sub ParseCommand(ByVal sCmd As String)
  If (Me.Visible = False) Then
    Me.Visible = True
  End If
   
  sCmd = Trim$(sCmd)
   
  strFilePath = sCmd
  strFilePath = Replace(strFilePath, Chr$(34), "")
  If LCase(Right$(strFilePath, 4)) = ".htp" Then
    frmPList.OpenPList strFilePath
    DoEvents
    frmPList.lstNames.ListIndex = 0: frmPList.lstNames_DblClick
  Else
    SOpenPlayFile
  End If
End Sub
