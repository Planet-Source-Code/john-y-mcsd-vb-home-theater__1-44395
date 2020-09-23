Attribute VB_Name = "modHTheater"
Public Const AliasName = "movie"
Global Path As String
Public strFilePath As String
Public strCommand As String
Public bHTEnd As Boolean

Sub Main()
  If App.PrevInstance Then
    Dim SaveTitle As String
    If Command$ <> "" Then SaveSetting App.Title, "Config", "Command", Command$
    SaveTitle = App.Title
    App.Title = ""
    AppActivate SaveTitle
    End
  End If

  Path = App.Path
  If Right$(Path, 1) <> "\" Then Path = Path & "\"
  
  frmMain.Show
End Sub
