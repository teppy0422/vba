Attribute VB_Name = "M30_Sound"
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Sub PlaySound(Optional SoundFile As String)
    Select Case SoundFile
    Case "けってい"
        SoundFile = "cursor2.wav"
    Case "もどる"
        SoundFile = "knock.wav"
    Case "とじる"
        SoundFile = "kaidan.wav"
    Case "じっこう"
        SoundFile = "enter.wav"
    Case "せんたく"
        SoundFile = "suiteki2.wav"
    Case "かんせい"
        SoundFile = "heal.wav"
    Case "とじる2"
        SoundFile = "Ring06.wav"
    Case "じっこう2" 'Ctrl & Enter
        SoundFile = "決定2.wav"
    Case "ばーじょんあっぷ"
        SoundFile = "m2 psi1.wav"
    Case Else
        Stop '登録されてないサウンド、サウンド気にしないなら左のStopを削除
    End Select
    
    Dim rs As Long, an
    SoundFile = myAddress(0, 1) & "\sound\" & SoundFile
    On Error Resume Next
    an = Dir(SoundFile)
    If err.number = 52 Then Exit Sub
    On Error GoTo 0
    If an = "" Then
        Beep
        Exit Sub
    End If
    rc = mciSendString("Play " & SoundFile, "", 0, 0)
End Sub
