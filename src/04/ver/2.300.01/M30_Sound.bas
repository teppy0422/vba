Attribute VB_Name = "M30_Sound"
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Sub PlaySound(Optional SoundFile As String)
    Select Case SoundFile
    Case "�����Ă�"
        SoundFile = "cursor2.wav"
    Case "���ǂ�"
        SoundFile = "knock.wav"
    Case "�Ƃ���"
        SoundFile = "kaidan.wav"
    Case "��������"
        SoundFile = "enter.wav"
    Case "���񂽂�"
        SoundFile = "suiteki2.wav"
    Case "���񂹂�"
        SoundFile = "heal.wav"
    Case "�Ƃ���2"
        SoundFile = "Ring06.wav"
    Case "��������2" 'Ctrl & Enter
        SoundFile = "����2.wav"
    Case "�΁[����񂠂���"
        SoundFile = "m2 psi1.wav"
    Case Else
        Stop '�o�^����ĂȂ��T�E���h�A�T�E���h�C�ɂ��Ȃ��Ȃ獶��Stop���폜
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
