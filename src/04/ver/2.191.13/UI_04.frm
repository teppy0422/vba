VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_04 
   Caption         =   "VerUp"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   OleObjectBlob   =   "UI_04.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






































































































Public thisVer As String
Public newVer As String

 Sub fjkajfdaljdfka()
    PlaySound "けってい"
    
    Call アドレスセット(ActiveWorkbook)
    Path = アドレス(0) & "\ver"
    If Dir(Path, vbDirectory) = "" Then MkDir (Path)

    Path = Path & "\" & Mid(ThisWorkbook.Name, 6, InStr(ThisWorkbook.Name, "_") - 6)
    If Dir(Path, vbDirectory) = "" Then MkDir (Path)
    
    myCount = VBC_Export(Path)
    Call Sheet_Ver_Export(Path)
    
    DoEvents
    
    If myCount = 0 Then
        MsgBox "エクスポート出来るコードがありませんでした。"
    Else
        MsgBox myCount & " 点のコードをエクスポートしました。"
    End If
    
    Unload UI_04
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Shift = 1 Then
        PlaySound "けってい"
        
        Call アドレスセット(ActiveWorkbook)
        Path = アドレス(0) & "\ver"
        If Dir(Path, vbDirectory) = "" Then MkDir (Path)
    
        Path = Path & "\" & Mid(ThisWorkbook.Name, 6, InStr(ThisWorkbook.Name, "_") - 6)
        If Dir(Path, vbDirectory) = "" Then MkDir (Path)
        Dim newVer As String: newVer = Mid(Path, InStrRev(Path, "\") + 1)
        
        myCount = VBC_Export(Path)
        Call Sheet_Ver_Export(Path)
        Call MakeShortcut(Path)
        DoEvents
        
        If myCount = 0 Then
            MsgBox "エクスポート出来るコードがありませんでした。"
        Else
            MsgBox myCount & " 点のコードをエクスポートしました。"
            Call ログ出力("test", "test", "VerExport = " & newVer)
        End If
        
        Unload UI_04
    End If
End Sub

Private Sub CommandButton2_Click()
    If CB0.Value = "" Then MsgBox "バージョンを選択して実行してください。": End
    If Left(ThisWorkbook.Name, Len(mySystemName)) <> mySystemName Then MsgBox "ファイル名が" & mySystemName & "から始まっている必要があります。" & vbCrLf & "名前を修正して下さい。": End
    newVer = CB0.Value
    aa = MsgBox("VerUpを実行します。" & vbCrLf & thisVer & " → " & newVer & vbCrLf & "処理の都合上、別ブックからの実行になります。バージョンアップの実行ボタンを押してください。", vbYesNo): If aa = vbNo Then End
    PlaySound "けってい"
    
    Call DeleteDefinedNames '名前の定義が重複したら警告出るから削除する
    mywb = ActiveWorkbook.FullName
    Workbooks.Open アドレス(0) & "\VerUp.xlsm"
    Set wb(0) = ActiveWorkbook
    
    With wb(0).Sheets("Sheet1")
        .Cells(1, 1) = アドレス(0) & "\ver\" & newVer
        .Cells(2, 1) = mywb
    End With
    
    Call ログ出力("test", "test", "VerUP" & thisVer & "→" & newVer)
    
    Unload UI_04
End Sub

Private Sub CommandButton4_Click()
    PlaySound ("もどる")
    Unload Me
    UI_Menu.Show
End Sub

Private Sub UserForm_Initialize()
    
    Dim buf As String, msg As String
    Dim 項目(1) As String
    Dim myDateTime
    Dim nowVer As String
    
    nowVer = ActiveWorkbook.Name
    nowVer = Mid(nowVer, InStr(nowVer, "+") + 1, InStr(nowVer, "_") - InStr(nowVer, "+") - 1)
    
    Me.Caption = nowVer
    Call アドレスセット(ActiveWorkbook)
    buf = Dir(アドレス(0) & "\ver\", vbDirectory)
    Do While buf <> ""
        If Replace(buf, ".", "") <> "" Then
            項目(0) = 項目(0) & "," & buf
            項目(1) = 項目(1) & "," & FileDateTime(アドレス(0) & "\ver\" & buf)
        End If
        buf = Dir()
    Loop
    項目(0) = Mid(項目(0), 2)
    項目(1) = Mid(項目(1), 2)
    Debug.Print msg
    
    項目0s = Split(項目(0), ",")
    項目1s = Split(項目(1), ",")
    With CB0
        .RowSource = ""
        For i = LBound(項目0s) To UBound(項目0s)
            .AddItem 項目0s(i)
            If 項目1s(i) > myDateTime Then myindex = i
            myDateTime = 項目1s(i)
        Next i
        .ListIndex = UBound(項目0s)
    End With
    
    newVer = CB0.Value
    thisVer = Mid(ThisWorkbook.Name, Len(mySystemName) + 1, InStr(ThisWorkbook.Name, "_") - Len(mySystemName) - 1)
    
    If thisVer = newVer Then
        messe.Caption = "バージョンは最新です"
    ElseIf thisVer < newVer Then
        messe.Caption = "新しいバージョンがあります"
        messe.ForeColor = RGB(255, 0, 0)
    Else
        messe.Caption = "このバージョンがより新しいです。" & vbCrLf & "エクスポートを実行してください。"
        messe.ForeColor = RGB(255, 0, 0)
        CommandButton2.Visible = False
    End If
    
End Sub
