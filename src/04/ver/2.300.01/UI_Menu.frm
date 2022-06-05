VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Menu 
   Caption         =   "menu"
   ClientHeight    =   6120
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7840
   OleObjectBlob   =   "UI_Menu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "UI_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














































Private clsForm As New clsUserForm
Private THEME As Long, THEMEgray1 As Long, THEMEgray2 As Long
Private THEMEwhite As Long

Private Sub initFormSetting()

    Me.BorderColor = THEME
    
    Me.Labeltitle.Top = 1
    Me.Labeltitle.Left = 1
    Me.Labeltitle.Width = Me.Width - 3
    Me.Labeltitle.BackColor = THEME
    
    Me.btnClose.Top = 1
    Me.btnClose.Left = Me.Labeltitle.Width - Me.btnClose.Width + 1
    
    Me.btnHelp.Top = 1
    Me.btnHelp.Left = Me.btnClose.Left - Me.btnHelp.Width - 3
    
    Me.myVerup.Top = 1
    
    Me.current.Top = 1
    Me.myVerup.BackColor = THEME
    Me.myVerup.ForeColor = white
    
    Me.Label0.ForeColor = black
    
End Sub

Private Sub NormalizeSet()
    Me.btnClose.BackColor = THEME
    Me.btnClose.ForeColor = clsForm.GetColor(white)
    Me.btnHelp.BackColor = THEME
    Me.btnHelp.ForeColor = clsForm.GetColor(white)
    Me.myVerup.BackColor = THEME
    Me.myVerup.ForeColor = clsForm.GetColor(white)
    Me.current.BackColor = THEME
    Me.current.ForeColor = clsForm.GetColor(white)
End Sub
Private Sub Normalizeset_tag()
        Me.tag1.ForeColor = clsForm.GetColor(gray02)
        Me.tag2.ForeColor = clsForm.GetColor(gray02)
End Sub
Private Sub current_Click()
    Shell "C:\Windows\explorer.exe " & ThisWorkbook.path, vbNormalFocus
End Sub

Private Sub current_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.current.BackColor = clsForm.GetColor(white)
    Me.current.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub Image5_Click()
    Me.MultiPage1.Value = 1
End Sub

Private Sub Image5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.ChangeCursor Hand
End Sub

Private Sub Image6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.ChangeCursor Hand
End Sub

Private Sub in01_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_00.Show
End Sub

Private Sub in01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in01.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub in02_Click()
    'If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_02.Show
End Sub

Private Sub in02_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in02.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub in03_Click()
    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_07.Show
End Sub

Private Sub in03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in03.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub in04_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_08.Show
End Sub

Private Sub in04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.in04.ForeColor = vbRed
    clsForm.ChangeCursor Hand
End Sub

Private Sub Label11_Click()
    
End Sub

Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Me.in01.ForeColor = THEMEgray2
        Me.in02.ForeColor = THEMEgray2
        Me.in03.ForeColor = THEMEgray2
        Me.in04.ForeColor = THEMEgray2
        Call NormalizeSet
End Sub

Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        Me.out01.ForeColor = THEMEgray2
        Me.out03.ForeColor = THEMEgray2
        Me.out04.ForeColor = THEMEgray2
        Me.out05.ForeColor = THEMEgray2
        Me.out06.ForeColor = THEMEgray2
        Me.out07.ForeColor = THEMEgray2
        Me.out08.ForeColor = THEMEgray2
        Me.out09.ForeColor = THEMEgray2
        
        Me.in01.ForeColor = THEMEgray2
        Me.in02.ForeColor = THEMEgray2
        Me.in03.ForeColor = THEMEgray2
        Me.in04.ForeColor = THEMEgray2
        Call NormalizeSet
End Sub

Private Sub MultiPage1_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call Normalizeset_tag
End Sub

Private Sub myVerup_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_04.Show
End Sub

Private Sub out01_Click()
    aa = MsgBox("これは検討中です。" & vbLf & "実行しますか?", vbYesNo, "回路マトリクス")
    If aa <> 6 Then Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    Call 回路マトリクス作成_徳島式
End Sub

Private Sub out01_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out01.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out03_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_06.Show
End Sub

Private Sub out03_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out03.ForeColor = vbRed
    clsForm.ChangeCursor Hand
End Sub

Private Sub out04_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out04.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out04_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    If Shift = 1 Then サンプル作成モード = True
    PlaySound ("けってい")
    Unload UI_Menu
    UI_01.Show
End Sub

Private Sub out05_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_03.Show
End Sub

Private Sub out05_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out05.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out06_Click()
    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_05.Show
End Sub

Private Sub out06_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out06.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out07_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_09.Show
End Sub

Private Sub out07_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out07.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out08_Click()
    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_10.Show
End Sub

Private Sub out09_Click()
'    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_11.Show
End Sub

Private Sub out08_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out08.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub out09_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.out09.ForeColor = THEMEwhite
    clsForm.ChangeCursor Hand
End Sub

Private Sub tag1_Click()
    Me.MultiPage1.Value = 0
End Sub

Private Sub tag1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.tag1.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub tag2_Click()
    Me.MultiPage1.Value = 1
End Sub

Private Sub tag2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.tag2.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.FormDrag Me.Name, Button
End Sub
Private Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.btnClose.BackColor = clsForm.GetColor(red)
    Me.btnClose.ForeColor = clsForm.GetColor(white)
    clsForm.ChangeCursor Hand
End Sub

Private Sub UserForm_Initialize()
    Set wb(0) = ThisWorkbook
    myIP = GetIPAddress
    addressSet wb(0)
    Me.Label_IP.Caption = myIP
    Call 参照不可があればそのフォルダを作成する
    Call 参照設定の変更
    
    Call connect_Server
    
    myVer = Base.verCheck(ThisWorkbook)
    Call ディレクトリ作成
    Call 必要ファイルの取得
    Call 最適化
    HDsize.Caption = checkSpace(myAddress(0, 1))
    'フォームデザイン
    On Error GoTo ErrHandler
    Static initCompleted As Boolean
    If initCompleted = False Then
        initCompleted = True
        THEME = clsForm.GetColor(TBLUE)         ' Choose theme colors
        THEMEgray1 = RGB(100, 100, 100)
        THEMEgray2 = RGB(220, 220, 220)
        THEMEwhite = RGB(255, 255, 255)
        THEMEred = RGB(255, 0, 0)
        
        clsForm.NonTitleBar Me.Name                      ' Set Flat style
        Call initFormSetting
    End If
    GoTo Finally
ErrHandler:
    Call MsgBox(err.Description, , "生準+:例外が発生しました。")
Finally:
    Me.startupposition = 2
    On Error GoTo 0
    Me.btnClose.BackColor = THEME
    Me.btnClose.ForeColor = clsForm.GetColor(white)
    Me.btnHelp.BackColor = THEME
    Me.btnHelp.ForeColor = clsForm.GetColor(white)
    Me.myVerup.BackColor = THEME
    Me.myVerup.ForeColor = clsForm.GetColor(white)
    Me.current.BackColor = THEME
    Me.current.ForeColor = clsForm.GetColor(white)
    
    Me.Labeltitle.Caption = "生産準備+" & myVer
    Dim FSO As New FileSystemObject
    'アドレスにアクセスできるか調べる
    With ActiveWorkbook.Sheets("設定")
        Dim アドレスb As Variant, myMsg As String
        For i = 0 To 2
            アドレスb = myAddress(i, 1)
            Select Case i
            Case 0, 1
                If FSO.FolderExists(アドレスb) = False Then
                    myMsg = myMsg & myAddress(i, 0) & " のフォルダが見つかりません" & vbCrLf
                Else
                    myMsg = myMsg & myAddress(i, 0) & " のフォルダが見つかりました" & vbCrLf
                End If
            Case 2
                If アドレスb = "" Then
                    myMsg = myMsg & "このIPでは " & myAddress(i, 0) & " の登録がされていません｡"
                Else
                    If FSO.FileExists(アドレスb) = False Then
                        myMsg = myMsg & myAddress(i, 0) & " のファイルが見つかりません" & vbCrLf
                    Else
                        myMsg = myMsg & myAddress(i, 0) & " のファイルが見つかりました" & vbCrLf
                    End If
                End If
            End Select
        Next i
    End With
    
    'アドレス確認の結果
    With Label0
        .Caption = myMsg
        If InStr(myMsg, "見つかりません") > 0 Then
            .ForeColor = RGB(255, 0, 0)
        Else
            .ForeColor = RGB(255, 255, 255)
        End If
    End With
    Debug.Print Label0.ForeColor
    Set FSO = Nothing
    'フィールド名のチェック
    Call fieldAdd("製品品番", "フィールド名_製品品番", 2)
    Call fieldAdd("PVSW_RLTF", "フィールド名_通常", 1)
    Call fieldAdd("PVSW_RLTF", "フィールド名_追加", 2)
    Call fieldAdd("PVSW_RLTF", "フィールド名_追加2", 2)
    
    Call 最適化もどす
End Sub
'**********************************
'top label
'**********************************
Private Sub labelTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call NormalizeSet
    clsForm.ChangeCursor Cross
End Sub

Private Sub labelTitle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.FormDrag Me.Name, Button
End Sub
'**********************************
'close button
'**********************************
Private Sub btnClose_Click()
    Unload Me
End Sub
'**********************************
'help button
'**********************************
Private Sub btnHelp_Click()
    buf = "http://10.7.1.35/nim_intra/40_program/plus/41_web/myweb/index.html "
    'If Dir(buf, vbDirectory) <> "" Then
        'buf = buf & "\myWeb\index.html"
        'IEの起動
        Dim objIE As Object '変数を定義します。
        Dim ieVerCheck As Variant
    
        Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
        Set objSFO = CreateObject("Scripting.FileSystemObject")
    
        ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
        
        'Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
        
        If ieVerCheck >= 11 Then
            Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10以降(たぶん)
        End If
        
        objIE.Visible = True      '可視、Trueで見えるようにします。
        
        '処理したいページを表示します。
       objIE.Navigate buf
       
       Set objIE = Nothing
    'End If
End Sub

Private Sub btnHelp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.btnHelp.BackColor = clsForm.GetColor(white)
    Me.btnHelp.ForeColor = THEME
    clsForm.ChangeCursor Hand
End Sub

Private Sub myVerup_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Me.myVerup.BackColor = clsForm.GetColor(red)
    Me.myVerup.ForeColor = clsForm.GetColor(white)
    clsForm.ChangeCursor Hand
End Sub
'**********************************
'excute button
'**********************************
Private Sub btnExcute_Click()
'    Me.btnExcute.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub btnExcute_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    clsForm.ChangeCursor Hand
End Sub

'**********************************
'bottom label
'**********************************
Private Sub labelBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call NormalizeSet
End Sub

Private Sub CommandButton5_Click()
    aa = MsgBox("これは検討中です。" & vbLf & "実行しますか?", vbYesNo, "回路マトリクス")
    If aa <> 6 Then Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    Call 回路マトリクス作成_徳島式
End Sub

Private Sub CommandButton8_Click()
    If Label0.ForeColor = 255 Then MsgBox "設定を確認してください", , "実行できません": Exit Sub
    PlaySound ("けってい")
    Unload UI_Menu
    UI_70.Show
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Call NormalizeSet
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then PlaySound "とじる"
End Sub

Private Sub UserForm_Terminate()
    Application.WindowState = xlMaximized
End Sub

Private Sub Version_Click()

End Sub
