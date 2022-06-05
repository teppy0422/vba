Attribute VB_Name = "M10_外部アプリ制御"
'ウインドウハンドルによる他アプリ制御
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
        (ByVal hParent As Long, ByVal hChildAfter As Long, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As Long
        '文字列をASC1文字毎渡すから遅い
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
        '文字列をそのまま渡すから速いけど、Classによっては使用不可
Declare Function SendMessageStr Lib "user32.dll" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal msg As Long, _
        ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDlgItem Lib "user32" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long

'定数_一般
Public Const WM_SETFOCUS = &H7
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLEN = &HE
Public Const WM_ALT = &H12
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104 'ALTとかこれ
Public Const WM_COMMAND = &H111&
Public Const WM_SYSCOMMAND = &H112
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203 'ダブルクリック
Public Const WM_IME_CHAR = &H286     '文字コード送信
Public Const WM_CLEAR = &H303
'WM_KEY*との違いはよく分からんけど、こっちでも動く
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291
'定数_リストボックス
Public Const LB_GETTEXT = &H189     '文字列
Public Const LB_GETTEXTLEN = &H18A  '文字列数
Public Const LB_GETCOUNT = &H18B    '要素数
Public Const LB_SETCURSEL = &H186   '指定した項目を選択
Public Const LB_SETTOPINDEX = &H197 '指定した項目をリスト上部に表示
'定数_コンボボックス
Public Const CB_GETTEXT = &H148
Public Const CB_GETTEXTLEN = &H149
Public Const CB_GETCOUNT = &H146
Public Const CB_SETCURSEL = &H14E
Public Const CB_SETTOPINDEX = &H15C
Public Const CB_SELECTSTRING = &H14D '使えるかわからん
Public Const CB_SHOWDROPDOWN = &H14F 'ドロップダウンリストの表示_0閉じる_1開く
'定数_ボタン
Public Const BN_CLICKED = 0&
Public Const BM_SETCHECK = &HF1      '0外す_1入れる
Public Const BM_CLICK = &HF5
Public Const BM_GETCHECK = &HF0      'ラジオorチェックボックスの状態を知る
'定数_その他
Public Const SC_CLOSE = &HF060
Public Const EM_SETSEL = &HB1
'定数_仮想キーコード
Public Const VK_MENU = &H12 'ALT
Public Const VK_RETURNE = &HD 'ENTER
'変数
Public 後引張り支援システムPath As String
Public myHND(10) As String
Public myHNDtemp As String
Dim i As Long
Dim Index, Ret, Rep As Integer

Sub Control_YcEditor()
    製品品番str = "8216136D40     "
    設変str = "test"
    Set myBook = ActiveWorkbook
    
    'Symbolデータの作成
    Call SQL_YcEditor_Symbol(RAN, myBook, 製品品番str)
    Dim i As Long
    '出力先Dirが無ければ作成
    Dim outPath(1) As String
    outPath(0) = myBook.Path & "\81_導通検査date_簡易"
    If Dir(outPath(0), vbDirectory) = "" Then MkDir outPath(0)
    outPath(1) = outPath(0) & "\" & Replace(製品品番str, " ", "")
    If Dir(outPath(1), vbDirectory) = "" Then MkDir outPath(1)
    '出力先bookを作成
    Set wb(3) = Workbooks.add
    Application.DisplayAlerts = False
    wb(3).SaveAs fileName:=outPath(1) & "\" & Replace(製品品番str, " ", "") & "_" & Replace(設変str, " ", "")
    Application.DisplayAlerts = True
    '出力sheetを作成
    wb(3).Worksheets.add
    wb(3).Sheets(1).Name = "Symbol"
    wb(3).Sheets(2).Name = "WH"
    
    With wb(3).Sheets("Symbol")
        .Activate
        .Cells.NumberFormat = "@"
        .Columns(1).NumberFormat = 0
        .Cells.Font.Name = "ＭＳ ゴシック"
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            For X = LBound(RAN, 1) To UBound(RAN, 1)
                .Cells(Y, X + 1) = RAN(X, Y)
            Next X
        Next Y
        '並び替え
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range(Rows(1), Rows(lastRow)).Sort key1:=Range("a1"), Order1:=xlAscending, Header:=xlNo
        Dim endPoint As Long
        endPoint = .Cells(lastRow, 1) + 200
        If endPoint > 1900 Then endPoint = 1900
    End With
    'WHデータの作成
    Call SQL_YcEditor_WH(RAN, myBook, 製品品番str)
    With wb(3).Sheets("WH")
        .Activate
        .Cells.NumberFormat = "@"
        .Cells.Font.Name = "ＭＳ ゴシック"
        For Y = LBound(RAN, 2) + 1 To UBound(RAN, 2)
            For X = LBound(RAN, 1) To UBound(RAN, 1)
                .Cells(Y, X + 1) = RAN(X, Y)
            Next X
        Next Y
        '並び替え
        lastRow = .Cells(.Rows.count, 1).End(xlUp).Row
        .Range(Rows(1), Rows(lastRow)).Sort key1:=Range("a1"), Order1:=xlAscending, Header:=xlNo
        Dim endKosei As Long
        endKosei = CLng(.Cells(lastRow, 1))
    End With
    'ボンダー、ジョイント、シールドドレン等の回路符号修正
    Stop
    
    'YcEditorに出力
    With myBook.Sheets("設定")
        For i = 0 To 10
            AppPath = .Cells.Find("YcEditor_exe", , , 1).Offset(i, 1)
            If Dir(AppPath) <> "" Then Exit For
        Next i
    End With
    '処理開始
    Call Control_アプリ起動(AppPath)
    'ハンドル取得
line00:
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    'ファイルを開く
    SetForegroundWindow myHND(0) 'ステップインだと最前面にならないから注意
    PostMessage myHND(0), WM_SYSKEYDOWN, VK_MENU, &H20380001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("F"), &H20210001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("O"), &H20180001
    '新規作成をクリック
    If myHND(0) = 0 Then GoTo line00
line01:
    myHND(1) = FindWindow("TfrmFile", "ファイル選択")
    If myHND(1) = 0 Then GoTo line01
    myHND(2) = FindWindowEx(myHND(1), 0&, "TButton", "新規")
    If myHND(2) = 0 Then GoTo line01
    PostMessage myHND(2), BM_CLICK, 0, 0
    '製品品番を入力
line02:
    myHND(3) = FindWindow("TForm", "新規ファイルの作成")
    If myHND(3) = 0 Then GoTo line02
    myHND(4) = FindWindowEx(myHND(3), 0&, "TEdit", vbNullString)
    If myHND(4) = 0 Then GoTo line02
    Call input_Message(myHND(4), Replace(製品品番str, " ", "") & "_" & Replace(設変str, " ", ""))
    'ENTER
line03:
    myHND(5) = FindWindowEx(myHND(3), 0&, "TButton", "OK")
    If myHND(5) = 0 Then GoTo line03
    Call Control_Click(myHND(5), "SEND", &H1&)
    'ヘッダー編集
    '製品品番
line04:
    myHND(6) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(7) = FindWindowEx(myHND(6), 0&, "TfrmHeader", vbNullString)
    myHND(8) = FindWindowEx(myHND(7), 0&, "TEdit", "00000000000000000000")
    If myHND(8) = 0 Then GoTo line04:
    SendMessage myHND(8), EM_SETSEL, 0, 20 '文字を選択
    SendMessage myHND(8), WM_CLEAR, 0, 0 '文字をクリア
    SendMessage myHND(8), WM_LBUTTONDOWN, 0&, 0& '選択
    SendMessage myHND(8), WM_LBUTTONUP, 0&, 0&
    a = SendMessageStr(myHND(8), WM_SETTEXT, 0&, Replace(製品品番str, " ", "") & "_" & Replace(設変str, " ", ""))
    'WH製品品番
    myHND(8) = FindWindowEx(myHND(7), 0&, "TEdit", "00000000")
    SendMessage myHND(8), EM_SETSEL, 0, 20 '文字を選択
    SendMessage myHND(8), WM_CLEAR, 0, 0 '文字をクリア
    SendMessage myHND(8), WM_LBUTTONDOWN, 0&, 0& '選択
    SendMessage myHND(8), WM_LBUTTONUP, 0&, 0&
    a = SendMessageStr(myHND(8), WM_SETTEXT, 0&, ByVal Right(Replace(製品品番str, " ", ""), 8))
    PostMessage myHND(8), WM_KEYDOWN, &HD, 0 'ENTERで確定
    'ｴﾝﾄﾞﾎﾟｲﾝﾄ
    myHND(8) = FindWindowEx(myHND(7), 0&, "TEdit", "100")
    SendMessage myHND(8), EM_SETSEL, 0, 20 '文字を選択
    SendMessage myHND(8), WM_CLEAR, 0, 0 '文字をクリア
    SendMessage myHND(8), WM_LBUTTONDOWN, 0&, 0& '選択
    SendMessage myHND(8), WM_LBUTTONUP, 0&, 0&
    a = SendMessageStr(myHND(8), WM_SETTEXT, 0&, "1900")
    PostMessage myHND(8), WM_KEYDOWN, &HD, 0 'ENTERで確定
    'YC機種
    myHND(8) = FindWindowEx(myHND(7), 0&, "TComboBox", vbNullString)
    Index = SendMessage(myHND(8), CB_GETCOUNT, 0, 0)
    For i = 0 To Index - 1
        myLen = SendMessage(myHND(8), CB_GETTEXTLEN, i, 0)
        Dim myStr As String: myStr = String(myLen, vbNullChar)
        Ret = SendMessageStr(myHND(8), CB_GETTEXT, i, myStr)
        a = SendMessage(myHND(8), CB_SETTOPINDEX, i, 0)
        b = SendMessage(myHND(8), CB_SETCURSEL, i, 0)
        c = SendMessage(myHND(8), BM_CLICK, i, 0)
        If i = 18 Then Exit For
    Next i
    PostMessage myHND(8), WM_KEYDOWN, &HD, 0 'ENTERで確定
    '検査ﾓｰﾄﾞ
    myHNDtemp = GetWindow(myHND(8), 2) '下のハンドル
    Index = SendMessage(myHNDtemp, CB_GETCOUNT, 0, 0)
    For i = 0 To Index - 1
        a = SendMessage(myHNDtemp, CB_SETTOPINDEX, i, 0)
        b = SendMessage(myHNDtemp, CB_SETCURSEL, i, 0)
        c = SendMessage(myHNDtemp, BM_CLICK, i, 0)
        If i = 2 Then Exit For
    Next i
    PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTERで確定
    'Symbolデータ編集
    Stop 'シンボルデータ編集をアクティブにする
    myHND(2) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(3) = FindWindowEx(myHND(2), 0&, "TfrmSymbol", "シンボルデータ編集")
    SetForegroundWindow myHND(3) 'ステップインだと最前面にならないから注意
    'PostMessage myHND(3), &H222, &H408E4, &H4081A
    myHND(4) = GetWindow(myHND(3), 5)
    myHND(5) = GetWindow(myHND(4), 1)
    myHNDtemp = myHND(5)
    'シンボルデータ編集に回路符号を与える
    Dim pageMax As Long: pageMax = 100
    With wb(3).Sheets("Symbol")
        For s = 1 To endPoint
            If s > pageMax Then
                myHNDtemp = myHND(5)
                pageMax = pageMax + 100
            End If
            Set myPoint = .Cells.Columns(1).Find(s, , , 1)
            If Not (myPoint Is Nothing) Then
                回路符号str = myPoint.Offset(0, 1)
                SetForegroundWindow myHNDtemp 'ステップインだと最前面にならないから注意
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '選択
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                'SendMessageStr myHNDtemp, WM_SETTEXT, 0&, 回路符号str
                Call input_Message(myHNDtemp, CStr(回路符号str))
            End If
            'Debug.Print Hex(myHNDtemp), s Mod 100
            'Call input_Message(myHNDtemp, CStr(回路符号str))
            'SendMessage myHNDtemp, &H281, &H1, &HC000000F
            'SendMessage myHNDtemp, &H281, &H0, &HC000000F
            'SendMessage myHNDtemp, &H1, 1, 0&
            'SendMessage myHNDtemp, BM_CLICK, 0, 0
            'bbb = PostMessage(myHNDtemp, WM_KEYUP, &HD, 0)
            PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTERで確定
            myHNDtemp = GetWindow(myHNDtemp, 3) '上のハンドル
            Sleep 50
        Next s
    End With
    Stop 'ここまで
    'W/Hデータ編集
    myHND(3) = FindWindowEx(myHND(2), 0&, "TfrmWH", "Ｗ／Ｈデータ編集")
    myHND(4) = GetWindow(myHND(3), 5)
    myHND(5) = GetWindow(myHND(4), 1)
    myHNDtemp = myHND(5)
    'WH編集に回路符号を与える
    pageMax = 60
    With wb(3).Sheets("WH")
        For s = 1 To endKosei
            Set mykosei = .Cells.Columns(1).Find(Format(s, "0000"), , , 1)
            If Not (mykosei Is Nothing) Then
                回路符号Astr = mykosei.Offset(0, 1)
                回路符号Bstr = mykosei.Offset(0, 2)
                SetForegroundWindow myHNDtemp 'ステップインだと最前面にならないから注意
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '選択
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                Call input_Message(myHNDtemp, CStr(回路符号Astr))
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTERで確定
                Sleep 50
                myHNDtemp = GetWindow(myHNDtemp, 3) '次のハンドル
                SetForegroundWindow myHNDtemp 'ステップインだと最前面にならないから注意
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '選択
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                Call input_Message(myHNDtemp, CStr(回路符号Bstr))
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTERで確定
                Sleep 50
            Else
                SetForegroundWindow myHNDtemp 'ステップインだと最前面にならないから注意
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '選択
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTERで確定
                Sleep 50
                myHNDtemp = GetWindow(myHNDtemp, 3) '次のハンドル
                SetForegroundWindow myHNDtemp 'ステップインだと最前面にならないから注意
                SendMessage myHNDtemp, WM_LBUTTONDOWN, 0&, 0& '選択
                SendMessage myHNDtemp, WM_LBUTTONUP, 0&, 0&
                PostMessage myHNDtemp, WM_KEYDOWN, &HD, 0 'ENTERで確定
                Sleep 50
            End If
            If s = pageMax Then
                myHNDtemp = myHND(5) '先頭のハンドル
                pageMax = pageMax + 60
            Else
                myHNDtemp = GetWindow(myHNDtemp, 3) '次のハンドル
            End If
        Next s
    End With
    
    Stop
    Stop
    
    a = SendMessage(myHND(0), WM_KEYDOWN, WM_ALT, 0)
    Call Control_Click(&H5&, "SEND", "ThunderRT6FormDC")
    Call Control_Click(&H2&, "SEND", "ThunderRT6FormDC")  '新伝送データ

    Call 後引張り支援システム_取り込み(myTextDir, ファイル名) '取込データファイル選択
    '待機
    '上書き確認、取込処理が完了しました。どちらか確認
    Do
        上書き確認 = 後引張り支援システム_処理確認("#32770", "処理確認", "はい(&Y)", &H6&)  '上書き確認
        
        取込処理完了 = 後引張り支援システム_処理確認("#32770", "処理確認", "OK", &H2&)  '取込処理が完了しました
        If 取込処理完了 <> 0 Then Exit Do
    Loop
    
    Call 後引張り支援システム_名前を付けて保存(&H1&, myTextDir & "\" & 管理ナンバー & "_KairoMat_3.txt")
    'エラーになるので再起動
    Call 後引張り支援システム_閉じる(&H1&)             '閉じる
    Sleep 1000
    Call Control_アプリ起動(後引張り支援システムPath)
    '部材所要量
    Call 後引張り支援システム_管理ナンバー選択(管理ナンバー)
    Call 後引張り支援システム_部材所要量表示
    Call 後引張り支援システム_名前を付けて保存(&H1&, myTextDir & "\" & 管理ナンバー & "_MRP.txt")
    Call 後引張り支援システム_閉じる2(&H1&)
    
    Set JJF = Nothing
    Set FSO = Nothing
    If myCount = 0 Then
        a = MsgBox("対象のファイルが見つからない為、処理が実行できませんでした。" & vbCrLf & _
                    "現在の場所に処理したいファイル(RFLT??-B?.txt)がある事を確認してから実行してください。" & vbCrLf & _
                    "" & vbCrLf & _
                    "現在の場所: " & ActiveWorkbook.Path, vbOKOnly, "PLUS+")
    Else
        a = MsgBox("処理が終了しました", vbOKOnly, "PLUS+")
    End If
End Sub

Public Sub 後引張り支援システム_部材所要量表示()
line10:
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "ThunderRT6Frame", "処理対象区分選択")
    If myHND(1) = 0& Then GoSub 取り込み待機
    'チェック入れる
    For i = 2 To 5
        myHND(2) = GetDlgItem(myHND(1), i)
        Rep = SendMessage(myHND(2), BM_SETCHECK, 1, 0)
    Next i
    '表示
    myHND(3) = GetDlgItem(myHND(0), &H8&)
    SendMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + &H8&, myHND(3)
    'ファイル出力
    myHND(4) = GetDlgItem(myHND(0), &H9&)
    PostMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + &H9&, myHND(4)
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub
Public Sub 後引張り支援システム_管理ナンバー選択(NMBナンバー)
line10:
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = FindWindowEx(myHND(0), 0&, "ThunderRT6Frame", "処理対象管理Ｎｏ選択")
    myHND(1) = FindWindowEx(myHND(0), 0&, "ThunderRT6TextBox", vbNullString)
    myHND(2) = FindWindowEx(myHND(0), 0&, "ThunderRT6ComboBox", vbNullString)
    myHND(3) = FindWindowEx(myHND(2), 0&, "Edit", vbNullString)
    If myHND(3) = 0& Then GoSub 取り込み待機

    Index = SendMessage(myHND(2), CB_GETCOUNT, 0, 0)
    For i = 0 To Index - 1
        myLen = SendMessage(myHND(2), CB_GETTEXTLEN, i, 0)
        Dim myStr As String: myStr = String(myLen, vbNullChar)
        Ret = SendMessageStr(myHND(2), CB_GETTEXT, i, myStr)
        Debug.Print myStr
        If NMBナンバー = myStr Then
            a = SendMessage(myHND(2), CB_SHOWDROPDOWN, 1, 0)
            b = SendMessage(myHND(2), CB_SETCURSEL, i, 0)
            c = SendMessage(myHND(2), WM_LBUTTONDOWN, i, 0)
            Exit For
        End If
    Next i
    
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Public Sub Control_アプリ起動(myPath)
line10:
        myHND(0) = FindWindow("TfrmMain", vbNullString)
        If myHND(0) = 0& Then GoSub アプリの起動
    Sleep 100
    Exit Sub
アプリの起動:
    
    On Error GoTo myErr
        ChDrive Left(myPath, 2)
        ChDir Left(myPath, InStrRev(myPath, "\") - 1)
        Shell myPath
    On Error GoTo 0
    
    myCount = myCount + 1: If myCount > 10 Then Stop
    GoTo line10
    
myErr:
    If Err.Number = 76 Or Err.Number = 53 Then
        MsgBox "シート[設定]のYcEditorのアドレスが正しくありません。" & vbCrLf & vbCrLf _
             & "YCEditor.exeの保存アドレスを確認して修正してください。"
    End If
    Sheets("設定").Activate
    Sheets("設定").Cells.Find("YcEditor_exe", , , 1).Offset(0, 1).Activate
    End
End Sub

Public Sub 後引張り支援システム_取り込み(取り込みファイルPath, ファイル名)
    'Driveに文字を渡す
    myDrive = Left(取り込みファイルPath, 1)
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = GetDlgItem(myHND(0), &H5&)
    myASC = Asc(myDrive)
    b = SendMessage(myHND(0), WM_IME_CHAR, myASC, 0)
    'ディレクトリの選択
    Dim myStr As String:
    temp = Split(取り込みファイルPath, "\")
    For i = LBound(temp) To UBound(temp)
        Dim myFolder As String
        If i = 0 Then myFolder = StrConv(temp(i), vbLowerCase) & "\" Else myFolder = temp(i)
        'Dirの参照
        myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
        myHND(0) = GetDlgItem(myHND(0), &H7&)
        Index = SendMessage(myHND(0), LB_GETCOUNT, 0, 0)
        For i2 = 0 To Index - 1
            myLen = SendMessage(myHND(0), LB_GETTEXTLEN, i2, 0)
            myStr = String(myLen, vbNullChar)
            Ret = SendMessageStr(myHND(0), LB_GETTEXT, i2, myStr)
            If StrConv(myStr, vbUpperCase) = StrConv(myFolder, vbUpperCase) Then
                c = SendMessage(myHND(0), LB_SETCURSEL, i2, 0)
                D = SendMessage(myHND(0), WM_LBUTTONDBLCLK, i2, 0)
                Exit For
            End If
        Next i2
line10:
    Next i
    'ファイルの選択
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = GetDlgItem(myHND(0), &H6&)
    Index = SendMessage(myHND(0), LB_GETCOUNT, 0, 0)
        For i3 = 0 To Index - 1
            myLen = SendMessage(myHND(0), LB_GETTEXTLEN, i3, 0)
            myStr = String(myLen, vbNullChar)
            Ret = SendMessageStr(myHND(0), LB_GETTEXT, i3, myStr)
            If myStr = ファイル名 Then
                c = SendMessage(myHND(0), LB_SETCURSEL, i3, 0)
                D = SendMessage(myHND(0), WM_LBUTTONDBLCLK, i3, 0)
            End If
        Next i3
    '取り込みボタンを押す
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(1) = GetDlgItem(myHND(0), &H9&)
    PostMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + &H9&, myHND(1)
Sleep 100
End Sub
Public Function 後引張り支援システム_処理確認(myClass, myCaption, myCaption2, myID)
    '上書きしてもいいですか?

    myHND(0) = FindWindow(myClass, myCaption)
    myHND(1) = FindWindowEx(myHND(0), 0&, vbNullString, myCaption2)
    後引張り支援システム_処理確認 = myHND(1)
    If myHND(1) = 0 Then Exit Function
    SendMessage myHND(0), WM_COMMAND, BN_CLICKED * &H10000 + myID, myHND(1)

End Function
Public Sub Control_Click(myHWND, 種類, CtrlID)
    Dim myDlg, myBtn, myStat As Long
    Dim myCount As Long
line10:
    SendMessage myHWND, BM_CLICK, 0, 0
    
    Exit Sub
    
    
    Stop
    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then Exit Sub
    Select Case 種類
        Case "SEND": SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
        Case "POST": PostMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
    End Select
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub

Public Sub 後引張り支援システム_クリック2(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    Dim myStr As String: myStr = String(15, vbNullChar)
line10:
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "ThunderRT6Frame", vbNullString)
    If myWHND = 0& Then GoSub 取り込み待機

    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then Exit Sub

    'クリック
    SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myStat = GetDlgItem(myWHND, &H6&)
    'ret = SendMessageStr(myStat, WM_GETTEXT, 15, myStr)
    Ret = SendMessage(myStat, LB_SETSEL, 1, -1)
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub


Public Sub 後引張り支援システム_クリック3(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    Dim myStr As String: myStr = String(15, vbNullChar)
line10:
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = FindWindowEx(myWHND, 0&, "ThunderRT6ListBox", vbNullString)
    If myWHND = 0& Then GoSub 取り込み待機

    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then Exit Sub

    'クリック
    SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn
    myWHND = FindWindow("ThunderRT6FormDC", vbNullString)
    myStat = GetDlgItem(myWHND, &H6&)
    'ret = SendMessageStr(myStat, WM_GETTEXT, 15, myStr)
    Ret = SendMessage(myStat, LB_SETSEL, 1, -1)
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub
Public Sub 後引張り支援システム_閉じる(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    
line10:
    myWHND = FindWindow("ThunderRT6FormDC", "後引張り支援システム")
    Rep = PostMessage(myWHND, WM_SYSCOMMAND, SC_CLOSE, 0)  '閉じる(続けて確認ウィンドウが開くので処理を待たないPOSTを使う)

    myWHND = FindWindow("#32770", "処理確認")
    If myWHND = 0& Then GoSub 取り込み待機
    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then GoSub 取り込み待機
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 1000 Then Stop 'postなので返り値が無い この数字大きくせないかんかも
GoTo line10
End Sub
Public Sub 後引張り支援システム_閉じる2(CtrlID)
    Dim myWHND, myDlg, myBtn, myStat As Long
    Dim myCount As Long
    
line10:
    
    myWHND = FindWindow("#32770", "部材所要量表示")
    If myWHND = 0& Then GoSub 取り込み待機
    myBtn = GetDlgItem(myWHND, &H2&)
    SendMessage myWHND, WM_COMMAND, BN_CLICKED * &H10000 + &H2&, myBtn
    
    myWHND = FindWindow("ThunderRT6FormDC", "後引張り支援システム")
    Rep = PostMessage(myWHND, WM_SYSCOMMAND, SC_CLOSE, 0)  '閉じる(続けて確認ウィンドウが開くので処理を待たないPOSTを使う)
Sleep 100
    myWHND = FindWindow("#32770", "部材所要量表示")
    If myWHND = 0& Then GoSub 取り込み待機
    myBtn = GetDlgItem(myWHND, CtrlID)
    If myBtn = 0& Then GoSub 取り込み待機
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Public Sub 後引張り支援システム_名前を付けて保存(CtrlID, 保存フルパス)
    Dim myWHND, myWHND2, myWHND3, lngRC, myDlg, myStat, myCount As Long
    Dim myText, myASC As String
line10:
    myWHND = FindWindow("#32770", "名前を付けて保存")
    myWHND = FindWindowEx(myWHND, 0&, "DUIViewWndClassName", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "DirectUIHWND", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "FloatNotifySink", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "ComboBox", vbNullString)
    myStat = FindWindowEx(myWHND, 0&, "Edit", vbNullString)
    If myWHND = 0& Then GoSub 取り込み待機
    
    For i = 1 To Len(保存フルパス)
        myText = Mid(保存フルパス, i, 1)
        myASC = Asc(myText)
        lngRC = SendMessage(myStat, WM_IME_CHAR, myASC, 0)
        Sleep 1
    Next i
    
    myWHND = FindWindow("#32770", "名前を付けて保存")
    myBtn = GetDlgItem(myWHND, CtrlID)
    'クリック
line20:
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
    If Rep <> 0 Then GoTo line20
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Sub Sample_コンボボックスの選択() 'Drive選択_FileListが切り替わらないので使用しない
    myHND(0) = FindWindow("ThunderRT6FormDC", vbNullString)
    myHND(0) = GetDlgItem(myHND(0), &H5&)
    
    Index = SendMessage(myHND(0), CB_GETCOUNT, 0, 0)
    
    For i = 0 To Index - 1
        myLen = SendMessage(myHND(0), CB_GETTEXTLEN, i, 0)
        Dim myStr As String: myStr = String(myLen, vbNullChar)
        Ret = SendMessageStr(myHND(0), CB_GETTEXT, i, myStr)
        a = SendMessage(myHND(0), CB_SETTOPINDEX, i, 0)
        b = SendMessage(myHND(0), CB_SETCURSEL, i, 0)
        c = SendMessage(myHND(0), BM_CLICK, i, 0)
    Next i
End Sub

Public Function input_Message(myStat, myMessage)
    
    For i = 1 To Len(myMessage)
        myText = Mid(myMessage, i, 1)
        myASC = Asc(myText)
        Sleep 10
        lngRC = SendMessage(myStat, WM_IME_CHAR, myASC, 0)
        If lngRC <> 0 Then Stop
    Next i
End Function

Public Sub Control_新規作成()
    Dim myWHND, myWHND2, myWHND3, lngRC, myDlg, myStat, myCount As Long
    Dim myText, myASC As String
line10:
    myWHND = FindWindow("TfrmFile", "ファイル選択")
    myWHND = FindWindowEx(myWHND, 0&, "TButton", "新規")
    myWHND = FindWindowEx(myWHND, 0&, "DirectUIHWND", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "FloatNotifySink", vbNullString)
    myWHND = FindWindowEx(myWHND, 0&, "ComboBox", vbNullString)
    myStat = FindWindowEx(myWHND, 0&, "Edit", vbNullString)
    If myWHND = 0& Then GoSub 取り込み待機
    
    For i = 1 To Len(保存フルパス)
        myText = Mid(保存フルパス, i, 1)
        myASC = Asc(myText)
        lngRC = SendMessage(myStat, WM_IME_CHAR, myASC, 0)
        Sleep 1
    Next i
    
    myWHND = FindWindow("#32770", "名前を付けて保存")
    myBtn = GetDlgItem(myWHND, CtrlID)
    'クリック
line20:
    Rep = SendMessage(myWHND, WM_COMMAND, BN_CLICKED * &H10000 + CtrlID, myBtn)
    If Rep <> 0 Then GoTo line20
Sleep 100
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 10 Then Stop
GoTo line10
End Sub

Public Function inputString(myStr)
    For e = 1 To Len(myStr)
        myText = Mid(myStr, e, 1)
        myASC = Asc(myText)
        lngRC = SendMessage(myHND(0), WM_IME_CHAR, myASC, 0)
    Next e
    Sleep 300
End Function

Public Sub clickButton(mySelect)
    Dim myCount As Long
    
line10:
    If myHND(0) = 0& Then GoSub 取り込み待機
    myHND(1) = FindWindow("ExToolBoxClass", vbNullString)
    If myHND(1) = 0& Then GoSub 取り込み待機
    Select Case mySelect
        Case "Enter"
            myHND(2) = FindWindowEx(myHND(1), 0&, "Button", "ENTER")
            myBtn = GetDlgItem(myHND(1), &H218)
            a = SendMessage(myBtn, BM_CLICK, 0, 0)
            Sleep 500
        Case "Home"
            myHND(2) = FindWindowEx(myHND(1), 0&, "Button", "HOME")
            myBtn = GetDlgItem(myHND(1), &H213)
            a = SendMessage(myBtn, BM_CLICK, 0, 0)
            Sleep 100
        Case "Pause"
            myHND(2) = FindWindowEx(myHND(1), 0&, "Button", "CLEAR")
            myBtn = GetDlgItem(myHND(1), &H217)
            a = SendMessage(myBtn, BM_CLICK, 0, 0)
        Sleep 100
    Case Else
        Stop
    End Select
    
Exit Sub
取り込み待機:
Sleep 300
myCount = myCount + 1: If myCount > 100 Then Stop
GoTo line10
End Sub

Sub Macro1()
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    Const WM_APP = &H8000

    Ret = PostMessage(myHND(0), WM_APP + 15620, 12, 20380001)
    
    SendMessage myHND(0), &H92, 0, &H19EE84
    SendMessage myHND(0), &H11F, &HFFFF0000, 0
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(2) = FindWindowEx(myHND(1), 0&, "TfrmSymbol", "シンボルデータ編集")
    myHND(3) = GetWindow(myHND(2), 5)
    myHND(4) = GetWindow(myHND(3), 1)
    myHNDtemp = myHND(4)
    PostMessage myHND(0), WM_IME_KEYDOWN, 164, 1
    PostMessage myHND(0), WM_IME_KEYDOWN, 70, 1
    PostMessage myHND(0), WM_IME_KEYUP, 70, 0
    PostMessage myHND(0), WM_IME_KEYUP, 164, 0
    AppActivate myHND(0)
    PostMessage myHND(0), &H106, VK_RMENU, 0
    Stop
    Stop
    For i = 1 To 13
        PostMessage myHND(0), WM_KEYDOWN, i, 0
        PostMessage myHND(0), WM_KEYUP, i, 0
    Next i
    Call input_Message(myHND(0), CStr(回路符号str))
End Sub

Sub test_notepad()
    myHND(0) = FindWindow("notepad", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "Edit", vbNullString)
    
    myHND(0) = FindWindow("TfrmMain", vbNullString)
    myHND(1) = FindWindowEx(myHND(0), 0&, "MDIClient", vbNullString)
    myHND(2) = FindWindowEx(myHND(1), 0&, "TfrmSymbol", "シンボルデータ編集")
    myHND(3) = GetWindow(myHND(2), 5)
    myHND(4) = GetWindow(myHND(3), 1)
    SetForegroundWindow myHND(0) 'デバッグでステップしたら最前面にならないから注意
    PostMessage myHND(0), WM_SYSKEYDOWN, VK_MENU, &H20380001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("F"), &H20210001
    PostMessage myHND(0), WM_SYSKEYDOWN, Asc("O"), &H20180001
End Sub
