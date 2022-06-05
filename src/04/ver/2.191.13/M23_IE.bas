Attribute VB_Name = "M23_IE"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'API_画像をダウンロード
Public Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    
Dim 部材詳細_タイトルRan As Range

Sub fajdlajf()

    Set area = Range("aa21:jk1142")
    
    For Each a In area
        If a.Value <> "" Or a.Value <> Empty Then
            If a.Borders(xlEdgeBottom).LineStyle <> xlContinuous Then
                a.Select
                Stop
            End If
        End If
    Next a

End Sub

Sub ie_通知書を取得(車種str)

    Dim 複線構成(1 To 20) As String
    Dim iD As String
    Dim myRyakuDir As String
    Dim mailURL(2) As String
    Dim 種類(2) As String

    With Sheets("設定")
        mailURL(0) = .Cells.Find("通知書アドレス_", , , 1).Offset(0, 1).Value
        mailURL(1) = .Cells.Find("通知書アドレス_", , , 1).Offset(1, 1).Value
        mailURL(2) = .Cells.Find("通知書アドレス_", , , 1).Offset(2, 1).Value
    End With
    種類(0) = "即"
    種類(1) = "設"
    種類(2) = "部"
    'マル即保管用のフォルダ
'    myRyakuDir = ActiveWorkbook.PAth & "\マル即"
'    If Dir(myRyakuDir, vbDirectory) = "" Then MkDir myRyakuDir
    'IEの起動
    Dim objIE As Object '変数を定義します
    Dim ieVerCheck As Variant
    Set objIE = CreateObject("InternetExplorer.Application")
    Set objSFO = CreateObject("Scripting.FileSystemObject")
'    Select Case Application.OperatingSystem
'    Case "Windows (32-bit) NT 6.01"
'        Set objIE = CreateObject("InternetExplorer.Application") 'オブジェクトを作成します。
'    Case Else
'        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
'    End Select
'    objIE.Visible = True
    
    ieVerCheck = Val(objSFO.GetFileVersion(objIE.FullName))
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10以降(たぶん)
    Else
        Set objIE = CreateObject("InternetExplorer.Application") '知らんけど
    End If
    
    objIE.Visible = True
    '上記で64-bit NT 6.01なのに32bitと判断される不具合の暫定対策
    On Error Resume Next
    objIE.Navigate mailURL(p)
    a = objIE.readystate
    b = objIE.busy
    Debug.Print Err.Number
    If Err.Number = -2147417848 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")
        objIE.Navigate mailURL(p)
    End If
    
    On Error GoTo 0
    
    '可視、Trueで見えるようにします
    '処理したいページを表示します
    '画面1 ログイン情報
'   objIE.document.all.Item(アカウントID).Value = アカウント
'   objIE.document.all.Item(パスID).Value = パス
'   objIE.document.all.Item("btnLogin").Click 'ログインクリック
'   Call ページ表示を待つ(objIE)
'   '画面2 使用注意情報
'   objIE.document.all.Item("btnOK").Click 'OKクリック
'   Call ページ表示を待つ(objIE)
'   '画面3 メインページ
'   objIE.document.all.Item("btnYzk").Click '矢崎品番からの検索
'   Call ページ表示を待つ(objIE)
'loop
   With ActiveSheet
        Dim key As Range: Set key = .Cells.Find("key_", , , 1)
        Dim key2 As Range: Set key2 = .Cells.Find("通知書№_", , , 1)
        Dim lastRow As Long: lastRow = .UsedRange.Rows.count
        Dim lastCol As Long: lastCol = .UsedRange.Columns.count
        Dim 通知書Row As Long: 通知書Row = key2.Row
        Dim 通知書Col As Long: 通知書Col = key2.Column
        Dim 日付Col As Long: 日付Col = .Rows(key2.Row).Find("日付_", , , 1).Column
        Dim 理由col As Long: 理由col = .Rows(key2.Row).Find("理由_", , , 1).Column
        Dim 変更要点col As Long: 変更要点col = .Cells.Find("変更要点_", , , 1).Column
        Dim 最終取得日Row As Long: 最終取得日Row = .Cells.Find("最終取得日", , , 1).Row
        Dim 車種Row As Long: 車種Row = .Cells.Find("車種_", , , 1).Row
        Dim 通知書 As String

        '通知書№の登録
        Dim 通知書RAN() As Variant, j As Long
        GoSub 通知書の登録

        '対象の製品品番の点数を計算
        For X = key.Offset(0, 1).Column To lastCol
            車種 = .Cells(車種Row, X)
            If 車種str = "" Or InStr(車種str, 車種) > 0 Then
                製品品番 = .Cells(key.Row, X)
                If 製品品番 <> "" Then
                    Total = Total + 1
                End If
            End If
        Next X
        
        Dim myText As String, mytext2 As String, myTextA As String, myTextTR As String
        Dim aa(6) As Long
        
        For X = key.Offset(0, 1).Column To lastCol
            For p = LBound(mailURL) To UBound(mailURL)
                車種 = .Cells(車種Row, X)
                If 車種str = "" Or InStr(車種str, 車種) > 0 Then
                    製品品番 = Replace(.Cells(key.Row, X), " ", "")
                    If 製品品番 <> "" Then
                        .Cells(key.Row, X).Select
                        '対象ページの表示
                        objIE.Navigate mailURL(p)
                        Call ページ表示を待つ(objIE)
                        '品番入力
                        Select Case p
                            Case 0, 1
                            objIE.document.all.Item("hinban").Value = 製品品番
                            hensu = 2
                            Case 2
                            objIE.document.all.Item("s_hinban").Value = 製品品番
                            hensu = 2
                        End Select
                        Call ページ表示を待つ(objIE)
                        
                        '検索クリック
                        Call ボタンクリック(objIE, "検索")
                        Call ページ表示を待つ(objIE)
                        
                        '画面情報の取得
                        For i = 0 To objIE.document.getElementsByTagName("tr").Length - hensu
                            Select Case p
                                Case 0
                                    myText = objIE.document.getElementsByTagName("tr")(i + 1).outerhtml
                                    If InStr(StrConv(myText, vbUpperCase), "HREF") > 0 Then
                                        URL = objIE.document.getElementsByTagName("a")(i - URL無いcount).href
                                        c = 0
                                    Else
                                        URL = ""
                                        c = 1
                                        URL無いcount = URL無いcount + 1
                                    End If
                                    a = 検索(myText, ">", 3 - c)
                                    b = 検索(myText, "<", 4 - c)
                                    通知書 = Mid(myText, a + 1, b - a - 1)
                                    a = 検索(myText, ">", 6 - c - c)
                                    b = 検索(myText, "<", 7 - c - c)
                                    日付 = CDate(Mid(myText, a + 1, b - a - 1))
                                    a = 検索(myText, ">", 10 - c - c)
                                    b = 検索(myText, "<", 11 - c - c)
                                    理由 = Mid(myText, a + 1, b - a - 1)
                                    a = 検索(myText, ">", 14 - c - c)
                                    b = 検索(myText, "<", 15 - c - c)
                                    設変 = Mid(myText, a + 1, b - a - 1)
                                    部品 = ""
                                Case 1
                                    myText = objIE.document.getElementsByTagName("tr")(i + 1).outerhtml
                                    'PDFのリンク先有無の確認_無い時がある→http://10.7.1.35/nim_intra/70_busyobetsu/30_sekkei/program/hentu/hinban_result.asp
                                    If InStr(StrConv(myText, vbUpperCase), "HREF") > 0 Then
                                        URL = objIE.document.getElementsByTagName("a")(i - URL無いcount).href
                                        c = 0
                                    Else
                                        URL = ""
                                        c = 1
                                        URL無いcount = URL無いcount + 1
                                    End If
                                    a = 検索(myText, ">", 3 - c)
                                    b = 検索(myText, "<", 4 - c)
                                    通知書 = Mid(myText, a + 1, b - a - 1)
                                    a = 検索(myText, ">", 6 - c - c)
                                    b = 検索(myText, "<", 7 - c - c)
                                    日付 = CDate(Mid(myText, a + 1, b - a - 1))
                                    理由 = "設計変更"
                                    a = 検索(myText, ">", 12 - c - c)
                                    b = 検索(myText, "<", 13 - c - c)
                                    設変 = Mid(myText, a + 1, b - a - 1)
                                    部品 = ""
                                Case 2
                                    myText = objIE.document.getElementsByTagName("tr")(i + 1).outerhtml
                                    If InStr(StrConv(myText, vbUpperCase), "HREF") > 0 Then
                                        
                                        URL = objIE.document.getElementsByTagName("a")(i - URL無いcount).href
                                        c = 0
                                    Else
                                        URL = ""
                                        c = 1
                                        URL無いcount = URL無いcount + 1
                                    End If
                                    a = 検索(myText, ">", 3 - c)
                                    b = 検索(myText, "<", 4 - c)
                                    通知書 = Mid(myText, a + 1, b - a - 1)
                                    a = 検索(myText, ">", 6 - c - c)
                                    b = 検索(myText, "<", 7 - c - c)
                                    日付 = CDate(Mid(myText, a + 1, b - a - 1))
                                    理由 = "部品変更"
                                    a = 検索(myText, ">", 14 - c - c)
                                    b = 検索(myText, "<", 15 - c - c)
                                    設変 = Mid(myText, a + 1, b - a - 1)
                                    a = 検索(myText, ">", 10 - c - c)
                                    b = 検索(myText, "<", 11 - c - c)
                                    部品 = Mid(myText, a + 1, b - a - 1)
                            End Select
                            
                            addRow = 0
                            '登録してるか確認
                            flg = False
                            For r = LBound(通知書RAN, 2) To UBound(通知書RAN, 2)
                                If 通知書 = 通知書RAN(0, r) And 種類(p) = 通知書RAN(2, r) Then
                                    addRow = 通知書RAN(1, r)
                                    Exit For
                                End If
                            Next r
                            
                            '無い場合登録
                            If addRow = 0 Then
                                flg = True
                                For r = LBound(通知書RAN, 2) To UBound(通知書RAN, 2)
                                    If 日付 < 通知書RAN(3, r) Then
                                        addRow = 通知書RAN(1, r)
                                        .Rows(addRow).Insert
                                        .Range(.Cells(key2.Row + 1, 1), .Cells(key2.Row + 1, key.Column)).Copy .Range(.Cells(addRow, 1), .Cells(addRow, key.Column))
                                        .Range(.Cells(addRow, 1), .Cells(addRow, key.Column)).ClearContents
                                        .Range(.Cells(addRow, key.Column + 1), .Cells(addRow, .Columns.count)).ClearFormats
                                        Exit For
                                        Stop
                                    End If
                                Next r
                            End If
                            
'                            '無い場合登録
'                            If addRow = 0 Then
'                                addRow = .Cells(.Rows.Count, key2.Column).End(xlUp).Row + 1
'                                ReDim Preserve 通知書RAN(3, UBound(通知書RAN, 2) + 1)
'                                通知書RAN(0, UBound(通知書RAN, 2)) = 通知書
'                                通知書RAN(1, UBound(通知書RAN, 2)) = addRow
'                                通知書RAN(2, UBound(通知書RAN, 2)) = 種類(p)
'                                .Cells(addRow, x).Select
'                            End If
                            '出力
                            If addRow = 0 Then
                                addRow = .Cells(.Rows.count, key2.Column).End(xlUp).Row + 1
                            End If
                            .Cells(addRow, key2.Column + 0) = 通知書
                            .Cells(addRow, key2.Column - 1) = 種類(p)
                            .Cells(addRow, key2.Column).NumberFormat = "@"
                            If URL <> "" Then
                                .Hyperlinks.add anchor:=.Cells(addRow, key2.Column), address:=URL, ScreenTip:="", TextToDisplay:=CStr(通知書)
                            Else
                                .Cells(addRow, key2.Column).Font.Underline = False
                            End If
                            Select Case p
                                Case 0
                                .Cells(addRow, key2.Column).Font.color = RGB(0, 0, 255)
                                .Cells(addRow, 理由col).Font.color = RGB(0, 0, 255)
                                .Cells(addRow, 変更要点col).Font.color = RGB(0, 0, 0)
                                設変 = Left(設変, 1) & Mid(設変, 3, 1) & Mid(設変, 5, 1)
                                Case 1
                                .Cells(addRow, key2.Column).Font.color = RGB(255, 0, 255)
                                .Cells(addRow, 理由col).Font.color = RGB(255, 0, 255)
                                .Cells(addRow, 変更要点col).Font.color = RGB(0, 0, 0)
                                設変 = Left(設変, 1) & Mid(設変, 3, 1) & Mid(設変, 5, 1)
                                Case 2
                                .Cells(addRow, key2.Column).Font.color = RGB(0, 100, 0)
                                .Cells(addRow, 理由col).Font.color = RGB(0, 100, 0)
                                .Cells(addRow, 変更要点col).Font.color = RGB(0, 100, 0)
                                .Cells(addRow, 変更要点col) = CStr(部品)
                                設変 = Left(設変, 1) & Mid(設変, 3, 1) & Mid(設変, 5, 1)
                            End Select
                            
                            .Cells(addRow, 理由col) = 理由
                            .Cells(addRow, 日付Col).NumberFormat = "yy/mm/dd"
                            .Cells(addRow, 日付Col) = 日付
                            .Cells(addRow, X).NumberFormat = "@"
                            .Cells(addRow, X).HorizontalAlignment = xlCenter
                            .Cells(addRow, X).VerticalAlignment = xlCenter
                            .Cells(addRow, X).Font.Bold = True
                            .Cells(addRow, X) = 設変
                            .Cells(addRow, X).Borders.Weight = xlThin
                            .Cells(addRow, X).Select
                            .Rows(addRow).RowHeight = 27
                            If flg = True Then GoSub 通知書の登録
                            
                        Next i
                        .Cells(最終取得日Row, X) = Date
                        If p = 0 Then
                            onetime = DateDiff("s", mytime, Time)
                            totaltime = totaltime + onetime
                            count = count + 1
                            counttime = totaltime / count
                            Application.StatusBar = "  " & count & "/" & Total & "  残り: " & Int(((Total - count) * counttime) / 60)
                            mytime = Time
                        End If
                     End If
                 End If
                 URL無いcount = 0
            Next p
        Next X
        '並び替えしたら書式がずれるので並び替えをしない
'        Stop
'        addRow = .Cells(Rows.Count, key2.Column).End(xlUp).Row
'        With .Sort.SortFields
'            .Clear
'            .Add key:=Range(Cells(key2.Row, 日付Col).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
''            .Add key:=Range(Cells(1, 優先2).Address), Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'        End With
'        .Sort.SetRange .Range(.Rows(key2.Row), .Rows(addRow))
'        .Sort.Header = xlYes
'        .Sort.MatchCase = False
'        .Sort.Orientation = xlTopToBottom
'        .Sort.Apply
'        '.Rows(key2.Row + 1 & ":" & addRow).Sort key1:=.Cells(key2.Row, 日付Col), order1:=xlAscending
        Application.StatusBar = False
        
        objIE.Quit
        Set objIE = Nothing
    End With
    
    MsgBox "更新が完了しました。"
    
Exit Sub

通知書の登録:

        ReDim 通知書RAN(3, 0): j = 0
        With ActiveSheet
            lastRow = .UsedRange.Rows.count
            For ii = key2.Row + 1 To lastRow
                If .Cells(ii, 通知書Col) <> "" Then
                    ReDim Preserve 通知書RAN(3, j)
                    通知書RAN(0, j) = .Cells(ii, 通知書Col)
                    通知書RAN(1, j) = ii
                    通知書RAN(2, j) = .Cells(ii, 通知書Col - 1)
                    通知書RAN(3, j) = .Cells(ii, 日付Col)
                    j = j + 1
                End If
            Next ii
        End With
Return

End Sub

Public Function a取得_コネクタ類_コネクタ極数(ByVal objIE As Object, iD, コネクタ極数)
  コネクタ極数 = ""
    検索文字 = "コネクタ極数"
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    コネクタ極数 = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a取得_略図ダウンロード(ByVal objIE As Object, myRyakuDir, 部品品番)
    検索文字 = "<img src="
    For a = 0 To 1
        '略図のボタンidが無ければ処理しない
        If InStr(objIE.document.all(0).outerhtml, "ctl01_dispRyaku_btnDraw") = 0 Then Exit Function
        objIE.document.all.Item("ctl01_dispRyaku_edtText").Value = 部品品番
        objIE.document.all.Item("ctl01_dispRyaku_rgpReverse_" & a).Click      '0=正面視 1=裏面視
        objIE.document.all.Item("ctl01_dispRyaku_cmbText")(3).Selected = True 'テキスト入力
        objIE.document.all.Item("ctl01_dispRyaku_chkOriginalSize").Checked = True     '描画
        objIE.document.all.Item("ctl01_dispRyaku_btnDraw").Click              '描画
        
        Call ページ表示を待つ(objIE)
        For X = 0 To objIE.document.all.tags("img").Length - 1  '要素の数
            データ = objIE.document.all.tags("img")(X).outerhtml
            aaa = InStr(データ, 検索文字)
            If aaa = 0 Then GoTo line0
            略図URL = "http://10.1.33.95/DesignSource" & Mid(データ, Len(検索文字) + 3)
            略図URL = Left(略図URL, Len(略図URL) - 2)
            略図保存PASS = myRyakuDir & "\" & 部品品番 & "_" & a & "_" & Format(X, "000") & ".emf"
            'ダウンロードの実行
            Ret = URLDownloadToFile(0, 略図URL, 略図保存PASS, 0, 0)
line0:
        Next X
    Next a
End Function
Public Function a取得_得意先品番(ByVal objIE As Object, iD, ByVal i As Long)
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML
    On Error GoTo 0
    Dim データs As Variant
    Dim タイトルAddCol As Long
    データs = Split(データ, vbLf)
    For Each データo In データs
        a = InStr(データo, "<th"): If a <> 0 Then GoTo line10
        aa = InStr(データo, Chr(34) & ">"): If aa = 0 Then GoTo line10
        aaa = Mid(データo, aa + 2)
        bb = InStr(aaa, "<"): If bb = 0 Then GoTo line10
        bbb = Left(aaa, bb - 1)
        cc = InStr(aaa, Chr(34) & ">"): If cc = 0 Then GoTo line10
        ccc = Mid(aaa, cc + 2)
        dd = InStr(ccc, "<"): If dd = 0 Then GoTo line10
        ddd = Left(ccc, dd - 1)
        得意先名 = Replace(bbb, "&nbsp;", "")
        得意先品番 = Replace(ddd, "&nbsp;", "")
    '部材詳細から探して項目が無ければ追加
    With Sheets("部材詳細")
        Set 得意先名find = 部材詳細_タイトルRan.Find(得意先名 & "_", LookAt:=xlWhole)
        If 得意先名find Is Nothing Then
            Dim タイトルRow As Long: タイトルRow = 部材詳細_タイトルRan.Row
             タイトルAddCol = .Cells(タイトルRow, .Columns.count).End(xlToLeft).Column + 1
            .Cells(タイトルRow - 1, タイトルAddCol) = "得意先名"
            .Cells(タイトルRow, タイトルAddCol) = 得意先名 & "_"
        Else
            タイトルAddCol = 得意先名find.Column
        End If
            .Cells(i, タイトルAddCol).NumberFormat = "@"
            .Cells(i, タイトルAddCol) = 得意先品番
    End With
line10:
    Next
End Function

Public Function a取得_チューブ外径(ByVal objIE As Object, iD, チューブ外径)
  チューブ外径 = ""
    検索文字 = "チューブ外径"
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    チューブ外径 = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a取得_チューブ内径(ByVal objIE As Object, iD, チューブ内径)
  チューブ内径 = ""
    検索文字 = "チューブ内径"
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    チューブ内径 = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a取得_チューブ長さ(ByVal objIE As Object, iD, チューブ長さ)
  チューブ長さ = ""
    検索文字 = "チューブ長さ"
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    チューブ長さ = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function

Public Function a取得_チューブ品種(ByVal objIE As Object, iD, チューブ品種)
  チューブ品種 = ""
    検索文字 = "チューブ品種"
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    チューブ品種 = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a取得_クランプタイプ(ByVal objIE As Object, クランプタイプ)
  クランプタイプ = ""
  
    検索文字 = "クランプタイプ"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdPtmIndivs").outertext 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = Mid(ccc, Len(検索文字) + 1)
    クランプタイプ = Replace(ddd, vbLf, "")
    
End Function

Public Function a取得_仕上がり外径(ByVal objIE As Object, 仕上がり外径)
  仕上がり外径 = ""
  
    検索文字 = "仕上がり外径"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdPtmIndivs").innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ";")
    ccc = InStr(bbb + 1, データ, ";")
    ddd = InStr(ccc + 1, データ, ";")
    eee = InStr(ddd + 1, データ, ">")
    zzz = InStr(eee + 1, データ, "<")
    仕上がり外径 = Mid(データ, eee + 1, zzz - eee - 1)
      
End Function

Public Function ページ表示を待つ(ByRef objIE As Object)

    While objIE.readystate <> 4 Or objIE.busy = True '.ReadyState <> 4の間まわる。
        DoEvents  '重いので嫌いな人居るけど。
        Sleep 1
        'Call 仮想キー入力(シフト)
    Wend
    
End Function

Public Function a取得_略図(ByVal objIE As Object, 略図URL, 略図数)
  略図URL = "": 略図数 = 0
  
    略図数 = objIE.document.Images.Length - 1
  
    For r = 1 To objIE.document.Images.Length - 1
  
        略図URL = objIE.document.Images(1).src
    Next r
  
      
End Function

Public Function a取得_部品種別(ByVal objIE As Object, 部品種別)
  部品種別 = ""
  
    検索文字 = "部品種別"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(ccc + 1, データ, "<")
    
    If aaa <> 0 Then 部品種別 = Mid(データ, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a取得_部品分類(ByVal objIE As Object, 部品分類)
  部品分類 = ""
  
    検索文字 = "部品分類"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(ccc + 1, データ, "<")
    
    If aaa <> 0 Then 部品分類 = Mid(データ, ccc + 1, zzz - ccc - 1)
      
End Function
Public Function a取得_部品名称(ByVal objIE As Object, 部品名称)
  部品名称 = ""
  
    検索文字 = "部品名称"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(ccc + 1, データ, "<")
    
    If aaa <> 0 Then 部品名称 = Mid(データ, ccc + 1, zzz - ccc - 1)
      
End Function
Public Function a取得_登録工場(ByVal objIE As Object, 登録工場)
  登録工場 = ""
  
    検索文字 = "登録工場"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(ccc + 1, データ, "<")
        
    If aaa <> 0 Then 登録工場 = Mid(データ, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a取得_名称(ByVal objIE As Object, 名称品名)
  名称品名 = "": データ = ""
  
    検索文字 = "名称"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdEmtrCommon").innerHTML 'JAIRS情報
    On Error GoTo 0
    
    If データ = "" Then
        検索文字 = "品名"
        On Error Resume Next
        データ = objIE.document.getElementById("ctl01_grdJairsCommon").innerHTML 'JAIRS情報
        On Error GoTo 0
    End If
    
    If データ = "" Then Stop '上記のどちらも見つからない
        
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(ccc + 1, データ, "<")
        
    If aaa <> 0 Then 名称品名 = Mid(データ, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a取得_部品色(ByVal objIE As Object, 部品色)
  部品色 = "": データ = ""
  
    検索文字 = "色"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdJairsSpecs").innerHTML 'JAIRSの仕様
    On Error GoTo 0
        
    'If データ = "" Then Stop '上記のどちらも見つからない
        
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(ccc + 1, データ, "<")
        
    If aaa <> 0 Then 部品色 = Mid(データ, ccc + 1, zzz - ccc - 1)
      
End Function

Public Function a取得_重量(ByVal objIE As Object, 重量)
  重量 = ""
  
    検索文字 = "重量"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdJairsSize").innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(ccc, "</td>")
    eee = Left(ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    重量 = Mid(eee, fff + 1, Len(eee) - fff + 1)
      
End Function

Public Function a取得_複線構成(ByVal objIE As Object, 複線構成)
  
    検索文字 = "追番"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdEmtrComp").innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, Len(検索文字) + aaa + 1, Len(データ))
    
    For i = 1 To 20
        ccc = InStr(bbb, "target")
        If ccc = 0 Then 複線構成(i) = "": GoTo line10
        ddd = Mid(bbb, ccc, Len(bbb))
        eee = InStr(ddd, ">")
        fff = InStr(ddd, "<")
        ggg = Mid(ddd, eee + 1, fff - eee - 1)
        複線構成(i) = ggg
        
        bbb = Mid(bbb, ccc + fff, Len(bbb))
line10:
    Next i
          
End Function

Public Function a取得_検索結果(ByVal objIE As Object, 検索結果, ByVal 部品品番)
    検索結果 = ""
    
    Dim リンク番号 As Long
    'NotFound確認
    データ = objIE.document.getElementById("ctl00_lblErrMsg").innerHTML
    検索結果 = データ
    If 検索結果 = "Not Found." Then Exit Function
    
    '見つかった点数を確認
    データ = objIE.document.getElementById("ctl00_grdList").innerHTML
    aaa = InStrRev(データ, "grdList")
    bbb = Mid(データ, aaa + 8, 100)
    zzz = InStr(bbb, "'")
    点数 = Mid(データ, aaa + 8, zzz - 1)
    
    '点数が複数ある場合、リンクをクリック
    If 点数 > 0 Then
    'リンクaaa = InStrRev(データ, ">" & Replace(部品品番, "-", "") & "<")
    'リンクbbb = Left(データ, リンクaaa)
    'リンクccc = InStrRev(リンクbbb, "grdList")
    'リンクアドレス = Mid(リンクbbb, リンクccc, 9 + Len(点数))
    'objIe.document.all.Item("javascript:__doPostBack('ctl00$grdList','grdList$0')").Click
    
    'リンク番号で開く(点数+4で検索する為、確実ではないかも)
    リンクaaa = InStrRev(データ, ">" & Replace(部品品番, "-", "") & "<")
    If リンクaaa <> 0 Then
        リンクbbb = Left(データ, リンクaaa)
        リンクccc = InStrRev(リンクbbb, "$")
        リンクzzz = InStrRev(リンクbbb, "'")
        リンク番号 = Mid(リンクbbb, リンクccc + 1, リンクzzz - (リンクccc + 1))
    Else
        検索結果 = "NotMatch"
    End If
    
    objIE.document.Links(4).Click
    
    End If
    
    Call ページ表示を待つ(objIE)
        '表示された品番と検索したい品番がマッチするか確認
        データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML
        aaa = InStr(データ, "ＹＢＭコード")
        aaa以下 = Mid(データ, aaa + 1, Len(データ) - aaa)
        bbb = InStr(aaa以下, ">")
        bbb以下 = Mid(aaa以下, bbb + 1, Len(aaa以下) - bbb)
        ccc = InStr(bbb以下, ">")
        ccc以下 = Mid(bbb以下, ccc + 1, Len(bbb以下) - ccc)
        zzz = InStr(ccc以下, "<")
        表示された部品品番 = Left(ccc以下, zzz - 1)
        '表示された部品品番 = ObjIE.Document.all.Item("ctl00_txtYbm").Value
        '表示された部品品番 = Replace(表示された部品品番, "%", "")
        '表示された部品品番 = Replace(表示された部品品番, "-", "")
        
        If Replace(表示された部品品番, "-", "") <> Replace(部品品番, "-", "") Then
            '検索した品番と表示された品番の照合
            If Replace(表示された部品品番, "-", "") Like "*" & Replace(部品品番, "-", "") Then
                検索結果 = "Found"
            Else
                Stop '検索した品番と表示された品番の後半が異なる
            End If
        Else
                検索結果 = "Found"
        End If
    
End Function

Public Function ボタンクリック(ByRef objIE As Object, buttonValue As String) '不要かも
    Dim objInput As Object
    
    For Each objInput In objIE.document.getElementsByTagName("input")
        If objInput.Value = buttonValue Then
            objInput.Click
            Exit For
        End If
    Next
End Function

Public Function 画面情報取得a(ByVal objIE As Object) '不要かも

Dim 改行数 As Long

    'ObjIE.document.getElementsByName("q")(0).Value = "あいう"
  For Each obj In objIE.document.all  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    With Sheets("ログ")
        nextGyo = .Range("a" & .Rows.count).End(xlUp).Row + 1
        値 = obj.innertext
        'Call 改行の回数を調べる(値, 改行数)
        'For a = 1 To 改行数
        .Range("a" & nextGyo) = 値
        .Range("b" & nextGyo) = "ID=" & obj.iD
        'Next a
    End With
  Next
  
End Function

Public Function 画面情報取得(ByVal objIE As Object) '不要かも

    'ObjIE.document.getElementsByName("q")(0).Value = "あいう"
  For Each obj In objIE.document.getElementsByTagName("a")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    Sheets("ログ").Range("a" & Sheets("ログ").Range("a" & Rows.count).End(xlUp).Row + 1) = "a_innertext=" & obj.innertext & "  " & "ID=" & obj.iD           'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
  Next
  
  For Each obj In objIE.document.getElementsByTagName("input")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    Sheets("ログ").Range("a" & Sheets("ログ").Range("a" & Rows.count).End(xlUp).Row + 1) = "input_innertext=" & obj.innertext & "  " & "ID=" & obj.iD           'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
  Next
  
  For Each obj In objIE.document.getElementsByTagName("btn")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    Sheets("ログ").Range("a" & Sheets("ログ").Range("a" & Rows.count).End(xlUp).Row + 1) = "btn_innertext=" & obj.innertext & "  " & "ID=" & obj.iD & " " & obj.Name         'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
  Next

End Function

Sub IE_open_sample() '参考
  
  j = 0
  
  Set objIE = CreateObject("InternetExplorer.Application")  'IEを開く際のお約束
  objIE.Visible = True                                      'IEを開く際のお約束
  objIE.Navigate "http://www.yahoo.co.jp/"                  '開きたいサイトのURLを指定
  
  Do While objIE.readystate <> 4                            'サイトが開かれるまで待つ（お約束）
    Do While objIE.busy = True                              'サイトが開かれるまで待つ（お約束）
    Loop
  Loop
  
  For Each obj In objIE.document.getElementsByTagName("a")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    If obj.innertext = "ファイナンス" Then                  'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
      obj.Click                                             '上記に該当するタグをクリック
      Exit For                                              '上記処理後、For Each　～　Nextを抜ける
    End If
  Next                                                      '次のタグを処理

  Sleep (1000)                                              '1秒待つ
  
  Do While objIE.readystate <> 4                            'サイトが開かれるまで待つ（お約束）
    Do While objIE.busy = True                              'サイトが開かれるまで待つ（お約束）
    
    Loop
  Loop
  
  For Each obj In objIE.document.getElementsByTagName("input")  '表示されているサイトのinputタグ一つずつを変数objにセット
                                                                '各inputタグ単位に以下の処理を実施
    If obj.iD = "searchText" Then                           'タグのid名が「searchText」の場合、以下の処理を実施
      obj.Value = "任天堂"                                  'テキストボックスに「任天堂」を挿入
    Else
      If obj.iD = "searchButton" Then                       'タグのid名が「searchButton」の場合、以下の処理を実施
        obj.Click                                           '該当のinputタグをクリック
        Exit For                                            '上記処理後、For Each　～　Nextを抜ける
      End If
    End If
  Next                                                      '次のタグを処理

End Sub

Public Function 検索a(ByVal objIE As Object, 検索文字, エレメント)

    On Error Resume Next
    データ = objIE.document.getElementById(エレメント).innerHTML 'PTM情報
    On Error GoTo 0
    aa = 検索(データ, 検索文字, 1)
    If aa = 0 Then Exit Function
    データa = Mid(データ, aa)
    bb = 検索(データa, "<", 3)
    データb = Left(データa, bb - 1)
    cc = InStrRev(データb, ">")
    検索a = Mid(データb, cc + 1)
    検索a = Replace(検索a, "&nbsp;", "")
      
End Function
Public Function 検索(ソース, 検索文字, ヒット数)
    Dim myCount As Long
    For i = 1 To Len(ソース)
        If 検索文字 = Mid(ソース, i, Len(検索文字)) Then
            myCount = myCount + 1
            If ヒット数 = myCount Then
                検索 = i
                Exit Function
            End If
        End If
    Next i
    
End Function
