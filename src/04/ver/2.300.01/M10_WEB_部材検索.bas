Attribute VB_Name = "M10_WEB_部材検索"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'API_画像をダウンロード
Public Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    
Dim 部材詳細_タイトルRan As Range

Function ieVerCheck() As Integer

  Set objIEA = CreateObject("InternetExplorer.Application")
  Set objSFO = CreateObject("Scripting.FileSystemObject")

  ieVerCheck = val(objSFO.GetFileVersion(objIEA.FullName))

  Set objIEA = Nothing
  Set objSFO = Nothing

End Function

Public Sub open_dsw()
Attribute open_dsw.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim 複線構成(1 To 20) As String
    Dim iD As String
    Dim myRyakuDir As String

    addressSet ThisWorkbook

    Dim gyo As Long: gyo = 10
    
    With wb(0).Sheets("WEB")
        アカウント = .Range("c" & gyo)
        アカウントID = .Range("d" & gyo)
        パス = .Range("e" & gyo)
        パスID = .Range("f" & gyo)
        ログインbtn = .Range("g" & gyo)
        アドレスstr = .Range("h" & gyo)
        ウィンドウ名 = .Range("i" & gyo)
        ブラウザ = .Range("j" & gyo)
    End With
'
'    With Sheets("A0_部材詳細")
'        Dim 部材詳細_タイトルRow As Long: 部材詳細_タイトルRow = .Cells.Find("部品品番_").Row
'        Set 部材詳細_タイトルRan = .Range(.Cells(部材詳細_タイトルRow, 1), .Cells(部材詳細_タイトルRow, .Columns.count))
'
'        タイトル文字 = "検索結果_,部品種別_,部品分類_,名称・品名_,色_,登録工場_,重量_,仕上がり外径_,略図数,略図URL,クランプタイプ_,チューブ品種_,チューブ内径_,コネクタ極数_,部品名称_,複線構成01,区分_,部品品番_,備考_,検索結果_,コネクタ色_,防水区分_,ロック位置寸法_,ロック方向区分_,端子一体型区分_,メッキ区分_,ファミリー_,オスメス_,チューブ外径_,チューブ長さ_"
'        タイトル文字s = Split(タイトル文字, ",")
'        'もしタイトル文字が無ければ追加
'        Dim addCol As Long, checkTitle As Variant, x As Long
'        For x = LBound(タイトル文字s) To UBound(タイトル文字s)
'            Set checkTitle = .Cells.Find(タイトル文字s(x), , , 1)
'            If checkTitle Is Nothing Then
'                addCol = .Cells(部材詳細_タイトルRow, .Columns.count).End(xlToLeft).Column + 1
'                .Cells(部材詳細_タイトルRow, addCol).Value = タイトル文字s(x)
'            End If
'        Next x
'
'        Dim myCol() As Long
'        ReDim myCol(UBound(タイトル文字s))
'        For i = LBound(タイトル文字s) To UBound(タイトル文字s)
'            myCol(i) = 部材詳細_タイトルRan.Find(タイトル文字s(i), , , 1).Column
'        Next i
'    End With
    
    '略図のダウンロード用のフォルダ
    If Dir(myRyakuDir, vbDirectory) = "" Then MkDir myRyakuDir
    'IEの起動
    Dim objIE As Object '変数を定義します。
    Dim ieVerCheck As Variant

    Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
    Set objSFO = CreateObject("Scripting.FileSystemObject")

    ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
    
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10以降(たぶん)
    End If
    
    objIE.Visible = True      '可視、Trueで見えるようにします。
    
    '処理したいページを表示します。
   objIE.Navigate アドレスstr
   Call ページ表示を待つ(objIE)
  
   '画面1 ログイン情報
   objIE.document.all.Item(アカウントID).Value = アカウント
   objIE.document.all.Item(パスID).Value = パス
   objIE.document.all.Item("btnLogin").Click 'ログインクリック
   Call ページ表示を待つ(objIE)
   '画面2 使用注意情報
   objIE.document.all.Item("btnOK").Click 'OKクリック
   Call ページ表示を待つ(objIE)
   '画面3 メインページ
   objIE.document.all.Item("btnYzk").Click '矢崎品番からの検索
   Call ページ表示を待つ(objIE)
'loop
    
    Set ws(0) = wb(0).ActiveSheet
   With ws(0)
         部品品番 = .Cells(ActiveCell.Row, .Cells.Find("部品品番", , , 1).Column).Value
'        lastgyo = .Cells(.Rows.count, myCol(17)).End(xlUp).Row
'        For i = 6 To lastgyo
'            If .Cells(i, myCol(19)) = "" Then
'                区分 = .Cells(i, myCol(16))
'                If Len(区分) = 1 Then
'                    部品品番 = .Cells(i, myCol(17))
                    '品番入力
                    objIE.document.all.Item("ctl00_txtYbm").Value = "%" & 部品品番 & "%"
                    Call ページ表示を待つ(objIE)
                    '検索クリック
                    objIE.document.all.Item("ctl00_btnSearch").Click
                    Call ページ表示を待つ(objIE)
                    '品番情報の取得
                    Call a取得_検索結果(objIE, 検索結果, 部品品番)
                    
                    If 検索結果 = "Not Found." Then
'                        .Cells(i, myCol(19)) = "NotFound"
                    ElseIf 検索結果 = "NotMatch" Then
'                        .Cells(i, myCol(19)) = "NotMatch"
                    Else
                        'PTM
                        部品種別 = 検索a(objIE, "部品種別", "ctl01_grdPtmCommn")
                        部品分類 = 検索a(objIE, "部品分類", "ctl01_grdPtmCommn")
                        部品名称 = 検索a(objIE, "部品名称", "ctl01_grdPtmCommn")
                        登録工場 = 検索a(objIE, "登録工場", "ctl01_grdPtmCommn")
                        'JAIRS
                        名称品名 = 検索a(objIE, "名称", "ctl01_grdEmtrCommon")
                        If 名称品名 = "" Then 名称品名 = 検索a(objIE, "品名", "ctl01_grdJairsCommon")
                        
                        部品色 = 検索a(objIE, "色", "ctl01_grdJairsSpecs")
                        ファミリー = 検索a(objIE, "ファミリー", "ctl01_grdJairsSpecs")
                        オスメス = 検索a(objIE, "オス/メス", "ctl01_grdJairsSpecs")
                        'JAIRS仕様
                        重量 = 検索a(objIE, "重量", "ctl01_grdJairsSize")
                        'タイプ = 検索a(objIE, "タイプ", "ctl01_grdJairsSpecs")
                        Call a取得_複線構成(objIE, 複線構成)
                        '略図
                        Call a取得_略図(objIE, 略図URL, 略図数)
                        '単線電線
                        仕上がり外径 = 検索a(objIE, "仕上がり外径", "ctl01_grdPtmIndivs")
                        'クランプタイプ
                        クランプタイプ = 検索a(objIE, "クランプタイプ", "ctl01_grdPtmIndivs")
                        'チューブ
                        チューブ品種 = 検索a(objIE, "チューブ品種", "ctl01_grdPtmIndivs")
                        チューブ長さ = 検索a(objIE, "チューブ長さ", "ctl01_grdPtmIndivs")
                        チューブ内径 = 検索a(objIE, "チューブ内径", "ctl01_grdPtmIndivs")
                        チューブ外径 = 検索a(objIE, "チューブ外径", "ctl01_grdPtmIndivs")
                        'コネクタ
                        コネクタ極数 = 検索a(objIE, "コネクタ極数", "ctl01_grdPtmIndivs")
                        コネクタ色 = 検索a(objIE, "コネクタ色", "ctl01_grdPtmIndivs")
                        コネクタ防水区分 = 検索a(objIE, "防水区分", "ctl01_grdPtmIndivs")
                        メッキ区分 = 検索a(objIE, "メッキ区分", "ctl01_grdPtmIndivs")
                        ロック位置寸法 = 検索a(objIE, "ロック位置寸法", "ctl01_grdPtmIndivs")
                        ロック方向区分 = 検索a(objIE, "ロック方向区分", "ctl01_grdPtmIndivs")
                        端子一体型区分 = 検索a(objIE, "端子一体型区分", "ctl01_grdPtmIndivs")
                        
                        '得意先品番
                        iD = "ctl01_grdJairsCustomers"
                        'Call a取得_得意先品番(objIE, iD, i)
                        '略図
                        iD = "ctl01_dispRyaku_btnDraw"
                        
'                        Call a取得_略図ダウンロード(objIE, アドレス(0) & "\202_略図", 部品品番, アドレスstr) '既に座標を調べた図が変更されたら再度座標を調べる必要があるので一時的にコメント行
'                        Call a取得_略図ダウンロード(objIE, アドレス(1) & "\202_略図", 部品品番, アドレスstr) '既に座標を調べた図が変更されたら再度座標を調べる必要があるので一時的にコメント行
                        
'                        .Cells(i, myCol(0)).Value = 検索結果
'                        .Cells(i, myCol(1)).Value = Replace(部品種別, "&nbsp;", " ")
'                        .Cells(i, myCol(2)).Value = Replace(部品分類, "&nbsp;", " ")
'                        .Cells(i, myCol(3)).Value = Replace(名称品名, "&nbsp;", " ")
'                        .Cells(i, myCol(4)).Value = Replace(部品色, "&nbsp;", " ")
'                        .Cells(i, myCol(5)).Value = Replace(登録工場, "&nbsp;", " ")
'                        .Cells(i, myCol(6)).Value = Replace(重量, "&nbsp;", " ")
'
'                        .Cells(i, myCol(7)).Value = 仕上がり外径
'                        .Cells(i, myCol(8)).Value = 略図数
'                        .Cells(i, myCol(9)).Value = 略図URL
'
'                        .Cells(i, myCol(10)).Value = クランプタイプ
'                        'チューブ
'                        .Cells(i, myCol(11)).Value = チューブ品種
'                        .Cells(i, myCol(12)).Value = チューブ内径
'                        .Cells(i, myCol(13)).Value = コネクタ極数
'                        .Cells(i, myCol(14)).Value = 部品名称
'                        For x = 1 To 20
'                            .Cells(i, 54 + myCol(15)).Value = 複線構成(x)
'                        Next x
'                        .Cells(i, myCol(20)).Value = コネクタ色
'                        .Cells(i, myCol(21)).Value = コネクタ防水区分
'                        .Cells(i, myCol(22)).Value = ロック位置寸法
'                        .Cells(i, myCol(23)).Value = ロック方向区分
'                        .Cells(i, myCol(24)).Value = 端子一体型区分
'                        .Cells(i, myCol(25)).Value = メッキ区分
'                        .Cells(i, myCol(26)).Value = ファミリー
'                        .Cells(i, myCol(27)).Value = オスメス
'
'                        .Cells(i, myCol(28)).Value = チューブ外径
'                        .Cells(i, myCol(29)).Value = チューブ長さ
                        
                    End If
'                    ActiveWindow.ScrollRow = i
'                    Sleep 1000
'                End If
'            End If
'        Next i
   End With

   Set objIE = Nothing
   Set objSFO = Nothing
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    
    コネクタ極数 = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a取得_略図ダウンロード(ByVal objIE As Object, myRyakuDir, 部品品番, アドレス)
    検索文字 = "<IMG SRC="
    For a = 0 To 1
        '略図のボタンidが無ければ処理しない
        'Stop 'もともとWin7なら0、Win10なら2。Win7でも2で動作確認済み
        If InStr(objIE.document.all(2).outerHTML, "ctl01_dispRyaku_btnDraw") = 0 Then Exit Function
        objIE.document.all.Item("ctl01_dispRyaku_edtText").Value = 部品品番
        objIE.document.all.Item("ctl01_dispRyaku_rgpReverse_" & a).Click      '0=正面視 1=裏面視
        objIE.document.all.Item("ctl01_dispRyaku_cmbText")(3).Selected = True 'テキスト入力
        objIE.document.all.Item("ctl01_dispRyaku_chkOriginalSize").Checked = True     '描画
        objIE.document.all.Item("ctl01_dispRyaku_btnDraw").Click              '描画
        
        Call ページ表示を待つ(objIE)
        For x = 0 To objIE.document.all.tags("img").Length - 1  '要素の数
            データ = objIE.document.all.tags("img")(x).outerHTML
            aaa = InStr(StrConv(データ, vbUpperCase), 検索文字)
            If aaa = 0 Then GoTo line0
            略図URL = Left(アドレス, InStrRev(アドレス, "/") - 1) & Mid(データ, Len(検索文字) + 3)
            
            略図URL = Left(略図URL, Len(略図URL) - 2)
            略図保存PASS = myRyakuDir & "\" & 部品品番 & "_" & a & "_" & Format(x, "000") & ".emf"
            'ダウンロードの実行
            Ret = URLDownloadToFile(0, 略図URL, 略図保存PASS, 0, 0)
line0:
        Next x
    Next a
End Function
Public Function a取得_得意先品番(ByVal objIE As Object, iD, ByVal i As Long)
    On Error Resume Next
    データ = objIE.document.getElementById(iD).innerHTML
    On Error GoTo 0
    Dim ii As Long
    Dim タイトルAddCol As Long
    For i = 1 To Len(データ)
        If Mid(データ, i, 3) = "<TD" Then
            得意先名 = "": 得意先品番 = ""
            flg = False: flg1 = False: flg2 = False: flg3 = False: flg4 = False
            For ii = i + 1 To Len(データ)
                If Mid(データ, ii, 1) = "<" Then
                    flg = False
                    flg1 = True
                End If
                If flg1 = False Then
                    If flg = True Then
                        得意先名 = 得意先名 & Mid(データ, ii, 1)
                    End If
                    
                    If Mid(データ, ii, 1) = ">" Then flg = True
                End If
                If flg1 = True Then
                   
                    If flg2 = True Then
                        If Mid(データ, ii, 1) = "<" Then
                            i = ii
                            flg4 = True
                            Exit For
                        End If
                        
                        If flg3 = True Then
                            得意先品番 = 得意先品番 & Mid(データ, ii, 1)
                        End If
                        If Mid(データ, ii, 1) = ">" Then flg3 = True
                    End If
                    If Mid(データ, ii, 3) = "<TD" Then flg2 = True
                End If
            Next ii
        End If
        
        If flg4 = True Then
            得意先名 = Replace(得意先名, "&nbsp;", "")
            得意先品番 = Replace(得意先品番, "&nbsp;", "")
                '部材詳細から探して項目が無ければ追加
            With Sheets("A0_部材詳細")
                Set 得意先名find = 部材詳細_タイトルRan.Find(得意先名 & "_", lookat:=xlWhole)
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
        End If
    Next i
    
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
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
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = InStrRev(Ccc, "</td>")
    eee = Left(Ccc, ddd - 1)
    fff = InStrRev(eee, ">")
    チューブ品種 = Replace(Mid(eee, fff + 1, Len(eee) - fff + 1), "&nbsp;", "")
End Function
Public Function a取得_クランプタイプ(ByVal objIE As Object, クランプタイプ)
  クランプタイプ = ""
  
    検索文字 = "クランプタイプ"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdPtmIndivs").outerText 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = InStr(1, データ, 検索文字)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    Ccc = Left(bbb, InStr(bbb, vbLf))
    ddd = Mid(Ccc, Len(検索文字) + 1)
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
    Ccc = InStr(bbb + 1, データ, ";")
    ddd = InStr(Ccc + 1, データ, ";")
    eee = InStr(ddd + 1, データ, ">")
    zzz = InStr(eee + 1, データ, "<")
    仕上がり外径 = Mid(データ, eee + 1, zzz - eee - 1)
      
End Function

Public Function ページ表示を待つ(ByRef objIE As Object)

    While objIE.ReadyState <> 4 Or objIE.Busy = True '.ReadyState <> 4の間まわる。
        DoEvents  '重いので嫌いな人居るけど。
        Call 仮想キー入力(シフト)
    Wend
    
End Function

Public Function a取得_略図(ByVal objIE As Object, 略図URL, 略図数)
  略図URL = "": 略図数 = 0
  
    略図数 = objIE.document.Images.Length - 1
  
    For r = 1 To objIE.document.Images.Length - 1
        略図URL = objIE.document.Images(1).src
    Next r
      
End Function

Public Function 検索a(ByVal objIE As Object, 検索文字, エレメント)
    On Error Resume Next
    データ = objIE.document.getElementById(エレメント).innerHTML 'PTM情報
    On Error GoTo 0
    aa = 検索(データ, 検索文字, 1)
    If aa = 0 Then Exit Function
    データa = Mid(データ, aa)
    bb = 検索(データa, "<", 3)
    データb = Left(データa, bb - 1)
    Cc = InStrRev(データb, ">")
    検索a = Mid(データb, Cc + 1)
    検索a = Replace(検索a, "&nbsp;", "")
End Function

Public Function a取得_部品分類(ByVal objIE As Object, 部品分類)
  部品分類 = ""
  
    検索文字 = "部品分類"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    Ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(Ccc + 1, データ, "<")
    
    If aaa <> 0 Then 部品分類 = Mid(データ, Ccc + 1, zzz - Ccc - 1)
      
End Function
Public Function a取得_部品名称(ByVal objIE As Object, 部品名称)
  部品名称 = ""
  
    検索文字 = "部品名称"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    Ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(Ccc + 1, データ, "<")
    
    If aaa <> 0 Then 部品名称 = Mid(データ, Ccc + 1, zzz - Ccc - 1)
      
End Function
Public Function a取得_登録工場(ByVal objIE As Object, 登録工場)
  登録工場 = ""
  
    検索文字 = "登録工場"
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML 'PTM情報
    aaa = InStr(1, データ, 検索文字)
    bbb = InStr(aaa + Len(検索文字) + 1, データ, ">")
    Ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(Ccc + 1, データ, "<")
        
    If aaa <> 0 Then 登録工場 = Mid(データ, Ccc + 1, zzz - Ccc - 1)
      
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
    Ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(Ccc + 1, データ, "<")
        
    If aaa <> 0 Then 名称品名 = Mid(データ, Ccc + 1, zzz - Ccc - 1)
      
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
    Ccc = InStr(bbb + 1, データ, ">")
    zzz = InStr(Ccc + 1, データ, "<")
        
    If aaa <> 0 Then 部品色 = Mid(データ, Ccc + 1, zzz - Ccc - 1)
      
End Function

Public Function a取得_重量(ByVal objIE As Object, 重量)
  重量 = ""
  
    検索文字 = "重量"
    On Error Resume Next
    データ = objIE.document.getElementById("ctl01_grdJairsSize").innerHTML 'JAIRS適用サイズ
    On Error GoTo 0
    If データ = "" Then Exit Function
    aaa = 検索(データ, 検索文字, 1)
    If aaa = 0 Then Exit Function
    bbb = Mid(データ, aaa)
    Ccc = 検索(bbb, "<", 3)
    ddd = Left(bbb, Ccc - 1)
    eee = InStrRev(ddd, ">")
    重量 = Mid(ddd, eee + 1)
      
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
        Ccc = InStr(bbb, "target")
        If Ccc = 0 Then 複線構成(i) = "": GoTo line10
        ddd = Mid(bbb, Ccc, Len(bbb))
        eee = InStr(ddd, ">")
        fff = InStr(ddd, "<")
        ggg = Mid(ddd, eee + 1, fff - eee - 1)
        複線構成(i) = ggg
        
        bbb = Mid(bbb, Ccc + fff, Len(bbb))
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
    
    'リストの検索結果数を確認
    Dim myCount As Long: myCount = 0
    データ = objIE.document.getElementById("ctl00_grdList").innerHTML '誰か指定したID内をForEachで参照する方法教えてください
    データsp = Split(データ, vbCrLf)
    For i = LBound(データsp) To UBound(データsp)
        Debug.Print データsp(i)
        If InStr(データsp(i), "javascript") > 0 Then
            myCount = myCount + 1
        End If
    Next i
    
    '検索結果が複数ある場合
    If myCount > 1 Then
        For Each objLink In objIE.document.getElementsByTagName("A")
            Debug.Print objLink.innerText
            If objLink.innerText = Replace(部品品番, "-", "") Then
                Debug.Print 部品品番, objLink.innerText, objLink.href
                'Debug.Print objLink.href
                objIE.Navigate objLink.href
                Exit For
    '        If objLink.innerText = anchorText Then
    '            objIE.navigate objLink.href
    '            Exit For
    '        End If
            ElseIf objLink.innerText = "450" & Replace(部品品番, "-", "") Then 'VS
                Debug.Print 部品品番, objLink.innerText, objLink.href
                'Debug.Print objLink.href
                objIE.Navigate objLink.href
                Exit For
            End If
        Next
    End If

'    '点数が複数ある場合、リンクをクリック
'    If 点数 > 0 Then
'        'リンクaaa = InStrRev(データ, ">" & Replace(部品品番, "-", "") & "<")
'        'リンクbbb = Left(データ, リンクaaa)
'        'リンクccc = InStrRev(リンクbbb, "grdList")
'        'リンクアドレス = Mid(リンクbbb, リンクccc, 9 + Len(点数))
'        'objIe.document.all.Item("javascript:__doPostBack('ctl00$grdList','grdList$0')").Click
'
'        'リンク番号で開く(点数+4で検索する為、確実ではないかも)
'        リンクaaa = InStrRev(データ, ">" & Replace(部品品番, "-", "") & "<")
'        If リンクaaa <> 0 Then
'            リンクbbb = left(データ, リンクaaa)
'            リンクccc = InStrRev(リンクbbb, "$")
'            リンクzzz = InStrRev(リンクbbb, "'")
'            リンク番号 = Mid(リンクbbb, リンクccc + 1, リンクzzz - (リンクccc + 1))
'        Else
'            検索結果 = "NotMatch"
'        End If
'
'        objIE.document.Links(4).Click
'
'    End If
    
    Call ページ表示を待つ(objIE)

    '表示された品番と検索したい品番がマッチするか確認
    データ = objIE.document.getElementById("ctl01_grdPtmCommn").innerHTML
        
    aa = 検索(データ, "ＹＢＭコード", 1)
    If aa = 0 Then Exit Function
    データa = Mid(データ, aa)
    bb = 検索(データa, "<", 3)
    データb = Left(データa, bb - 1)
    Cc = InStrRev(データb, ">")
    v = Mid(データb, Cc + 1)
    表示された部品品番 = Replace(v, "&nbsp;", "")
        
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
        Debug.Print objInput.Value
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
        値 = obj.innerText
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
    Sheets("ログ").Range("a" & Sheets("ログ").Range("a" & Rows.count).End(xlUp).Row + 1) = "a_innertext=" & obj.innerText & "  " & "ID=" & obj.iD           'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
  Next
  
  For Each obj In objIE.document.getElementsByTagName("input")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    Sheets("ログ").Range("a" & Sheets("ログ").Range("a" & Rows.count).End(xlUp).Row + 1) = "input_innertext=" & obj.innerText & "  " & "ID=" & obj.iD           'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
  Next
  
  For Each obj In objIE.document.getElementsByTagName("btn")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    Sheets("ログ").Range("a" & Sheets("ログ").Range("a" & Rows.count).End(xlUp).Row + 1) = "btn_innertext=" & obj.innerText & "  " & "ID=" & obj.iD & " " & obj.Name         'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
  Next

End Function

Sub IE_open_sample() '参考
  
  j = 0
  
  Set objIE = CreateObject("InternetExplorer.Application")  'IEを開く際のお約束
  objIE.Visible = True                                      'IEを開く際のお約束
  objIE.Navigate "http://www.yahoo.co.jp/"                  '開きたいサイトのURLを指定
  
  Do While objIE.ReadyState <> 4                            'サイトが開かれるまで待つ（お約束）
    Do While objIE.Busy = True                              'サイトが開かれるまで待つ（お約束）
    Loop
  Loop
  
  For Each obj In objIE.document.getElementsByTagName("a")  '表示されているサイトのアンカータグ一つずつを変数objにセット
                                                            '各アンカータグ単位に以下の処理を実施
    If obj.innerText = "ファイナンス" Then                  'アンカータグの表示内容が「ファイナンス」の場合に以下の処理を実施
      obj.Click                                             '上記に該当するタグをクリック
      Exit For                                              '上記処理後、For Each　～　Nextを抜ける
    End If
  Next                                                      '次のタグを処理

  Sleep (1000)                                              '1秒待つ
  
  Do While objIE.ReadyState <> 4                            'サイトが開かれるまで待つ（お約束）
    Do While objIE.Busy = True                              'サイトが開かれるまで待つ（お約束）
    
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

Sub エクスポート_部材詳細1273()
    Dim Timer2 As Single
    Timer2 = Timer
'    Dim 再出力flg As Boolean
'    再出力flg = True
    
    Set wb(0) = ActiveWorkbook
    Call addressSet(wb(0))

    With wb(0).Sheets("A0_部材詳細")
        'myRanにワードと列番号と出力行の値入れて、行毎に値を変えながらテキスト出力してみるmyRan(,2)に行毎の値入れる ←VBAのjoinは二次配列で実行できないからやめた
        Dim myWords As String, myWords2 As String, myWords2Col2 As Long, myWordsSP
        myWords = "部品品番_,検索結果_,コネクタ極数_,部品種別_,部品分類_,クランプタイプ_,備考_,防水区分_,色_,メッキ区分_,ファミリー_,オスメス_"
        myWordsSP = Split(myWords, ",")
        Dim myRan(), r As Long
        ReDim myRan(UBound(myWordsSP), 1)
        For r = LBound(myRan) To UBound(myRan)
            myRan(r, 0) = myWordsSP(r)                                             'myWordsSPを入れる
            myRan(r, 1) = .Cells.Find(myWordsSP(r), , , 1).Column    '部材一覧での列番号
            If myWordsSP(r) = "色_" Then myWords2Col2 = r
        Next r
        'コネクタ色がブランクの場合は色を使用する為の追記
        Dim myword2 As String, myWords2Col As Long, コネクタ色 As String
        myWords2 = "コネクタ色_"
        myWords2Col = .Cells.Find(myWords2, , , 1).Column
        
        Dim lastRow As Long, key, i As Long, 検索結果str As String, 出力flg As Boolean
        Set key = .Cells.Find(myRan(0, 0), , , 1)
        lastRow = .Cells(.Rows.count, key.Column).End(xlUp).Row
        For i = key.Row + 1 To lastRow
            '出力flg = .Cells(i, myRan(0, 1)).Value: If 出力flg = True And 再出力flg = False Then GoTo line20
            検索結果str = .Cells(i, myRan(1, 1)).Value: If 検索結果str <> "Found" Then GoTo line20
            コネクタ色 = .Cells(i, myWords2Col).Value
            'text出力する
            Dim 部品品番str As String: 部品品番str = .Cells(i, myRan(0, 1)).Value
            Dim FSO As New FileSystemObject ' FileSystemObject
            Dim TS As TextStream            ' TextStream
            Dim strREC As String            ' 書き出すレコード内容
            Set TS = FSO.CreateTextFile(fileName:=myAddress(0) & "\300_部材詳細\" & 部品品番str & ".txt", overwrite:=True)
            Dim text1 As String, myValue As String
            Dim IntFlNo As Integer: IntFlNo = FreeFile
            TS.WriteLine Join(myWordsSP, ",")
            For r = LBound(myRan) To UBound(myRan)  'myWordsの0と1の要素は無視する
                If myWords2Col2 = r Then
                    'myValue = PTMorJCMP(コネクタ色, .Cells(i, myRan(r, 1)))
                Else
                    myValue = .Cells(i, myRan(r, 1))
                End If
                text1 = text1 & "," & myValue
            Next r
            text1 = Mid(text1, 2)
            TS.WriteLine text1
            TS.Close
            text1 = ""
            Set TS = Nothing
            Set FSO = Nothing
line20:
        Next i
        Debug.Print Round(Timer - Timer2, 1) & "s"
    End With
End Sub



Sub ie_test()

    Dim 複線構成(1 To 20) As String
    Dim iD As String
    Dim myRyakuDir As String, gyo As Long

    Dim myBook As Workbook: Set myBook = ActiveWorkbook
    Call addressSet(myBook)

    gyo = 10
    With Sheets("WEB")
        アカウント = .Range("c" & gyo)
        アカウントID = .Range("d" & gyo)
        パス = .Range("e" & gyo)
        パスID = .Range("f" & gyo)
        ログインbtn = .Range("g" & gyo)
        アドレスstr = .Range("h" & gyo)
        ウィンドウ名 = .Range("i" & gyo)
        ブラウザ = .Range("j" & gyo)
    End With
        
    With Sheets("A0_部材詳細")
        Dim 部材詳細_タイトルRow As Long: 部材詳細_タイトルRow = .Cells.Find("部品品番_").Row
        Set 部材詳細_タイトルRan = .Range(.Cells(部材詳細_タイトルRow, 1), .Cells(部材詳細_タイトルRow, .Columns.count))
        
        タイトル文字 = "検索結果_,部品種別_,部品分類_,名称・品名_,色_,登録工場_,重量_,仕上がり外径_,略図数,略図URL,クランプタイプ_,チューブ品種_,チューブ内径×外径-長さ_,コネクタ極数_,部品名称_,複線構成01,区分_,部品品番_,備考_,検索結果_,コネクタ色_,防水区分_,ロック位置寸法_,ロック方向区分_,端子一体型区分_,メッキ区分_,ファミリー_,オスメス_"
        
        タイトル文字s = Split(タイトル文字, ",")
        Dim myCol() As Long
        ReDim myCol(UBound(タイトル文字s))
        For i = LBound(タイトル文字s) To UBound(タイトル文字s)
            myCol(i) = 部材詳細_タイトルRan.Find(タイトル文字s(i), , , 1).Column
        Next i
    End With
    
    '略図のダウンロード用のフォルダ
    If Dir(myRyakuDir, vbDirectory) = "" Then MkDir myRyakuDir
    'IEの起動
    Dim objIE As Object '変数を定義します。
    Dim ieVerCheck As Variant

    Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
    Set objSFO = CreateObject("Scripting.FileSystemObject")

    ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
    
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10以降(たぶん)
    End If
    
    objIE.Visible = True      '可視、Trueで見えるようにします。
    
    '処理したいページを表示します。
   objIE.Navigate アドレスstr
   Call ページ表示を待つ(objIE)
  
   '画面1 ログイン情報
   objIE.document.all.Item(アカウントID).Value = アカウント
   objIE.document.all.Item(パスID).Value = パス
   objIE.document.all.Item("btnLogin").Click 'ログインクリック
   Call ページ表示を待つ(objIE)
   '画面2 使用注意情報
   objIE.document.all.Item("btnOK").Click 'OKクリック
   Call ページ表示を待つ(objIE)
   '画面3 メインページ
   objIE.document.all.Item("btnYzk").Click '矢崎品番からの検索
   Call ページ表示を待つ(objIE)
'loop
   With Sheets("A0_部材詳細")
        lastgyo = .Cells(.Rows.count, myCol(17)).End(xlUp).Row
        For i = 6 To lastgyo
            If .Cells(i, myCol(19)) = "" Then
                区分 = .Cells(i, myCol(16))
                If Len(区分) = 1 Then
                    部品品番 = .Cells(i, myCol(17))
                    '品番入力
                    objIE.document.all.Item("ctl00_txtYbm").Value = "%" & 部品品番 & "%"
                    Call ページ表示を待つ(objIE)
                    '検索クリック
                    objIE.document.all.Item("ctl00_btnSearch").Click
                    Call ページ表示を待つ(objIE)
                    '品番情報の取得
                    Call a取得_検索結果(objIE, 検索結果, 部品品番)
                    
                    If 検索結果 = "Not Found." Then
                        .Cells(i, myCol(19)) = "NotFound"
                    ElseIf 検索結果 = "NotMatch" Then
                        .Cells(i, myCol(19)) = "NotMatch"
                    Else
                        'PTM
                        部品種別 = 検索a(objIE, "部品種別", "ctl01_grdPtmCommn")
                        部品分類 = 検索a(objIE, "部品分類", "ctl01_grdPtmCommn")
                        部品名称 = 検索a(objIE, "部品名称", "ctl01_grdPtmCommn")
                        登録工場 = 検索a(objIE, "登録工場", "ctl01_grdPtmCommn")
                        'JAIRS
                        名称品名 = 検索a(objIE, "名称", "ctl01_grdEmtrCommon")
                        If 名称品名 = "" Then 名称品名 = 検索a(objIE, "品名", "ctl01_grdJairsCommon")
                        
                        部品色 = 検索a(objIE, "色", "ctl01_grdJairsSpecs")
                        ファミリー = 検索a(objIE, "ファミリー", "ctl01_grdJairsSpecs")
                        オスメス = 検索a(objIE, "オス/メス", "ctl01_grdJairsSpecs")
                        'JAIRS仕様
                        重量 = 検索a(objIE, "重量", "ctl01_grdJairsSize")
                        Call a取得_複線構成(objIE, 複線構成)
                        '略図
                        Call a取得_略図(objIE, 略図URL, 略図数)
                        '単線電線
                        仕上がり外径 = 検索a(objIE, "仕上がり外径", "ctl01_grdPtmIndivs")
                        'クランプタイプ
                        クランプタイプ = 検索a(objIE, "クランプタイプ", "ctl01_grdPtmIndivs")
                        'チューブ
                        チューブ品種 = 検索a(objIE, "チューブ品種", "ctl01_grdPtmIndivs")
                        チューブ長さ = 検索a(objIE, "チューブ長さ", "ctl01_grdPtmIndivs")
                        チューブ内径 = 検索a(objIE, "チューブ内径", "ctl01_grdPtmIndivs")
                        チューブ外径 = 検索a(objIE, "チューブ外径", "ctl01_grdPtmIndivs")
                        
                        'コネクタ
                        コネクタ極数 = 検索a(objIE, "コネクタ極数", "ctl01_grdPtmIndivs")
                        コネクタ色 = 検索a(objIE, "コネクタ色", "ctl01_grdPtmIndivs")
                        コネクタ防水区分 = 検索a(objIE, "防水区分", "ctl01_grdPtmIndivs")
                        メッキ区分 = 検索a(objIE, "メッキ区分", "ctl01_grdPtmIndivs")
                        ロック位置寸法 = 検索a(objIE, "ロック位置寸法", "ctl01_grdPtmIndivs")
                        ロック方向区分 = 検索a(objIE, "ロック方向区分", "ctl01_grdPtmIndivs")
                        端子一体型区分 = 検索a(objIE, "端子一体型区分", "ctl01_grdPtmIndivs")
                        
                        '得意先品番
                        iD = "ctl01_grdJairsCustomers"
                        Call a取得_得意先品番(objIE, iD, i)
                        '略図
                        iD = "ctl01_dispRyaku_btnDraw"
                        
                        Call a取得_略図ダウンロード(objIE, myAddress(0) & "\202_略図", 部品品番, アドレスstr) '既に座標を調べた図が変更されたら再度座標を調べる必要があるので一時的にコメント行
                        Call a取得_略図ダウンロード(objIE, myAddress(1) & "\202_略図", 部品品番, アドレスstr) '既に座標を調べた図が変更されたら再度座標を調べる必要があるので一時的にコメント行
                        
                        
                        .Cells(i, myCol(0)).Value = 検索結果
                        .Cells(i, myCol(1)).Value = Replace(部品種別, "&nbsp;", " ")
                        .Cells(i, myCol(2)).Value = Replace(部品分類, "&nbsp;", " ")
                        .Cells(i, myCol(3)).Value = Replace(名称品名, "&nbsp;", " ")
                        .Cells(i, myCol(4)).Value = Replace(部品色, "&nbsp;", " ")
                        .Cells(i, myCol(5)).Value = Replace(登録工場, "&nbsp;", " ")
                        .Cells(i, myCol(6)).Value = Replace(重量, "&nbsp;", " ")
                        
                        .Cells(i, myCol(7)).Value = 仕上がり外径
                        .Cells(i, myCol(8)).Value = 略図数
                        .Cells(i, myCol(9)).Value = 略図URL
                        
                        .Cells(i, myCol(10)).Value = クランプタイプ
                        'チューブ
                        .Cells(i, myCol(11)).Value = チューブ品種
                        .Cells(i, myCol(12)).Value = チューブ内径 & "×" & チューブ外径 & "-" & チューブ長さ
                        .Cells(i, myCol(13)).Value = コネクタ極数
                        .Cells(i, myCol(14)).Value = 部品名称
                        For x = 1 To 20
                            .Cells(i, 54 + myCol(15)).Value = 複線構成(x)
                        Next x
                        .Cells(i, myCol(20)).Value = コネクタ色
                        .Cells(i, myCol(21)).Value = コネクタ防水区分
                        .Cells(i, myCol(22)).Value = ロック位置寸法
                        .Cells(i, myCol(23)).Value = ロック方向区分
                        .Cells(i, myCol(24)).Value = 端子一体型区分
                        .Cells(i, myCol(25)).Value = メッキ区分
                        .Cells(i, myCol(26)).Value = ファミリー
                        .Cells(i, myCol(27)).Value = オスメス
                    End If
                    ActiveWindow.ScrollRow = i
                End If
            End If
        Next i
   End With

   Set objIE = Nothing
   Set objSFO = Nothing
End Sub

Public Function dsw_open() As Variant
    
    addressSet ThisWorkbook
    
    Dim getStrings As String, getSplit As Variant, g() As Variant, i As Long
    getStrings = "分類,サイト名,アカウント,アカウントID,パス,パスID,ログインbt,アドレス"
    getSplit = Split(getStrings, ",")
    ReDim g(UBound(getSplit))
    For i = LBound(getSplit) To UBound(getSplit)
        g(i) = wb(0).Sheets("WEB").Cells.Find(getSplit(i), , , 1, , , 1).Offset(1, 0)
    Next i
    
    With Sheets("WEB")
        アカウント = g(2)
        アカウントID = g(3)
        パス = g(4)
        パスID = g(5)
        ログインbtn = g(6)
        アドレスstr = g(7)
    End With
    
    'IEの起動
    Dim objIE As Object
    Dim ieVerCheck As Variant
    
    Set objIE = CreateObject("InternetExplorer.Application") 'EXCEL=32bit,6.01=win7?
    Set objSFO = CreateObject("Scripting.FileSystemObject")
    
    ieVerCheck = val(objSFO.GetFileVersion(objIE.FullName))
    
    Debug.Print Application.OperatingSystem, Application.Version, ieVerCheck
    
    If ieVerCheck >= 11 Then
        Set objIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}") 'Win10以降(たぶん)
    End If
    
    objIE.Visible = True      'Trueで見えるようにする
    
    objIE.Navigate アドレスstr
    Call ページ表示を待つ(objIE)
    
    '画面1 ログイン情報
    objIE.document.all.Item(アカウントID).Value = アカウント
    objIE.document.all.Item(パスID).Value = パス
    objIE.document.all.Item("btnLogin").Click 'ログインクリック
    Call ページ表示を待つ(objIE)
    '画面2 使用注意情報
    objIE.document.all.Item("btnOK").Click 'OKクリック
    Call ページ表示を待つ(objIE)
    '画面3 メインページ
    objIE.document.all.Item("btnYzk").Click '矢崎品番からの検索
    Call ページ表示を待つ(objIE)
    
    Set dsw_open = objIE
    
    Set objIE = Nothing
    Set objSFO = Nothing
   
End Function

Public Function dsw_search(ByVal objIE As Object, ByVal searchWord As String) As Variant
    
    '品番入力
    objIE.document.all.Item("ctl00_txtYbm").Value = "%" & searchWord & "%"
    Call ページ表示を待つ(objIE)
    '検索クリック
    objIE.document.all.Item("ctl00_btnSearch").Click
    Call ページ表示を待つ(objIE)
    '品番情報の取得
    Call a取得_検索結果(objIE, 検索結果, searchWord)
                    
    If 検索結果 = "Not Found." Then
        DSW = "False"
    Else
        Dim FieldStrings As String, i As Long, fieldStringSplit As Variant, a As Long
        FieldStrings = "部品品番_,検索結果_,コネクタ極数_,部品種別_,部品分類_,クランプタイプ_,備考_," & _
            "防水区分_,色_,メッキ区分_,ファミリー_,オスメス_,チューブ内径_,チューブ外径_,チューブ長さ_,名称・品名_"
        fieldStringSplit = Split(FieldStrings, ",")
        a = UBound(fieldStringSplit)
        Dim myArray() As Variant
        ReDim myArray(a, 1)
        For i = LBound(myArray) To UBound(myArray)
            myArray(i, 0) = fieldStringSplit(i)
        Next i
        
        myArray(0, 1) = searchWord
        myArray(1, 1) = 検索結果
        myArray(6, 1) = "" '備考
        'PTM
        myArray(3, 1) = 検索a(objIE, "部品種別", "ctl01_grdPtmCommn")
        myArray(4, 1) = 検索a(objIE, "部品分類", "ctl01_grdPtmCommn")
        部品名称 = 検索a(objIE, "部品名称", "ctl01_grdPtmCommn")
        登録工場 = 検索a(objIE, "登録工場", "ctl01_grdPtmCommn")
        'JAIRS
        名称品名 = 検索a(objIE, "名称", "ctl01_grdEmtrCommon")
        If 名称品名 = "" Then 名称品名 = 検索a(objIE, "品名", "ctl01_grdJairsCommon")
        myArray(15, 1) = 名称品名
        
        部品色 = 検索a(objIE, "色", "ctl01_grdJairsSpecs")
        myArray(10, 1) = 検索a(objIE, "ファミリー", "ctl01_grdJairsSpecs")
        myArray(11, 1) = 検索a(objIE, "オス/メス", "ctl01_grdJairsSpecs")
        'JAIRS仕様
        重量 = 検索a(objIE, "重量", "ctl01_grdJairsSize")
        Call a取得_複線構成(objIE, 複線構成)
        '略図
        Call a取得_略図(objIE, 略図URL, 略図数)
        '単線電線
        仕上がり外径 = 検索a(objIE, "仕上がり外径", "ctl01_grdPtmIndivs")
        'クランプタイプ
        myArray(5, 1) = 検索a(objIE, "クランプタイプ", "ctl01_grdPtmIndivs")
        'チューブ
        チューブ品種 = 検索a(objIE, "チューブ品種", "ctl01_grdPtmIndivs")
        myArray(12, 1) = 検索a(objIE, "チューブ内径", "ctl01_grdPtmIndivs")
        myArray(13, 1) = 検索a(objIE, "チューブ外径", "ctl01_grdPtmIndivs")
        myArray(14, 1) = 検索a(objIE, "チューブ長さ", "ctl01_grdPtmIndivs")
        
        'コネクタ
        myArray(2, 1) = 検索a(objIE, "コネクタ極数", "ctl01_grdPtmIndivs")
        コネクタ色 = 検索a(objIE, "コネクタ色", "ctl01_grdPtmIndivs")
        コネクタ色 = Mid(コネクタ色, 4)
        myArray(7, 1) = 検索a(objIE, "防水区分", "ctl01_grdPtmIndivs")
        If コネクタ色 <> "" Then 部品色 = コネクタ色
        myArray(8, 1) = 部品色
        myArray(9, 1) = 検索a(objIE, "メッキ区分", "ctl01_grdPtmIndivs")
        ロック位置寸法 = 検索a(objIE, "ロック位置寸法", "ctl01_grdPtmIndivs")
        ロック方向区分 = 検索a(objIE, "ロック方向区分", "ctl01_grdPtmIndivs")
        端子一体型区分 = 検索a(objIE, "端子一体型区分", "ctl01_grdPtmIndivs")
        
        '得意先品番
'        iD = "ctl01_grdJairsCustomers"
'        Call a取得_得意先品番(objIE, iD, i)
        '略図
        Dim アドレスstr As String
        アドレスstr = Left(objIE.locationurl, InStrRev(objIE.locationurl, "/"))
        iD = "ctl01_dispRyaku_btnDraw"
        Call a取得_略図ダウンロード(objIE, myAddress(1, 1) & "\202_略図", searchWord, アドレスstr) '既に座標を調べた図が変更されたら再度座標を調べる必要があるので一時的にコメント行
        dsw_search = myArray
        
    End If

   Set objIE = Nothing

End Function





